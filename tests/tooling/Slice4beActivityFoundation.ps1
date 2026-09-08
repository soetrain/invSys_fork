# Supplemental boundary checks use only Admin-generated disposable fixtures.
# Payloads, identifiers, credentials and paths never enter result output.
function Get-ActivityRead([string]$Id) {
    [string](Run 'invSys.Core.xlam' 'modActivity.ReadActivityRecord' @($Id))
}
function Save-ActivityFixtureBody([string]$Path,[string]$Body) {
    $sha=[Security.Cryptography.SHA256]::Create()
    try { $hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Body))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
    [IO.File]::WriteAllText($Path,($Body.Substring(0,$Body.Length-1)+',"ContentSha256":"'+$hash+'"}'),[Text.UTF8Encoding]::new($false))
}
function Add-ActivityFixtureTable($Workbook,[string]$Name,[string[]]$Headers,[object[]]$Rows) {
    $sheet=$Workbook.Worksheets.Add()
    $sheet.Name=$Name.Substring(3)
    for($i=0;$i -lt $Headers.Count;$i++) { $sheet.Cells.Item(1,$i+1).Value2=$Headers[$i] }
    for($r=0;$r -lt $Rows.Count;$r++) {
        for($c=0;$c -lt $Headers.Count;$c++) { Set-ActivityFixtureCell $sheet.Cells.Item($r+2,$c+1) $Rows[$r][$c] }
    }
    $range=$sheet.Range($sheet.Cells.Item(1,1),$sheet.Cells.Item($Rows.Count+1,$Headers.Count))
    $table=$sheet.ListObjects.Add(1,$range,$null,1)
    $table.Name=$Name
    return $table
}
function Set-ActivityFixtureCell($Cell,$Value) {
    if($Value -is [bool]) { $Cell.Value2=[bool]$Value }
    elseif($Value -is [double] -or $Value -is [int]) { $Cell.Value2=[double]$Value }
    else { $Cell.Value2=[string]$Value }
}
function Test-Slice4beActivityFoundation($Fixture,$Other) {
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $core.Name='TestActivityWire'
    $core.CodeModule.AddFromString(@'
Public Function JsonRoundTrip(ByVal text As String) As String
    Dim record As Object
    Set record = modTrainingJson.DecodeObject(text)
    If record Is Nothing Then Exit Function
    JsonRoundTrip = modTrainingJson.EncodeObject(record)
End Function
'@)
    $sample='{"items":[{"text":"\u00E9\uD83D\uDE00"},2,true],"after":"value"}'
    Check 'Activity.Json.NestedUnicodeRoundTrip' ([string](Run 'invSys.Core.xlam' 'TestActivityWire.JsonRoundTrip' @($sample)) -ceq $sample)
    foreach($case in @(@('Duplicate','{"a":1,"a":2}'),@('Trailing','{"a":1}x'),@('LeadingZero','{"a":01}'),@('TrailingComma','{"a":[1,]}'))) {
        Check ('Activity.Json.Reject'+$case[0]) ([string](Run 'invSys.Core.xlam' 'TestActivityWire.JsonRoundTrip' @($case[1])) -eq '')
    }
    Check 'Activity.Hash.KnownVector' ([string](Run 'invSys.Core.xlam' 'modTrainingWire.Sha256' @('abc')) -ceq 'ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad')
    $files=@(Get-Slice4beActivityFiles $Fixture)
    $path=@($files | Where-Object { ([IO.File]::ReadAllText($_) | ConvertFrom-Json).ControlId -eq 'ADMIN_SETTINGS_SAVE_VALUE' })[0]
    $original=[IO.File]::ReadAllText($path); $record=$original | ConvertFrom-Json
    $id=[string]$record.RecordId
    Check 'Activity.Read.ValidRecord' ((Get-ActivityRead $id) -ceq ('OK|'+$original))
    $missing=[guid]::NewGuid().ToString()
    $beforeCount=$files.Count
    Check 'Activity.Read.MissingDoesNotCreate' ((Get-ActivityRead $missing).StartsWith('UNAVAILABLE|') -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $beforeCount)
    try {
        # Keep a well-formed 64-hex digest so rejection proves a content mismatch,
        # rather than merely rejecting a malformed integrity field.
        $changedHash=if($record.ContentSha256.StartsWith('0')){'1'}else{'0'}
        $changedHash+=$record.ContentSha256.Substring(1)
        [IO.File]::WriteAllText($path,$original.Replace([string]$record.ContentSha256,$changedHash))
        Check 'Activity.Read.CorruptHashRejected' ((Get-ActivityRead $id).StartsWith('UNAVAILABLE|'))
        $body=$original -replace ',"ContentSha256":"[a-f0-9]{64}"}$','}'
        $cases=@(
            @('NumericString',$body.Replace('"SchemaVersion":1','"SchemaVersion":"1"')),
            @('InvalidTimestamp',($body -replace '"OccurredAtUTC":"[^"]+"','"OccurredAtUTC":"2026-99-99T88:77:66.000Z"')),
            @('InvalidPolicy',$body.Replace('"PolicyVersion":0','"PolicyVersion":"invalid"')),
            @('OrdinalWithoutSequence',$body.Replace('"Ordinal":0','"Ordinal":4')),
            @('CrossWarehouse',$body.Replace('"WarehouseId":"'+$Fixture.Warehouse+'"','"WarehouseId":"'+$Other.Warehouse+'"'))
        )
        foreach($case in $cases) {
            Save-ActivityFixtureBody $path $case[1]
            Check ('Activity.Read.Reject'+$case[0]) ((Get-ActivityRead $id).StartsWith('UNAVAILABLE|'))
        }
    } finally { [IO.File]::WriteAllText($path,$original,[Text.UTF8Encoding]::new($false)) }
    SelectTarget $Other
    Check 'Activity.Read.OtherWarehouseUnavailable' ((Get-ActivityRead $id).StartsWith('UNAVAILABLE|'))
    SelectTarget $Fixture
    $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    $action=[string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @('ADMIN_SETTINGS_SAVE_VALUE',$context))
    Check 'Activity.Context.ValidBegin' ($action -ne '')
    $ok=[bool](Run 'invSys.Core.xlam' 'modActivity.FinishAction' @($action,'UNCHANGED'))
    $snapshot=@{}; foreach($p in @(Get-Slice4beActivityFiles $Fixture)) { $snapshot[$p]=(Get-FileHash -LiteralPath $p).Hash }
    $again=[bool](Run 'invSys.Core.xlam' 'modActivity.FinishAction' @($action,'UNCHANGED'))
    $same=$snapshot.Count -eq @(Get-Slice4beActivityFiles $Fixture).Count
    foreach($p in $snapshot.Keys) { $same=$same -and $snapshot[$p] -eq (Get-FileHash -LiteralPath $p).Hash }
    Check 'Activity.Append.SameRecordIdempotent' ($ok -and $again -and $same)
    Check 'Activity.Append.ConflictingOutcomeRejected' (-not [bool](Run 'invSys.Core.xlam' 'modActivity.FinishAction' @($action,'COMPLETED')))
    $next=[string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @('ADMIN_SETTINGS_SAVE_VALUE',$context))
    Check 'Activity.Append.RepeatedActionDistinct' ($next -ne '' -and $next -ne $action)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
    [void](Run 'invSys.Operations.xlam' 'TestD5Uom.HoldForm')
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    SelectTarget $Fixture
    $count=@(Get-Slice4beActivityFiles $Fixture).Count
    Check 'Activity.Context.ReauthenticatedBeginRejected' ([string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @('ADMIN_SETTINGS_SAVE_VALUE',$context)) -eq '')
    Check 'Activity.Context.ReauthenticatedFinishRejected' (-not [bool](Run 'invSys.Core.xlam' 'modActivity.FinishAction' @($next,'UNCHANGED')))
    Check 'Activity.Context.StaleSessionNoAppend' (@(Get-Slice4beActivityFiles $Fixture).Count -eq $count)
    $configHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','623'))
    Check 'Activity.Context.StaleSettingsSessionCannotSave' (-not $ok -and $configHash -eq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    $stage=$excel.Workbooks.Add()
    $configHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $ok=[bool](Run 'invSys.Operations.xlam' 'TestD5Uom.HeldRoundTrip' @($stage.Name))
    Check 'Activity.Context.StaleProductionSessionCannotPublish' (-not $ok -and $configHash -eq (Get-FileHash -LiteralPath $Fixture.Config).Hash -and $stage.Worksheets.Item('invSys UOM Catalog').ListObjects.Count -eq 1)
    $stage.Close($false)
    Test-Slice4beActivityPolicy $Fixture $id
}
function Test-Slice4beActivityPolicy($Fixture,[string]$RecordId) {
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $cfg=$null
    try {
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        $meta=Add-ActivityFixtureTable $cfg 'tblEventTrackingPolicies' @('PolicyVersion','SchemaVersion','CatalogVersion','CreatedAtUTC','CreatedByUserId','DefaultView','ViewerActionPathCaptureEnabled','AdminViewerEventLoggingEnabled','Operator Extra') @(
            ,@(1.0,1.0,1.0,'2026-09-07T12:00:00.000Z','config-admin','How-To',$false,$true,'preserve'))
        $cfg.Save(); $cfg.Close($false); $cfg=$null
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        $before=@(Get-Slice4beActivityFiles $Fixture).Count
        $ok=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','621'))
        $status=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.LastStatus')
        Check 'Activity.Policy.PartialPairDisablesTrackingOnly' ($ok -and $status.Contains('Tracking unavailable') -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $before)
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        $rows=Add-ActivityFixtureTable $cfg 'tblEventTrackingControls' @('PolicyVersion','ControlId','Collect','Visible','SequenceEligible','Operator Extra') @(
            @(1.0,'ADMIN_SETTINGS_SAVE_VALUE',$false,$false,$true,'preserve'),
            @(1.0,'PRODUCTION_UOM_RETRIEVE',$true,$true,$true,'preserve'))
        $cfg.Save(); $cfg.Close($false); $cfg=$null
        $hash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        Check 'Activity.Policy.HiddenRecordNotReturned' ((Get-ActivityRead $RecordId) -eq 'UNAVAILABLE|Hidden by policy.')
        $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
        Check 'Activity.Policy.DisabledControlNotRecorded' ([string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @('ADMIN_SETTINGS_SAVE_VALUE',$context)) -eq '')
        Check 'Activity.Policy.ReadPreservesConfigBytes' ($hash -eq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        $meta=Table $cfg 'tblEventTrackingPolicies'; $rows=Table $cfg 'tblEventTrackingControls'
        $rows.DataBodyRange.Cells.Item(1,3).Value2=$true
        $rows.DataBodyRange.Cells.Item(1,4).Value2=$true
        $cfg.Save(); $cfg.Close($false); $cfg=$null
        Check 'Activity.Policy.ValidVisibleRecordReturned' ((Get-ActivityRead $RecordId).StartsWith('OK|'))
        if ($CheckReceivingActivity) {
            $beforeNewControl = @(Get-Slice4beActivityFiles $Fixture).Count
            $newControl = [string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @('RECEIVING_CONFIRM_WRITES',$context))
            Check 'Receiving.Policy.OlderCatalogDoesNotEnableNewControl' ($newControl -eq '' -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $beforeNewControl)
        }
        if ($CheckReceivingStagingActivity) {
            $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
            $meta=Table $cfg 'tblEventTrackingPolicies'; $rows=Table $cfg 'tblEventTrackingControls'
            $meta.ListColumns.Item('CatalogVersion').DataBodyRange.Cells.Item(1,1).Value2=2.0
            $row=$rows.ListRows.Add()
            foreach($field in @('Collect','Visible','SequenceEligible')) { $row.Range.Cells.Item(1,$rows.ListColumns.Item($field).Index).Value2=$true }
            $row.Range.Cells.Item(1,$rows.ListColumns.Item('PolicyVersion').Index).Value2=1.0
            $row.Range.Cells.Item(1,$rows.ListColumns.Item('ControlId').Index).Value2='RECEIVING_CONFIRM_WRITES'
            $cfg.Save(); $cfg.Close($false); $cfg=$null
            Check 'Coverage.Policy.CatalogTwoRemainsValid' ((Get-ActivityRead $RecordId).StartsWith('OK|'))
            foreach($control in @('RECEIVING_ADD_SELECTED','DISPOSITION_ADD_SELECTED','DISPOSITION_CONFIRM')) {
                $count=@(Get-Slice4beActivityFiles $Fixture).Count
                $id=[string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @($control,$context))
                Check ('Coverage.Policy.OlderCatalogDoesNotEnable.'+$control) ($id -eq '' -and @(Get-Slice4beActivityFiles $Fixture).Count -eq $count)
            }
            $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
            (Table $cfg 'tblEventTrackingControls').ListRows.Item(3).Delete()
            (Table $cfg 'tblEventTrackingPolicies').ListColumns.Item('CatalogVersion').DataBodyRange.Cells.Item(1,1).Value2=1.0
            $cfg.Save(); $cfg.Close($false); $cfg=$null
        }
        if ($CheckReceivingLocalActivity) { Test-ReceivingLocalOlderPolicy $Fixture $RecordId $context }
        foreach($case in @(@('InvalidTimestamp','CreatedAtUTC','2026-99-99T88:77:66.000Z'),@('InvalidFlag','ViewerActionPathCaptureEnabled','yes'),@('UnknownCatalog','CatalogVersion',999.0))) {
            $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
            $meta=Table $cfg 'tblEventTrackingPolicies'
            $cell=$meta.ListColumns.Item($case[1]).DataBodyRange.Cells.Item(1,1)
            $prior=$cell.Value2; Set-ActivityFixtureCell $cell $case[2]
            $cfg.Save(); $cfg.Close($false); $cfg=$null
            Check ('Activity.Policy.Reject'+$case[0]) ([string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @('ADMIN_SETTINGS_SAVE_VALUE',$context)) -eq '')
            $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
            $meta=Table $cfg 'tblEventTrackingPolicies'
            Set-ActivityFixtureCell $meta.ListColumns.Item($case[1]).DataBodyRange.Cells.Item(1,1) $prior
            $cfg.Save(); $cfg.Close($false); $cfg=$null
        }
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        $meta=Table $cfg 'tblEventTrackingPolicies'
        $newRow=$meta.ListRows.Add()
        for($column=1;$column -le $meta.ListColumns.Count;$column++) {
            Set-ActivityFixtureCell $newRow.Range.Cells.Item(1,$column) $meta.ListRows.Item(1).Range.Cells.Item(1,$column).Value2
        }
        $newRow.Range.Cells.Item(1,1).Value2=2.0
        $cfg.Save(); $cfg.Close($false); $cfg=$null
        Check 'Activity.Policy.IncompleteLatestNeverFallsBack' ([string](Run 'invSys.Core.xlam' 'modActivity.BeginAction' @('ADMIN_SETTINGS_SAVE_VALUE',$context)) -eq '')
    } finally {
        if($null -ne $cfg) { $cfg.Close($false) }
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        [void](Run 'invSys.Core.xlam' 'modConfig.LoadConfig' @($Fixture.Warehouse,'S1'))
    }
}
