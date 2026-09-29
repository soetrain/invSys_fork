# Test-only adapters; actual packaged Edit handler and owner boundaries stay intact.
function Install-ProductionUomActivityProbe {
    $operations=$packages['invSys.Operations.xlam'].VBProject
    $operations.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function UomSuppressedForTest(ByVal kind As String) As String
    On Error GoTo Done
    If kind = "Loading" Then mLoading = True Else mDesignerActionInProgress = True
    mBtnUomCatalogSend_Click
Done:
    mLoading = False: mDesignerActionInProgress = False
    UomSuppressedForTest = mTxtStatus.Text
End Function
Public Function UomCaptionForTest() As String
    UomCaptionForTest = mBtnUomCatalogSend.Caption
End Function
'@)
    $operations.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function UomSuppressed(ByVal kind As String) As String
    UomSuppressed = mForm.UomSuppressedForTest(kind)
End Function
Public Function UomCaption() As String
    UomCaption = mForm.UomCaptionForTest()
End Function
Public Sub UomProtect(ByVal workbookName As String)
    Application.Workbooks(workbookName).Protect Structure:=True
End Sub
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function UomPolicyForTest(ByVal enabled As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(enabled))
    UomPolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function UomTerminalForTest(ByVal recordJson As String) As Boolean
    Dim record As Object
    Set record = modTrainingJson.DecodeObject(recordJson)
    UomTerminalForTest = modEvaluationMatches.CommandCompleted(record)
End Function
'@)
}

function Test-ProductionUomActivity($Fixture,$Other) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function AdapterDiagnostic([string]$Case) {
        $state=[string](Probe 'UomAdapterStateForTest')
        if($state -cnotmatch '^(True|False)\|(-?[0-9]+)$'){throw 'Invalid fixed UOM adapter diagnostic.'}
        $entry=[pscustomobject]@{Case=$Case;UTC=[DateTimeOffset]::UtcNow.ToString('o');FormEntered=$Matches[1] -ceq 'True';AdapterError=[long]$Matches[2]}
        $entry|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'uom-adapter-diagnostic.jsonl')
        return $entry
    }
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function SetPolicy([bool]$Enabled){
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.UomPolicyForTest' @($Enabled))){throw 'UOM policy fixture unavailable; not product RED.'}
        SelectTarget $Fixture 'config-producer'
    }
    function Pair([string[]]$Before,[string]$Outcome,[string]$Case,[string]$Actor='config-producer'){
        $raw=@(Files|Where-Object {$_ -cnotin $Before}|ForEach-Object {[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object {$_|ConvertFrom-Json})
        $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$terminal=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $paired=$records.Count -eq 2 -and $attempt.Count -eq 1 -and $terminal.Count -eq 1
        Check ($Case+'.Pair') $paired
        $context=$paired;$redacted=$paired;$integrity=$paired;$linked=$false;$owner=$false;$completion=$false
        foreach($record in $records){
            $context=$context -and $record.ControlId -ceq 'PRODUCTION_UOM_EDIT' -and $record.OwnerId -ceq 'PRODUCTION_UOM_STAGING' -and $record.UserId -ceq $Actor -and $record.WarehouseId -ceq $Fixture.Warehouse -and $record.StationId -ceq 'S1' -and $record.CatalogVersion -eq 16
            $redacted=$redacted -and @($record.SourceEventRefs).Count -eq 0
        }
        foreach($value in $raw){
            foreach($secret in @($canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash')){
                $encoded=ConvertTo-Json -InputObject $secret -Compress
                if($value.Contains($secret) -or $value.Contains($encoded.Substring(1,$encoded.Length-2))){$redacted=$false}
            }
            $match=[regex]::Match($value,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $hash -ceq $match.Groups[1].Value
        }
        if($paired){
            $linked=$attempt[0].ActivityId -cne '' -and $attempt[0].ActivityId -ceq $terminal[0].ActivityId -and $attempt[0].RecordId -cne $terminal[0].RecordId
            $severity=switch($Outcome){'REJECTED'{'Warning'} 'DENIED'{'Blocked'} 'FAILED'{'Error'} default{'Info'}}
            $effect=if($Outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
            $owner=$attempt[0].DataEffect -ceq 'Unknown' -and $attempt[0].Severity -ceq 'Info' -and $terminal[0].Severity -ceq $severity -and $terminal[0].DataEffect -ceq $effect -and $terminal[0].EventCode -ceq ('PRODUCTION_UOM_EDIT_'+$Outcome)
            $requested=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.UomTerminalForTest' @(($attempt[0]|ConvertTo-Json -Depth 20 -Compress)))
            $finished=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.UomTerminalForTest' @(($terminal[0]|ConvertTo-Json -Depth 20 -Compress)))
            $completion=-not $requested -and $finished -eq ($Outcome -in @('OPENED','REUSED'))
        }
        Check ($Case+'.Context') $context
        Check ($Case+'.NoEnteredDataOrSources') $redacted
        Check ($Case+'.ContentIntegrity') $integrity
        Check ($Case+'.LinkedDistinctRecords') $linked
        Check ($Case+'.ExactOwnerFact') $owner
        Check ($Case+'.ExactCommandCompletion') $completion
    }
    $canary='UOMACTIVITY'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null
    SetPolicy $true
    if($UomAdapterDiagnostic){
        # Staging's finally block has closed its test form. Calibrate the numeric
        # diagnostic against that absent form without dispatching a control.
        $tracePath=Join-Path $reportRoot 'uom-adapter-entry.txt'
        [void](Probe 'UomAdapterTraceInitForTest' @($tracePath))
        [void](Probe 'UomAdapterTraceCaseForTest' @('MissingFormCalibration'))
        $before=@(Files);$calibration=[string](Probe 'SendUom');$state=AdapterDiagnostic 'MissingFormCalibration'
        Check 'UomAdapterDiagnostic.MissingFormCaptured' ($calibration -ceq 'ADAPTER_ERROR|91' -and -not $state.FormEntered -and $state.AdapterError -eq 91)
        Check 'UomAdapterDiagnostic.CalibrationIsNotAction' (@(Files).Count -eq $before.Count)
        $trace=Get-Content -LiteralPath $tracePath
        Check 'UomAdapterDiagnostic.DurableMissingFormCaptured' (@($trace|Where-Object{$_ -cmatch '^\d+\|MissingFormCalibration\|AdapterFailed\|91$'}).Count -eq 1 -and @($trace|Where-Object{$_ -cmatch '\|MissingFormCalibration\|FormEntered\|'}).Count -eq 0)
        [void](Probe 'UomAdapterTraceCaseForTest' @('Activity'))
    }
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(14))).Split("`n")|Where-Object{$_})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(15))).Split("`n")|Where-Object{$_})
    Check 'UomActivity.Catalog15Extends14' ($old.Count -eq 79 -and $new.Count -eq 80 -and @($old|Where-Object{$_ -cnotin $new}).Count -eq 0 -and @($new|Sort-Object -Unique).Count -eq 80)
    foreach($id in $old){
        $before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,14))
        $after=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,15))
        Check ('UomActivity.CatalogPreserved.'+$id) ($before -ne '' -and $before -ceq $after)
    }
    foreach($outcome in @('COMPLETED','CONFIRMED','STAGED','VALIDATED','APPLIED')){
        Check ('UomActivity.Unsupported.'+$outcome) ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Outcome' @('PRODUCTION_UOM_EDIT',$outcome)) -ceq '')
    }
    try {
        $book=$excel.Workbooks.Add();$decoy=$excel.Workbooks.Add()
        $before=@(Files);[void](Probe 'OpenDesigner' @($book.Name))
        Check 'UomActivity.InitializationNotAction' (@(Files).Count -eq $before.Count)
        $json=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @('PRODUCTION_UOM_EDIT',15))
        $definition=if($json){$json|ConvertFrom-Json}else{$null}
        Check 'UomActivity.FixedMetadata' ($null -ne $definition -and $definition.Caption -ceq [string](Probe 'UomCaption') -and $definition.Surface -ceq 'Operations > Production > Production Settings > UOM Catalog' -and $definition.Class -ceq 'Command' -and $definition.Role -ceq 'Production' -and $definition.Capability -ceq 'PROD_POST')
        $configPin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        $decoy.Activate();$before=@(Files);$notice=[string](Probe 'SendUom')
        $sheet=$book.Worksheets.Item('invSys UOM Catalog');$table=$sheet.ListObjects.Item('tblInvSysUomCatalog')
        Check 'UomActivity.Opened.CapturedWorkbook' ($sheet.ListObjects.Count -eq 1 -and $decoy.Worksheets.Count -eq 1)
        Pair $before 'OPENED' 'UomActivity.Opened'
        $extra=$table.ListColumns.Add();$extra.Name='Local Annotation';$extra.DataBodyRange.Value2=$canary
        $before=@(Files);$notice=[string](Probe 'SendUom')
        Check 'UomActivity.Reused.PreservesLocalData' ([string]$extra.DataBodyRange.Cells.Item(1,1).Value2 -ceq $canary)
        Pair $before 'REUSED' 'UomActivity.Reused'
        $table.Unlist();$before=@(Files);$notice=[string](Probe 'SendUom')
        Pair $before 'REUSED' 'UomActivity.Reopened'
        $table=$sheet.ListObjects.Item('tblInvSysUomCatalog');$table.ListColumns.Item('Notes').Name='Local Notes'
        $before=@(Files);$cells=$sheet.UsedRange.Formula|ConvertTo-Json -Compress -Depth 5;$notice=[string](Probe 'SendUom')
        Check 'UomActivity.Rejected.PreservesStaging' (($sheet.UsedRange.Formula|ConvertTo-Json -Compress -Depth 5) -ceq $cells)
        Pair $before 'REJECTED' 'UomActivity.Rejected'
        Check 'UomActivity.StagingNeverPublishesConfig' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
        [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
        foreach($mode in @('DENIED','FAILED','Busy','Loading','Target','Session','SignedOut','ClosedWorkbook')){
            SelectTarget $Fixture $(if($mode -ceq 'DENIED'){'config-reader'}else{'config-producer'})
            if($UomAdapterDiagnostic){[void](Probe 'UomAdapterTraceCaseForTest' @($mode))}
            $book=$excel.Workbooks.Add()
            # Keep the deliberately retained test form alive independently of
            # its binding; the real launcher-owned form has its own close gate.
            if($mode -ceq 'ClosedWorkbook'){$decoy.Activate()}
            [void](Probe 'OpenDesigner' @($book.Name))
            if($mode -ceq 'FAILED'){[void](Probe 'UomProtect' @($book.Name))}
            if($mode -ceq 'Target'){SelectTarget $Other 'config-producer'}
            if($mode -ceq 'Session'){SelectTarget $Fixture 'config-producer'}
            if($mode -ceq 'SignedOut'){[void](Run 'invSys.Core.xlam' 'modAuth.SignOut')}
            if($mode -ceq 'ClosedWorkbook'){$book.Close($false);$book=$null}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            $notice=if($mode -in @('Busy','Loading')){[string](Probe 'UomSuppressed' @($mode))}else{[string](Probe 'SendUom')}
            if($UomAdapterDiagnostic -and $mode -notin @('Busy','Loading')){
                $state=AdapterDiagnostic $mode
                Check ('UomAdapterDiagnostic.'+$mode+'.FormEnteredWithoutAdapterError') ($state.FormEntered -and $state.AdapterError -eq 0)
            }
            if($mode -in @('DENIED','FAILED')){Pair $before $mode ('UomActivity.'+$mode) $(if($mode -ceq 'DENIED'){'config-reader'}else{'config-producer'})}
            else {Check ('UomActivity.Guard.'+$mode+'.NoRedirectedActivity') (@(Files).Count -eq $before.Count -and @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count)}
            if($null -ne $book){Check ('UomActivity.Guard.'+$mode+'.NoStagingMutation') ($book.Worksheets.Count -eq 1)}
            if($mode -in @('Target','Session','SignedOut','ClosedWorkbook')){Check ('UomActivity.Guard.'+$mode+'.RefusalVisible') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))}
            [void](Probe 'CloseDesigner');if($null -ne $book){$book.Close($false);$book=$null}
        }
        if($UomAdapterDiagnostic){[void](Probe 'UomAdapterTraceCaseForTest' @('Activity'))}
        SetPolicy $false
        $book=$excel.Workbooks.Add();[void](Probe 'OpenDesigner' @($book.Name));$before=@(Files);$notice=[string](Probe 'SendUom')
        Check 'UomActivity.TrackingOff.ActionContinues' ($book.Worksheets.Count -eq 2 -and @(Files).Count -eq $before.Count)
        [void](Probe 'CloseDesigner');$book.Close($false);$book=$null
        SetPolicy $true
        $parent=Join-Path $Fixture.Root 'Training\Activity';$blocked=Join-Path $parent $Fixture.Warehouse;$held=$blocked+'-uom-held'
        foreach($path in @($blocked,$held)){if(-not [IO.Path]::GetFullPath($path).StartsWith([IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Activity fault escaped disposable fixture.'}}
        if(Test-Path -LiteralPath $held){throw 'Preserve existing held fixture.'}
        $moved=$false
        try {
            if(Test-Path -LiteralPath $blocked){Move-Item -LiteralPath $blocked -Destination $held;$moved=$true}
            [IO.File]::WriteAllText($blocked,'Blocked disposable UOM activity path')
            $book=$excel.Workbooks.Add();[void](Probe 'OpenDesigner' @($book.Name));$before=@(Files);$notice=[string](Probe 'SendUom')
            Check 'UomActivity.StoreUnavailable.ActionContinues' ($book.Worksheets.Count -eq 2)
            Check 'UomActivity.StoreUnavailable.VisibleNotice' $notice.Contains('Tracking unavailable')
            Check 'UomActivity.StoreUnavailable.NoFallbackRecords' (@(Files).Count -eq $before.Count)
        }finally{
            if(Test-Path -LiteralPath $blocked -PathType Leaf){Remove-Item -LiteralPath $blocked}
            if($moved){Move-Item -LiteralPath $held -Destination $blocked}
        }
    } finally {
        [void](Probe 'CloseDesigner')
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
