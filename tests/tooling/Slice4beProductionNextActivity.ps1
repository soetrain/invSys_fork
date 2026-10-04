# D18: actual Next Batch handler, independent owner state and journal records.
function Install-ProductionNextActivityProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function NextActivityActForTest(ByVal guard As String) As Boolean
    Dim priorLoading As Boolean, priorBusy As Boolean
    priorLoading = mLoading: priorBusy = mDesignerActionInProgress
    If guard = "Loading" Then mLoading = True
    If guard = "Busy" Then mDesignerActionInProgress = True
    mBtnManagerNext_Click
    NextActivityActForTest = (mLoading = (priorLoading Or guard = "Loading")) And _
        (mDesignerActionInProgress = (priorBusy Or guard = "Busy"))
    mLoading = priorLoading: mDesignerActionInProgress = priorBusy
End Function
Public Function NextActivityWorksheetPrepareForTest(ByVal canary As String) As Boolean
    Dim ws As Worksheet, lo As ListObject, headers As Variant, col As Long
    If Not RunWorksheetPrepareForTest() Then Exit Function
    If RunWorksheetStageForTest("ALLOCATE", "Quantity", canary) <> "READY" Then Exit Function
    Set ws = mOperatorWorkbook.Worksheets("Production")
    ws.ListObjects("InventoryPalette_RunFixture").Name = "proc_NextActivity_palette"
    headers = Array("System_Key", "PROCESS", "OUTPUT", "REAL OUTPUT", "BATCH", "RECALL CODE", "Operator Extra", "Operator Formula")
    For col = 0 To UBound(headers): ws.Cells(1, 31 + col).Value2 = headers(col): Next col
    Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("AE1:AL2"), , xlYes)
    lo.Name = "ProductionOutput"
    lo.DataBodyRange.Cells(1, 1).Value2 = mRunSheetKey
    lo.DataBodyRange.Cells(1, 2).Value2 = canary
    lo.DataBodyRange.Cells(1, 3).Value2 = canary
    lo.DataBodyRange.Cells(1, 4).Value2 = 4#
    lo.DataBodyRange.Cells(1, 5).Value2 = 2#
    lo.DataBodyRange.Cells(1, 6).Value2 = canary
    lo.DataBodyRange.Cells(1, 7).Value2 = canary
    lo.DataBodyRange.Cells(1, 8).Formula = "=1+2"
    NextActivityWorksheetPrepareForTest = Not modProductionReusableRun.ReusableRunIsLoaded()
End Function
Public Function NextActivityWorksheetFactForTest(ByVal fact As String, ByVal canary As String) As Boolean
    Dim lo As ListObject
    Set lo = mOperatorWorkbook.Worksheets("Production").ListObjects("ProductionOutput")
    Select Case fact
        Case "Output": NextActivityWorksheetFactForTest = CellByHeader(lo, 1, "REAL OUTPUT") = "" And _
            CellByHeader(lo, 1, "RECALL CODE") = "" And CellByHeader(lo, 1, "BATCH") = "3"
        Case "Identity": NextActivityWorksheetFactForTest = StrComp(CellByHeader(lo, 1, "System_Key"), mRunSheetKey, vbBinaryCompare) = 0
        Case "Custom": NextActivityWorksheetFactForTest = CellByHeader(lo, 1, "Operator Extra") = canary And _
            lo.DataBodyRange.Cells(1, ProductionColumnIndex(lo, "Operator Formula")).Formula = "=1+2"
        Case "Palette"
            Set lo = mOperatorWorkbook.Worksheets("Production").ListObjects("proc_NextActivity_palette")
            NextActivityWorksheetFactForTest = CellByHeader(lo, 1, "System_Key") = "" And _
                CellByHeader(lo, 1, "Operator Extra") = canary And lo.DataBodyRange.Cells(1, 4).Formula = "=1+2"
    End Select
End Function
Public Function NextActivityMissingWorksheetForTest() As Boolean
    Dim alerts As Boolean, number As Long, description As String
    alerts = Application.DisplayAlerts
    On Error GoTo Failed
    Application.DisplayAlerts = False
    mOperatorWorkbook.Worksheets("Production").Delete
    Application.DisplayAlerts = alerts
    NextActivityMissingWorksheetForTest = Not WorkbookHasSheet(mOperatorWorkbook, "Production") And _
        Not modProductionReusableRun.ReusableRunIsLoaded()
    Exit Function
Failed:
    number = Err.Number: description = Err.Description
    Application.DisplayAlerts = alerts
    Err.Raise number, "NextActivityMissingWorksheetForTest", description
End Function
Public Function NextActivityNoFalseSuccessForTest() As Boolean
    NextActivityNoFalseSuccessForTest = mTxtStatus.Text <> "Next Batch completed."
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mNextActivityNest As Boolean')
    $line=$adapter.ProcStartLine('NextBaselineOwnerHit',0)
    $end=$line+$adapter.ProcCountLines('NextBaselineOwnerHit',0)
    $anchor=@(for($i=$line;$i -lt $end;$i++){if($adapter.Lines($i,1).Trim() -ceq 'mNextBaselineEntries = mNextBaselineEntries + 1'){$i+1}})
    if($anchor.Count -ne 1){throw 'Owner counter anchor unavailable; not product RED.'}
    $adapter.InsertLines($anchor[0],'    If mNextActivityNest Then mNextActivityNest = False: mForm.NextActivityActForTest ""')
    $adapter.AddFromString(@'
Public Function NextActivityAct(ByVal guard As String) As Boolean
    mNextBaselineEntries = 0: mNextActivityNest = (guard = "Nested")
    NextActivityAct = mForm.NextActivityActForTest(guard)
End Function
Public Function NextActivityWorksheetPrepare(ByVal canary As String) As Boolean
    NextActivityWorksheetPrepare = mForm.NextActivityWorksheetPrepareForTest(canary)
End Function
Public Function NextActivityWorksheetFact(ByVal fact As String, ByVal canary As String) As Boolean
    NextActivityWorksheetFact = mForm.NextActivityWorksheetFactForTest(fact, canary)
End Function
Public Function NextActivityMissingWorksheet() As Boolean
    NextActivityMissingWorksheet = mForm.NextActivityMissingWorksheetForTest()
End Function
Public Function NextActivityNoFalseSuccess() As Boolean
    NextActivityNoFalseSuccess = mForm.NextActivityNoFalseSuccessForTest()
End Function
'@)
}

function Test-ProductionNextActivity($Fixture,$Other,$Book,$Decoy,[string]$Canary) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Files {@(Get-Slice4beActivityFiles $Fixture)}
    function CanonicalPins {
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File | Where-Object Name -match '\.invSys\.Data\.(Inventory|Designs)\.'){
            $stream=[IO.File]::Open($file.FullName,'Open','Read','ReadWrite')
            try{$pins[$file.FullName]=(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
        }
        if($pins.Count -lt 2){throw 'Canonical fixture files unavailable; not product RED.'}
        return $pins
    }
    function CanonicalSame($Before){
        $after=CanonicalPins
        if($Before.Count -ne $after.Count){return $false}
        foreach($path in $Before.Keys){if(-not $after.ContainsKey($path) -or $after[$path] -cne $Before[$path]){return $false}}
        return $true
    }
    function Pair([string[]]$Before,[string]$Outcome,[string]$Label,[string]$Actor='config-producer'){
        $raw=@(Files|Where-Object{$_ -cnotin $Before}|ForEach-Object{[IO.File]::ReadAllText($_)})
        $records=@($raw|ForEach-Object{$_|ConvertFrom-Json})
        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $Outcome)
        $pair=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
        $context=$pair;$safe=$pair;$integrity=$pair;$linked=$false;$facts=$false;$terminal=$false
        foreach($r in $records){
            $context=$context -and $r.ControlId -ceq 'PRODUCTION_RUN_NEXT_BATCH' -and $r.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $r.UserId -ceq $Actor -and $r.WarehouseId -ceq $Fixture.Warehouse -and $r.StationId -ceq 'S1' -and $r.CatalogVersion -eq 26
            $safe=$safe -and @($r.SourceEventRefs).Count -eq 0
        }
        foreach($value in $raw){
            foreach($secret in @($Canary,$Fixture.Secret,(CredentialHash $Fixture.Secret),$Fixture.Root,'mBtn','PinHash','DEMO-RAW-BLACK-TEA')){
                $encoded=ConvertTo-Json -InputObject $secret -Compress
                if($value.Contains($secret) -or $value.Contains($encoded.Substring(1,$encoded.Length-2))){$safe=$false}
            }
            $match=[regex]::Match($value,',"ContentSha256":"([a-f0-9]{64})"\}$')
            if(-not $match.Success){$integrity=$false;continue}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($value.Substring(0,$match.Index)+'}'))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            $integrity=$integrity -and $hash -ceq $match.Groups[1].Value
        }
        if($pair){
            $linked=$first[0].ActivityId -cne '' -and $first[0].ActivityId -ceq $last[0].ActivityId -and $first[0].RecordId -cne $last[0].RecordId -and $activities.Add([string]$first[0].ActivityId)
            $severity=if($Outcome -ceq 'REJECTED'){'Warning'}elseif($Outcome -ceq 'FAILED'){'Error'}elseif($Outcome -ceq 'DENIED'){'Blocked'}else{'Info'}
            $effect=if($Outcome -ceq 'FAILED'){'Unknown'}else{'Unchanged'}
            $facts=$first[0].DataEffect -ceq 'Unknown' -and $last[0].Severity -ceq $severity -and $last[0].DataEffect -ceq $effect -and $last[0].EventCode -ceq ('PRODUCTION_RUN_NEXT_BATCH_'+$Outcome)
            $terminal=-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($first[0]|ConvertTo-Json -Depth 20 -Compress))) -and [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($last[0]|ConvertTo-Json -Depth 20 -Compress))) -eq ($Outcome -ceq 'STAGED')
        }
        Check ($Label+'.AttemptAndOutcome') $pair
        Check ($Label+'.ExactContext') $context
        Check ($Label+'.RedactedAndNoSourceReferences') $safe
        Check ($Label+'.Integrity') $integrity
        Check ($Label+'.DistinctLinkedAttempt') $linked
        Check ($Label+'.OwnerOutcomeFacts') $facts
        Check ($Label+'.ExactCommandTerminal') $terminal
    }
    Test-ProductionNextCatalog
    $activities=[Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Authorized tracking policy unavailable; not product RED.'}
    SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
    foreach($guard in @('Normal','Loading','Busy','Nested')){
        if(-not [bool](Probe 'CompleteBaselinePrepare' @($true)) -or -not [bool](Probe 'CompleteBaselineAct') -or -not [bool](Probe 'CompleteBaselineCompleted')){throw 'Completed batch unavailable; not product RED.'}
        $state=[string](Probe 'RunLocalOwnerState');$projection=[string](Probe 'RunLocalState');$before=Files;$label='NextActivity.'+$guard;$pins=CanonicalPins
        $Decoy.Activate()
        Check ($label+'.ActualHandlerAndGuardRestoration') ([bool](Probe 'NextActivityAct' @($guard)))
        Check ($label+'.CanonicalBytesPreserved') (CanonicalSame $pins)
        if($guard -cin @('Loading','Busy')){
            Check ($label+'.NoOwnerEntry') ([int](Probe 'NextBaselineEntries') -eq 0)
            Check ($label+'.OwnerPreserved') ([string](Probe 'RunLocalOwnerState') -ceq $state)
            Check ($label+'.ProjectionPreserved') ([string](Probe 'RunLocalState') -ceq $projection)
            Check ($label+'.NoRecords') (@(Files|Where-Object{$_ -cnotin $before}).Count -eq 0)
        }else{
            Check ($label+'.OneOwnerEntry') ([int](Probe 'NextBaselineEntries') -eq 1)
            Check ($label+'.IndependentNextBatchReady') ([bool](Probe 'NextBaselineReady'))
            Pair $before 'STAGED' $label
        }
        if($guard -ceq 'Normal'){
            $state=[string](Probe 'RunLocalOwnerState');$before=Files
            Check 'NextActivity.PriorSuccess.HandlerReturned' ([bool](Probe 'NextActivityAct' @('')))
            Check 'NextActivity.PriorSuccess.OwnerUnchanged' ([string](Probe 'RunLocalOwnerState') -ceq $state)
            Pair $before 'REJECTED' 'NextActivity.PriorSuccess'
        }
    }
    # The existing native observer dismisses only notifications from this test Excel.
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'),[ref]$tokens,[ref]$errors)
    $definition=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Start-DialogCaptureAndDismiss'},$false)
    if($errors.Count -or $null -eq $definition){throw 'Native observer unavailable; not product RED.'}
    . ([scriptblock]::Create($definition.Extent.Text))
    if(-not [bool](Probe 'NextActivityWorksheetPrepare' @($Canary))){throw 'Worksheet fixture unavailable; not product RED.'}
    $processes=@(Get-Process EXCEL);if($processes.Count -ne 1){throw 'Isolated Excel owner unavailable.'}
    $stop=Join-Path $runRoot 'next-worksheet-dialog-stop';$observer=Start-DialogCaptureAndDismiss -ExcelProcessId $processes[0].Id -TimeoutSeconds 30 -StopPath $stop
    $before=Files;$pins=CanonicalPins
    try{Check 'NextActivity.Worksheet.HandlerReturned' ([bool](Probe 'NextActivityAct' @('')))}finally{
        [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
        $notification=@(Receive-Job $observer -ErrorAction SilentlyContinue) -join "`n"
        Check 'NextActivity.Worksheet.NativeNotice' ($observer.State -eq 'Completed' -and $notification.Contains('Next Batch ready. Inventory selections cleared for unchecked processes.'))
        if($observer.State -ne 'Completed'){Stop-Job $observer};Remove-Job $observer
    }
    foreach($fact in @('Output','Identity','Custom','Palette')){Check ('NextActivity.Worksheet.'+$fact) ([bool](Probe 'NextActivityWorksheetFact' @($fact,$Canary)))}
    Check 'NextActivity.Worksheet.CanonicalBytesPreserved' (CanonicalSame $pins)
    Pair $before 'STAGED' 'NextActivity.Worksheet'
    if(-not [bool](Probe 'NextActivityMissingWorksheet')){throw 'Missing worksheet fixture unavailable; not product RED.'}
    $before=Files
    $stop=Join-Path $runRoot 'next-missing-dialog-stop';$observer=Start-DialogCaptureAndDismiss -ExcelProcessId $processes[0].Id -TimeoutSeconds 30 -StopPath $stop
    try{Check 'NextActivity.MissingWorksheet.HandlerReturned' ([bool](Probe 'NextActivityAct' @('')))}finally{
        [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
        $null=Receive-Job $observer -ErrorAction SilentlyContinue
        if($observer.State -ne 'Completed'){Stop-Job $observer};Remove-Job $observer
    }
    Check 'NextActivity.MissingWorksheet.NoFalseSuccess' ([bool](Probe 'NextActivityNoFalseSuccess'))
    Pair $before 'FAILED' 'NextActivity.MissingWorksheet'
    SelectTarget $Fixture 'config-reader';[void](Probe 'RunLocalReopen' @($Book.Name))
    if(-not [bool](Probe 'CheckBaselineReusableStage' @('Selected'))){throw 'Denied actor fixture unavailable; not product RED.'}
    $state=[string](Probe 'RunLocalOwnerState');$before=Files
    Check 'NextActivity.Permission.HandlerReturned' ([bool](Probe 'NextActivityAct' @('')))
    Check 'NextActivity.Permission.NoOwnerEntry' ([int](Probe 'NextBaselineEntries') -eq 0)
    Check 'NextActivity.Permission.OwnerPreserved' ([string](Probe 'RunLocalOwnerState') -ceq $state)
    Pair $before 'DENIED' 'NextActivity.Permission' 'config-reader'
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($false))){throw 'Disabled policy fixture unavailable; not product RED.'}
    SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
    if(-not [bool](Probe 'CompleteBaselinePrepare' @($true)) -or -not [bool](Probe 'CompleteBaselineAct') -or -not [bool](Probe 'CompleteBaselineCompleted')){throw 'Disabled-policy owner fixture unavailable; not product RED.'}
    $before=Files
    Check 'NextActivity.Disabled.HandlerReturned' ([bool](Probe 'NextActivityAct' @('')))
    Check 'NextActivity.Disabled.OwnerSucceeded' ([bool](Probe 'NextBaselineReady'))
    Check 'NextActivity.Disabled.NoRecords' (@(Files|Where-Object{$_ -cnotin $before}).Count -eq 0)
}

function Test-ProductionNextCatalog {
    $id='PRODUCTION_RUN_NEXT_BATCH'
    $old=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(25))).Split("`n")|Where-Object{$_})
    $new=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(26))).Split("`n")|Where-Object{$_})
    Check 'NextCatalog.Extends25Exactly' ($old.Count -eq 128 -and $new.Count -eq 129 -and @($new|Select-Object -Unique).Count -eq 129 -and ($new[0..127] -join '|') -ceq ($old -join '|') -and $new[-1] -ceq $id)
    $same=$true
    foreach($prior in $old){$before=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($prior,25));$after=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($prior,26));$same=$same -and $before -cne '' -and $after -ceq $before}
    Check 'NextCatalog.All128PriorDefinitionsPreserved' $same
    $excluded=$true
    foreach($version in 1..25){$excluded=$excluded -and [string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,$version)) -ceq ''}
    Check 'NextCatalog.ExcludedFromHistoricalVersions' $excluded
    $wire=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Definition' @($id,26));$record=if($wire){$wire|ConvertFrom-Json}else{$null}
    Check 'NextCatalog.ExactControlContract' ($null -ne $record -and $record.ControlId -ceq $id -and $record.OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $record.Class -ceq 'Command' -and $record.Role -ceq 'Production' -and $record.Caption -ceq 'Next Batch' -and $record.Surface -ceq 'Operations > Production > Production Run - List' -and $record.Capability -ceq 'PROD_POST')
    foreach($code in @('REQUESTED','DENIED','REJECTED','FAILED','STAGED','CONFIRMED')){
        $supported=$code -cne 'CONFIRMED'
        Check ('NextCatalog.EmptyReferences.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,$code,'[]')) -eq $supported)
        $json=@{ControlId=$id;OwnerId='PRODUCTION_RUN_LOCAL';CatalogVersion=26;OutcomeCode=$code}|ConvertTo-Json -Compress
        Check ('NextCatalog.Terminal.'+$code) ([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @($json)) -eq ($code -ceq 'STAGED'))
    }
    foreach($kind in @('Inventory','Designs')){
        $json=ConvertTo-Json -InputObject @(@{WarehouseId='CATALOG_TEST';SourceKind=$kind;EventId='Source_A';SubmissionState='Submitted'}) -Compress
        Check ('NextCatalog.RejectSource.'+$kind) (-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.References' @($id,'STAGED',$json)))
    }
}
