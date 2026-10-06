# D18: optional recording cannot change Print's owner result or warehouse authority.
function Install-ProductionPrintPolicyProbe {
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function PrintPolicyForTest(ByVal mode As String) As Boolean
    Dim context As String, version As Long, request As String, report As String
    Dim model As Object, row As Variant, disabled As Object, id As Variant
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(mode <> "Off")
    If mode = "ControlOff" Then
        For Each row In model("Controls")
            If row("ControlId") = "PRODUCTION_RUN_PRINT" Then row("Collect") = False
        Next row
    ElseIf mode = "UserOff" Then
        For Each id In Array("config-producer", "config-reader")
            Set disabled = CreateObject("Scripting.Dictionary")
            disabled.Add "UserId", CStr(id): disabled.Add "Record", False
            model("Users").Add disabled
        Next id
    End If
    PrintPolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, _
        modTrainingJson.EncodeObject(model), report)
End Function
'@)
}

function Test-ProductionPrintPolicy($Fixture,$Other,$Book,$Sheet,$Decoy){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function Fingerprint($Worksheet){ConvertTo-Json -Compress -Depth 10 -InputObject @($Worksheet.UsedRange.Formula)}
    function Pins([string]$Root,[switch]$Canonical){
        $state=@{}
        if(Test-Path -LiteralPath $Root){foreach($file in Get-ChildItem -LiteralPath $Root -Recurse -File){
            if(-not $Canonical -or $file.Name -match '\.invSys\.Data\.(Inventory|Designs)\.'){ $state[$file.FullName]=Hash $file.FullName }
        }}
        $state
    }
    function Same($Before,$After){
        if($Before.Count -ne $After.Count){return $false}
        foreach($file in $Before.Keys){if(-not $After.ContainsKey($file) -or $After[$file] -cne $Before[$file]){return $false}}
        return $true
    }
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Print policy fixture must remain inside the owned test root.'}
    $training=Join-Path $Fixture.Root 'Training'
    $blocked=Join-Path $training ('Activity/'+$Fixture.Warehouse);$held=$blocked+'-print-policy-held'
    foreach($path in @($blocked,$held,$Fixture.Config)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Print policy path escaped its disposable fixture.'}
    }
    if(Test-Path -LiteralPath $held){throw 'Preserve existing held Print policy fixture.'}
    SelectTarget $Fixture
    $originalConfig=[IO.File]::ReadAllBytes($Fixture.Config);$originalPin=Hash $Fixture.Config
    $oldTraining=Pins $training;$otherPins=RestartPins $Other.Root
    $canonical=Pins $Fixture.Root -Canonical
    if(@($canonical.Keys|Where-Object {$_ -match '\.invSys\.Data\.Inventory\.'}).Count -ne 1){throw 'Canonical Inventory prerequisite missing; not product RED.'}
    $source=Fingerprint $Sheet;$foreign=Fingerprint $Decoy.Worksheets.Item('Production')
    $historical=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(28))).Split("`n")|Where-Object {$_})
    try{
        foreach($policy in @('Off','ControlOff','UserOff','Older','Invalid','StoreUnavailable','Recovery')){
            $moved=$false;$marker=$false
            try{
                SelectTarget $Fixture
                if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.PrintPolicyForTest' @($policy))){throw 'Print policy save prerequisite failed; not product RED.'}
                if($policy -cin @('Older','Invalid')){
                    $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
                    try{
                        $headers=Table $cfg 'tblEventTrackingPolicies';$controls=Table $cfg 'tblEventTrackingControls'
                        if($policy -ceq 'Older'){
                            $headers.ListColumns.Item('CatalogVersion').DataBodyRange.Value2=28.0
                            for($i=$controls.ListRows.Count;$i -ge 1;$i--){
                                if([string]$controls.ListRows.Item($i).Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -cnotin $historical){$controls.ListRows.Item($i).Delete()}
                            }
                        }else{$headers.ListColumns.Item('SchemaVersion').DataBodyRange.Value2=999.0}
                        $cfg.Save()
                    }finally{$cfg.Close($false)}
                }
                if($policy -cin @('ControlOff','UserOff','Older')){
                    if(-not ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_CLOSE'))).StartsWith('True|True|')){throw 'Independent enabled-control prerequisite failed; not product RED.'}
                }
                $policyPin=Hash $Fixture.Config
                if($policy -ceq 'StoreUnavailable'){
                    if(-not (Test-Path -LiteralPath $blocked -PathType Container)){throw 'Existing activity store required; not product RED.'}
                    Move-Item -LiteralPath $blocked -Destination $held;$moved=$true
                    [IO.File]::WriteAllText($blocked,'Disposable Print activity store intentionally unavailable.');$marker=$true
                }
                foreach($case in @('Allowed','PreviewFailure','Denied')){
                    $actor=if($case -ceq 'Denied'){'config-reader'}else{'config-producer'}
                    SelectTarget $Fixture $actor
                    $effective=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_RUN_PRINT'))
                    $expected=if($policy -cin @('Off','ControlOff','UserOff')){'True|False|'}elseif($policy -cin @('Older','Invalid')){'False|False|'}else{'True|True|'}
                    if(-not $effective.StartsWith($expected)){throw 'Effective Print policy prerequisite mismatch; not product RED.'}
                    $before=Pins $training
                    [void](Probe 'OpenDesigner' @($Book.Name));[void](Probe 'RunLocalShowAndCapture' @($Book.Name,'PRINT'))
                    [void](Probe 'ResetPrintPreviewForTest');[void](Probe 'PrintPreviewFailureForTest' @($case -ceq 'PreviewFailure'))
                    $report=$Book.Worksheets.Item('RecallCodesPrint');$reportBefore=Fingerprint $report
                    $label='PrintPolicy.'+$policy+'.'+$case
                    Check ($label+'.SetupNotUserAction') (Same $before (Pins $training))
                    $Decoy.Activate();$stop=Join-Path $runRoot ('print-policy-'+$policy+'-'+$case)
                    $observer=Start-DialogCaptureAndDismiss -ExcelProcessId (@(Get-Process EXCEL)[0].Id) -TimeoutSeconds 30 -StopPath $stop
                    try{Check ($label+'.ActualHandlerReturned') ([bool](Probe 'PrintEntryAct' @('Policy')))}finally{
                        [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
                        Receive-Job $observer -ErrorAction SilentlyContinue|Out-Null
                        if($observer.State -ne 'Completed'){Stop-Job $observer;Remove-Job $observer;throw 'Print policy observer failed; not product RED.'}
                        Remove-Job $observer
                    }
                    $entries=if($case -ceq 'Denied'){0}else{1}
                    foreach($fact in @('PrintOwnerEntries','PrintReportReads','PrintPreviewCountForTest')){Check ($label+'.'+$fact) ([int](Probe $fact) -eq $entries)}
                    Check ($label+'.GuardsRestored') ([bool](Probe 'PrintEntryFact' @('GuardsRestored')))
                    $message=switch($case){'Denied'{'Production permission changed. Reopen Production before continuing.'} 'PreviewFailure'{'BTN_PRINT_CODES failed: Print preview unavailable (fixture).'} default{'Print preview closed.'}}
                    $notice=switch($policy){'Older'{'Tracking unavailable: the saved policy does not include this control.'} 'Invalid'{'Tracking unavailable: the saved tracking policy is invalid.'} 'StoreUnavailable'{'Tracking unavailable: the training record could not be saved.'} default{''}}
                    if($notice){$message+=' '+$notice}
                    Check ($label+'.ExactOwnerMessageAndNotice') ([string](Probe 'PrintStatus') -ceq $message)
                    $after=Pins $training
                    if($policy -ceq 'Recovery'){
                        $added=@($after.Keys|Where-Object {-not $before.ContainsKey($_)})
                        $records=@($added|ForEach-Object {[IO.File]::ReadAllText($_)|ConvertFrom-Json})
                        $outcome=switch($case){'Denied'{'DENIED'} 'PreviewFailure'{'FAILED'} default{'PREVIEW_RETURNED'}}
                        $first=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$last=@($records|Where-Object OutcomeCode -CEQ $outcome)
                        $pair=$records.Count -eq 2 -and $first.Count -eq 1 -and $last.Count -eq 1
                        Check ($label+'.RecordingRecoversExactPair') ($pair -and $first[0].ActivityId -ceq $last[0].ActivityId -and @($records|Where-Object {$_.ControlId -cne 'PRODUCTION_RUN_PRINT' -or $_.UserId -cne $actor -or $_.WarehouseId -cne $Fixture.Warehouse -or @($_.SourceEventRefs).Count}).Count -eq 0)
                        $prior=$true;foreach($file in $before.Keys){$prior=$prior -and $after.ContainsKey($file) -and $after[$file] -ceq $before[$file]}
                        Check ($label+'.PriorEvidenceImmutable') $prior
                    }else{Check ($label+'.NoFalseFallbackOrChangedRecords') (Same $before $after)}
                    Check ($label+'.NoPolicyRepairOrRewrite') ((Hash $Fixture.Config) -ceq $policyPin)
                    Check ($label+'.CanonicalBytesPreserved') (Same $canonical (Pins $Fixture.Root -Canonical))
                    Check ($label+'.SourceKeysAndCustomValuesPreserved') ((Fingerprint $Sheet) -ceq $source)
                    Check ($label+'.DecoyPreserved') ((Fingerprint $Decoy.Worksheets.Item('Production')) -ceq $foreign)
                    Check ($label+'.OtherWarehousePreserved') (RestartPinsEqual $otherPins $Other.Root)
                    if($case -ceq 'Denied'){Check ($label+'.ReportPreserved') ((Fingerprint $report) -ceq $reportBefore)}
                    if($policy -cin @('StoreUnavailable','Recovery')){CaptureOwnedFormByCaptionEvidence 'Production' ('print-policy-'+$policy.ToLowerInvariant()+'-'+$case.ToLowerInvariant()+'.png')}
                    [void](Probe 'PrintPreviewFailureForTest' @($false));[void](Probe 'RunLocalSafeClose')
                }
            }finally{
                if($marker){Remove-Item -LiteralPath $blocked -Force}
                if($moved){Move-Item -LiteralPath $held -Destination $blocked}
                if(@($excel.Workbooks|Where-Object {$_.FullName -ieq $Fixture.Config}).Count){throw 'Close owned Config before restoring bytes.'}
                [IO.File]::WriteAllBytes($Fixture.Config,$originalConfig)
            }
            Check ('PrintPolicy.'+$policy+'.ConfigBytesRestored') ((Hash $Fixture.Config) -ceq $originalPin)
        }
        $after=Pins $training;$retained=$true
        foreach($file in $oldTraining.Keys){$retained=$retained -and $after.ContainsKey($file) -and $after[$file] -ceq $oldTraining[$file]}
        Check 'PrintPolicy.OlderTrainingEvidenceImmutable' $retained
    }finally{
        [void](Probe 'PrintPreviewFailureForTest' @($false));[void](Probe 'RunLocalSafeClose')
        SelectTarget $Fixture
    }
}
