# Optional logging may fail; the actual authorized completion still owns business effects.
function Install-ProductionCompletePolicyProbe {
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function CompletePolicyNotice(ByVal policy As String) As Boolean
    Dim notice As String, text As String
    Select Case policy
        Case "Off": notice = ""
        Case "Older": notice = "Tracking unavailable: the saved policy does not include this control."
        Case "Invalid": notice = "Tracking unavailable: the saved tracking policy is invalid."
        Case "StoreUnavailable": notice = "Tracking unavailable: the training record could not be saved."
        Case Else: Exit Function
    End Select
    text = mForm.RunAllocationStatus()
    If notice = "" Then
        CompletePolicyNotice = (InStr(1, text, "Tracking unavailable:", vbBinaryCompare) = 0)
    Else
        CompletePolicyNotice = (Right$(text, Len(notice) + 1) = " " & notice)
    End If
End Function
'@)
}

function Test-ProductionCompletePolicy($Fixture,$Other,$Book,$Decoy,[string]$Canary) {
    function TrainingPins($Target){
        $pins=@{};$training=Join-Path $Target.Root 'Training'
        if(Test-Path -LiteralPath $training){foreach($file in Get-ChildItem -LiteralPath $training -Recurse -File){$pins[$file.FullName]=Hash $file.FullName}}
        $pins
    }
    function SamePins($Before,$After){
        if($Before.Count -ne $After.Count){return $false}
        foreach($file in $Before.Keys){if(-not $After.ContainsKey($file) -or $After[$file] -cne $Before[$file]){return $false}}
        return $true
    }
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\'
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Completion policy fixture escaped its owned root.'}
    $blocked=Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse)
    $held=$blocked+'-complete-policy-held'
    foreach($path in @($blocked,$held,$Fixture.Config)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Completion policy path escaped its fixture.'}
    }
    if(Test-Path -LiteralPath $held){throw 'Preserve existing held completion fixture.'}
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config);$configPin=Hash $Fixture.Config
    $prior=TrainingPins $Fixture;$otherPins=TrainingPins $Other
    $historical=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(27))).Split("`n")|Where-Object{$_})
    foreach($policy in @('Off','Older','Invalid','StoreUnavailable')){
        try{
            SelectTarget $Fixture
            if($policy -ceq 'Off'){
                if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($false))){throw 'Disabled completion policy unavailable; not product RED.'}
            }elseif($policy -cin @('Older','Invalid')){
                $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
                try{
                    $headers=Table $cfg 'tblEventTrackingPolicies';$controls=Table $cfg 'tblEventTrackingControls'
                    if($policy -ceq 'Older'){
                        $headers.ListColumns.Item('CatalogVersion').DataBodyRange.Value2=27.0
                        for($i=$controls.ListRows.Count;$i -ge 1;$i--){
                            if([string]$controls.ListRows.Item($i).Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -cnotin $historical){$controls.ListRows.Item($i).Delete()}
                        }
                    }else{$headers.ListColumns.Item('SchemaVersion').DataBodyRange.Value2=999.0}
                    $cfg.Save()
                }finally{$cfg.Close($false)}
            }
            $effective=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_RUN_COMPLETE'))
            $valid=if($policy -ceq 'Off'){$effective.StartsWith('True|False|')}elseif($policy -ceq 'StoreUnavailable'){$effective.StartsWith('True|True|')}else{$effective.StartsWith('False|False|')}
            if(-not $valid){throw 'Requested completion policy is not effective; not product RED.'}
            if($policy -ceq 'Older' -and -not ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_RUN_CHECK_IN'))).StartsWith('True|True|')){throw 'Historical policy is not independently valid; not product RED.'}
            Check ('CompletePolicy.'+$policy+'.EffectiveFixture') $true
            $policyPin=Hash $Fixture.Config
            foreach($mode in @('Reusable','Worksheet')){
                $work=$null;$moved=$false;$marker=$false
                try{
                    SelectTarget $Fixture 'config-producer'
                    $work=$excel.Workbooks.Add();$sheet=$work.Worksheets.Item(1)
                    $sheet.Cells.Item(2,1).Value2=$Canary;$sheet.Cells.Item(2,2).Formula='=1+2'
                    [void](Probe 'RunLocalReopen' @($work.Name))
                    if($mode -ceq 'Reusable'){$ready=[bool](Probe 'CompleteBaselinePrepare' @($true))}
                    else{
                        $ready=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CompleteWorksheetInventoryForTest' @($work.Name))
                        if($ready){$ready=[bool](Probe 'CompleteWorksheetPrepare' @($Canary))}
                    }
                    if(-not $ready){throw ('Completion policy owner preparation unavailable: '+$policy+'/'+$mode+'; not product RED.')}
                    [void](Probe 'RunLocalShowAndCapture' @($work.Name,'CHECK_IN'));$Decoy.Activate()
                    if($policy -ceq 'StoreUnavailable'){
                        if(-not (Test-Path -LiteralPath $blocked -PathType Container)){throw 'Existing activity store required for fault test.'}
                        Move-Item -LiteralPath $blocked -Destination $held;$moved=$true
                        [IO.File]::WriteAllText($blocked,'Disposable completion logging store unavailable.');$marker=$true
                    }
                    $before=TrainingPins $Fixture;$label='CompletePolicy.'+$policy+'.'+$mode
                    if($mode -ceq 'Reusable'){[void](Probe 'CompleteSubmissionArm' @('','','ObserveOnly'))}
                    Check ($label+'.ActualHandlerReturned') ([bool](Probe 'CompleteEntryAct' @('')))
                    Check ($label+'.GuardsRestored') ([bool](Probe 'CompleteEntryFact' @('GuardsRestored')))
                    if($mode -ceq 'Reusable'){
                        Check ($label+'.OwnerCompleted') ([bool](Probe 'CompleteBaselineCompleted'))
                        Check ($label+'.ExactInputConsumedOnce') ([bool](Owner 'CompleteBaselineBalancesForTest' @($true)))
                        Check ($label+'.ExactFreshOutput') ([bool](Owner 'CompleteBaselineOutputForTest'))
                        $ids=@([string](Probe 'CompleteSubmissionEvent'),[string](Probe 'CompleteSubmissionOutputEvent'))
                    }else{
                        foreach($fact in @('ExactInput','FreshOutput','Log','OutputCustom','CheckCustom')){Check ($label+'.'+$fact) ([bool](Probe 'CompleteWorksheetFact' @($fact,$Canary)))}
                        $ids=([string](Probe 'CompleteWorksheetEvents')).Split("`n")
                    }
                    Check ($label+'.TwoDistinctWriterAcknowledgments') ($ids.Count -eq 2 -and $ids[0] -cne '' -and $ids[1] -cne '' -and $ids[0] -cne $ids[1])
                    Check ($label+'.ExactTrackingNotice') ([bool](Probe 'CompletePolicyNotice' @($policy)))
                    Check ($label+'.NoFallbackOrChangedTrainingRecords') (SamePins $before (TrainingPins $Fixture))
                    Check ($label+'.NoRedirectedTrainingRecords') (SamePins $otherPins (TrainingPins $Other))
                    Check ($label+'.NoPolicyRepairOrRewrite') ((Hash $Fixture.Config) -ceq $policyPin)
                    Check ($label+'.CustomValueAndFormula') ($sheet.Cells.Item(2,1).Value2 -ceq $Canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
                    Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
                    if($policy -ceq 'StoreUnavailable'){CaptureOwnedFormByCaptionEvidence 'Production' ('complete-policy-store-'+$mode.ToLowerInvariant()+'.png')}
                }finally{
                    [void](Probe 'CompleteSubmissionReset');[void](Probe 'CloseDesigner')
                    if($null -ne $work){$work.Close($false)}
                    if($marker){Remove-Item -LiteralPath $blocked}
                    if($moved){Move-Item -LiteralPath $held -Destination $blocked}
                }
            }
        }finally{
            if(@($excel.Workbooks|Where-Object{$_.FullName -ieq $Fixture.Config}).Count){throw 'Close owned Config before restoring its bytes.'}
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        }
        Check ('CompletePolicy.'+$policy+'.ConfigBytesRestored') ((Hash $Fixture.Config) -ceq $configPin)
    }
    $after=TrainingPins $Fixture;$retained=$true
    foreach($file in $prior.Keys){$retained=$retained -and $after.ContainsKey($file) -and $after[$file] -ceq $prior[$file]}
    Check 'CompletePolicy.PriorTrainingRecordsImmutable' $retained
    SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
}
