# Optional tracking faults must not change the actual Next Batch owner result.
function Install-ProductionNextPolicyProbe {
    $store=$packages['invSys.Core.xlam'].VBProject.VBComponents.Item('modActivityStore').CodeModule
    $start=$store.ProcStartLine('Append',0);$end=$start+$store.ProcCountLines('Append',0)
    $anchor=@(for($i=$start;$i -lt $end;$i++){if($store.Lines($i,1).Trim() -ceq 'If Not ValidBody(target, recordId, body) Then Exit Function'){$i+1}})
    if($anchor.Count -ne 1){throw 'Activity append boundary unavailable; not product RED.'}
    $store.InsertLines($anchor[0],'    If NextPolicyFaultForTest(body) Then Err.Raise vbObjectError + 7871, , "Injected Next Batch terminal store failure."')
    $store.InsertLines(1,'Private mNextPolicyTerminalFault As Boolean, mNextPolicyFaultCalls As Long')
    $store.AddFromString(@'
Public Sub NextPolicyArmForTest(ByVal enabled As Boolean)
    mNextPolicyTerminalFault = enabled: mNextPolicyFaultCalls = 0
End Sub
Private Function NextPolicyFaultForTest(ByVal body As String) As Boolean
    Dim value As Object
    If Not mNextPolicyTerminalFault Then Exit Function
    Set value = modTrainingJson.DecodeObject(body)
    If CStr(value("ControlId")) <> "PRODUCTION_RUN_NEXT_BATCH" Then Exit Function
    If CStr(value("OutcomeCode")) = "REQUESTED" Then Exit Function
    mNextPolicyFaultCalls = mNextPolicyFaultCalls + 1
    NextPolicyFaultForTest = True
End Function
Public Function NextPolicyFaultCallsForTest() As Long
    NextPolicyFaultCallsForTest = mNextPolicyFaultCalls
End Function
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function NextPolicyCatalogVersion() As Long
    NextPolicyCatalogVersion = modActivityCatalog.CATALOG_VERSION
End Function
Public Sub NextPolicyFaultArm(ByVal enabled As Boolean)
    modActivityStore.NextPolicyArmForTest enabled
End Sub
Public Function NextPolicyFaultCalls() As Long
    NextPolicyFaultCalls = modActivityStore.NextPolicyFaultCallsForTest()
End Function
'@)
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function NextPolicyNotice(ByVal policy As String) As Boolean
    Dim notice As String, text As String
    Select Case policy
        Case "Off": notice = ""
        Case "Older": notice = "Tracking unavailable: the saved policy does not include this control."
        Case "Invalid": notice = "Tracking unavailable: the saved tracking policy is invalid."
        Case "StoreUnavailable", "TerminalFailure": notice = "Tracking unavailable: the training record could not be saved."
        Case Else: Exit Function
    End Select
    text = mForm.RunAllocationStatus()
    If notice = "" Then
        NextPolicyNotice = (InStr(1, text, "Tracking unavailable:", vbBinaryCompare) = 0)
    Else
        NextPolicyNotice = (Right$(text, Len(notice) + 1) = " " & notice)
    End If
End Function
'@)
}

function Test-ProductionNextPolicy($Fixture,$Other,$Book,$Decoy,[string]$Canary){
    . (Join-Path $PSScriptRoot 'Slice4beProductionNextPaths.ps1')
    function SavedHash([string]$Path){Hash $Path}
    function PathCheck([string]$Name,[bool]$Passed){Check ($Name.Replace('InstructionPaths.','')) $Passed}
    function TrainingPins($Target){
        $pins=@{};$training=Join-Path $Target.Root 'Training'
        if(Test-Path -LiteralPath $training){foreach($file in Get-ChildItem -LiteralPath $training -Recurse -File){$pins[$file.FullName]=Hash $file.FullName}}
        return $pins
    }
    function Retained($Before,$After){
        foreach($file in $Before.Keys){if(-not $After.ContainsKey($file) -or $After[$file] -cne $Before[$file]){return $false}}
        return $true
    }
    SelectTarget $Fixture
    if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Current Next Batch policy unavailable; not product RED.'}
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config);$configPin=Hash $Fixture.Config
    $root=[IO.Path]::GetFullPath($Fixture.Root).TrimEnd('\')+'\';$owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    if(-not $root.StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Next policy fixture escaped its owned root.'}
    $blocked=Join-Path $Fixture.Root ('Training/Activity/'+$Fixture.Warehouse);$held=$blocked+'-next-policy-held'
    foreach($path in @($blocked,$held,$Fixture.Config)){if(-not [IO.Path]::GetFullPath($path).StartsWith($root,[StringComparison]::OrdinalIgnoreCase)){throw 'Next policy path escaped its fixture.'}}
    if(Test-Path -LiteralPath $held){throw 'Preserve existing held Next policy fixture.'}
    $prior=TrainingPins $Fixture;$otherPins=TrainingPins $Other
    $historical=@(([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Ids' @(25))).Split("`n")|Where-Object{$_})
    foreach($policy in @('Off','Older','Invalid','StoreUnavailable','TerminalFailure')){
        try{
            SelectTarget $Fixture
            if($policy -ceq 'Off'){
                if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($false))){throw 'Disabled Next policy unavailable; not product RED.'}
            }elseif($policy -cin @('Older','Invalid')){
                $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
                try{
                    $headers=Table $cfg 'tblEventTrackingPolicies';$controls=Table $cfg 'tblEventTrackingControls'
                    if($policy -ceq 'Older'){
                        $headers.ListColumns.Item('CatalogVersion').DataBodyRange.Value2=25.0
                        for($i=$controls.ListRows.Count;$i -ge 1;$i--){if([string]$controls.ListRows.Item($i).Range.Cells.Item(1,$controls.ListColumns.Item('ControlId').Index).Value2 -cnotin $historical){$controls.ListRows.Item($i).Delete()}}
                    }else{$headers.ListColumns.Item('SchemaVersion').DataBodyRange.Value2=999.0}
                    $cfg.Save()
                }finally{$cfg.Close($false)}
            }
            $effective=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_RUN_NEXT_BATCH'))
            $valid=if($policy -ceq 'Off'){$effective.StartsWith('True|False|')}elseif($policy -cin @('StoreUnavailable','TerminalFailure')){$effective.StartsWith('True|True|')}else{$effective.StartsWith('False|False|')}
            if(-not $valid){throw 'Requested Next policy is not effective; not product RED.'}
            if($policy -ceq 'Older' -and -not ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_RUN_LOAD'))).StartsWith('True|True|')){throw 'Older policy is not independently valid; not product RED.'}
            Check ('NextPolicy.'+$policy+'.EffectiveFixture') $true
            $policyPin=Hash $Fixture.Config
            foreach($mode in @('Reusable','Worksheet')){
                $work=$null;$moved=$false;$marker=$false
                try{
                    SelectTarget $Fixture 'config-producer';$work=$excel.Workbooks.Add();$sheet=$work.Worksheets.Item(1)
                    $sheet.Cells.Item(2,1).Value2=$Canary;$sheet.Cells.Item(2,2).Formula='=1+2'
                    [void](Probe 'RunLocalReopen' @($work.Name));Prepare-NextPath $work $Canary $mode
                    if($policy -ceq 'StoreUnavailable'){
                        if(-not (Test-Path -LiteralPath $blocked -PathType Container)){throw 'Existing activity store required for Next fault test.'}
                        Move-Item -LiteralPath $blocked -Destination $held;$moved=$true
                        [IO.File]::WriteAllText($blocked,'Disposable Next Batch tracking store unavailable.');$marker=$true
                    }
                    [void](Run 'invSys.Core.xlam' 'TestShippingCatalog.NextPolicyFaultArm' @($policy -ceq 'TerminalFailure'))
                    $before=TrainingPins $Fixture;$activityBefore=@(Get-Slice4beActivityFiles $Fixture);$canonical=Get-NextPathAuthorityPins $Fixture
                    $label='NextPolicy.'+$policy+'.'+$mode;$Decoy.Activate()
                    Invoke-NextPath $Fixture $Canary $mode $label $canonical
                    Check ($label+'.ExactTrackingNotice') ([bool](Probe 'NextPolicyNotice' @($policy)))
                    $after=TrainingPins $Fixture;$fresh=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $activityBefore})
                    Check ($label+'.PriorTrainingRecordsImmutable') (Retained $before $after)
                    if($policy -ceq 'TerminalFailure'){
                        $rows=@($fresh|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json})
                        $requested=$rows.Count -eq 1
                        if($requested){$requested=$rows[0].ControlId -ceq 'PRODUCTION_RUN_NEXT_BATCH' -and $rows[0].OutcomeCode -ceq 'REQUESTED' -and @($rows[0].SourceEventRefs).Count -eq 0 -and -not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadTerminalForTest' @(($rows[0]|ConvertTo-Json -Depth 20 -Compress)))}
                        Check ($label+'.RealTerminalAppendFailureOnce') ([int](Run 'invSys.Core.xlam' 'TestShippingCatalog.NextPolicyFaultCalls') -eq 1)
                        Check ($label+'.IncompleteAttemptNeverTerminal') $requested
                        Check ($label+'.OnlyOneNewTrainingRecord') ($after.Count -eq $before.Count+1)
                    }else{Check ($label+'.NoFallbackOrNewTrainingRecord') ($fresh.Count -eq 0 -and $after.Count -eq $before.Count)}
                    $otherAfter=TrainingPins $Other
                    Check ($label+'.OtherWarehouseTrainingPreserved') ($otherPins.Count -eq $otherAfter.Count -and (Retained $otherPins $otherAfter))
                    Check ($label+'.PolicyBytesPreserved') ((Hash $Fixture.Config) -ceq $policyPin)
                    Check ($label+'.CustomValueAndFormula') ($sheet.Cells.Item(2,1).Value2 -ceq $Canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2')
                    Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
                    if($policy -cin @('StoreUnavailable','TerminalFailure')){CaptureOwnedFormByCaptionEvidence 'Production' ('next-policy-'+$policy.ToLowerInvariant()+'-'+$mode.ToLowerInvariant()+'.png')}
                }finally{
                    [void](Run 'invSys.Core.xlam' 'TestShippingCatalog.NextPolicyFaultArm' @($false))
                    [void](Probe 'CloseDesigner');if($null -ne $work){$work.Close($false)}
                    if($marker){Remove-Item -LiteralPath $blocked};if($moved){Move-Item -LiteralPath $held -Destination $blocked}
                }
            }
        }finally{
            if(@($excel.Workbooks|Where-Object {$_.FullName -ieq $Fixture.Config}).Count){throw 'Close owned Config before restoring bytes.'}
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        }
        Check ('NextPolicy.'+$policy+'.ConfigBytesRestored') ((Hash $Fixture.Config) -ceq $configPin)
    }
    Check 'NextPolicy.PriorTrainingHistoryImmutable' (Retained $prior (TrainingPins $Fixture))
    SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
}
