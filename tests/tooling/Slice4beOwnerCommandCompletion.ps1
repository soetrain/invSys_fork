# D18 completion classification through real recorded commands, the expectation
# editor and Evaluate. Inputs/results remain in memory; reports contain checks.
function Test-OwnerCommandCompletion($Fixture,$BoxingEvidence) {
    if(-not ('OwnerCompletionResources' -as [type])){
        Add-Type @'
using System;using System.Runtime.InteropServices;using System.Text;using System.Collections.Generic;
public static class OwnerCompletionResources {
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")]public static extern uint GetGuiResources(IntPtr p,uint flag);
 public delegate bool EnumProc(IntPtr h,IntPtr p);
 [DllImport("user32.dll")]static extern bool EnumWindows(EnumProc callback,IntPtr p);
 [DllImport("user32.dll")]static extern bool EnumChildWindows(IntPtr parent,EnumProc callback,IntPtr p);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]static extern int GetClassName(IntPtr h,StringBuilder text,int count);
 public static Dictionary<string,int> Classes(uint owner){
  var counts=new Dictionary<string,int>();var seen=new HashSet<IntPtr>();
  EnumProc collect=(h,p)=>{uint id;GetWindowThreadProcessId(h,out id);if(id==owner&&seen.Add(h)){var name=new StringBuilder(256);GetClassName(h,name,256);var key=name.ToString();if(!counts.ContainsKey(key))counts[key]=0;counts[key]++;}return true;};
  EnumWindows((h,p)=>{uint id;GetWindowThreadProcessId(h,out id);if(id==owner){collect(h,p);EnumChildWindows(h,collect,p);}return true;},IntPtr.Zero);return counts;
 }
}
'@
    }
    [uint32]$ownedProcess=0
    [void][OwnerCompletionResources]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$ownedProcess)
    $ownedCreated=(Get-Process -Id $ownedProcess).StartTime.ToUniversalTime().Ticks
    function ObserveOwnerResources([string]$Stage){
        $process=Get-Process -Id $ownedProcess
        if($process.StartTime.ToUniversalTime().Ticks -ne $ownedCreated){throw 'Generated process identity changed.'}
        $gdi=[OwnerCompletionResources]::GetGuiResources($process.Handle,0)
        $user=[OwnerCompletionResources]::GetGuiResources($process.Handle,1)
        [pscustomobject]@{Stage=$Stage;Gdi=$gdi;User=$user;PrivateBytes=$process.PrivateMemorySize64;Classes=[OwnerCompletionResources]::Classes($ownedProcess)}|
            ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'owner-completion-resources.jsonl')
        if($gdi -eq 0 -or $user -eq 0 -or $gdi -ge 9000 -or $user -ge 9000){throw 'Generated Excel GUI capacity unavailable; not terminal-map RED.'}
    }
    ObserveOwnerResources 'Entry'
    if($null -eq $BoxingEvidence -or -not $BoxingEvidence.ActionPathId -or @($BoxingEvidence.Observations).Count -ne 8){
        throw 'Actual stopped Boxing evidence is unavailable; not diagnostic RED.'
    }
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $configHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $entryVisible=[bool]$excel.Visible
    $priorActivity=ActivityPins
    $priorJournal=RestartPins $journalRoot
    $canary='DIAG'+[guid]::NewGuid().ToString('N').Substring(0,12).ToUpperInvariant()
    $uomPath='';$sequence=''
    function ReadOwnerUoms { [string](Run 'invSys.Core.xlam' 'modUomSettings.GetConfiguredUomsPackedText') }
    function ReopenOwnerSettings {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
    }
    try {
        $excel.Visible=$true
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
        OpenRecordingViewer
        $beforeStart=RestartPins $journalRoot
        if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Actual UOM recording Start unavailable.'}
        $starts=@(Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File|Where-Object {-not $beforeStart.ContainsKey($_.FullName)}|ForEach-Object {
            [IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json
        }|Where-Object RecordType -CEQ 'Start')
        if($starts.Count -ne 1){throw 'Actual UOM recording identity unavailable.'}
        $sequence=[string]$starts[0].SequenceId
        $cases=@(
            @{Name='AddCompleted';Action='Add';Value=$canary;Outcome='COMPLETED';Setup=''},
            @{Name='AddUnchanged';Action='Add';Value=$canary;Outcome='UNCHANGED';Setup=''},
            @{Name='RemoveCompleted';Action='Remove';Value=$canary;Outcome='COMPLETED';Setup=''},
            @{Name='RemoveUnchanged';Action='Remove';Value=$canary;Outcome='UNCHANGED';Setup='StaleSelection'},
            @{Name='ResetCompleted';Action='Reset';Value='';Outcome='COMPLETED';Setup='Nondefault'},
            @{Name='ResetUnchanged';Action='Reset';Value='';Outcome='UNCHANGED';Setup=''},
            @{Name='ResetCancelled';Action='Reset';Value='';Outcome='CANCELLED';Setup=''},
            @{Name='AddRejected';Action='Add';Value='';Outcome='REJECTED';Setup=''}
        )
        $ordinal=0
        foreach($case in $cases){
            $ordinal++
            $setupPins=ActivityPins
            if($case.Setup -cne ''){
                if(-not [bool](Run 'invSys.Core.xlam' 'modUomSettings.AddConfiguredUom' @($canary))){throw 'Generated UOM owner setup failed.'}
            }
            ReopenOwnerSettings
            ObserveOwnerResources ('Settings.'+$case.Name)
            if($case.Setup -ceq 'StaleSelection'){
                # The actual form retains its old selection list while the owning
                # test fixture removes that entry. No fake outcome is injected.
                if(-not [bool](Run 'invSys.Core.xlam' 'modUomSettings.RemoveConfiguredUom' @($canary))){throw 'Generated stale UOM selection setup failed.'}
            }
            Check ('OwnerCompletion.Uom.'+$case.Name+'.SetupIsNotAUserClick') ((ActivityPins).Count -eq $setupPins.Count -and (PinsRetained $setupPins))
            $before=@(Get-Slice4beActivityFiles $Fixture)
            $beforeConfig=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            if($case.Action -ceq 'Reset'){
                $choice=if($case.Outcome -ceq 'CANCELLED'){'No'}else{'Yes'}
                [void](Invoke-AdminUomResetChoice $choice ('owner-completion-'+$case.Name.ToLowerInvariant()+'.png'))
            }else{
                [void](Run 'invSys.Admin.xlam' 'TestD5Commands.UomActivityAction' @($case.Action,$case.Value))
            }
            $owner=ReadOwnerUoms
            $same=(Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $beforeConfig
            $ownerCorrect=switch($case.Name){
                'AddCompleted' {$canary -cin $owner.Split('|') -and -not $same}
                'RemoveCompleted' {$canary -cnotin $owner.Split('|') -and -not $same}
                'ResetCompleted' {$owner -ceq 'EA|LB|LBS|OZ|KG|G|GAL|QT|PT|L|ML|CS' -and -not $same}
                'AddUnchanged' {$same -and $canary -cin $owner.Split('|')}
                'RemoveUnchanged' {$same -and $canary -cnotin $owner.Split('|')}
                default {$same}
            }
            Check ('OwnerCompletion.Uom.'+$case.Name+'.IndependentOwnerResult') ([bool]$ownerCorrect)
            $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {[IO.File]::ReadAllText($_)|ConvertFrom-Json})
            $attempts=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED')
            $outcomes=@($records|Where-Object OutcomeCode -CEQ $case.Outcome)
            $control='ADMIN_UOM_'+$case.Action.ToUpperInvariant()
            $observed=$records.Count -eq 2 -and $attempts.Count -eq 1 -and $outcomes.Count -eq 1
            if($observed){
                $observed=$attempts[0].ControlId -ceq $control -and $outcomes[0].ControlId -ceq $control -and
                    $attempts[0].ActivityId -ceq $outcomes[0].ActivityId -and $attempts[0].RecordId -cne $outcomes[0].RecordId -and
                    $attempts[0].SequenceId -ceq $sequence -and $outcomes[0].SequenceId -ceq $sequence -and
                    $attempts[0].Ordinal -eq $ordinal -and $outcomes[0].Ordinal -eq $ordinal -and @($outcomes[0].SourceEventRefs).Count -eq 0
            }
            Check ('OwnerCompletion.Uom.'+$case.Name+'.ExactOriginalRecordedPair') $observed
            if(-not $ownerCorrect -or -not $observed){throw 'Actual UOM owner/recording fixture incomplete; not evaluator RED.'}
        }
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
        Check 'OwnerCompletion.Uom.ActualStop' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
        $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
        if($closed.Count -ne 1 -or $closed[0].Lifecycle -cne 'Stopped' -or @($closed[0].Observations).Count -ne 16){throw 'Actual UOM stopped run unavailable.'}
        $uomPath=[string]$closed[0].ActionPathId
        Check 'OwnerCompletion.Uom.EightActionsCompleteIntegrityChain' ($closed[0].ActionCount -eq 8 -and (JournalChain $sequence 18))
        CloseRecordingViewer
        $published=Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest'
        if($published -isnot [bool] -or -not $published){throw 'Ordinary completion-fixture publication unavailable.'}
        Check 'OwnerCompletion.ActualAdminPublication' $true
        OpenRecordingViewer
        if(-not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))){throw 'Actual completion-fixture Viewer refresh unavailable.'}
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File){$pins[$file.FullName]=Get-ShippingActivityHash $file.FullName}
        $evaluationRoot=[IO.Path]::GetFullPath((Join-Path $journalRoot 'Evaluations')).TrimEnd('\')+'\'
        $checks=@()
        foreach($case in $cases){
            $positive=$case.Outcome -cin @('COMPLETED','UNCHANGED')
            $checks+=@{Name='Uom.'+$case.Name;Path=$uomPath;Control='ADMIN_UOM_'+$case.Action.ToUpperInvariant();Outcome=$case.Outcome;Positive=$positive}
        }
        foreach($verb in @('MAKE','UNBOX')){
            $checks+=@{Name='Boxing.'+$verb+'.Confirmed';Path=$BoxingEvidence.ActionPathId;Control='BOXING_'+$verb;Outcome='CONFIRMED';Positive=$true}
            $checks+=@{Name='Boxing.'+$verb+'.Rejected';Path=$BoxingEvidence.ActionPathId;Control='BOXING_'+$verb;Outcome='REJECTED';Positive=$false}
        }
        foreach($case in $checks){
            ObserveOwnerResources ('Before.'+$case.Name)
            if((Select-EvaluationRun $case.Path) -cne 'SELECTED'){throw 'Actual recorded task selection unavailable.'}
            $steps=@(,@($case.Control,$case.Outcome,'True'))
            $ready=Set-EvaluationDraft $steps 0 'CommandCompleted'
            Check ('OwnerCompletion.'+$case.Name+'.ActualExpectationEditor') $ready
            if(-not $ready){throw 'Existing outcome/editor fixture unavailable; not terminal-map RED.'}
            $beforeResults=@(EvaluationFiles|ForEach-Object FullName)
            $invoked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
            $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
            $text=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $beforeResults})
            if($invoked -cne 'DELIVERED' -or $fresh.Count -ne 1){throw 'Actual Evaluate did not save one result; not terminal-map RED.'}
            $result=[IO.File]::ReadAllText($fresh[0].FullName)|ConvertFrom-Json
            $matched=@($result.Matches)
            Check ('OwnerCompletion.'+$case.Name+'.MatchesExactObservedOwnerFact') ($matched.Count -eq 1 -and
                $matched[0].ControlId -ceq $case.Control -and $matched[0].OutcomeCode -ceq $case.Outcome -and
                @($result.MissingSteps).Count -eq 0 -and @($result.FailedSteps).Count -eq 0 -and
                @($result.UnavailableSteps).Count -eq 0 -and $result.Publication.Availability -ceq 'Loaded' -and
                @($result.ReasonCodes).Count -eq 1 -and $result.ReasonCodes[0] -cin @('COMMAND_COMPLETED','TERMINAL_NOT_COMPLETED'))
            $expected=if($case.Positive){'Concluded'}else{'Failed'}
            $expectedCaption=if($case.Positive){'Conclusion observed'}else{'Failed'}
            $correct=$result.ResultState -ceq $expected -and $status.StartsWith($expectedCaption,[StringComparison]::Ordinal)
            if($case.Positive){$correct=$correct -and $text.Contains('Command completed; Domain application not asserted')}
            Check ('OwnerCompletion.'+$case.Name+'.ConclusionFromOwnerFact') $correct
            Check ('OwnerCompletion.'+$case.Name+'.ExactRunAndIntent') ($result.ActionPathId -ceq $case.Path -and
                $result.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted' -and @($result.ExpectedConclusion.Steps).Count -eq 1 -and
                $result.ExpectedConclusion.Steps[0].ControlId -ceq $case.Control -and $result.ExpectedConclusion.Steps[0].RequiredOutcome -ceq $case.Outcome)
            Capture-BoxingFormEvidence 'Action Paths' ('owner-completion-'+$case.Name.ToLowerInvariant()+'.png') ('OwnerCompletion.'+$case.Name+'.VisibleResult')
            ObserveOwnerResources ('After.'+$case.Name)
        }
        $preserved=$true
        foreach($path in $pins.Keys){$preserved=$preserved -and (Test-Path -LiteralPath $path) -and (Get-ShippingActivityHash $path) -ceq $pins[$path]}
        $newFiles=@(Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object {-not $pins.ContainsKey($_.FullName)})
        $onlyResults=$newFiles.Count -eq $checks.Count
        foreach($file in $newFiles){$onlyResults=$onlyResults -and $file.FullName.StartsWith($evaluationRoot,[StringComparison]::OrdinalIgnoreCase) -and $file.Extension -ceq '.json'}
        Check 'OwnerCompletion.EvaluationPreservesAllOriginalBytes' $preserved
        Check 'OwnerCompletion.OnlySeparateEvaluationResultsAppended' $onlyResults
        Check 'OwnerCompletion.PriorActivityPreserved' (PinsRetained $priorActivity)
        $journalsPreserved=$true
        foreach($path in $priorJournal.Keys){$journalsPreserved=$journalsPreserved -and (Get-FileHash -LiteralPath $path).Hash -ceq $priorJournal[$path]}
        Check 'OwnerCompletion.PriorJournalsPreserved' $journalsPreserved
    } finally {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
        CloseRecordingViewer
        foreach($book in @($excel.Workbooks)){
            if([string]::Equals($book.FullName,$Fixture.Config,[StringComparison]::OrdinalIgnoreCase)){$book.Close($false)}
        }
        # Only the already Admin-generated, disposable fixture is restored.
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        SelectTarget $Fixture 'config-admin'
        [void](Run 'invSys.Core.xlam' 'modConfig.Reload')
        $excel.Visible=$entryVisible
    }
    Check 'OwnerCompletion.GeneratedConfigRestoredExactly' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configHash)
}
