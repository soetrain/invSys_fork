# D18 exact Settings terminal facts through actual expectation and Evaluate handlers.
function Write-SettingsHostEvidence([string]$Stage,[bool]$ReadExcel=$true) {
    $row=[ordered]@{Stage=$Stage;ObservedUTC=[DateTimeOffset]::UtcNow.ToString('o');ReadOnlyInspection=$true}
    if($ReadExcel){
        $books=$null;$process=$null
        try{
            if(-not ('SettingsHostOwner' -as [type])){
                Add-Type @'
using System;using System.Runtime.InteropServices;
public static class SettingsHostOwner {
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")]public static extern uint GetGuiResources(IntPtr h,uint flag);
}
'@
            }
            $window=$excel.Hwnd
            if($null -eq $window -or [long]$window -eq 0){throw 'Native identity unavailable.'}
            [uint32]$owner=0
            [void][SettingsHostOwner]::GetWindowThreadProcessId([IntPtr]$window,[ref]$owner)
            if($owner -notin $initialExcelProcessIds){throw 'Native owner changed.'}
            $process=Get-Process -Id $owner
            $books=$excel.Workbooks;$count=$books.Count
            if($count -isnot [int]){throw 'Typed workbook count unavailable.'}
            $row.ProcessId=$owner;$row.CreatedUTC=$process.StartTime.ToUniversalTime().ToString('o')
            $row.NativeWindow=[long]$window;$row.WorkbookCount=$count;$row.TypedCountVerified=$true
            $row.CPU=$process.CPU;$row.PrivateBytes=$process.PrivateMemorySize64;$row.Handles=$process.HandleCount
            $row.Gdi=[SettingsHostOwner]::GetGuiResources($process.Handle,0);$row.User=[SettingsHostOwner]::GetGuiResources($process.Handle,1)
        }catch{$row.TypedCountVerified=$false;$row.InspectionHResult=$_.Exception.HResult}
        finally{
            if($null -ne $books){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($books)}
            if($null -ne $process){$process.Dispose()}
        }
    }
    try{[pscustomobject]$row|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'settings-host-evidence.jsonl')}catch{}
}

function Save-SettingsDiagnosticRun($Evidence,[string]$Name,[string]$Sequence,[int]$Actions) {
    $interrupted=$Name -ceq 'SaveCancelled.Tracking'
    if(-not $interrupted -and (RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED'){throw 'Actual Settings diagnostic Stop unavailable.'}
    $closed=@(RecordingJournal $Sequence|Where-Object RecordType -CEQ 'Close')
    if($interrupted){
        $valid=$closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Incomplete' -and $closed[0].ReasonCode -ceq 'POLICY_CHANGED' -and
            $closed[0].ActionCount -eq 1 -and @($closed[0].Observations).Count -eq 1 -and
            $closed[0].Observations[0].OutcomeCode -ceq 'REQUESTED' -and (JournalChain $Sequence 3)
    }else{
        $valid=$closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq $Actions -and
            @($closed[0].Observations).Count -eq (2*$Actions) -and (JournalChain $Sequence (2*$Actions+2))
    }
    Check ('SettingsDiagnostic.Recording.'+$Name+'.ExactOriginalLifecycle') $valid
    if(-not $valid){throw 'Original Settings recording lifecycle unavailable; not classifier RED.'}
    $Evidence[$Name]=$closed[0]
}

function Test-SettingsDiagnostics($Fixture,$Evidence) {
    CloseSettingsEditors
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Ordinary publication unavailable; not classifier RED.'}
    Check 'SettingsDiagnostic.ActualAdminPublication' $true
    OpenRecordingViewer
    if(-not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))){throw 'Actual Viewer refresh unavailable.'}
    $cases=@()
    foreach($action in @(SettingsActivityCases)){
        $cases+=@{Name=$action.Id+'.'+$action.Outcome;Run='Main';Control=$action.Id;Outcome=$action.Outcome;State='Concluded';Reason='COMMAND_COMPLETED'}
    }
    foreach($id in @('ADMIN_PATH_PREFERENCE_SAVE','VIEWER_PATH_PREFERENCE_SAVE')){
        $cases+=@{Name=$id+'.UNCHANGED';Run='Unchanged';Control=$id;Outcome='UNCHANGED';State='Concluded';Reason='COMMAND_COMPLETED'}
    }
    foreach($section in @('Tracking','Detail')){
        $id=if($section -ceq 'Tracking'){'ADMIN_TRACKING_SAVE'}else{'ADMIN_DETAIL_SAVE'}
        foreach($failure in @('StaleVersion','SaveCancelled')){
            $outcome=if($failure -ceq 'StaleVersion'){'REJECTED'}else{'FAILED'}
            if($failure -ceq 'SaveCancelled' -and $section -ceq 'Tracking'){
                $cases+=@{Name=$id+'.CancelledWriteInterrupted';Run=$failure+'.'+$section;Control=$id;Outcome='REQUESTED';State='Incomplete';Reason='CAPTURE_INCOMPLETE'}
            }else{
                $cases+=@{Name=$id+'.'+$outcome;Run=$failure+'.'+$section;Control=$id;Outcome=$outcome;State='Failed';Reason='TERMINAL_NOT_COMPLETED'}
            }
        }
        $cases+=@{Name=$id+'.DENIED';Run='SaveDenial';Control=$id;Outcome='DENIED';State='Failed';Reason='TERMINAL_NOT_COMPLETED'}
    }
    foreach($run in @('FailedProfileRead','SaveDenial')){
        $cases+=@{Name=$run+'.ReloadFailed';Run=$run;Control='ADMIN_DETAIL_RELOAD';Outcome='FAILED';State='Failed';Reason='TERMINAL_NOT_COMPLETED'}
    }
    foreach($id in @('ADMIN_DETAIL_SAVE','ADMIN_DETAIL_SELECT_FIELD','ADMIN_TRACKING_RESET','VIEWER_PATH_PREFERENCE_SAVE')){
        $cases+=@{Name=$id+'.REQUESTED';Run='Main';Control=$id;Outcome='REQUESTED';State='Failed';Reason='TERMINAL_NOT_COMPLETED'}
    }
    $cases+=@{Name='PolicySave.Interrupted';Run='PolicySave';Control='ADMIN_TRACKING_SAVE';Outcome='REQUESTED';State='Incomplete';Reason='CAPTURE_INCOMPLETE'}
    foreach($action in @((SettingsActivityCases)|Where-Object {$_.Id -cin @('ADMIN_DETAIL_SAVE','ADMIN_TRACKING_RESET','ADMIN_DETAIL_SELECT_FIELD','ADMIN_DETAIL_RELOAD','VIEWER_PATH_PREFERENCE_SAVE')})){
        $cases+=@{Name=$action.Id+'.NoDomainEvidence';Run='Main';Control=$action.Id;Outcome=$action.Outcome;State='Incomplete';Reason='SOURCE_UNAVAILABLE';Kind='SourceEventsApplied'}
    }
    # In the real main run, Detail Save precedes Reset. A later unsaved Reset
    # cannot stand in for the Save required after that Reset.
    $cases+=@{Name='StagedResetCannotReplaceLaterSave';Run='Main';Control='ADMIN_DETAIL_RESET';Outcome='STAGED';State='Failed';Reason='REQUIRED_STEP_MISSING';RequireSave=$true}
    $pins=@{}
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    $evaluationRoot=[IO.Path]::GetFullPath((Join-Path $journalRoot 'Evaluations')).TrimEnd('\')+'\'
    $caseNumber=0
    foreach($case in $cases){
        if($caseNumber % 10 -eq 0){Write-SettingsHostEvidence ('BeforeEvaluation.'+$caseNumber)}
        $caseNumber++
        if(-not $Evidence.ContainsKey($case.Run)){throw 'Original diagnostic recording missing.'}
        $source=$Evidence[$case.Run]
        $path=[string]$source.ActionPathId
        if((Select-EvaluationRun $path) -cne 'SELECTED'){throw 'Actual recorded Settings task selection unavailable.'}
        $kind=if($case.ContainsKey('Kind')){$case.Kind}else{'CommandCompleted'}
        $steps=@(,@($case.Control,$case.Outcome,'True'));$terminal=0
        if($case.ContainsKey('RequireSave')){$steps+=,@('ADMIN_DETAIL_SAVE','COMPLETED','True');$terminal=1}
        $label='SettingsDiagnostic.'+$case.Name
        $ready=Set-EvaluationDraft $steps $terminal $kind
        Check ($label+'.ActualExpectationEditor') $ready
        if(-not $ready){throw 'Existing Settings outcome/editor fixture unavailable; not classifier RED.'}
        $before=@(EvaluationFiles|ForEach-Object FullName)
        $invoked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
        $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
        $display=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
        $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
        if($invoked -cne 'DELIVERED' -or $fresh.Count -ne 1){throw 'Actual Evaluate did not append one result; not classifier RED.'}
        $result=[IO.File]::ReadAllText($fresh[0].FullName)|ConvertFrom-Json
        $matches=@($result.Matches)
        $matched=$matches.Count -eq 1 -and $matches[0].ControlId -ceq $case.Control -and $matches[0].OutcomeCode -ceq $case.Outcome
        if($matched){
            $original=@($source.Observations|Where-Object {$_.ActivityId -ceq $matches[0].ActivityId -and $_.OutcomeCode -ceq $case.Outcome})
            $matched=$original.Count -eq 1 -and $original[0].ControlId -ceq $case.Control -and
                $original[0].Ordinal -eq $matches[0].Ordinal -and @($original[0].SourceEventRefs).Count -eq 0
        }
        Check ($label+'.MatchesExactOriginalOwnerFact') ($matched -and @($result.FailedSteps).Count -eq 0 -and @($result.UnavailableSteps).Count -eq 0 -and $result.Publication.Availability -ceq 'Loaded')
        $expectedCaption=switch($case.State){'Concluded'{'Conclusion observed'} 'Incomplete'{'Incomplete evidence'} default{'Failed'}}
        $correct=$result.ResultState -ceq $case.State -and $status.StartsWith($expectedCaption,[StringComparison]::Ordinal) -and
            $case.Reason -cin @($result.ReasonCodes)
        if($case.State -ceq 'Concluded'){$correct=$correct -and $display.Contains('Command completed; Domain application not asserted')}
        Check ($label+'.ExactTerminalClassification') $correct
        Check ($label+'.ExactRunAndAuthoredIntent') ($result.ActionPathId -ceq $path -and
            $result.JournalRecordId -ceq $source.RecordId -and $result.JournalSha256 -ceq $source.ContentSha256 -and
            $result.ExpectedConclusion.TerminalKind -ceq $kind -and @($result.ExpectedConclusion.Steps).Count -eq $steps.Count -and
            $result.ExpectedConclusion.Steps[0].ControlId -ceq $case.Control -and $result.ExpectedConclusion.Steps[0].RequiredOutcome -ceq $case.Outcome)
        if($case.ContainsKey('RequireSave')){
            Check ($label+'.MissingRequiredSaveExplained') (@($result.MissingSteps).Count -eq 1 -and $display.Contains('Missing expected step') -and $display.Contains('Save Detail Profile'))
        }else{Check ($label+'.NoMissingExpectedStep') (@($result.MissingSteps).Count -eq 0)}
        Check ($label+'.DoesNotInventDomainEvidence') (@($result.TerminalSources).Count -eq 0)
        CaptureOwnedFormByCaptionEvidence 'Action Paths' ('settings-diagnostic-'+$case.Name.ToLowerInvariant()+'.png')
        Check ($label+'.VisibleResult') $true
    }
    $preserved=$true
    foreach($path in $pins.Keys){$preserved=$preserved -and (Test-Path -LiteralPath $path) -and (Get-FileHash -LiteralPath $path).Hash -ceq $pins[$path]}
    $newFiles=@(Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object {-not $pins.ContainsKey($_.FullName)})
    $onlyResults=$newFiles.Count -eq $cases.Count
    foreach($file in $newFiles){$onlyResults=$onlyResults -and $file.FullName.StartsWith($evaluationRoot,[StringComparison]::OrdinalIgnoreCase) -and $file.Extension -ceq '.json'}
    Check 'SettingsDiagnostic.EvaluationPreservesOriginalBytes' $preserved
    Check 'SettingsDiagnostic.OnlySeparateEvaluationResultsAppended' $onlyResults
    Write-SettingsHostEvidence 'AfterEvaluations'
    CloseRecordingViewer
}
