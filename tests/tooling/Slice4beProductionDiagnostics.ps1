# No terminal inference: explicitly authored expected outcomes against original runs.
function Test-ProductionDiagnostics($Cases,$Source,[string]$EventsPath) {
    $checks=@($Cases)
    foreach($designer in @('Process','Recipe')){
        $valid=@($Cases|Where-Object {$_.Designer -ceq $designer -and $_.Outcome -ceq 'VALIDATED'})[0]
        $checks+=@{Name=$designer+'.Validate.NoDomainEvidence';Id=$valid.Id;Outcome='VALIDATED';Kind='SourceEventsApplied'}
        $checks+=@{Name=$designer+'.Validate.RequestIsNotCompletion';Id=$valid.Id;Outcome='REQUESTED';Kind='CommandCompleted'}
    }
    $publicationHash=(Get-FileHash $EventsPath).Hash
    foreach($case in $checks){
        if((Select-EvaluationRun ([string]$Source.ActionPathId)) -cne 'SELECTED'){throw 'Original Production run selection unavailable.'}
        $kind=if($case.ContainsKey('Kind')){$case.Kind}else{'CommandCompleted'}
        $ready=Set-EvaluationDraft @(,@($case.Id,$case.Outcome,'True')) 0 $kind -StopAtMissingChoice
        $label='ProductionDiagnostic.'+$case.Name
        Check ($label+'.ActualExpectationEditor') $ready
        # A missing supported outcome choice is behavioral RED; keep other cases reachable.
        if(-not $ready){continue}
        $before=@(EvaluationFiles|ForEach-Object FullName)
        $invoked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
        $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
        $display=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
        $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
        if($invoked -cne 'DELIVERED' -or $fresh.Count -ne 1){throw 'Actual Evaluate did not append one result.'}
        $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
        $matches=@($result.Matches);$matched=$matches.Count -eq 1
        if($matched){
            $origin=@($Source.Observations|Where-Object {$_.ActivityId -ceq $matches[0].ActivityId -and $_.OutcomeCode -ceq $case.Outcome})
            $matched=$origin.Count -eq 1 -and $origin[0].ControlId -ceq $case.Id -and $origin[0].Ordinal -eq $matches[0].Ordinal -and @($origin[0].SourceEventRefs).Count -eq 0
        }
        Check ($label+'.ExactOriginalOwnerMatch') $matched
        $positive=$kind -ceq 'CommandCompleted' -and $case.Outcome -cin @('STAGED','VALIDATED')
        $state=if($positive){'Concluded'}elseif($kind -ceq 'SourceEventsApplied'){'Incomplete'}else{'Failed'}
        $reason=if($positive){'COMMAND_COMPLETED'}elseif($kind -ceq 'SourceEventsApplied'){'SOURCE_UNAVAILABLE'}else{'TERMINAL_NOT_COMPLETED'}
        $caption=if($positive){'Conclusion observed'}elseif($state -ceq 'Incomplete'){'Incomplete evidence'}else{'Failed'}
        $correct=$result.ResultState -ceq $state -and $reason -cin @($result.ReasonCodes) -and $status.StartsWith($caption,[StringComparison]::Ordinal)
        if($positive){$correct=$correct -and $display.Contains('Command completed; Domain application not asserted')}
        Check ($label+'.ExactTerminalClassification') $correct
        Check ($label+'.ExactRunAndIntent') ($result.JournalRecordId -ceq $Source.RecordId -and $result.JournalSha256 -ceq $Source.ContentSha256 -and $result.ExpectedConclusion.TerminalKind -ceq $kind -and @($result.ExpectedConclusion.Steps).Count -eq 1 -and $result.ExpectedConclusion.Steps[0].ControlId -ceq $case.Id -and $result.ExpectedConclusion.Steps[0].RequiredOutcome -ceq $case.Outcome)
        if($CaptureProductionDesignerPaths -and ($case.Outcome -ceq 'VALIDATED' -or $case.Outcome -ceq 'REJECTED')){
            CaptureOwnedFormEvidence 'Action Paths' ('diagnostic-'+$case.Name.ToLowerInvariant()+'.png') ([long](ExpectationControl '' 'WindowHandle' '' 'frmActionPaths'))
        }
    }
    Check 'ProductionDiagnostic.EvaluatePreservesPublication' ((Get-FileHash $EventsPath).Hash -ceq $publicationHash)
}
