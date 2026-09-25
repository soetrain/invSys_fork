# The same saved run is evaluated through the actual expectation editor and button.
function Test-ProductionLifecycleDiagnostics($Cases,$Source,$Publication,[string]$EventsPath) {
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationEvidence.ps1')
    $publicationHash=(Get-FileHash $EventsPath).Hash
    $checks=@($Cases)
    foreach($case in $Cases|Where-Object Mode -CEQ 'BeforeAppend'){
        $checks+=@{Name=$case.Name+'.Request';Id=$case.Id;Mode='Request';Outcome='REQUESTED';Record=$case.Attempt;WriterId=$case.WriterId}
    }
    foreach($case in $checks){foreach($kind in @('CommandCompleted','SourceEventsApplied')){
        $label='LifecycleDiagnostic.'+$case.Name+'.'+$kind
        if((Select-EvaluationRun ([string]$Source.ActionPathId)) -cne 'SELECTED'){throw 'Original run selection unavailable.'}
        $steps=@(,@($case.Id,$case.Outcome,'True'))
        # Two FAILED occurrences have different owning facts; consume both in order.
        if($case.Mode -ceq 'AfterAppend'){$steps+=,@($case.Id,$case.Outcome,'True')}
        $ready=Set-EvaluationDraft $steps ($steps.Count-1) $kind -StopAtMissingChoice
        Check ($label+'.ActualExpectationEditor') $ready
        if(-not $ready){continue}
        $before=@(EvaluationFiles|ForEach-Object FullName)
        $invoked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
        $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
        $display=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
        $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
        if($invoked -cne 'DELIVERED' -or $fresh.Count -ne 1){throw 'Actual Evaluate did not append exactly one result.'}
        $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
        $matched=@($result.Matches).Count -eq $steps.Count
        if($matched){$last=$result.Matches[$steps.Count-1];$matched=$last.ActivityId -ceq $case.Record.ActivityId -and $last.Ordinal -eq $case.Record.Ordinal -and $last.ControlId -ceq $case.Id -and $last.OutcomeCode -ceq $case.Outcome}
        Check ($label+'.ExactOriginalTerminal') $matched
        $state='Incomplete';$reason='SOURCE_UNAVAILABLE';$caption='Incomplete evidence'
        if($kind -ceq 'CommandCompleted'){
            $state='Failed';$reason='TERMINAL_NOT_COMPLETED';$caption='Failed'
            if($case.Mode -ceq 'Confirmed'){$state='Concluded';$reason='COMMAND_COMPLETED';$caption='Conclusion observed'}
        } elseif($case.Mode -ceq 'Confirmed'){$state='Concluded';$reason='SOURCE_APPLIED';$caption='Conclusion observed'}
        elseif($case.Mode -ceq 'Pending'){$state='Awaiting';$reason='SOURCE_PENDING';$caption='Awaiting published result'}
        $correct=$result.ResultState -ceq $state -and $reason -cin @($result.ReasonCodes) -and $status.StartsWith($caption,[StringComparison]::Ordinal)
        if($kind -ceq 'CommandCompleted' -and $state -ceq 'Concluded'){$correct=$correct -and $display.Contains('Command completed; Domain application not asserted')}
        Check ($label+'.ExactConclusion') $correct
        Check ($label+'.OriginalRunAndAuthoredIntent') ($result.JournalRecordId -ceq $Source.RecordId -and $result.JournalSha256 -ceq $Source.ContentSha256 -and $result.ExpectedConclusion.TerminalKind -ceq $kind -and @($result.ExpectedConclusion.Steps).Count -eq $steps.Count)
        $refs=@($result.TerminalSources)
        $retained=$refs.Count -eq 0
        if($kind -ceq 'SourceEventsApplied' -and $case.Mode -cin @('AfterAppend','Pending','Confirmed')){
            $retained=$refs.Count -eq 1
            if($retained){
                $ref=$refs[0];$originalRef=$case.Record.SourceEventRefs[0]
                $owner=if($case.Mode -ceq 'Confirmed'){'Applied'}elseif($case.Mode -ceq 'Pending'){'Awaiting'}else{'Unavailable'}
                $retained=$ref.WarehouseId -ceq $originalRef.WarehouseId -and $ref.SourceKind -ceq 'Designs' -and $ref.EventId -ceq $originalRef.EventId -and $ref.SubmissionState -ceq $originalRef.SubmissionState -and $ref.OwnerStatus -ceq $owner -and @($ref.SystemKeys).Count -eq 0
                if($owner -ceq 'Applied'){
                    $group=@($Publication.Groups|Where-Object {$_.Source -ceq 'Designs' -and $_.SourceId -ceq $ref.EventId})
                    $lineBody=[ordered]@{Lines=@($group[0].Lines)}|ConvertTo-Json -Depth 16 -Compress
                    $retained=$retained -and $group.Count -eq 1 -and $ref.LineCount -eq @($group[0].Lines).Count -and $ref.LinesSha256 -ceq (EvaluationSha $lineBody)
                }else{$retained=$retained -and $ref.LineCount -eq 0 -and $ref.LinesSha256 -ceq ''}
            }
        }
        Check ($label+'.ExactSourceEvidenceWithoutInventoryKeys') $retained
        Check ($label+'.ReadOnlyDiagnostic') ((ExpectationControl 'txtPathEvaluation' 'Locked' '' 'frmActionPaths') -ceq 'True')
    }}
    # A diagnostic must also prove the ordered six-action series, not just isolated steps.
    foreach($kind in @('CommandCompleted','SourceEventsApplied')){
        [void](Select-EvaluationRun ([string]$Source.ActionPathId))
        $confirmed=@($Cases|Where-Object Mode -CEQ 'Confirmed')
        $steps=@(foreach($case in $confirmed){,@($case.Id,'CONFIRMED','True')})
        $ready=Set-EvaluationDraft $steps 5 $kind -StopAtMissingChoice
        Check ('LifecycleDiagnostic.Series.'+$kind+'.ActualExpectationEditor') $ready
        if(-not $ready){continue}
        $before=@(EvaluationFiles|ForEach-Object FullName)
        [void](ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
        $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
        if($fresh.Count -ne 1){throw 'Series evaluation unavailable.'}
        $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
        Check ('LifecycleDiagnostic.Series.'+$kind+'.Concluded') ($result.ResultState -ceq 'Concluded')
        Check ('LifecycleDiagnostic.Series.'+$kind+'.OriginalOrderAndExtraAttempts') ((@($result.Matches.ActivityId) -join '|') -ceq (@($confirmed|ForEach-Object {$_.Record.ActivityId}) -join '|') -and @($result.ExtraActivityIds).Count -eq 18)
    }
    Check 'LifecycleDiagnostic.EvaluateDoesNotRepublish' ((Get-FileHash $EventsPath).Hash -ceq $publicationHash)
}
