# Use the immutable Receiving guide with two independently recorded operator runs.
function Test-ReceivingGuideComparison($Guide,$Source,$Staging) {
    $evaluationRoot=Join-Path $journalRoot 'Evaluations'
    $guidePins=BoundPins (Join-Path $journalRoot 'Guides')
    $runIds=@($Source.ActionPathId)
    function Receive([string]$Action) {
        [string](Run 'invSys.Operations.xlam' 'modTS_Received.ReplaySourceForTest' @($Action))
    }
    foreach($stage in 'complete','missing-confirm'){
        $complete=$stage -ceq 'complete'
        $prefix='ReceivingGuideComparison.'+$stage+'.'
        Deliver 'frmActionPaths' 'btnClose' 'Click'
        $prior=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json' | ForEach-Object Name)
        if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Could not start the independent Receiving recording.'}
        if((Receive 'Refresh') -notmatch '^[1-9][0-9]*$' -or (Receive 'Add') -cne 'True'){throw 'Receiving did not stage the observed training receipt.'}
        if($complete -and (Receive 'Confirm') -cne 'True'){throw 'Receiving did not confirm the observed training receipt.'}
        if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED'){throw 'Could not stop the independent Receiving recording.'}
        if(-not $complete){[void](Receive 'Clear')}
        Check ($prefix+'StagingCleared') ($Staging.ListRows.Count -eq 0)
        $closed=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json' |
            Where-Object Name -CNotIn $prior | ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json} |
            Where-Object RecordType -CEQ 'Close')
        if($closed.Count -ne 1 -or $closed[0].ActionPathId -cin $runIds){throw 'A separate stopped recording was required.'}
        $observed=$closed[0]
        $runIds+=$observed.ActionPathId
        $confirmed=@($observed.Observations|Where-Object {$_.ControlId -ceq 'RECEIVING_CONFIRM_WRITES' -and $_.OutcomeCode -ceq 'CONFIRMED'})
        Check ($prefix+'OriginalConfirmMatchesOperatorAction') ($confirmed.Count -eq [int]$complete)
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'The observed Receiving result could not be published.'}
        CloseRecordingViewer
        OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $observed.ActionPathId) -cne 'SELECTED'){throw 'Could not select the new observed recording.'}
        BoundOpen
        BoundSelect $Guide
        Deliver 'frmActionPathLibrary' 'btnUseGuideForRun' 'Click'
        Deliver 'frmActionPathLibrary' 'btnCloseGuides' 'Click'
        $before=BoundPins $evaluationRoot
        Deliver 'frmActionPaths' 'btnEvaluatePath' 'Click'
        $added=@(Get-ChildItem -LiteralPath $evaluationRoot -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)})
        if($added.Count -ne 1){throw 'Expected exactly one saved evaluation for the observed run.'}
        $result=Get-Content -Raw -LiteralPath $added[0].FullName|ConvertFrom-Json
        Check ($prefix+'ExactGuideAndObservedJournal') ($result.Guide.ContentSha256 -ceq $Guide.ContentSha256 -and $result.Guide.Version -eq $Guide.Version -and $result.ActionPathId -ceq $observed.ActionPathId -and $result.JournalSha256 -ceq $observed.ContentSha256)
        if($complete){
            Check ($prefix+'AppliedObservedReceiptConcludes') ($result.ResultState -ceq 'Concluded' -and $confirmed.Count -eq 1 -and @($result.TerminalSources|Where-Object {$_.OwnerStatus -ceq 'Applied' -and $_.EventId -cin @($confirmed[0].SourceEventRefs|ForEach-Object EventId)}).Count -eq 1)
        }else{
            Check ($prefix+'CannotBorrowGuideSourceConfirmation') ($result.ResultState -ceq 'Failed' -and $Guide.ExpectedConclusion.TerminalStepId -cin @($result.MissingSteps))
        }
        Deliver 'frmActionPaths' 'btnViewActionPath' 'Click'
        foreach($method in 'How-To','Diagnostic','Compare both'){
            Deliver 'frmActionPathView' 'cboActionPathView' 'Write' $method
            $pair=BoundControl 'lblActionPathPair' 'Label' '' 'frmActionPathView'
            $how=BoundControl 'txtActionPathHowTo' 'Text' '' 'frmActionPathView'
            $diagnostic=BoundControl 'txtActionPathDiagnostic' 'Text' '' 'frmActionPathView'
            $status=if($complete){$diagnostic.Contains('Conclusion observed')}else{$diagnostic.Contains('Missing expected step') -and -not $diagnostic.Contains('Conclusion observed')}
            Check ($prefix+$method+'.PairTextAndConclusion') ($pair.Contains($Guide.ContentSha256) -and $pair.Contains($observed.ActionPathId) -and $how.Contains($Guide.Instructions) -and $diagnostic.Contains($result.EvaluationId) -and $status)
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ('receiving-guide-'+$stage+'-'+$method.Replace(' ','-')+'.png')
        }
        if((BoundControl 'txtActionPathDiagnostic' 'ViewportBottom' '' 'frmActionPathView') -cne 'DELIVERED'){throw 'The diagnostic result could not be scrolled into view.'}
        CaptureOwnedFormByCaptionEvidence 'Action Path view' ('receiving-guide-'+$stage+'-result.png')
        Deliver 'frmActionPathView' 'btnCloseActionPathView' 'Click'
    }
    Check 'ReceivingGuideComparison.OriginalGuidePreserved' (BoundSame $guidePins (BoundPins (Join-Path $journalRoot 'Guides')))
}
