# D18 behavioral expectations through the packaged editor and Evaluate handlers.
# Fixed authored intent is separate from the actual saved observation stream.
function Select-EvaluationRun([string]$Id) {
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Open',''))
    [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Select',$Id))
}

function Set-EvaluationDraft($Steps,[int]$Terminal,[string]$Kind,[string]$Form='frmActionPaths',[switch]$StopAtMissingChoice) {
    $entry=if($Form -eq 'frmInventoryViewer'){'btnRecordingExpectation'}else{'btnExpectedConclusion'}
    $ok=(ExpectationControl $entry 'Click' '' $Form) -ceq 'DELIVERED'
    # Remove a captured/default draft through the real editor, without changing
    # its source recording. A missing control is an observation, not a fake seam.
    $count=ExpectationControl 'lstExpectedSteps' 'Count'
    if($count -match '^[0-9]+$'){
        for($i=[int]$count-1;$i -ge 0;$i--){
            [void](ExpectationControl 'lstExpectedSteps' 'Index' ([string]$i))
            $removed=ExpectationControl 'btnRemoveExpectedStep' 'Click'
            $ok=$ok -and $removed -ceq 'DELIVERED'
        }
    }
    foreach($step in $Steps){
        $control=ExpectationControl 'cboExpectedControl' 'Select' $step[0]
        $outcome=ExpectationControl 'cboExpectedOutcome' 'Select' $step[1]
        if($StopAtMissingChoice -and ($control -cne 'DELIVERED' -or $outcome -cne 'DELIVERED')){
            [void](ExpectationControl 'btnCancelExpectation' 'Click')
            return $false
        }
        $retry=ExpectationControl 'chkExpectedRetry' 'Boolean' $step[2]
        $add=ExpectationControl 'btnAddExpectedStep' 'Click'
        $ok=$ok -and $control -ceq 'DELIVERED' -and $outcome -ceq 'DELIVERED' -and $retry -ceq 'DELIVERED' -and $add -ceq 'DELIVERED'
    }
    if($Steps.Count -gt 0){$selected=ExpectationControl 'cboTerminalStep' 'Index' ([string]$Terminal);$ok=$ok -and $selected -ceq 'DELIVERED'}
    $kindSelected=ExpectationControl 'cboTerminalKind' 'Select' $Kind
    $used=ExpectationControl 'btnUseExpectation' 'Click'
    return ($ok -and $kindSelected -ceq 'DELIVERED' -and $used -ceq 'DELIVERED')
}

function Assert-EvaluationStatus([string]$Name,[string]$Expected,[bool]$Ready=$true) {
    $invoked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
    $status=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
    Check ('EvaluationContract.'+$Name) ($Ready -and $invoked -ceq 'DELIVERED' -and $status.StartsWith($Expected,[StringComparison]::Ordinal))
}

function Test-EvaluationLoadedPublication([string]$Stage) {
    # Previous stage's explicit expectation and selected run remain loaded.
    # A later file is not evidence until the operator explicitly refreshes.
    $before=(Get-FileHash -LiteralPath $eventsPath).Hash
    Assert-EvaluationStatus ($Stage+'.NewPublicationDoesNotAdvanceLoadedEvidence') 'Awaiting published result'
    Check ('EvaluationContract.'+$Stage+'.EvaluateDoesNotRepublish') ((Get-FileHash -LiteralPath $eventsPath).Hash -ceq $before)
}

function Test-EvaluationContracts {
    $originalPath=$pathId
    $cases=@(
        @{Name='NoneCannotConclude';Steps=@();Terminal=0;Kind='None';Status='Incomplete evidence'},
        @{Name='CommandCompletion';Steps=@(,@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'));Terminal=0;Kind='CommandCompleted';Status='Conclusion observed'},
        @{Name='RepeatedRequiredActionsNeedDistinctOccurrences';Steps=@(@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'),@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'),@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'));Terminal=2;Kind='CommandCompleted';Status='Failed'},
        @{Name='RequiredOrderCannotBeReconstructed';Steps=@(@('RECEIVING_CONFIRM_WRITES','PENDING','True'),@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'),@('RECEIVING_CONFIRM_WRITES','PENDING','True'));Terminal=2;Kind='SourceEventsApplied';Status='Failed'}
    )
    foreach($case in $cases){
        if((Select-EvaluationRun $originalPath) -cne 'SELECTED'){throw 'Existing Operations run selection failed.'}
        $ready=Set-EvaluationDraft $case.Steps $case.Terminal $case.Kind
        Assert-EvaluationStatus $case.Name $case.Status $ready
        if($case.Name -eq 'CommandCompletion'){
            $text=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
            Check 'EvaluationContract.CommandConclusionDoesNotAssertDomainApplication' ($ready -and $text.Contains('Command completed; Domain application not asserted'))
        }
        if($case.Name -eq 'RepeatedRequiredActionsNeedDistinctOccurrences'){
            $text=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
            Check 'EvaluationContract.MissingExpectedStepExplainedInResult' ($ready -and $text.Contains('Missing expected step') -and $text.Contains('Save Value'))
        }
    }
    [void](Select-EvaluationRun $originalPath)
    $terminalSteps=@(@('RECEIVING_CONFIRM_WRITES','PENDING','True'),@('RECEIVING_CONFIRM_WRITES','PENDING','True'))
    $ready=Set-EvaluationDraft $terminalSteps 1 'SourceEventsApplied'
    Assert-EvaluationStatus 'TwoDistinctSubmissionsConcludeOnlyFromAllTerminalReferences' 'Conclusion observed' $ready
    $publication=[IO.File]::ReadAllBytes($eventsPath)
    $body=[Text.Encoding]::UTF8.GetString($publication)
    try {
        [IO.File]::WriteAllText($eventsPath,$body.Replace('"ContentSha256":"','"ContentSha256":"0'),[Text.UTF8Encoding]::new($false))
        Assert-EvaluationStatus 'EvaluateUsesValidatedLoadWithoutRereadingChangedFile' 'Conclusion observed' $ready
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
        $beforeStaleEvaluation=@(EvaluationFiles|ForEach-Object FullName)
        Assert-EvaluationStatus 'FailedRefreshInvalidatesConclusion' 'Incomplete evidence' $ready
        $staleResults=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $beforeStaleEvaluation})
        $staleReferences=$false
        if($staleResults.Count -eq 1){
            $staleResult=Get-Content -LiteralPath $staleResults[0].FullName -Raw|ConvertFrom-Json
            $staleReferences=$staleResult.ResultState -ceq 'Incomplete' -and $staleResult.Publication.Availability -ceq 'Stale' -and
                @($staleResult.TerminalSources).Count -eq $submissionIds.Count
            foreach($id in $submissionIds){
                $ref=@($staleResult.TerminalSources|Where-Object EventId -CEQ $id)
                $staleReferences=$staleReferences -and $ref.Count -eq 1 -and $ref[0].SubmissionState -ceq 'Submitted' -and
                    $ref[0].OwnerStatus -ceq 'Unavailable' -and $ref[0].LineCount -eq 0 -and $ref[0].LinesSha256 -ceq '' -and @($ref[0].SystemKeys).Count -eq 0
            }
        }
        Check 'EvaluationContract.StaleResultRetainsEveryTerminalReferenceAsUnavailable' $staleReferences
        [IO.File]::WriteAllBytes($eventsPath,$publication)
        Assert-EvaluationStatus 'RestoringFileWithoutRefreshDoesNotClearStaleEvidence' 'Incomplete evidence' $ready
    } finally {
        [IO.File]::WriteAllBytes($eventsPath,$publication)
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
    }
    Assert-EvaluationStatus 'ExplicitRefreshRestoresSupportedConclusion' 'Conclusion observed' $ready
    # A later policy hides a required Admin observation in an already loaded run.
    [void](Select-EvaluationRun $originalPath)
    $command=@(,@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'))
    $ready=Set-EvaluationDraft $command 0 'CommandCompleted'
    try {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PublishedReadVisibilityForTest' @($false))){throw 'Actual policy restriction fixture failed.'}
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
        Assert-EvaluationStatus 'CurrentRestrictionCannotUsePreviouslyLoadedObservation' 'Incomplete evidence' $ready
    } finally {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.PublishedReadVisibilityForTest' @($true))){throw 'Actual policy restoration fixture failed.'}
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
    }
    [void](Select-EvaluationRun $originalPath)
    $missing=@(@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'),@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'),@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'))
    $ready=Set-EvaluationDraft $missing 2 'CommandCompleted'
    Assert-EvaluationStatus 'HistoricalEligibleCaptureSurvivesLaterPolicyVersions' 'Failed' $ready

    # New explicit run: a real rejected Admin Save followed by a real successful
    # retry. No source observation is fabricated or rewritten to create failure.
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Actual recorder baseline is unavailable.'}
    $captureReady=Set-EvaluationDraft $command 0 'CommandCompleted' 'frmInventoryViewer'
    Check 'EvaluationContract.ActiveRecordingCanStageExpectedConclusion' $captureReady
    $before=ActivityPins
    $configBefore=Get-ReceivingFixtureHash $Fixture.Config
    try {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        $accepted=[bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','not-a-number'))
    } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    $records=@(foreach($file in Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File){if(-not $before.ContainsKey($file.Name)){Get-Content -LiteralPath $file.FullName -Raw|ConvertFrom-Json}})
    $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$rejected=@($records|Where-Object OutcomeCode -CEQ 'REJECTED')
    if($accepted -or $records.Count -ne 2 -or $attempt.Count -ne 1 -or $rejected.Count -ne 1){throw 'Actual rejected-action fixture failed; not evaluator RED.'}
    Check 'EvaluationContract.RejectedOwnerActionPreservesConfigBytes' ((Get-ReceivingFixtureHash $Fixture.Config) -ceq $configBefore)
    $retry=SaveRecordedSetting '671'
    $retrySequence=[string]$retry.Attempt.SequenceId
    Check 'EvaluationContract.RejectedThenSuccessfulOwningActionsAreDistinct' ((HasSequence $retry 2) -and $attempt[0].SequenceId -ceq $retrySequence -and $attempt[0].ActivityId -cne $retry.Attempt.ActivityId)
    [void](RecordingControl 'Stop Recording' 'Click')
    $close=@(RecordingJournal $retrySequence|Where-Object RecordType -CEQ 'Close')
    if($close.Count -ne 1 -or -not (JournalChain $retrySequence 6)){throw 'Real retry journal fixture failed.'}
    $retryPath=[string]$close[0].ActionPathId
    $captured=$false
    if($null -ne $close[0].PSObject.Properties['ExpectedConclusion']){
        $definition=$close[0].ExpectedConclusion
        $captured=$close[0].SchemaVersion -eq 2 -and $definition.SchemaVersion -eq 1 -and $definition.TerminalKind -ceq 'CommandCompleted' -and
            @($definition.Steps).Count -eq 1 -and $definition.Steps[0].ControlId -ceq 'ADMIN_SETTINGS_SAVE_VALUE' -and
            $definition.Steps[0].RequiredOutcome -ceq 'COMPLETED' -and $definition.Steps[0].RetryAllowed -eq $true -and
            $definition.Steps[0].StepId -ceq $definition.TerminalStepId
    }
    Check 'EvaluationContract.StopFreezesExplicitSchema2Expectation' ($captureReady -and $captured)
    $retryPins=RestartPins $journalRoot
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
    if((Select-EvaluationRun $retryPath) -cne 'SELECTED'){throw 'Actual retry recording is not selectable.'}
    $summary=ExpectationControl 'lblExpectationSummary' 'Caption' '' 'frmActionPaths'
    Check 'EvaluationContract.CapturedExpectationProvenanceVisible' ($captureReady -and $summary.Contains('Captured expectation'))
    Assert-EvaluationStatus 'CapturedRetryAllowedCanConclude' 'Conclusion observed' $captureReady
    foreach($allowed in @('False','True')){
        [void](Select-EvaluationRun $retryPath)
        $steps=@(,@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED',$allowed))
        $ready=Set-EvaluationDraft $steps 0 'CommandCompleted'
        $expected=if($allowed -eq 'True'){'Conclusion observed'}else{'Failed'}
        Assert-EvaluationStatus ('RetryAllowed'+$allowed) $expected $ready
        if($allowed -eq 'False'){
            $text=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
            Check 'EvaluationContract.MismatchedExpectedOutcomeExplainedInResult' ($ready -and $text.Contains('Outcome mismatch') -and $text.Contains('Save Value'))
        }
        $observed=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Evidence',''))
        Check ('EvaluationContract.Retry'+$allowed+'RetainsRejectedAndSuccessfulOccurrences') ($observed.Contains($attempt[0].ActivityId) -and $observed.Contains($retry.Attempt.ActivityId) -and $observed.Contains('REJECTED') -and $observed.Contains('COMPLETED'))
    }
    $unchanged=$true
    foreach($path in $retryPins.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $retryPins[$path]){$unchanged=$false}}
    Check 'EvaluationContract.AnalysisNeverRewritesCapturedExpectationOrObservations' $unchanged
    [void](Select-EvaluationRun $retryPath)
    $opened=ExpectationControl 'btnExpectedConclusion' 'Click' '' 'frmActionPaths'
    $draftCount=ExpectationControl 'lstExpectedSteps' 'Count'
    [void](Select-EvaluationRun $originalPath)
    [void](ExpectationControl 'btnUseExpectation' 'Click')
    $staleCount=ExpectationControl 'lstExpectedSteps' 'CountAnyVisibility'
    $summary=ExpectationControl 'lblExpectationSummary' 'Caption' '' 'frmActionPaths'
    Check 'EvaluationContract.RunSwitchDiscardsNonemptyAnalysisDraft' ($opened -ceq 'DELIVERED' -and $draftCount -ceq '1' -and
        ($staleCount -ceq 'MISSING' -or $staleCount -ceq '0') -and -not $summary.Contains('This evaluation'))
    $reopened=ExpectationControl 'btnExpectedConclusion' 'Click' '' 'frmActionPaths'
    Check 'EvaluationContract.NewRunNeverInheritsPreviousAnalysisSteps' ($reopened -ceq 'DELIVERED' -and
        (ExpectationControl 'lstExpectedSteps' 'Count') -ceq '0')
    [void](ExpectationControl 'btnCancelExpectation' 'Click')
    # Explicit cancel is a lifecycle outcome even when no expectation is chosen.
    [void](RecordingControl 'Start Recording' 'Click')
    $cancelled=SaveRecordedSetting '672'
    [void](RecordingControl 'Cancel Recording' 'Click')
    $cancelClose=@(RecordingJournal $cancelled.Attempt.SequenceId|Where-Object RecordType -CEQ 'Close')
    if($cancelClose.Count -ne 1 -or $cancelClose[0].Lifecycle -cne 'Cancelled'){throw 'Actual cancelled-run fixture failed.'}
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Refresh',''))
    [void](Select-EvaluationRun $cancelClose[0].ActionPathId)
    Assert-EvaluationStatus 'CancelledRunCannotBecomeSuccessful' 'Cancelled'
    # Ordinary Viewer user may evaluate a permitted other actor's recording.
    CloseRecordingViewer
    SelectTarget $Fixture 'config-reader'
    OpenRecordingViewer
    [void](Select-EvaluationRun $retryPath)
    $ready=Set-EvaluationDraft $command 0 'CommandCompleted'
    Assert-EvaluationStatus 'OrdinaryViewerCanEvaluatePermittedRunWithoutGuideMaintenance' 'Conclusion observed' $ready
    $loadedResult=ExpectationControl 'txtPathEvaluation' 'Value' '' 'frmActionPaths'
    $loadedStatus=ExpectationControl 'lblEvaluationStatus' 'Caption' '' 'frmActionPaths'
    $hadConclusion=$loadedResult -cne 'MISSING' -and $loadedResult.Length -gt 0 -and
        $loadedStatus.StartsWith('Conclusion observed',[StringComparison]::Ordinal)
    $beforeSignOut=@(EvaluationFiles).Count
    [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
    $invoked=ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths'
    # A hidden/disabled button is insufficient: inspect retained text even when
    # hidden, or require that the form was unloaded. No private text is reported.
    $text=ExpectationControl 'txtPathEvaluation' 'ValueAnyVisibility' '' 'frmActionPaths'
    Check 'EvaluationContract.SignOutClearsLoadedConclusion' ($ready -and $hadConclusion -and ($text -ceq '' -or $text -ceq 'MISSING') -and @(EvaluationFiles).Count -eq $beforeSignOut)
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    OpenRecordingViewer
}
