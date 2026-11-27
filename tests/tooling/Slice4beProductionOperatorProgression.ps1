# Record two deliberate owner attempts so both guide steps name the same control.
# Stable StepIds and fresh claims must distinguish them in Run all and Verify.
function Test-OperatorStepProgression {
    PrepareBook
    OpenRecordingViewer
    $journals=BoundPins $destRoot
    Delivered (RecordingControl 'Start Recording' 'Click')
    [void](Probe 'SetEnteredActual' @(''))
    OrdinaryClick
    [void](Probe 'SetEnteredActual' @('1'))
    OrdinaryClick
    Delivered (RecordingControl 'Stop Recording' 'Click')
    $closed=@(Get-ChildItem $destRoot -File -Filter '*.json'|Where-Object {!$journals.ContainsKey($_.FullName)}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Close')
    if($closed.Count -ne 1 -or $closed[0].ActionCount -ne 2){throw 'Exactly two ordinary recorded completions are required for progression.'}
    CloseRecordingViewer
    SelectTarget $dest
    if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Two-step source publication unavailable.'}
    OpenRecordingViewer
    Delivered (BoundLibrary 'Open')
    if((BoundLibrary 'Select' $closed[0].ActionPathId) -cne 'SELECTED'){throw 'Two-step source recording unavailable.'}
    Delivered (BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths')
    if((Author 'lstGuideSteps' 'Rows') -cne '2'){throw 'Two independent ordinary action rows required.'}
    Delivered (Author 'txtGuideName' 'Write' 'Complete two prepared Processes')
    Delivered (Author 'txtGuideSubject' 'Write' 'Production completion sequencing')
    Delivered (Author 'txtGuideInstructions' 'Write' 'Bind and complete each prepared Process separately. Refresh Events and check its owner evidence before continuing.')
    Delivered (Author 'btnGuideExpectedConclusion' 'Click')
    foreach($i in 1..2){
        Delivered (BoundControl 'cboExpectedControl' 'Write' $control 'frmActionPathExpectation')
        Delivered (BoundControl 'cboExpectedOutcome' 'Write' 'CONFIRMED' 'frmActionPathExpectation')
        Delivered (BoundControl 'btnAddExpectedStep' 'Click' '' 'frmActionPathExpectation')
    }
    if((BoundControl 'cboTerminalStep' 'Select' '1' 'frmActionPathExpectation') -cne 'SELECTED'){throw 'Second expected terminal unavailable.'}
    Delivered (BoundControl 'cboTerminalKind' 'Write' 'SourceEventsApplied' 'frmActionPathExpectation')
    Delivered (BoundControl 'btnUseExpectation' 'Click' '' 'frmActionPathExpectation')
    $guideFolder=Join-Path $destRoot 'Guides'
    $prior=BoundPins $guideFolder
    Delivered (Author 'btnSaveGuide' 'Click')
    $files=@(Get-ChildItem $guideFolder -File -Filter '*.json'|Where-Object {!$prior.ContainsKey($_.FullName)})
    if($files.Count -ne 1){throw 'Fresh two-step guide required.'}
    $two=ReadGuideExpectationRecord $files[0].FullName
    if($two.Steps.Count -ne 2 -or $two.Steps[0].StepId -ceq $two.Steps[1].StepId -or @($two.Steps|Where-Object ControlId -CNE $control).Count){throw 'Two distinct guide steps for Complete Run required.'}
    Delivered (Author 'btnCancelGuide' 'Click')
    BoundOpen
    BoundSelect $two
    Delivered (BoundControl 'btnConfigureExecution' 'Click')
    Delivered (BoundControl 'btnSaveExecutionProfile' 'Click' '' 'frmActionPathExecution')
    CloseRecordingViewer
    SelectTarget $dest 'config-producer'
    PrepareBook
    OpenRecordingViewer
    Delivered (BoundLibrary 'Open')
    BoundOpen
    BoundSelect $two
    Delivered (BoundControl 'btnRunHowTo' 'Click')
    StartCase
    ClickRun 'btnBindProduction'
    OrdinaryClick
    $first=LastRun
    ClickRun 'btnCheckProduction'
    Pass 'LaterStepWaitsOnFirstUnpublishedCompletion' ($first.OperatorClaims.Count -eq 1 -and $first.OperatorClaims[0].OwnerValidated -and $first.Steps.Count -eq 0 -and (LastRun).ContentSha256 -ceq $first.ContentSha256 -and (BoundControl 'btnNextRunStep' 'Click' '' $runner) -ceq 'DISABLED')
    Delivered (BoundControl 'btnRefresh' 'Click' '' 'frmInventoryViewer')
    ClickRun 'btnCheckProduction'
    $proved=LastRun
    if($proved.State -cne 'Running' -or $proved.Steps.Count -ne 1 -or $proved.Steps[0].StepId -cne $two.Steps[0].StepId -or $proved.Steps[0].State -cne 'OperatorCompleted'){throw 'First attributable completion must release Next Step.'}
    $calls=Run 'invSys.Operations.xlam' 'TestCompleteReplay.Counts'
    ClickRun 'btnNextRunStep'
    $second=LastRun
    ClickRun 'btnCheckProduction'
    Pass 'NextIdenticalControlRequiresFreshStepAndOwner' ($second.Steps.Count -eq 1 -and $second.OperatorClaims.Count -eq 2 -and $second.OperatorClaims[1].StepId -ceq $two.Steps[1].StepId -and $second.OperatorClaims[1].BindingId -cne $first.OperatorClaims[0].BindingId -and !$second.OperatorClaims[1].Kind -and !$second.OperatorClaims[1].ActivityId -and (LastRun).ContentSha256 -ceq $second.ContentSha256 -and (Run 'invSys.Operations.xlam' 'TestCompleteReplay.Counts') -ceq $calls)
    ClickRun 'btnStopRun'
    $result=VerifySaved
    Pass 'VerifyKeepsFirstCreditAndMissingSecondStep' ($result.ResultState -ceq 'Failed' -and $result.Matches.Count -eq 1 -and $result.Matches[0].ActivityId -ceq $first.OperatorClaims[0].ActivityId -and $result.MissingSteps.Count -eq 1 -and $result.MissingSteps[0] -ceq $two.ExpectedConclusion.Steps[1].StepId -and 'REQUIRED_STEP_MISSING' -cin $result.ReasonCodes -and (Run 'invSys.Operations.xlam' 'TestCompleteReplay.Counts') -ceq $calls)
    ClickRun 'btnCloseRun'
}
