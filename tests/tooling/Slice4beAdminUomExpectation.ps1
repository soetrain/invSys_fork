# Actual native Reset cancellation, actual expectation editor, actual Stop handler.
function Test-AdminUomExpectedCancellation {
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'UOM expectation recording fixture unavailable.'}
    $first=SaveRecordedSetting '683'
    $uomSequence=[string]$first.Attempt.SequenceId
    $before=ActivityPins
    $configPin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    try {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ShowSettings')
        [void](Invoke-AdminUomResetChoice 'No' 'expectation-uom-reset-no.png')
    } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    $fresh=@(foreach($file in Get-ChildItem -LiteralPath $activityRoot -Filter '*.json' -File){
        if(-not $before.ContainsKey($file.Name)){Get-Content -LiteralPath $file.FullName -Raw|ConvertFrom-Json}
    })
    $cancel=@($fresh|Where-Object OutcomeCode -CEQ 'CANCELLED')
    $observed=$fresh.Count -eq 2 -and $cancel.Count -eq 1
    if($observed){$observed=$cancel[0].ControlId -ceq 'ADMIN_UOM_RESET' -and $cancel[0].SequenceId -ceq $uomSequence -and $cancel[0].DataEffect -ceq 'Unchanged' -and @($cancel[0].SourceEventRefs).Count -eq 0}
    Check 'AdminUom.Expectation.ActualCancellationObserved' $observed
    Check 'AdminUom.Expectation.CancelPreservesConfigBytes' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
    if(-not $observed){throw 'Native cancellation observation fixture is incomplete; not editor-choice RED.'}
    [void](SaveRecordedSetting '684')
    $editorPins=ActivityPins
    $configPin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    [void](ExpectationControl 'btnRecordingExpectation' 'Click' '' 'frmInventoryViewer')
    Check 'AdminUom.Expectation.RegisteredResetIsSelectable' ((ExpectationControl 'cboExpectedControl' 'Select' 'ADMIN_UOM_RESET') -ceq 'DELIVERED')
    $selectable=(ExpectationControl 'cboExpectedOutcome' 'Select' 'CANCELLED') -ceq 'DELIVERED'
    Check 'AdminUom.Expectation.CancelledOutcomeIsSelectable' $selectable
    if($selectable){
        [void](ExpectationControl 'btnAddExpectedStep' 'Click')
        [void](ExpectationControl 'cboExpectedControl' 'Select' 'ADMIN_SETTINGS_SAVE_VALUE')
        [void](ExpectationControl 'cboExpectedOutcome' 'Select' 'COMPLETED')
        [void](ExpectationControl 'btnAddExpectedStep' 'Click')
        [void](ExpectationControl 'cboTerminalStep' 'Index' '1')
        [void](ExpectationControl 'cboTerminalKind' 'Select' 'CommandCompleted')
    }
    [void](ExpectationControl 'cboExpectedControl' 'Select' 'ADMIN_UOM_ADD')
    Check 'AdminUom.Expectation.AddDoesNotOfferCancellation' ((ExpectationControl 'cboExpectedOutcome' 'Select' 'CANCELLED') -ceq 'CHOICE_UNAVAILABLE')
    [void](ExpectationControl 'cboExpectedControl' 'Select' 'ADMIN_UOM_RESET')
    if($selectable){[void](ExpectationControl 'cboExpectedOutcome' 'Select' 'CANCELLED')}
    Capture-ExpectationEditorEvidence 'UomCancelled'
    if($selectable){[void](ExpectationControl 'btnUseExpectation' 'Click')}
    else{[void](ExpectationControl 'btnCancelExpectation' 'Click')}
    $after=ActivityPins
    $same=$after.Count -eq $editorPins.Count
    foreach($key in $editorPins.Keys){$same=$same -and $after.ContainsKey($key) -and $after[$key] -ceq $editorPins[$key]}
    Check 'AdminUom.Expectation.AuthoringDoesNotPerformWork' ($same -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
    [void](RecordingControl 'Stop Recording' 'Click')
    $close=@(RecordingJournal $uomSequence|Where-Object RecordType -CEQ 'Close')
    if($close.Count -ne 1){throw 'UOM expectation actual Stop fixture did not close.'}
    $definition=$close[0].ExpectedConclusion
    $steps=@($definition.Steps)
    $retained=$steps.Count -eq 2
    if($retained){$retained=$steps[0].ControlId -ceq 'ADMIN_UOM_RESET' -and $steps[0].RequiredOutcome -ceq 'CANCELLED' -and $steps[1].ControlId -ceq 'ADMIN_SETTINGS_SAVE_VALUE' -and $definition.TerminalStepId -ceq $steps[1].StepId -and $definition.TerminalKind -ceq 'CommandCompleted'}
    Check 'AdminUom.Expectation.AuthoredCancelledStepSurvivesActualStop' $retained
    $now=ActivityPins
    Check 'AdminUom.Expectation.OriginalCancelledObservationUnchanged' ($now.ContainsKey($cancel[0].RecordId+'.json') -and $now[$cancel[0].RecordId+'.json'] -ceq $editorPins[$cancel[0].RecordId+'.json'])
}
