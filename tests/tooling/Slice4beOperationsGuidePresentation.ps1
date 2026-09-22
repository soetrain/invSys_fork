# D18 visible comparison of a real Receiving task before and after owner application.
# The existing fixture supplies original actions, publication and processor evidence.
function Test-OperationsGuidePresentation([string]$Stage,[ref]$GuideState) {
    $Guide=$GuideState.Value
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    function View([string]$Name,[string]$Action,[string]$Value='') {
        BoundControl $Name $Action $Value 'frmActionPathView'
    }
    function Deliver([string]$Form,[string]$Name,[string]$Action,[string]$Value='') {
        if((BoundControl $Name $Action $Value $Form) -cne 'DELIVERED'){throw 'Accepted Operations guide fixture handler is unavailable.'}
    }
    function Capture([string]$Name) {
        CaptureOwnedFormByCaptionEvidence 'Action Path view' ('operations-paired-'+$Stage.ToLowerInvariant()+'-'+$Name+'.png')
    }
    $prefix='OperationsGuidePresentation.'+$Stage+'.'
    $steps=@(@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'),@('RECEIVING_CONFIRM_WRITES','PENDING','True'),@('RECEIVING_CONFIRM_WRITES','PENDING','True'),@('ADMIN_SETTINGS_SAVE_VALUE','COMPLETED','True'))
    $source=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
    if($source.Count -ne 1 -or $source[0].ActionPathId -cne $pathId -or $source[0].Version -ne 20){throw 'Real Operations recording fixture is unavailable.'}
    $source=$source[0]
    $guideRoot=Join-Path $journalRoot 'Guides'
    if($null -eq $Guide){
        $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
        if($Stage -cne 'Pending' -or $allowed -isnot [bool] -or -not $allowed){throw 'Generated guide-author fixture capability or stage is invalid.'}
        $before=BoundPins $guideRoot
        $activity=ActivityPins;$config=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        $publication=(Get-FileHash -LiteralPath $eventsPath).Hash
        Deliver 'frmActionPaths' 'btnCreateGuide' 'Click'
        Deliver 'frmActionPathGuide' 'txtGuideName' 'Write' 'Receive items and verify applied source events'
        Deliver 'frmActionPathGuide' 'txtGuideInstructions' 'Write' 'Stage items and confirm writes. Refresh published evidence and verify every source event before concluding.'
        Deliver 'frmActionPathGuide' 'btnGuideExpectedConclusion' 'Click'
        foreach($step in $steps){
            Deliver 'frmActionPathExpectation' 'cboExpectedControl' 'Write' $step[0]
            Deliver 'frmActionPathExpectation' 'cboExpectedOutcome' 'Write' $step[1]
            Deliver 'frmActionPathExpectation' 'chkExpectedRetry' 'Check' $step[2]
            Deliver 'frmActionPathExpectation' 'btnAddExpectedStep' 'Click'
        }
        if((BoundControl 'cboTerminalStep' 'Select' '2' 'frmActionPathExpectation') -cne 'SELECTED'){throw 'Explicit guide terminal fixture cannot select its step.'}
        Deliver 'frmActionPathExpectation' 'cboTerminalKind' 'Write' 'SourceEventsApplied'
        Deliver 'frmActionPathExpectation' 'btnUseExpectation' 'Click'
        Deliver 'frmActionPathGuide' 'btnSaveGuide' 'Click'
        $fresh=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)})
        if($fresh.Count -ne 1){throw 'Actual Operations guide Save did not create one immutable fixture.'}
        $Guide=ReadGuideExpectationRecord $fresh[0].FullName
        if($null -eq $Guide){throw 'Actual Operations guide wire validation failed.'}
        Check 'OperationsGuidePresentation.AuthoredBeforeApplicationWithExplicitIntent' ($Guide.ExpectedConclusion.TerminalKind -ceq 'SourceEventsApplied' -and @($Guide.ExpectedConclusion.Steps).Count -eq 4 -and $Guide.ExpectedConclusion.TerminalStepId -ceq $Guide.ExpectedConclusion.Steps[2].StepId -and @($published.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -cin $submissionIds}).Count -eq 0)
        Check 'OperationsGuidePresentation.GuidePreservesActualSourceJournalAndObservations' ($Guide.SourceRun.ActionPathId -ceq $source.ActionPathId -and $Guide.SourceRun.ContentSha256 -ceq $source.ContentSha256 -and ($Guide.Observations|ConvertTo-Json -Depth 12 -Compress) -ceq ($source.Observations|ConvertTo-Json -Depth 12 -Compress))
        Check 'OperationsGuidePresentation.AuthoringPreservesActivityConfigAndPublication' ((PinsRetained $activity) -and (ActivityPins).Count -eq $activity.Count -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config -and (Get-FileHash -LiteralPath $eventsPath).Hash -ceq $publication)
        Deliver 'frmActionPathGuide' 'btnCancelGuide' 'Click'
    }
    if((BoundLibrary 'Select' $pathId) -cne 'SELECTED'){throw 'Actual Operations run cannot be explicitly selected.'}
    $beforePair=BoundPins $journalRoot
    BoundOpen;BoundSelect $Guide
    Deliver 'frmActionPathLibrary' 'btnUseGuideForRun' 'Click'
    Deliver 'frmActionPathLibrary' 'btnCloseGuides' 'Click'
    Check ($prefix+'ExplicitPairDoesNotEvaluateOrWriteTraining') (BoundSame $beforePair (BoundPins $journalRoot))
    $beforeResults=@(EvaluationFiles|ForEach-Object FullName)
    Deliver 'frmActionPaths' 'btnEvaluatePath' 'Click'
    $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $beforeResults})
    if($fresh.Count -ne 1){throw 'Separate actual Evaluate did not create one result.'}
    $result=Get-Content -LiteralPath $fresh[0].FullName -Raw|ConvertFrom-Json
    $expected=if($Stage -ceq 'Applied'){'Concluded'}else{'Awaiting'}
    Check ($prefix+'ExactGuideRunAndOwnerState') ($result.ExpectationSource -ceq 'Guide expectation' -and $result.Guide.ActionPathId -ceq $Guide.ActionPathId -and $result.Guide.ContentSha256 -ceq $Guide.ContentSha256 -and $result.Guide.Version -eq 1 -and $result.ActionPathId -ceq $source.ActionPathId -and $result.JournalSha256 -ceq $source.ContentSha256 -and $result.ResultState -ceq $expected)
    $exact=@($result.TerminalSources).Count -eq 4
    foreach($id in $submissionIds){
        $terminal=@($result.TerminalSources|Where-Object EventId -CEQ $id)
        $group=@($published.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq $id})
        $exact=$exact -and $terminal.Count -eq 1 -and $terminal[0].WarehouseId -ceq $Fixture.Warehouse -and $terminal[0].SubmissionState -ceq 'Submitted'
        if($group.Count -eq 1){
            $keys=@($group[0].Lines|ForEach-Object System_Key)
            $exact=$exact -and $terminal[0].OwnerStatus -ceq 'Applied' -and ($terminal[0].SystemKeys -join '|') -ceq ($keys -join '|')
        }else{
            $exact=$exact -and $terminal[0].OwnerStatus -ceq 'Awaiting' -and @($terminal[0].SystemKeys).Count -eq 0
        }
    }
    Check ($prefix+'EveryExactSourceReferenceAndAppliedKey') $exact
    $training=BoundPins $journalRoot;$activity=ActivityPins
    $config=(Get-FileHash -LiteralPath $Fixture.Config).Hash;$publication=(Get-FileHash -LiteralPath $eventsPath).Hash
    $publishCount=Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest'
    $authorityCount=Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest'
    try {
        Deliver 'frmActionPaths' 'btnViewActionPath' 'Click'
        $pair=View 'lblActionPathPair' 'Label'
        $how=View 'txtActionPathHowTo' 'Text';$diagnostic=View 'txtActionPathDiagnostic' 'Text'
        Check ($prefix+'VisiblePairNamesExactGuideAndRecordedTask') ($pair.Contains([string]$Guide.ActionPathId) -and $pair.Contains([string]$Guide.ContentSha256) -and $pair.Contains([string]$source.ActionPathId))
        Check ($prefix+'AuthoredInstructionsRemainSeparateFromObservations') ($how.Contains('Authored instruction') -and $how.Contains([string]$Guide.Instructions) -and (View 'txtActionPathHowTo' 'Locked') -ceq 'True' -and (View 'txtActionPathDiagnostic' 'Locked') -ceq 'True')
        $ordered=$true;$position=-1
        foreach($action in @($source.Observations|Where-Object OutcomeCode -CEQ 'REQUESTED')){
            $next=$diagnostic.IndexOf([string]$action.ActivityId,[StringComparison]::Ordinal)
            $ordered=$ordered -and $next -gt $position;$position=$next
        }
        Check ($prefix+'OriginalObservedActionsRemainInOrder') $ordered
        $status=if($Stage -ceq 'Applied'){'Conclusion observed'}else{'Awaiting published result'}
        Check ($prefix+'DiagnosticShowsExactSavedResultWithoutEarlyConclusion') ($diagnostic.Contains([string]$result.EvaluationId) -and $diagnostic.Contains($status) -and ($Stage -ceq 'Applied' -or -not $diagnostic.Contains('Conclusion observed')))
        foreach($method in @('How-To','Diagnostic','Compare both')){
            Deliver 'frmActionPathView' 'cboActionPathView' 'Write' $method
            $howVisible=((View 'txtActionPathHowTo' 'State') -split '\|')[0]
            $diagnosticVisible=((View 'txtActionPathDiagnostic' 'State') -split '\|')[0]
            Check ($prefix+'Method.'+$method) ((View 'cboActionPathView' 'Selected') -ceq $method -and $howVisible -ceq $(if($method -ceq 'Diagnostic'){'False'}else{'True'}) -and $diagnosticVisible -ceq $(if($method -ceq 'How-To'){'False'}else{'True'}) -and (View 'lblActionPathPair' 'Label') -ceq $pair -and (View 'txtActionPathHowTo' 'Text') -ceq $how -and (View 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
            Capture $method.Replace(' ','-').ToLowerInvariant()
        }
        if((View 'txtActionPathDiagnostic' 'ViewportBottom') -cne 'DELIVERED'){throw 'Actual paired diagnostic viewport is unavailable.'}
        Capture 'result'
        Check ($prefix+'AllTerminalSourcesVisibleInDiagnostic') (@($submissionIds|Where-Object {-not $diagnostic.Contains([string]$_)}).Count -eq 0)
        Deliver 'frmActionPathView' 'btnRefreshActionPathView' 'Click'
        Check ($prefix+'RefreshRetainsExactPairAndResult') ((View 'lblActionPathPair' 'Label') -ceq $pair -and (View 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
        Check ($prefix+'ViewsPreserveAllTrainingAndActivityBytes') ((BoundSame $training (BoundPins $journalRoot)) -and (PinsRetained $activity) -and (ActivityPins).Count -eq $activity.Count)
        Check ($prefix+'ViewsPreserveConfigPublicationAndAuthorityCounters') ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $config -and (Get-FileHash -LiteralPath $eventsPath).Hash -ceq $publication -and (Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publishCount -and (Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest') -eq $authorityCount)
    } finally {
        if((View '' 'Count') -ceq '1'){[void](View 'btnCloseActionPathView' 'Click')}
    }
    # Continue the original evaluator matrix with its independent analysis intent.
    if((BoundLibrary 'Select' $pathId) -cne 'SELECTED' -or -not (Set-EvaluationDraft $steps 2 'SourceEventsApplied')){throw 'Original Operations evaluation fixture cannot be restored.'}
    $GuideState.Value=$Guide
}
