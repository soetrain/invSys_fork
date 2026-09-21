# D18: an explicitly selected observed run never comes from a guide's source run.
# Call after the existing guide-expectation fixture saved both immutable versions.
function Test-GuideEvaluation($Fixture,$Other,$FirstGuide,$SecondGuide,$GuideSource) {
    function BoundControl([string]$Name,[string]$Action,[string]$Value='', [string]$Form='frmActionPathLibrary') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($Form,$Name,$Action,$Value))
    }
    function BoundLibrary([string]$Action,[string]$Value='') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @($Action,$Value))
    }
    function BoundOpen {[void](BoundControl 'btnPublishedGuides' 'Click' '' 'frmActionPaths')}
    function BoundSelect($Guide) {
        $key=[string]$Guide.ActionPathId+'|'+[string]$Guide.Version+'|'+[string]$Guide.ContentSha256
        $keys=@((BoundControl 'lstPublishedGuides' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        $index=[Array]::IndexOf($keys,$key)
        if($index -lt 0){throw 'Accepted exact-version reader fixture is unavailable; not guide-binding RED.'}
        if((BoundControl 'lstPublishedGuides' 'Select' ([string]$index)) -cne 'SELECTED'){throw 'Accepted guide selection fixture failed.'}
    }
    function BoundSummary {BoundControl 'lblExpectationSummary' 'Label' '' 'frmActionPaths'}
    function BoundPins([string]$Root) {
        $pins=@{}
        if(Test-Path -LiteralPath $Root){foreach($file in Get-ChildItem -LiteralPath $Root -Recurse -File){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}}
        return $pins
    }
    function BoundSame($Before,$After) {
        if($Before.Count -ne $After.Count){return $false}
        foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $Before[$path] -cne $After[$path]){return $false}}
        return $true
    }
    function BoundResults {
        if(Test-Path -LiteralPath $evaluationRoot){Get-ChildItem -LiteralPath $evaluationRoot -File -Filter '*.json'}
    }
    function BoundEvaluate {
        $before=@(BoundResults|ForEach-Object FullName)
        if((BoundControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths') -cne 'DELIVERED'){throw 'Accepted Evaluate action is unavailable; not guide-binding RED.'}
        $fresh=@(BoundResults|Where-Object {$_.FullName -cnotin $before})
        if($fresh.Count -ne 1){return $null}
        Get-Content -LiteralPath $fresh[0].FullName -Raw|ConvertFrom-Json
    }
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    SetRecordingPolicy $true
    OpenRecordingViewer
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Accepted observed-run fixture cannot start.'}
    $observedFirst=SaveRecordedSetting '672';$observedSecond=SaveRecordedSetting '673'
    if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED'){throw 'Accepted observed-run fixture cannot stop.'}
    $observedJournal=@(RecordingJournal $observedFirst.Attempt.SequenceId|Sort-Object Version)
    if($observedJournal.Count -ne 6 -or -not (JournalChain $observedFirst.Attempt.SequenceId 6)){throw 'Observed-run fixture is not a complete real journal.'}
    $observed=$observedJournal[-1]
    if($observed.ActionPathId -ceq $GuideSource.ActionPathId -or $observed.ExpectedConclusion.TerminalKind -cne 'None'){throw 'Observed run must be separate and have no inferred expectation.'}
    $guideRoot=Join-Path $journalRoot 'Guides';$evaluationRoot=Join-Path $journalRoot 'Evaluations'
    $emptyRuns=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Where-Object {$_.RecordType -ceq 'Close' -and $_.Lifecycle -ceq 'Stopped' -and @($_.Observations).Count -eq 0})
    if($emptyRuns.Count -eq 0){throw 'Accepted guide-expectation fixture supplied no stopped empty recording.'}
    $guidePins=BoundPins $guideRoot;$trainingBefore=BoundPins $journalRoot;$activityBefore=ActivityPins
    $configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $publicationBefore=Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest'
    if($publicationBefore -isnot [int]){throw 'Publication counter is not typed.'}
    try {
        CloseRecordingViewer
        SelectTarget $Fixture 'config-reader'
        $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-reader',$Fixture.Warehouse,'S1')
        if($allowed -isnot [bool] -or $allowed){throw 'Ordinary reader fixture is not isolated from maintenance.'}
        OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Accepted recording library fixture failed.'}
        BoundOpen
        Check 'GuideEvaluation.NoSelectedRunDisablesApply' ((BoundControl 'btnUseGuideForRun' 'State') -ceq 'True|False')
        [void](BoundControl 'btnCloseGuides' 'Click')
        if((BoundLibrary 'Select' $observed.ActionPathId) -cne 'SELECTED'){throw 'Accepted observed-run selection fixture failed.'}
        BoundOpen;BoundSelect $FirstGuide
        $present=(BoundControl 'btnUseGuideForRun' 'State') -ceq 'True|True'
        Check 'GuideEvaluation.OrdinaryReaderCanApplyToExplicitRun' $present
        $runLabel=BoundControl 'lblGuideObservedRun' 'Label'
        Check 'GuideEvaluation.ReaderNamesCapturedObservedRun' ($present -and $runLabel.Contains([string]$observed.ActionPathId) -and $runLabel -match '(?i)version\s*:?\s*6\b')
        $used=(BoundControl 'btnUseGuideForRun' 'Click') -ceq 'DELIVERED'
        $firstSummary=BoundSummary
        Check 'GuideEvaluation.ActualHandlerStagesExactFirstGuide' ($used -and $firstSummary.Contains('Guide expectation') -and $firstSummary.Contains([string]$FirstGuide.ActionPathId) -and $firstSummary -match '(?i)version\s*:?\s*1\b')
        Check 'GuideEvaluation.ApplyingDoesNotWriteTrainingOrActivity' ($used -and (BoundSame $trainingBefore (BoundPins $journalRoot)) -and (PinsRetained $activityBefore) -and (ActivityPins).Count -eq $activityBefore.Count)
        BoundSelect $SecondGuide
        Check 'GuideEvaluation.BrowsingDoesNotReplaceExplicitIntent' ($used -and (BoundSummary) -ceq $firstSummary)
        [void](BoundControl 'btnRefreshGuides' 'Click')
        Check 'GuideEvaluation.RefreshDoesNotReplaceExplicitIntent' ($used -and (BoundSummary) -ceq $firstSummary)
        [void](BoundControl 'btnCloseGuides' 'Click')
        Check 'GuideEvaluation.ClosingReaderRetainsStagedIntent' ($used -and (BoundSummary) -ceq $firstSummary)
        $firstResult=BoundEvaluate
        $exact=$null -ne $firstResult -and $firstResult.ExpectationSource -ceq 'Guide expectation'
        Check 'GuideEvaluation.SeparateEvaluateUsesExactGuideVersionHash' ($used -and $exact -and $firstResult.Guide.ActionPathId -ceq $FirstGuide.ActionPathId -and $firstResult.Guide.Version -eq 1 -and $firstResult.Guide.ContentSha256 -ceq $FirstGuide.ContentSha256)
        Check 'GuideEvaluation.ResultUsesObservedJournalNotGuideSource' ($used -and $exact -and $firstResult.ActionPathId -ceq $observed.ActionPathId -and $firstResult.JournalRecordId -ceq $observed.RecordId -and $firstResult.JournalSha256 -ceq $observed.ContentSha256)
        Check 'GuideEvaluation.MatchesOriginalObservedActionsAndGuideSteps' ($used -and $exact -and ($firstResult.Matches.ActivityId -join '|') -ceq (@($observedFirst.Attempt.ActivityId,$observedSecond.Attempt.ActivityId) -join '|') -and ($firstResult.Matches.StepId -join '|') -ceq ($FirstGuide.ExpectedConclusion.Steps.StepId -join '|'))
        Check 'GuideEvaluation.GuideIntentCanConcludeOnlySelectedRun' ($used -and $exact -and $firstResult.ResultState -ceq 'Concluded' -and ($firstResult.ExpectedConclusion|ConvertTo-Json -Depth 12 -Compress) -ceq ($FirstGuide.ExpectedConclusion|ConvertTo-Json -Depth 12 -Compress))
        $guidePath=Join-Path $guideRoot ($FirstGuide.ActionPathId+'.1.json')
        $guideBytes=[IO.File]::ReadAllBytes($guidePath)
        $savedResultPins=BoundPins $evaluationRoot
        try {
            [IO.File]::WriteAllText($guidePath,'{}',[Text.UTF8Encoding]::new($false))
            [void](BoundLibrary 'Refresh')
            Check 'GuideEvaluation.SavedReadRevalidatesGuideIntegrity' ($used -and (BoundControl 'txtPathEvaluation' 'Text' '' 'frmActionPaths') -ceq '' -and (BoundSame $savedResultPins (BoundPins $evaluationRoot)))
        } finally {[IO.File]::WriteAllBytes($guidePath,$guideBytes)}
        [void](BoundLibrary 'Refresh')
        BoundOpen;BoundSelect $SecondGuide
        $secondUsed=(BoundControl 'btnUseGuideForRun' 'Click') -ceq 'DELIVERED'
        Check 'GuideEvaluation.ApplyingAnotherVersionClearsDisplayedOldResult' ($secondUsed -and (BoundControl 'txtPathEvaluation' 'Text' '' 'frmActionPaths') -ceq '')
        $secondResult=BoundEvaluate
        $secondExact=$null -ne $secondResult -and $secondResult.ExpectationSource -ceq 'Guide expectation'
        Check 'GuideEvaluation.AnotherVersionAppendsItsExactDefinition' ($secondUsed -and $secondExact -and $secondResult.Guide.Version -eq 2 -and $secondResult.Guide.ContentSha256 -ceq $SecondGuide.ContentSha256 -and @($secondResult.Matches).Count -eq 1 -and $secondResult.ExpectedConclusion.TerminalStepId -ceq $SecondGuide.ExpectedConclusion.TerminalStepId)
        [void](BoundControl 'btnExpectedConclusion' 'Click' '' 'frmActionPaths')
        $editorIds=BoundControl 'lstExpectedSteps' 'Values' '' 'frmActionPathExpectation'
        Check 'GuideEvaluation.AnalysisEditorStartsFromStagedGuideIntent' ($secondUsed -and $editorIds -ceq ([string]$SecondGuide.ExpectedConclusion.Steps[0].StepId+"`n"))
        [void](BoundControl 'btnUseExpectation' 'Click' '' 'frmActionPathExpectation')
        $analysisResult=BoundEvaluate
        Check 'GuideEvaluation.ExplicitAnalysisBecomesIndependentIntent' ($secondUsed -and $null -ne $analysisResult -and $analysisResult.ExpectationSource -ceq 'This evaluation' -and @($analysisResult.Guide.PSObject.Properties).Count -eq 0)
        $noneGuides=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json'|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Where-Object {$_.ExpectedConclusion.TerminalKind -ceq 'None'})
        if($noneGuides.Count -eq 0){throw 'Earlier actual Save fixture supplied no None guide; not guide-binding RED.'}
        BoundOpen;BoundSelect $noneGuides[0]
        $noneUsed=(BoundControl 'btnUseGuideForRun' 'Click') -ceq 'DELIVERED'
        $noneResult=BoundEvaluate
        Check 'GuideEvaluation.NoneGuideDoesNotInferSuccessfulSourceConclusion' ($noneUsed -and $null -ne $noneResult -and $noneResult.ExpectationSource -ceq 'Guide expectation' -and $noneResult.ExpectedConclusion.TerminalKind -ceq 'None' -and @($noneResult.ExpectedConclusion.Steps).Count -eq 0 -and $noneResult.ResultState -ceq 'Incomplete')
        if((BoundLibrary 'Select' $emptyRuns[0].ActionPathId) -cne 'SELECTED'){throw 'Accepted empty observed-run selection failed.'}
        BoundOpen;BoundSelect $FirstGuide
        $emptyUsed=(BoundControl 'btnUseGuideForRun' 'Click') -ceq 'DELIVERED'
        $emptyResult=BoundEvaluate
        Check 'GuideEvaluation.EmptyObservedRunCannotBorrowSuccessfulGuideSource' ($emptyUsed -and $null -ne $emptyResult -and $emptyResult.ActionPathId -ceq $emptyRuns[0].ActionPathId -and $emptyResult.ResultState -ceq 'Failed' -and @($emptyResult.Matches).Count -eq 0 -and @($emptyResult.MissingSteps).Count -eq 2)
        if((BoundLibrary 'Select' $observed.ActionPathId) -cne 'SELECTED'){throw 'Accepted observed-run reselection failed.'}
        BoundOpen;BoundSelect $FirstGuide
        $beforeCorrupt=BoundSummary
        $guidePath=Join-Path $guideRoot ($FirstGuide.ActionPathId+'.1.json')
        $guideBytes=[IO.File]::ReadAllBytes($guidePath)
        try {
            [IO.File]::WriteAllText($guidePath,'{}',[Text.UTF8Encoding]::new($false))
            [void](BoundControl 'btnUseGuideForRun' 'Click')
            Check 'GuideEvaluation.CorruptSelectedGuideCannotReplaceIntent' ($present -and (BoundSummary) -ceq $beforeCorrupt -and (BoundControl 'lblPublishedGuideStatus' 'Label') -match '(?i)incomplete|unavailable|corrupt')
        } finally {[IO.File]::WriteAllBytes($guidePath,$guideBytes)}
        if([Convert]::ToBase64String([IO.File]::ReadAllBytes($guidePath)) -cne [Convert]::ToBase64String($guideBytes)){throw 'Disposable guide corruption was not restored.'}
        BoundOpen;BoundSelect $FirstGuide
        $beforeInvalidUse=(BoundControl 'btnUseGuideForRun' 'Click') -ceq 'DELIVERED'
        $beforeInvalidEvaluation=BoundPins $evaluationRoot
        try {
            [IO.File]::WriteAllText($guidePath,'{}',[Text.UTF8Encoding]::new($false))
            [void](BoundControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
            Check 'GuideEvaluation.IntegrityLossBeforeEvaluatePreventsSilentFallback' ($beforeInvalidUse -and (BoundSame $beforeInvalidEvaluation (BoundPins $evaluationRoot)) -and (BoundControl 'lblEvaluationStatus' 'Label' '' 'frmActionPaths') -match '(?i)incomplete|unavailable')
        } finally {[IO.File]::WriteAllBytes($guidePath,$guideBytes)}
        if([Convert]::ToBase64String([IO.File]::ReadAllBytes($guidePath)) -cne [Convert]::ToBase64String($guideBytes)){throw 'Evaluated guide fixture bytes were not restored.'}
        BoundOpen
        BoundSelect $FirstGuide
        [void](BoundLibrary 'Select' $GuideSource.ActionPathId)
        $beforeStale=BoundSummary;$resultsBeforeStale=BoundPins $evaluationRoot
        [void](BoundControl 'btnUseGuideForRun' 'Click')
        Check 'GuideEvaluation.RunChangeDoesNotRetargetPendingReader' ($present -and (BoundSummary) -ceq $beforeStale -and (BoundControl 'lblPublishedGuideStatus' 'Label') -match '(?i)changed|reopen|unavailable')
        Check 'GuideEvaluation.StaleUseDoesNotAppendEvaluation' (BoundSame $resultsBeforeStale (BoundPins $evaluationRoot))
        BoundOpen;BoundSelect $FirstGuide
        $rebound=(BoundControl 'btnUseGuideForRun' 'Click') -ceq 'DELIVERED'
        Check 'GuideEvaluation.ExplicitReopenBindsIntendedRun' ($rebound -and (BoundControl 'lblGuideObservedRun' 'Label').Contains([string]$GuideSource.ActionPathId) -and (BoundSummary).Contains('Guide expectation'))
        SelectTarget $Other 'config-reader'
        [void](BoundControl 'btnUseGuideForRun' 'Click')
        Check 'GuideEvaluation.ContextChangeClearsReaderBinding' ($present -and (BoundControl 'lblGuideObservedRun' 'Label') -cin @('','MISSING') -and (BoundControl 'btnUseGuideForRun' 'State') -cne 'True|True')
        Check 'GuideEvaluation.GuidesAndActivityRemainUnchanged' ((BoundSame $guidePins (BoundPins $guideRoot)) -and (PinsRetained $activityBefore) -and (ActivityPins).Count -eq $activityBefore.Count)
        $journalUnchanged=$true
        foreach($path in $trainingBefore.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $trainingBefore[$path]){$journalUnchanged=$false}}
        Check 'GuideEvaluation.PriorTrainingFilesRemainByteIdentical' $journalUnchanged
        Check 'GuideEvaluation.NoConfigurationWriteOrPublication' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore -and (Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publicationBefore)
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
        OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $observed.ActionPathId) -cne 'SELECTED'){throw 'Accepted policy fixture selection failed.'}
        BoundOpen;BoundSelect $FirstGuide
        $policyUsed=(BoundControl 'btnUseGuideForRun' 'Click') -ceq 'DELIVERED'
        $beforePolicyEvaluation=BoundPins $evaluationRoot
        try {
            SaveGuideExpectationVisibility $false
            [void](BoundControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
            Check 'GuideEvaluation.CurrentPolicyLossPreventsGuideFallback' ($policyUsed -and (BoundSame $beforePolicyEvaluation (BoundPins $evaluationRoot)) -and (BoundControl 'txtPathEvaluation' 'Text' '' 'frmActionPaths') -ceq '')
            [void](BoundControl 'btnUseGuideForRun' 'Click')
            Check 'GuideEvaluation.RestrictedGuideExpectationCannotBeApplied' ($present -and (BoundControl 'lblPublishedGuideStatus' 'Label') -match '(?i)hidden|restricted|incomplete|unavailable')
        } finally {SaveGuideExpectationVisibility $true}
        Check 'GuideEvaluation.PolicyExercisePreservesGuidesAndSavedResults' ((BoundSame $guidePins (BoundPins $guideRoot)) -and (BoundSame $beforePolicyEvaluation (BoundPins $evaluationRoot)))
    } finally {
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
    }
}
