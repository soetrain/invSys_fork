# D18 paired presentation; fixtures come from the existing real recording/guide handlers.
# No new form or presentation implementation is injected by this test.
function Test-GuidePresentation($Fixture,$Other,$FirstGuide,$SecondGuide,$GuideSource,$Observed,$FirstAction,$SecondAction) {
    function ViewControl([string]$Name,[string]$Action,[string]$Value='') {
        BoundControl $Name $Action $Value 'frmActionPathView'
    }
    function SettingsControl([string]$Name,[string]$Action,[string]$Value='') {
        BoundControl $Name $Action $Value 'frmEventTrackingSettings'
    }
    function OpenView {BoundControl 'btnViewActionPath' 'Click' '' 'frmActionPaths'}
    function OpenViewSettings {
        if((BoundControl 'btnSettings' 'Click' '' 'frmInventoryViewer') -cne 'DELIVERED'){throw 'Accepted Operations Settings fixture is unavailable.'}
    }
    function SaveViewPreference([string]$Choice) {
        if((SettingsControl 'cmbPreferredActionPathView' 'Write' $Choice) -cne 'DELIVERED' -or
           (SettingsControl 'btnSaveMyPreference' 'Click') -cne 'DELIVERED' -or
           (SettingsControl 'lblPreferenceStatus' 'Label') -cne 'Your Action Path preference was saved.'){
            throw 'Accepted personal preference save fixture failed; not presentation RED.'
        }
    }
    CloseRecordingViewer
    SelectTarget $Fixture 'config-reader'
    OpenRecordingViewer
    OpenViewSettings
    $priorChoice=SettingsControl 'cmbPreferredActionPathView' 'Selected'
    if($priorChoice -cnotin @('Use warehouse default','How-To','Diagnostic','Compare both')){throw 'Accepted preference read fixture failed.'}
    try {
        SaveViewPreference 'Compare both'
        [void](SettingsControl 'btnClose' 'Click')
        if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Accepted recording library fixture failed.'}
        Check 'GuidePresentation.NoPairDisablesEntry' ((BoundControl 'btnViewActionPath' 'State' '' 'frmActionPaths') -ceq 'True|False')
        if((BoundLibrary 'Select' $Observed.ActionPathId) -cne 'SELECTED'){throw 'Accepted selected-run fixture failed.'}
        BoundOpen;BoundSelect $FirstGuide
        if((BoundControl 'btnUseGuideForRun' 'Click') -cne 'DELIVERED'){throw 'Accepted guide/run application fixture failed.'}
        [void](BoundControl 'btnCloseGuides' 'Click')
        $saved=BoundEvaluate
        if($null -eq $saved -or $saved.ResultState -cne 'Concluded' -or $saved.Guide.ContentSha256 -cne $FirstGuide.ContentSha256){throw 'Accepted exact-pair evaluation fixture failed.'}
        $before=BoundPins $journalRoot;$activityBefore=ActivityPins
        $configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        $publishBefore=Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest'
        if($publishBefore -isnot [int]){throw 'Publication counter is unavailable.'}
        $libraryEvidence=BoundControl 'txtPathEvidence' 'Text' '' 'frmActionPaths'
        $libraryResult=BoundControl 'txtPathEvaluation' 'Text' '' 'frmActionPaths'
        Check 'GuidePresentation.PairedReaderEntryEnabled' ((BoundControl 'btnViewActionPath' 'State' '' 'frmActionPaths') -ceq 'True|True')
        $opened=(OpenView) -ceq 'DELIVERED' -and (ViewControl '' 'Count') -ceq '1'
        Check 'GuidePresentation.ActualHandlerOpensView' $opened
        Check 'GuidePresentation.ApprovedCaption' ($opened -and (ViewControl '' 'Caption') -ceq 'Action Path view')
        Check 'GuidePresentation.ThreeFixedChoices' ($opened -and (ViewControl 'cboActionPathView' 'Values') -ceq "How-To`nDiagnostic`nCompare both`n")
        Check 'GuidePresentation.FreshOpenUsesSavedPersonalChoice' ($opened -and (ViewControl 'cboActionPathView' 'Selected') -ceq 'Compare both')
        $pair=ViewControl 'lblActionPathPair' 'Label'
        Check 'GuidePresentation.ExactGuideAndDifferentObservedRunNamed' ($opened -and $pair.Contains([string]$FirstGuide.ActionPathId) -and $pair.Contains([string]$FirstGuide.ContentSha256) -and $pair.Contains([string]$Observed.ActionPathId) -and $pair -match '(?i)version\s*:?\s*1\b' -and $pair -match '(?i)version\s*:?\s*6\b')
        $howTo=ViewControl 'txtActionPathHowTo' 'Text';$diagnostic=ViewControl 'txtActionPathDiagnostic' 'Text'
        Check 'GuidePresentation.AuthoredGuideInstructionsLabelled' ($opened -and $howTo.Contains('Authored instruction') -and $howTo.Contains([string]$FirstGuide.Name))
        $firstIndex=$diagnostic.IndexOf([string]$FirstAction.Attempt.ActivityId,[StringComparison]::Ordinal)
        $secondIndex=$diagnostic.IndexOf([string]$SecondAction.Attempt.ActivityId,[StringComparison]::Ordinal)
        $sourceIds=@($GuideSource.Observations|ForEach-Object ActivityId|Select-Object -Unique)
        if($sourceIds.Count -eq 0){throw 'Accepted guide source fixture has no original actions.'}
        $hasSource=$false;foreach($id in $sourceIds){if($diagnostic.Contains([string]$id)){$hasSource=$true}}
        Check 'GuidePresentation.DiagnosticUsesOriginalObservedActionsInOrder' ($opened -and $firstIndex -ge 0 -and $secondIndex -gt $firstIndex -and -not $hasSource)
        Check 'GuidePresentation.ExactSavedResultWithComparisonStates' ($opened -and $diagnostic.Contains([string]$saved.EvaluationId) -and $diagnostic.Contains('Matched at action') -and $diagnostic.Contains('Additional observed actions: 0') -and $diagnostic.Contains('Domain application not asserted'))
        Check 'GuidePresentation.BothEvidencePanesLocked' ($opened -and (ViewControl 'txtActionPathHowTo' 'Locked') -ceq 'True' -and (ViewControl 'txtActionPathDiagnostic' 'Locked') -ceq 'True')
        foreach($method in @('How-To','Diagnostic','Compare both')){
            $selected=(ViewControl 'cboActionPathView' 'Write' $method) -ceq 'DELIVERED'
            $howState=ViewControl 'txtActionPathHowTo' 'State';$diagnosticState=ViewControl 'txtActionPathDiagnostic' 'State'
            Check ('GuidePresentation.Method.'+$method) ($opened -and $selected -and (ViewControl 'cboActionPathView' 'Selected') -ceq $method -and
                ($howState -split '\|')[0] -ceq $(if($method -ceq 'Diagnostic'){'False'}else{'True'}) -and
                ($diagnosticState -split '\|')[0] -ceq $(if($method -ceq 'How-To'){'False'}else{'True'}))
            Check ('GuidePresentation.SwitchPreservesPair.'+$method) ($opened -and (ViewControl 'lblActionPathPair' 'Label') -ceq $pair -and
                (ViewControl 'txtActionPathHowTo' 'Text') -ceq $howTo -and (ViewControl 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
        }
        foreach($layout in @('Minimum','Default','Larger','Restored')){
            Check ('GuidePresentation.Layout.'+$layout) ($opened -and (ViewControl '' 'Fit' $layout) -ceq 'True')
        }
        [void](ViewControl 'cboActionPathView' 'Write' 'Diagnostic')
        [void](ViewControl 'btnRefreshActionPathView' 'Click')
        Check 'GuidePresentation.RefreshRetainsMethodPairAndResult' ($opened -and (ViewControl 'cboActionPathView' 'Selected') -ceq 'Diagnostic' -and (ViewControl 'lblActionPathPair' 'Label') -ceq $pair -and (ViewControl 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
        [void](OpenView)
        Check 'GuidePresentation.RepeatedEntryReusesOneView' ($opened -and (ViewControl '' 'Count') -ceq '1')
        Check 'GuidePresentation.OriginalLibraryEvidenceRemainsAvailable' ($opened -and (BoundControl 'txtPathEvidence' 'Text' '' 'frmActionPaths') -ceq $libraryEvidence -and (BoundControl 'txtPathEvaluation' 'Text' '' 'frmActionPaths') -ceq $libraryResult)
        [void](ViewControl 'btnCloseActionPathView' 'Click')
        Check 'GuidePresentation.CloseReleasesInstance' ($opened -and (ViewControl '' 'Count') -ceq '0')
        [void](OpenView)
        Check 'GuidePresentation.ReopenUsesSavedChoiceNotUnsavedSwitch' ($opened -and (ViewControl 'cboActionPathView' 'Selected') -ceq 'Compare both')
        Check 'GuidePresentation.ViewsDoNotAppendOrRewriteTraining' (BoundSame $before (BoundPins $journalRoot))
        Check 'GuidePresentation.ViewsDoNotWriteActivityConfigOrPublish' ((PinsRetained $activityBefore) -and (ActivityPins).Count -eq $activityBefore.Count -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore -and (Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publishBefore)
        $guidePath=Join-Path $guideRoot ($FirstGuide.ActionPathId+'.1.json');$bytes=[IO.File]::ReadAllBytes($guidePath)
        try {
            [IO.File]::WriteAllText($guidePath,'{}',[Text.UTF8Encoding]::new($false))
            [void](ViewControl 'btnRefreshActionPathView' 'Click')
            Check 'GuidePresentation.RefreshClearsCorruptGuideEvidence' ($opened -and (ViewControl 'txtActionPathHowTo' 'Text') -ceq '' -and (ViewControl 'txtActionPathDiagnostic' 'Text') -ceq '' -and (ViewControl 'lblActionPathViewStatus' 'Label') -match '(?i)unavailable|incomplete')
        } finally {[IO.File]::WriteAllBytes($guidePath,$bytes)}
        [void](ViewControl 'btnCloseActionPathView' 'Click');[void](OpenView)
        [void](BoundLibrary 'Select' $GuideSource.ActionPathId)
        [void](ViewControl 'cboActionPathView' 'Write' 'Diagnostic')
        Check 'GuidePresentation.RunChangeCannotRetargetPendingView' ($opened -and (ViewControl 'txtActionPathDiagnostic' 'Text') -cin @('','MISSING') -and (ViewControl 'lblActionPathPair' 'Label') -cin @('','MISSING'))
        Check 'GuidePresentation.RunChangeDisablesUnpairedEntry' ($opened -and (BoundControl 'btnViewActionPath' 'State' '' 'frmActionPaths') -ceq 'True|False')
        if((BoundLibrary 'Select' $Observed.ActionPathId) -cne 'SELECTED'){throw 'Accepted run reselection failed.'}
        BoundOpen;BoundSelect $SecondGuide
        if((BoundControl 'btnUseGuideForRun' 'Click') -cne 'DELIVERED'){throw 'Accepted second-guide application failed.'}
        [void](BoundControl 'btnCloseGuides' 'Click');[void](OpenView)
        Check 'GuidePresentation.ExplicitEntryBindsSecondGuideVersion' ($opened -and (ViewControl 'lblActionPathPair' 'Label').Contains([string]$SecondGuide.ContentSha256))
        $unevaluated=ViewControl 'txtActionPathDiagnostic' 'Text'
        Check 'GuidePresentation.NewGuideNeverBorrowsOldSavedConclusion' ($opened -and -not $unevaluated.Contains([string]$saved.EvaluationId) -and -not $unevaluated.Contains('Conclusion observed') -and $unevaluated -match '(?i)not evaluated|no .*evaluation|choose Evaluate')
        SelectTarget $Other 'config-reader'
        [void](ViewControl 'btnRefreshActionPathView' 'Click')
        Check 'GuidePresentation.ContextLossClearsRetainedContent' ($opened -and (ViewControl 'txtActionPathHowTo' 'Text') -cin @('','MISSING') -and (ViewControl 'txtActionPathDiagnostic' 'Text') -cin @('','MISSING'))
        Check 'GuidePresentation.InvalidationsPreserveEveryTrainingByte' (BoundSame $before (BoundPins $journalRoot))
    } finally {
        CloseRecordingViewer
        SelectTarget $Fixture 'config-reader'
        OpenRecordingViewer;OpenViewSettings
        SaveViewPreference $priorChoice
        CloseRecordingViewer
    }
}
