# Additional visible regressions use actual recorded actions and explicit Evaluate.
function Test-GuidePresentationScenarios($Fixture,$FirstGuide,$SecondGuide,$Observed) {
    [pscustomobject]@{File='Slice4beGuidePresentationScenarios.ps1';Sha256=(Get-FileHash -LiteralPath (Join-Path $PSScriptRoot 'Slice4beGuidePresentationScenarios.ps1')).Hash}|
        ConvertTo-Json|Set-Content (Join-Path $reportRoot 'paired-scenario-test-source.json')
    function ScenarioControl([string]$Name,[string]$Action,[string]$Value='') {
        try {BoundControl $Name $Action $Value 'frmActionPathView'}
        catch {
            Write-Host ('Paired scenario control failed: '+$Name+'; action='+$Action)
            throw
        }
    }
    function ScenarioPair($Guide,$Recording) {
        Write-Host 'Paired scenario: close preceding view'
        if((ScenarioControl '' 'Count') -ceq '1'){[void](ScenarioControl 'btnCloseActionPathView' 'Click')}
        Write-Host 'Paired scenario: select observed recording'
        if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $Recording.ActionPathId) -cne 'SELECTED'){throw 'Accepted scenario recording selection failed.'}
        Write-Host 'Paired scenario: select exact published guide'
        BoundOpen;BoundSelect $Guide
        Write-Host 'Paired scenario: use selected guide'
        if((BoundControl 'btnUseGuideForRun' 'Click') -cne 'DELIVERED'){throw 'Accepted scenario guide application failed.'}
        [void](BoundControl 'btnCloseGuides' 'Click')
        Write-Host 'Paired scenario: evaluate explicit pair'
        $result=BoundEvaluate
        if($null -eq $result){throw 'Accepted scenario Evaluate did not append its result.'}
        Write-Host 'Paired scenario: open paired view'
        if((BoundControl 'btnViewActionPath' 'Click' '' 'frmActionPaths') -cne 'DELIVERED' -or (ScenarioControl '' 'Count') -cne '1'){throw 'Paired view fixture did not open.'}
        [void](ScenarioControl 'cboActionPathView' 'Write' 'Compare both')
        return $result
    }
    function CaptureScenario([string]$Name) {
        $beforeCapture=BoundPins $journalRoot
        CaptureOwnedFormByCaptionEvidence 'Action Path view' ('paired-scenario-'+$Name.ToLowerInvariant()+'.png')
        if((ScenarioControl 'txtActionPathDiagnostic' 'Text') -cne ''){
            if((ScenarioControl 'txtActionPathDiagnostic' 'ViewportBottom') -cne 'DELIVERED'){throw 'Scenario diagnostic viewport is unavailable.'}
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ('paired-scenario-'+$Name.ToLowerInvariant()+'-result.png')
            [void](ScenarioControl 'txtActionPathDiagnostic' 'ViewportTop')
        }
        Check ('GuidePresentation.ScenarioCapture.'+$Name) $true
        Check ('GuidePresentation.ScenarioCaptureReadOnly.'+$Name) (BoundSame $beforeCapture (BoundPins $journalRoot))
    }
    function ScenarioDisplayPreserves([string]$Name,$Before) {
        [void](ScenarioControl 'btnRefreshActionPathView' 'Click')
        foreach($method in @('How-To','Diagnostic','Compare both')){[void](ScenarioControl 'cboActionPathView' 'Write' $method)}
        Check ('GuidePresentation.ScenarioReadOnly.'+$Name) (BoundSame $Before (BoundPins $journalRoot))
    }
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    OpenRecordingViewer
    try {
        $empty=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Where-Object {$_.RecordType -ceq 'Close' -and $_.Lifecycle -ceq 'Stopped' -and @($_.Observations).Count -eq 0})
        $none=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json'|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Where-Object {$_.ExpectedConclusion.TerminalKind -ceq 'None'})
        if($empty.Count -eq 0 -or $none.Count -eq 0){throw 'Accepted empty-run/None-guide scenario fixtures are missing.'}
        $missing=ScenarioPair $FirstGuide $empty[0]
        if($missing.ResultState -cne 'Failed' -or @($missing.MissingSteps).Count -ne 2){throw 'Existing evaluator did not produce the missing-step fixture.'}
        $text=ScenarioControl 'txtActionPathDiagnostic' 'Text'
        Check 'GuidePresentation.ScenarioMissingIdentifiesExpectedSteps' ($text.Contains('Missing expected step') -and $text.Contains([string]$missing.MissingSteps[0]) -and $text.Contains([string]$missing.MissingSteps[1]) -and -not $text.Contains('Conclusion observed'))
        ScenarioDisplayPreserves 'Missing' (BoundPins $journalRoot);CaptureScenario 'Missing'

        $extra=ScenarioPair $SecondGuide $Observed
        if(@($extra.Matches).Count -ne 1 -or @($extra.ExtraActivityIds).Count -ne 1){throw 'Existing evaluator did not produce the extra-action fixture.'}
        $text=ScenarioControl 'txtActionPathDiagnostic' 'Text'
        Check 'GuidePresentation.ScenarioExtraRetainsOriginalActionIdentity' ($text.Contains('Additional observed actions: 1') -and $text.Contains('Additional observed action: '+[string]$extra.ExtraActivityIds[0]) -and $text.Contains([string]$extra.Matches[0].ActivityId))
        $extraPins=BoundPins $journalRoot
        Check 'GuidePresentation.ScenarioExtraMinimumLayout' ((ScenarioControl '' 'Fit' 'Minimum') -ceq 'True')
        ScenarioDisplayPreserves 'Extra' $extraPins;CaptureScenario 'Extra'
        if((ScenarioControl '' 'Fit' 'Default') -cne 'True'){throw 'Scenario layout restoration failed.'}

        $noExpectation=ScenarioPair $none[0] $Observed
        if($noExpectation.ResultState -cne 'Incomplete' -or $noExpectation.ExpectedConclusion.TerminalKind -cne 'None'){throw 'Existing evaluator did not produce the None fixture.'}
        $text=ScenarioControl 'txtActionPathDiagnostic' 'Text'
        Check 'GuidePresentation.ScenarioNoneHasInstructionsWithoutConclusion' ((ScenarioControl 'txtActionPathHowTo' 'Text').Contains('Authored instruction') -and $text.Contains('Incomplete evidence') -and -not $text.Contains('Conclusion observed'))
        ScenarioDisplayPreserves 'None' (BoundPins $journalRoot);CaptureScenario 'None'

        [void](ScenarioControl 'btnCloseActionPathView' 'Click')
        if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Actual rejected-action recording cannot start.'}
        $activityBefore=ActivityPins;$configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            $accepted=Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','not-a-number')
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
        $records=@(foreach($file in Get-ChildItem -LiteralPath $activityRoot -File -Filter '*.json'){if(-not $activityBefore.ContainsKey($file.Name)){Get-Content -LiteralPath $file.FullName -Raw|ConvertFrom-Json}})
        $attempt=@($records|Where-Object OutcomeCode -CEQ 'REQUESTED');$rejected=@($records|Where-Object OutcomeCode -CEQ 'REJECTED')
        if($accepted -isnot [bool] -or $accepted -or $records.Count -ne 2 -or $attempt.Count -ne 1 -or $rejected.Count -ne 1 -or $attempt[0].ActivityId -cne $rejected[0].ActivityId){throw 'Actual rejected-action fixture failed.'}
        if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED' -or -not (JournalChain $attempt[0].SequenceId 4)){throw 'Rejected-action journal did not close correctly.'}
        Check 'GuidePresentation.RejectedScenarioPreservesConfig' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore)
        $failedRun=@(RecordingJournal $attempt[0].SequenceId|Where-Object RecordType -CEQ 'Close')
        if($failedRun.Count -ne 1){throw 'Rejected-action journal identity is ambiguous.'}
        [void](BoundLibrary 'Refresh')
        $failed=ScenarioPair $FirstGuide $failedRun[0]
        if($failed.ResultState -cne 'Failed' -or @($failed.FailedSteps).Count -ne 1){throw 'Existing evaluator did not produce the failed-outcome fixture.'}
        $text=ScenarioControl 'txtActionPathDiagnostic' 'Text'
        Check 'GuidePresentation.ScenarioFailureRetainsRejectedActionAndStep' ($text.Contains('Outcome mismatch') -and $text.Contains('REJECTED') -and $text.Contains([string]$attempt[0].ActivityId) -and $text.Contains([string]$failed.FailedSteps[0]) -and -not $text.Contains('Conclusion observed'))
        ScenarioDisplayPreserves 'Failed' (BoundPins $journalRoot);CaptureScenario 'Failed'

        $historical=ScenarioPair $FirstGuide $Observed
        $historicalPins=BoundPins $journalRoot
        $context=Run 'invSys.Core.xlam' 'modActivity.CaptureContext'
        try {
            SetRecordingPolicy $false
            [void](ScenarioControl 'btnRefreshActionPathView' 'Click')
            $text=ScenarioControl 'txtActionPathDiagnostic' 'Text'
            Check 'GuidePresentation.CaptureOffRetainsHistoricalEvidenceWithNotice' ((Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -ceq $context -and $text.Contains([string]$historical.EvaluationId) -and (ScenarioControl 'lblActionPathViewStatus' 'Label').Contains('capture is off'))
            ScenarioDisplayPreserves 'CaptureOff' $historicalPins;CaptureScenario 'CaptureOff'
        } finally {SetRecordingPolicy $true}
        try {
            SaveGuideExpectationVisibility $false
            [void](ScenarioControl 'btnRefreshActionPathView' 'Click')
            Check 'GuidePresentation.CurrentVisibilityClearsSameSessionRetainedContent' ((Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -ceq $context -and (ScenarioControl 'txtActionPathHowTo' 'Text') -ceq '' -and (ScenarioControl 'txtActionPathDiagnostic' 'Text') -ceq '' -and (ScenarioControl 'lblActionPathViewStatus' 'Label') -match '(?i)unavailable|incomplete')
            Check 'GuidePresentation.CurrentVisibilityPreservesHistoricalTraining' (BoundSame $historicalPins (BoundPins $journalRoot))
            CaptureScenario 'Restricted'
        } finally {SaveGuideExpectationVisibility $true}
        if($CheckGuidePresentationAvailability){
            . (Join-Path $PSScriptRoot 'Slice4beGuidePresentationAvailability.ps1')
            Test-GuidePresentationAvailability $Fixture $FirstGuide
        }
    } finally {CloseRecordingViewer}
}
