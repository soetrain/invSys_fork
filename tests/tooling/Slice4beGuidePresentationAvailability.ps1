# D18 presentation of historical capture gaps and current unreadable policy.
# Requires the existing TrackingPolicy probes installed before any forms open.
function Test-GuidePresentationAvailability($Fixture,$FirstGuide) {
    function SetRequiredControlCapture([bool]$Enabled) {
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyControl' @('ADMIN_SETTINGS_SAVE_VALUE',$Enabled))
            $saved=Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicySave'
            if($saved -isnot [bool] -or -not $saved){throw 'Actual required-control policy fixture did not save.'}
            $verified=([string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyRequest'))|ConvertFrom-Json
            $row=@($verified.Controls|Where-Object ControlId -CEQ 'ADMIN_SETTINGS_SAVE_VALUE')
            if($row.Count -ne 1 -or $row[0].Collect -isnot [bool] -or $row[0].Visible -isnot [bool] -or $row[0].SequenceEligible -isnot [bool] -or
                $row[0].Collect -ne $Enabled -or $row[0].Visible -ne $Enabled -or $row[0].SequenceEligible -ne $Enabled){throw 'Required-control policy fixture did not retain all three flags.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    }
    try {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
        $request=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingPolicyRequest')
        $policy=$request|ConvertFrom-Json
        $required=@($policy.Controls|Where-Object ControlId -CEQ 'ADMIN_SETTINGS_SAVE_VALUE')
        if($required.Count -ne 1 -or -not $required[0].Collect -or -not $required[0].Visible -or -not $required[0].SequenceEligible){throw 'Accepted policy fixture must start with the required control enabled.'}
    } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    if((ScenarioControl '' 'Count') -ceq '1'){[void](ScenarioControl 'btnCloseActionPathView' 'Click')}
    $before=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|ForEach-Object FullName)
    try {
        SetRequiredControlCapture $false
        $activityBefore=ActivityPins
        if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Historical gap fixture recording could not start.'}
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            $saved=Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('BatchSize','674')
            if($saved -isnot [bool] -or -not $saved){throw 'Actual uncaptured Settings action failed.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
        # Executing an excluded command closes the active recorder as Incomplete;
        # a later Stop must not be used to relabel that capture as Stopped.
        $entries=@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'|Where-Object FullName -CNotIn $before|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json})
        $closed=@($entries|Where-Object RecordType -CEQ 'Close')
        $gapFacts=[ordered]@{JournalEntries=$entries.Count;CloseEntries=$closed.Count;Incomplete=$false;TrackingUnavailable=$false;Observations=-1;ActivityUnchanged=(BoundSame $activityBefore (ActivityPins))}
        if($closed.Count -eq 1){
            $gapFacts.Incomplete=$closed[0].Lifecycle -ceq 'Incomplete'
            $gapFacts.TrackingUnavailable=$closed[0].ReasonCode -ceq 'TRACKING_UNAVAILABLE'
            $gapFacts.Observations=@($closed[0].Observations).Count
        }
        $gapFacts|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'paired-historical-gap-fixture.json')
        $gapValid=$gapFacts.JournalEntries -eq 2 -and $gapFacts.CloseEntries -eq 1 -and $gapFacts.Incomplete -and $gapFacts.TrackingUnavailable -and $gapFacts.Observations -eq 0 -and $gapFacts.ActivityUnchanged
        Check 'GuidePresentation.DisabledRequiredCommandInterruptsRecording' $gapValid
        if(-not $gapValid){throw 'Historical gap fixture must retain the actual Incomplete closure and excluded collection; see sanitized count/flag evidence.'}
    } finally {SetRequiredControlCapture $true}
    $unavailable=ScenarioPair $FirstGuide $closed[0]
    if($unavailable.ResultState -cne 'Incomplete' -or @($unavailable.UnavailableSteps).Count -ne 2 -or @($unavailable.MissingSteps).Count -ne 0 -or @($unavailable.Matches).Count -ne 0){throw 'Existing evaluator did not produce the historical unavailable-step fixture.'}
    $text=ScenarioControl 'txtActionPathDiagnostic' 'Text'
    Check 'GuidePresentation.HistoricalUnavailableNamesBothExpectedSteps' ($text.Contains('Evidence unavailable') -and $text.Contains([string]$unavailable.UnavailableSteps[0]) -and $text.Contains([string]$unavailable.UnavailableSteps[1]) -and $text.Contains([string]$unavailable.EvaluationId) -and -not $text.Contains('Conclusion observed'))
    ScenarioDisplayPreserves 'Unavailable' (BoundPins $journalRoot);CaptureScenario 'Unavailable'

    $configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $trainingBefore=BoundPins $journalRoot
    $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
    try {
        $table=Table $book 'tblWarehouseConfig'
        $cell=$table.ListColumns.Item('WarehouseName').DataBodyRange.Cells.Item(1,1)
        $cell.Value2='unsaved paired policy fixture'
        $value=$cell.Value2
        [void](ScenarioControl 'btnRefreshActionPathView' 'Click')
        Check 'GuidePresentation.UnreadablePolicyClearsPanesAndDisablesMethod' ((ScenarioControl 'txtActionPathHowTo' 'Text') -ceq '' -and (ScenarioControl 'txtActionPathDiagnostic' 'Text') -ceq '' -and (ScenarioControl 'cboActionPathView' 'State') -ceq 'True|False' -and (ScenarioControl 'lblActionPathViewStatus' 'Label') -match '(?i)unavailable|incomplete')
        CaptureScenario 'UnreadablePolicy'
        Check 'GuidePresentation.UnreadablePolicyPreservesUnsavedWorkbook' (-not $book.Saved -and $cell.Value2 -ceq $value)
    } finally {$book.Close($false)}
    Check 'GuidePresentation.UnreadablePolicyPreservesConfigAndTrainingBytes' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore -and (BoundSame $trainingBefore (BoundPins $journalRoot)))
    [void](ScenarioControl 'btnRefreshActionPathView' 'Click')
    Check 'GuidePresentation.RestoredPolicyReadsSameHistoricalResult' ((ScenarioControl 'txtActionPathDiagnostic' 'Text').Contains([string]$unavailable.EvaluationId))
}
