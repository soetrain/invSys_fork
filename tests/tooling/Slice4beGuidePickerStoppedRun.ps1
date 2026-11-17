# D18: a stopped recording remains selected while the published-action picker
# and both real Create guide controls are used. The fixture creates actual
# tracked Admin actions and publishes them through the packaged callback.
function Test-GuidePickerStoppedRun($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    function Control([string]$Form,[string]$Name,[string]$Action,[string]$Value='') {
        BoundControl $Name $Action $Value $Form
    }
    function Picker([string]$Name,[string]$Action,[string]$Value='') {
        Control 'frmGuideActionPicker' $Name $Action $Value
    }
    function Editor([string]$Name,[string]$Action,[string]$Value='') {
        Control 'frmActionPathGuide' $Name $Action $Value
    }

    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    SetRecordingPolicy $true -SaveValueOnly
    OpenRecordingViewer
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED') {
        throw 'Actual Start Recording control unavailable; not picker RED.'
    }
    $recorded=@(for($i=810;$i -le 814;$i++){SaveRecordedSetting ([string]$i)})
    if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED') {
        throw 'Actual Stop Recording control unavailable; not picker RED.'
    }
    $sequence=[string]$recorded[0].Attempt.SequenceId
    if($recorded.Count -ne 5 -or -not (HasSequence $recorded[0] 1) -or
       -not (HasSequence $recorded[4] 5) -or -not (JournalChain $sequence 12)) {
        throw 'Stopped actual recording fixture invalid; not picker RED.'
    }
    $entries=@(RecordingJournal $sequence | Sort-Object Version)
    $pathId=[string]$entries[0].ActionPathId
    CloseRecordingViewer

    # Match the operator's busy Events publication with distinct real actions.
    for($i=815;$i -le 831;$i++){[void](SaveRecordedSetting ([string]$i))}
    $published=Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest'
    $publication=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $valid=Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($publication,$Fixture.Warehouse)
    if($published -isnot [bool] -or -not $published -or $valid -isnot [bool] -or -not $valid) {
        throw 'Actual published Events fixture unavailable; not picker RED.'
    }
    $model=Get-Content -LiteralPath $publication -Raw | ConvertFrom-Json
    $groups=@($model.Groups | Where-Object {$_.Source -ceq 'Activity'})
    if($groups.Count -lt 22){throw 'Published Events lacks the real action groups; not picker RED.'}
    $activityBefore=ActivityPins

    try {
        OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $pathId) -cne 'SELECTED') {
            throw 'Selected stopped recording unavailable; not picker RED.'
        }
        Check 'GuidePickerStoppedRun.SelectedStoppedRecording' ((Control 'frmActionPaths' 'btnCreateGuide' 'State') -ceq 'True|True')
        [void](Control 'frmActionPaths' 'btnChooseGuideActions' 'Click')
        $ids=@((Picker 'lstGuideActions' 'Values') -split "`n" | Where-Object {$_ -cne ''})
        if((Picker '' 'Count') -cne '1' -or $ids.Count -lt 22) {
            throw 'Visible picker lacks the loaded action candidates; not picker RED.'
        }
        $timer=[Diagnostics.Stopwatch]::StartNew()
        $toggle=Picker 'lstGuideActions' 'Toggle' '0'
        $timer.Stop()
        $toggleSeconds=$timer.Elapsed.TotalSeconds
        [pscustomobject]@{Stage='FirstToggle';Seconds=$toggleSeconds;Candidates=$ids.Count} |
            ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $reportRoot 'guide-picker-stopped-run-timing.json')
        Check 'GuidePickerStoppedRun.FirstActionSelectsThroughPicker' ($toggle -ceq 'SELECTED' -and (Picker 'lstGuideActions' 'SelectedValues') -ceq ($ids[0]+"`n"))
        Check 'GuidePickerStoppedRun.FirstActionResponsive' ($toggleSeconds -lt 5)

        $timer=[Diagnostics.Stopwatch]::StartNew()
        $secondToggle=Picker 'lstGuideActions' 'Toggle' '1'
        $timer.Stop()
        $secondToggleSeconds=$timer.Elapsed.TotalSeconds
        Check 'GuidePickerStoppedRun.RepeatedActionSelectsThroughPicker' ($secondToggle -ceq 'SELECTED' -and @((Picker 'lstGuideActions' 'SelectedValues') -split "`n") -ccontains $ids[1])
        Check 'GuidePickerStoppedRun.RepeatedActionResponsive' ($secondToggleSeconds -lt 5)

        [void](Control 'frmActionPaths' 'btnCreateGuide' 'Click')
        $runEditor=(Editor '' 'Count') -ceq '1'
        Check 'GuidePickerStoppedRun.RunCreateOpensEditorWithPickerPresent' ($runEditor -and (Picker '' 'Count') -ceq '1' -and (Editor 'lblGuideSource' 'Label').Contains($pathId))
        if(-not $runEditor){throw 'Actual guide editor unavailable; not typing-performance RED.'}
        $beforeStep=Editor 'lstGuideSteps' 'Selected'
        $timer=[Diagnostics.Stopwatch]::StartNew()
        $stepSelect=Editor 'lstGuideSteps' 'SelectDirect' '1'
        $timer.Stop()
        $stepSelectSeconds=$timer.Elapsed.TotalSeconds
        $afterStep=Editor 'lstGuideSteps' 'Selected'
        Check 'GuidePickerStoppedRun.GuideStepSelectionRetained' ($stepSelect -ceq 'SELECTED' -and $beforeStep -cne '' -and $afterStep -cne '' -and $afterStep -cne $beforeStep)
        Check 'GuidePickerStoppedRun.GuideStepSelectionResponsive' ($stepSelectSeconds -lt 2.5)
        # The test-only Type action exercises each real Change handler inside Excel.
        $inputTimes=@{}; $inputLengths=@{}
        foreach($spec in @(@('txtGuideName','Training guide title'),@('txtGuideInstructions','Follow the recorded steps'),@('txtGuideStepInstruction','Check the saved action'))) {
            $name=[string]$spec[0]; $value=[string]$spec[1]
            $timer=[Diagnostics.Stopwatch]::StartNew()
            if((Editor $name 'Type' $value) -cne 'DELIVERED'){
                throw ('Actual guide editor text handler unavailable: '+$name+'; not typing-performance RED.')
            }
            $timer.Stop()
            $inputTimes[$name]=$timer.Elapsed.TotalSeconds
            $inputLengths[$name]=$value.Length
            Check ('GuidePickerStoppedRun.'+$name+'.IncrementalTextPreserved') ((Editor $name 'Text') -ceq $value)
            Check ('GuidePickerStoppedRun.'+$name+'.IncrementalInputResponsive') ($inputTimes[$name] -lt 2.5)
        }
        [pscustomobject]@{Stage='IncrementalEditorInput';SecondToggleSeconds=$secondToggleSeconds;StepSelectSeconds=$stepSelectSeconds;NameCharacters=$inputLengths['txtGuideName'];NameSeconds=$inputTimes['txtGuideName'];InstructionsCharacters=$inputLengths['txtGuideInstructions'];InstructionsSeconds=$inputTimes['txtGuideInstructions'];StepCharacters=$inputLengths['txtGuideStepInstruction'];StepSeconds=$inputTimes['txtGuideStepInstruction']} |
            ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $reportRoot 'guide-editor-input-timing.json')
        [void](Editor 'txtGuideName' 'Write' 'Stopped run guide')
        [void](Editor 'btnSaveGuide' 'Click')
        $guideRoot=Join-Path $journalRoot 'Guides'
        $saved=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.1.json' -ErrorAction SilentlyContinue)
        $runGuide=$null
        if($saved.Count -eq 1){$runGuide=ReadGuideExpectationRecord $saved[0].FullName}
        Check 'GuidePickerStoppedRun.RunSavePublishesImmutableGuide' ($null -ne $runGuide -and $runGuide.Name -ceq 'Stopped run guide' -and
            $runGuide.SourceRun.ActionPathId -ceq $pathId -and $runGuide.SourceRun.JournalVersion -eq 12 -and @($runGuide.Steps).Count -eq 5)

        [void](Picker 'btnCreateSelectedGuide' 'Click')
        $directEditor=(Editor '' 'Count') -ceq '1'
        Check 'GuidePickerStoppedRun.PickerCreateRebindsSingleEditor' ($directEditor -and (Editor 'lblGuideSource' 'Label') -match 'No recorded sequence; authored order is not execution evidence')
        [void](Editor 'txtGuideName' 'Write' 'Selected published action guide')
        [void](Editor 'btnSaveGuide' 'Click')
        $saved=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.1.json' -ErrorAction SilentlyContinue)
        $directGuide=@($saved | ForEach-Object {ReadGuideExpectationRecord $_.FullName} | Where-Object {$null -ne $_ -and $_.Name -ceq 'Selected published action guide'})
        Check 'GuidePickerStoppedRun.PickerSavePublishesSeparateImmutableGuide' ($directGuide.Count -eq 1 -and $saved.Count -eq 2 -and @($directGuide[0].SourceRun.PSObject.Properties).Count -eq 0)

        [void](Picker 'btnCancelGuideActions' 'Click')
        Check 'GuidePickerStoppedRun.CancelClosesPickerAndOwnedEditor' ((Picker '' 'Count') -ceq '0' -and (Editor '' 'Count') -ceq '0')
        Check 'GuidePickerStoppedRun.CancelRetainsBothPublishedGuides' (@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.1.json').Count -eq 2)
        Check 'GuidePickerStoppedRun.SourceActivityBytesUnchanged' (PinsRetained $activityBefore)
    } finally {
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
    }
}
