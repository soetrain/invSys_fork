# D18: a stopped recording remains selected while the published-action picker
# and both real Create guide controls are used. The fixture creates actual
# tracked Admin actions and publishes them through the packaged callback.
function Test-GuidePickerStoppedRun($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideInputNative.ps1')
    function Control([string]$Form,[string]$Name,[string]$Action,[string]$Value='') {
        BoundControl $Name $Action $Value $Form
    }
    function Picker([string]$Name,[string]$Action,[string]$Value='') {
        Control 'frmGuideActionPicker' $Name $Action $Value
    }
    function Editor([string]$Name,[string]$Action,[string]$Value='') {
        Control 'frmActionPathGuide' $Name $Action $Value
    }
    function GuideWrites { [long](Run 'invSys.Core.xlam' 'TestGuideInputWriteCount.ReadCount') }
    function ResetGuideWrites { [void](Run 'invSys.Core.xlam' 'TestGuideInputWriteCount.ResetCount') }

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
        if((Editor '' 'Place') -cne 'PLACED'){throw 'Guide input form could not be placed on the current display; not behavioral RED.'}
        $beforeStep=Editor 'lstGuideSteps' 'Selected'
        $timer=[Diagnostics.Stopwatch]::StartNew()
        $stepSelect=Editor 'lstGuideSteps' 'SelectDirect' '1'
        $timer.Stop()
        $stepSelectSeconds=$timer.Elapsed.TotalSeconds
        $afterStep=Editor 'lstGuideSteps' 'Selected'
        Check 'GuidePickerStoppedRun.GuideStepSelectionRetained' ($stepSelect -ceq 'SELECTED' -and $beforeStep -cne '' -and $afterStep -cne '' -and $afterStep -cne $beforeStep)
        Check 'GuidePickerStoppedRun.GuideStepSelectionResponsive' ($stepSelectSeconds -lt 2.5)
        # The test-only Type action exercises each real Change handler inside Excel.
        $inputTimes=@{}; $inputLengths=@{}; $blurWrites=@{}
        foreach($spec in @(@('txtGuideName','Training guide title'),@('txtGuideTags','training, review'),@('txtGuideInstructions','Follow the recorded steps'),@('txtGuideStepInstruction','Check the saved action'))) {
            $name=[string]$spec[0]; $value=[string]$spec[1]
            if((Focus-GuideControlNative $name) -cne 'FOCUSED'){throw ('Native guide focus fixture unavailable: '+$name+'; not buffering RED.')}
            ResetGuideWrites
            $timer=[Diagnostics.Stopwatch]::StartNew()
            if((Type-GuideTextNative $name $value) -cne 'DELIVERED'){
                throw ('Actual guide editor text handler unavailable: '+$name+'; not typing-performance RED.')
            }
            $timer.Stop()
            $inputTimes[$name]=$timer.Elapsed.TotalSeconds
            $inputLengths[$name]=$value.Length
            Check ('GuidePickerStoppedRun.'+$name+'.IncrementalTextPreserved') ((Editor $name 'Text') -ceq $value)
            Check ('GuidePickerStoppedRun.'+$name+'.IncrementalInputResponsive') ($inputTimes[$name] -lt 2.5)
            Check ('GuidePickerStoppedRun.'+$name+'.TypingDoesNotWriteCore') ((GuideWrites) -eq 0)
            if($name -cne 'txtGuideStepInstruction'){
                if($name -ceq 'txtGuideName'){
                    [void](Focus-GuideControlNative 'txtGuideTags')
                    $newFocus=Editor '' 'Active'
                } else {
                    $newFocus=Move-GuideFocusByTabNative
                }
                $blurWrites[$name]=GuideWrites
                Check ('GuidePickerStoppedRun.'+$name+'.OrdinaryBlurFlushesOnce') ($newFocus -cne $name -and $blurWrites[$name] -eq 1)
            }
        }
        [pscustomobject]@{Stage='IncrementalEditorInput';SecondToggleSeconds=$secondToggleSeconds;StepSelectSeconds=$stepSelectSeconds;NameCharacters=$inputLengths['txtGuideName'];NameSeconds=$inputTimes['txtGuideName'];TagsCharacters=$inputLengths['txtGuideTags'];TagsSeconds=$inputTimes['txtGuideTags'];InstructionsCharacters=$inputLengths['txtGuideInstructions'];InstructionsSeconds=$inputTimes['txtGuideInstructions'];StepCharacters=$inputLengths['txtGuideStepInstruction'];StepSeconds=$inputTimes['txtGuideStepInstruction']} |
            ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $reportRoot 'guide-editor-input-timing.json')
        $stepBeforeSwitch=$afterStep
        ResetGuideWrites
        $switch=Editor 'lstGuideSteps' 'SelectDirect' '2'
        $stepAfterSwitch=Editor 'lstGuideSteps' 'Selected'
        $stepSwitchWrites=GuideWrites
        [pscustomobject]@{Stage='InputBoundaries';BlurWrites=$blurWrites;StepSwitchWrites=$stepSwitchWrites;ActiveAfterSwitch=(Editor '' 'Active')} |
            ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $reportRoot 'guide-input-boundary-counts.json')
        Check 'GuidePickerStoppedRun.DirtyStepSwitchFlushesOnce' ($switch -ceq 'SELECTED' -and $stepAfterSwitch -cne $stepBeforeSwitch -and $stepSwitchWrites -eq 1)
        Check 'GuidePickerStoppedRun.DirtyStepSwitchLoadsNewStep' ((Editor 'txtGuideStepInstruction' 'Text') -ceq '')
        if((Editor 'txtGuideName' 'Focus') -cne 'FOCUSED'){throw 'Guide name focus before Save unavailable.'}
        ResetGuideWrites
        [void](Editor 'txtGuideName' 'Write' 'Stopped run guide')
        [void](Editor 'btnSaveGuide' 'Click')
        Check 'GuidePickerStoppedRun.SaveWhileNameFocusedFlushesOnce' ((GuideWrites) -eq 1)
        $guideRoot=Join-Path $journalRoot 'Guides'
        $saved=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.1.json' -ErrorAction SilentlyContinue)
        $runGuide=$null
        if($saved.Count -eq 1){$runGuide=ReadGuideExpectationRecord $saved[0].FullName}
        if($null -eq $runGuide){throw 'Saved run guide prerequisite invalid; not dirty-boundary RED.'}
        $firstGuideHash=(Get-FileHash -LiteralPath $saved[0].FullName).Hash
        Check 'GuidePickerStoppedRun.RunSavePublishesImmutableGuide' ($null -ne $runGuide -and $runGuide.Name -ceq 'Stopped run guide' -and
            $runGuide.SourceRun.ActionPathId -ceq $pathId -and $runGuide.SourceRun.JournalVersion -eq 12 -and @($runGuide.Steps).Count -eq 5)
        $boundStep=@();$newStep=@()
        if($null -ne $runGuide){
            $boundStep=@($runGuide.Steps | Where-Object {$_.StepId -ceq $stepBeforeSwitch})
            $newStep=@($runGuide.Steps | Where-Object {$_.StepId -ceq $stepAfterSwitch})
        }
        Check 'GuidePickerStoppedRun.SavePersistsExactTextAndPreviousStep' ($null -ne $runGuide -and
            $runGuide.Instructions -ceq 'Follow the recorded steps' -and (@($runGuide.Tags) -join '|') -ceq 'training|review' -and
            $boundStep.Count -eq 1 -and $boundStep[0].Instruction -ceq 'Check the saved action' -and
            $newStep.Count -eq 1 -and $newStep[0].Instruction -ceq '')

        if((Focus-GuideControlNative 'txtGuideStepInstruction') -cne 'FOCUSED'){throw 'Native step blur fixture unavailable; not behavioral RED.'}
        ResetGuideWrites
        [void](Type-GuideTextNative 'txtGuideStepInstruction' 'Unpublished temporary step text')
        $stepBlurFocus=Move-GuideFocusByTabNative
        $stepBlurWrites=GuideWrites
        Check 'GuidePickerStoppedRun.txtGuideStepInstruction.OrdinaryBlurFlushesOnce' ($stepBlurFocus -cne 'txtGuideStepInstruction' -and $stepBlurWrites -eq 1)

        # Reorder must flush the text bound to the old StepId before the list
        # changes selection. Save proves the exact assignment in the next version.
        [void](Editor 'txtGuideStepInstruction' 'Write' 'Reordered step instruction')
        ResetGuideWrites
        $moved=Editor 'btnGuideStepUp' 'Click'
        Check 'GuidePickerStoppedRun.DirtyReorderFlushesPreviousStepOnce' ($moved -ceq 'DELIVERED' -and
            (GuideWrites) -eq 1 -and (Editor 'lstGuideSteps' 'Selected') -ceq $stepAfterSwitch -and
            (Editor 'txtGuideStepInstruction' 'Text') -ceq 'Reordered step instruction')
        [void](Editor 'btnSaveGuide' 'Click')
        $revisedPath=Join-Path $guideRoot ($runGuide.ActionPathId+'.2.json')
        $revised=ReadGuideExpectationRecord $revisedPath
        if($null -eq $revised){throw 'Reordered guide version could not be read; not dirty-assignment RED.'}
        $revisedGuideHash=(Get-FileHash -LiteralPath $revisedPath).Hash
        $revisedStep=@($revised.Steps|Where-Object {$_.StepId -ceq $stepAfterSwitch})
        Check 'GuidePickerStoppedRun.DirtyReorderSaveKeepsExactStepAssignment' ($null -ne $revised -and
            $revisedStep.Count -eq 1 -and $revisedStep[0].Instruction -ceq 'Reordered step instruction' -and
            @($revised.Steps)[1].StepId -ceq $stepAfterSwitch -and
            (Get-FileHash -LiteralPath $saved[0].FullName).Hash -ceq $firstGuideHash)

        [void](Editor 'txtGuideStepInstruction' 'Write' 'Remove this unsaved step')
        ResetGuideWrites
        $removed=Editor 'btnRemoveGuideStep' 'Click'
        $remaining=@((Editor 'lstGuideSteps' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        Check 'GuidePickerStoppedRun.DirtyRemovalFlushesBoundStepBeforeRemoval' ($removed -ceq 'DELIVERED' -and
            (GuideWrites) -eq 1 -and $remaining -cnotcontains $stepAfterSwitch -and
            (Get-FileHash -LiteralPath $revisedPath).Hash -ceq $revisedGuideHash)

        [void](Editor 'txtGuideTags' 'Write' 'expectation, pending')
        ResetGuideWrites
        $expectation=Editor 'btnGuideExpectedConclusion' 'Click'
        $expectationOpened=(Control 'frmActionPathExpectation' '' 'Count') -ceq '1'
        $expectationWrites=GuideWrites
        $expectationCancelled=Control 'frmActionPathExpectation' 'btnCancelExpectation' 'Click'
        Check 'GuidePickerStoppedRun.DirtyExpectationEntryFlushesAndReturnRetainsText' ($expectation -ceq 'DELIVERED' -and
            $expectationOpened -and $expectationWrites -eq 1 -and $expectationCancelled -ceq 'DELIVERED' -and
            (Editor 'txtGuideTags' 'Text') -ceq 'expectation, pending' -and
            (Get-FileHash -LiteralPath $revisedPath).Hash -ceq $revisedGuideHash)

        [void](Editor 'txtGuideName' 'Write' 'Rebound guide draft')
        ResetGuideWrites
        $reused=Control 'frmActionPaths' 'btnCreateGuide' 'Click'
        Check 'GuidePickerStoppedRun.SameSourceReentryRetainsDirtyLocalBuffer' ($reused -ceq 'DELIVERED' -and
            (GuideWrites) -eq 0 -and (Editor '' 'Count') -ceq '1' -and
            (Editor 'txtGuideName' 'Text') -ceq 'Rebound guide draft' -and
            (Get-FileHash -LiteralPath $revisedPath).Hash -ceq $revisedGuideHash)

        $rebound=Picker 'btnCreateSelectedGuide' 'Click'
        Check 'GuidePickerStoppedRun.DirtyRebindFlushesOnceBeforeReload' ($rebound -ceq 'DELIVERED' -and
            (GuideWrites) -eq 1 -and (Get-FileHash -LiteralPath $revisedPath).Hash -ceq $revisedGuideHash)
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

        # Actual Cancel/X event order, including stale authority, must not
        # turn dirty form text into a Core write or a published version.
        [void](BoundLibrary 'Open')
        [void](BoundLibrary 'Select' $pathId)
        [void](Control 'frmActionPaths' 'btnCreateGuide' 'Click')
        if((Editor '' 'Place') -cne 'PLACED'){throw 'Guide X form could not be placed on the current display; not behavioral RED.'}
        if((Editor '' 'Count') -cne '1' -or (Editor 'txtGuideName' 'Focus') -cne 'FOCUSED'){
            throw 'Dirty guide window-X fixture unavailable; not behavioral RED.'
        }
        ResetGuideWrites
        [void](Editor 'txtGuideName' 'Type' 'Close without writing')
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        $closed=Close-GuideWindowXNative
        $xEditors=Editor '' 'Count';$xWrites=GuideWrites
        [pscustomobject]@{Stage='AfterWindowX';WindowClosed=$closed;LoadedEditors=$xEditors;WriteCalls=$xWrites} |
            ConvertTo-Json -Compress | Set-Content -LiteralPath (Join-Path $reportRoot 'guide-input-close-state.json')
        Check 'GuidePickerStoppedRun.StaleDirtyWindowXClosesWithoutCoreWriteOrPublication' ($closed -and
            $xEditors -ceq '0' -and $xWrites -eq 0 -and
            @(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json').Count -eq 3)

        [void](BoundLibrary 'Close')
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $pathId) -cne 'SELECTED'){
            throw 'Dirty guide Cancel source unavailable; not behavioral RED.'
        }
        $cancelCreate=Control 'frmActionPaths' 'btnCreateGuide' 'Click'
        if((Editor '' 'Place') -cne 'PLACED'){throw 'Guide Cancel form could not be placed on the current display; not behavioral RED.'}
        $cancelEditorCount=Editor '' 'Count'
        $cancelNameState=Editor 'txtGuideName' 'State'
        $cancelFocus=if($cancelEditorCount -ceq '1'){Editor 'txtGuideName' 'Focus'}else{'NO EDITOR'}
        if($cancelEditorCount -cne '1' -or $cancelFocus -cne 'FOCUSED'){
            throw ("Dirty guide Cancel fixture unavailable; create=$cancelCreate; editor count=$cancelEditorCount; name state=$cancelNameState; focus=$cancelFocus; library status=$(BoundLibrary 'Status'); not behavioral RED.")
        }
        ResetGuideWrites
        [void](Editor 'txtGuideName' 'Type' 'Discard this guide text')
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        $cancelled=Click-GuideCancelNative
        Check 'GuidePickerStoppedRun.StaleDirtyCancelClosesWithoutCoreWriteOrPublication' ($cancelled -and
            (Editor '' 'Count') -ceq '0' -and (GuideWrites) -eq 0 -and
            @(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json').Count -eq 3)

        CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED' -or (BoundLibrary 'Select' $pathId) -cne 'SELECTED' -or
            (Control 'frmActionPaths' 'btnCreateGuide' 'Click') -cne 'DELIVERED' -or
            (Editor '' 'Count') -cne '1'){
            throw 'Dirty stale Save source unavailable; not behavioral RED.'
        }
        [void](Editor 'txtGuideName' 'Write' 'Do not publish after sign-out')
        ResetGuideWrites
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        [void](Editor 'btnSaveGuide' 'Click')
        Check 'GuidePickerStoppedRun.StaleDirtySaveRefusesAndClearsEvidence' ((GuideWrites) -eq 1 -and
            (Editor 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and
            (Editor 'lblGuideStatus' 'Label') -match '(?i)unavailable|changed|reopen' -and
            @(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json').Count -eq 3)
    } finally {
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
    }
}
