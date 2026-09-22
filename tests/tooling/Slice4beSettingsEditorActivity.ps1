# D18 Settings observations require disposable probes compiled before forms.
# Source/fixture failures are distinct from missing-observation behavioral RED.
function Test-SettingsEditorActivity($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beSettingsEditorCases.ps1')
    . (Join-Path $repo 'tests/tooling/Slice4beRecordingFixture.ps1')
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $originalConfig=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $otherConfig=(Get-FileHash -LiteralPath $Other.Config).Hash
    $entryVisible=[bool]$excel.Visible
    $anchor=$null
    $editors=@{Admin=$false;Operations=$false}
    $settingsEvidence=if($CheckSettingsDiagnostics){@{}}else{$null}
    function SettingsAct([string]$Section,[string]$Action,[string]$Value='') {
        if($Section -ceq 'Operations'){
            [string](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.SettingsActivityAction' @($Action,$Value))
        }else{
            [string](Run 'invSys.Admin.xlam' 'TestD5Commands.SettingsActivityAction' @($Section,$Action,$Value))
        }
    }
    function SettingsOwnerState {
        ([string](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.State'))|ConvertFrom-Json
    }
    function CloseSettingsEditors {
        [void](Run 'invSys.Operations.xlam' 'modOperationsTrackingSettings.CloseSettings')
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')
        $editors.Admin=$false;$editors.Operations=$false
    }
    function OpenSettingsEditor([string]$Section) {
        if($Section -ceq 'Operations'){
            if(-not $editors.Operations){
                if(-not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.SettingsActivityOpen')){throw 'Actual Viewer Settings entry unavailable; not activity RED.'}
                $editors.Operations=$true
            }
        }else{
            if(-not $editors.Admin){
                [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
                $editors.Admin=$true
            }
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.SettingsActivitySection' @($Section))
        }
    }
    function StartSettingsRecording {
        $before=@(if(Test-Path -LiteralPath $journalRoot){Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File|ForEach-Object FullName})
        if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Actual Settings recording Start unavailable.'}
        $starts=@(Get-ChildItem -LiteralPath $journalRoot -Filter '*.json' -File|Where-Object {$_.FullName -cnotin $before}|ForEach-Object {
            [IO.File]::ReadAllText($_.FullName)|ConvertFrom-Json
        }|Where-Object RecordType -CEQ 'Start')
        if($starts.Count -ne 1){throw 'Actual Settings recording identity unavailable.'}
        [string]$starts[0].SequenceId
    }
    try {
        $excel.Visible=$true
        $anchor=$excel.Workbooks.Add()
        $anchor.Activate()
        CloseRecordingViewer
        CloseSettingsEditors
        SelectTarget $Fixture 'config-admin'
        Test-SettingsEditorCatalog
        $setup=ActivityPins
        if(-not [bool](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.SetupRecordingPolicy')){throw 'Generated recording policy setup failed.'}
        if(-not [bool](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.SetupPreference' @('Use warehouse default'))){throw 'Generated preference setup failed.'}
        OpenRecordingViewer
        OpenSettingsEditor 'Tracking'
        Check 'SettingsActivity.DirectSetupAndInitializationAreNotUserClicks' ((ActivityPins).Count -eq $setup.Count -and (PinsRetained $setup))
        if($SettingsSafetyOnly){
            $setup=ActivityPins
            if(-not [bool](Run 'invSys.Core.xlam' 'TestSettingsOwnerState.SetupProfile')){throw 'Generated profile setup failed.'}
            Check 'SettingsActivity.Safety.DirectProfileSetupIsNotUserClick' ((ActivityPins).Count -eq $setup.Count -and (PinsRetained $setup))
            . (Join-Path $PSScriptRoot 'Slice4beSettingsEditorSafety.ps1')
            Test-SettingsEditorSafety $Fixture
            return
        }
        $sequence=StartSettingsRecording
        $ordinal=0
        $captured=@{}
        foreach($case in @(SettingsActivityCases)){
            $ordinal++
            $opening=ActivityPins
            OpenSettingsEditor $case.Section
            Check ('SettingsActivity.'+$case.Id+'.OpeningDoesNotInventClick') ((ActivityPins).Count -eq $opening.Count -and (PinsRetained $opening))
            $before=@(Get-Slice4beActivityFiles $Fixture)
            $ownerBefore=SettingsOwnerState
            $configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            $requestBefore=''
            if($case.Section -cin @('Tracking','Detail')){$requestBefore=SettingsAct $case.Section 'Request'}
            [void](SettingsAct $case.Section $case.Action $case.Value)
            $ownerAfter=SettingsOwnerState
            $same=(Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore
            $ownerCorrect=$false
            switch($case.Owner){
                'ProfileAppend' {
                    $ownerCorrect=-not $same -and $ownerAfter.ProfileVersion -eq ($ownerBefore.ProfileVersion+1) -and
                        $ownerAfter.ProfileRequest -ceq $requestBefore -and $ownerAfter.PolicyRequest -ceq $ownerBefore.PolicyRequest
                }
                'PreferenceSave' {
                    $choice=SettingsAct $case.Section 'Choice'
                    $ownerCorrect=$same -and $ownerAfter.Preference -ceq $choice -and $ownerAfter.Preference -cne $ownerBefore.Preference
                }
                'PolicyStaging' {
                    $staged=SettingsAct 'Tracking' 'Request'
                    $ownerCorrect=$same -and $staged -cne $requestBefore -and $ownerAfter.PolicyRequest -ceq $ownerBefore.PolicyRequest
                }
                'ProfileStaging' {
                    $staged=SettingsAct 'Detail' 'Request'
                    $ownerCorrect=$same -and $staged -cne $requestBefore -and $ownerAfter.ProfileRequest -ceq $ownerBefore.ProfileRequest
                }
                'PolicyDefaults' {
                    $expected=[string](Run 'invSys.Core.xlam' 'modTrackingPolicySettings.DefaultRequest')
                    $ownerCorrect=$same -and (SettingsAct 'Tracking' 'Request') -ceq $expected
                }
                'ProfileDefaults' {
                    $expected=[string](Run 'invSys.Core.xlam' 'modEventDetailSettings.DefaultRequest')
                    $ownerCorrect=$same -and (SettingsAct 'Detail' 'Request') -ceq $expected
                }
                'PolicyRead' {$ownerCorrect=$same -and (SettingsAct 'Tracking' 'Request') -ceq $ownerAfter.PolicyRequest}
                'ProfileRead' {$ownerCorrect=$same -and (SettingsAct 'Detail' 'Request') -ceq $ownerAfter.ProfileRequest}
                'PreferenceRead' {$ownerCorrect=$same -and (SettingsAct $case.Section 'Choice') -ceq $ownerAfter.Preference}
                'PreferenceDefaults' {$ownerCorrect=$same -and (SettingsAct $case.Section 'Choice') -ceq 'Use warehouse default' -and $ownerAfter.Preference -ceq $ownerBefore.Preference}
                'PreferenceStaging' {$ownerCorrect=$same -and (SettingsAct $case.Section 'Choice') -ceq $case.Value -and $ownerAfter.Preference -ceq $ownerBefore.Preference}
                'Selection' {
                    $field=if($case.Action -ceq 'SelectFamily'){'Family'}else{'Selected'}
                    $ownerCorrect=$same -and (SettingsAct $case.Section $field) -ceq $case.Value -and (SettingsAct $case.Section 'Request') -ceq $requestBefore
                }
            }
            Check ('SettingsActivity.'+$case.Id+'.IndependentOwnerResult') $ownerCorrect
            if(-not $ownerCorrect){throw 'Existing Settings owner fixture invalid; missing observations are the expected RED.'}
            AssertSettingsPair $Fixture $before $case $sequence $ordinal
            if(-not $captured.ContainsKey($case.Section)){
                $capturePins=ActivityPins
                $captureConfig=(Get-FileHash -LiteralPath $Fixture.Config).Hash
                $title=if($case.Section -ceq 'Operations'){'Event Tracking Settings'}else{'invSys Settings'}
                CaptureOwnedFormByCaptionEvidence $title ('settings-activity-'+$case.Section.ToLowerInvariant()+'.png')
                Check ('SettingsActivity.Visible.'+$case.Section) $true
                Check ('SettingsActivity.CaptureOnly.'+$case.Section) ((ActivityPins).Count -eq $capturePins.Count -and (PinsRetained $capturePins) -and
                    (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $captureConfig)
                $captured[$case.Section]=$true
            }
        }
        CloseSettingsEditors
        Check 'SettingsActivity.ActualStop' ((RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED')
        $closed=@(RecordingJournal $sequence|Where-Object RecordType -CEQ 'Close')
        Check 'SettingsActivity.TwentyFiveActionsCompleteStoppedRun' ($closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq 25 -and @($closed[0].Observations).Count -eq 50 -and (JournalChain $sequence 52))
        if($null -ne $settingsEvidence){$settingsEvidence.Main=$closed[0]}
        # The policy-save record is tested separately so interrupted evidence
        # cannot accidentally satisfy the preceding complete-run assertion.
        OpenSettingsEditor 'Tracking'
        $policySequence=StartSettingsRecording
        $before=@(Get-Slice4beActivityFiles $Fixture)
        $ownerBefore=SettingsOwnerState
        [void](SettingsAct 'Tracking' 'Save')
        $ownerAfter=SettingsOwnerState
        Check 'SettingsActivity.PolicySave.IndependentVersionAppend' ($ownerAfter.PolicyVersion -eq ($ownerBefore.PolicyVersion+1))
        $records=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {[IO.File]::ReadAllText($_)|ConvertFrom-Json})
        Check 'SettingsActivity.PolicySave.OnlyOriginalEligibleRequest' ($records.Count -eq 1 -and $records[0].ControlId -ceq 'ADMIN_TRACKING_SAVE' -and $records[0].OutcomeCode -ceq 'REQUESTED' -and $records[0].SequenceId -ceq $policySequence)
        $closed=@(RecordingJournal $policySequence|Where-Object RecordType -CEQ 'Close')
        Check 'SettingsActivity.PolicySave.InterruptedCannotConclude' ($closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Incomplete' -and $closed[0].ReasonCode -ceq 'POLICY_CHANGED')
        if($null -ne $settingsEvidence){$settingsEvidence.PolicySave=$closed[0]}
        . (Join-Path $PSScriptRoot 'Slice4beSettingsEditorSaveCases.ps1')
        Test-SettingsEditorSaveCases $Fixture $settingsEvidence
        . (Join-Path $PSScriptRoot 'Slice4beSettingsEditorSafety.ps1')
        Test-SettingsEditorSafety $Fixture $settingsEvidence
        if($CheckSettingsDiagnostics){Test-SettingsDiagnostics $Fixture $settingsEvidence}
        foreach($loss in @('SignedOut','Reauthenticated','OtherTarget')){
            foreach($section in @('Tracking','Detail','Preference','Operations')){
                foreach($action in @('Reset','Reload')){
                    CloseSettingsEditors
                    CloseRecordingViewer
                    SelectTarget $Fixture 'config-admin'
                    OpenRecordingViewer
                    OpenSettingsEditor $section
                    # Establish deliberately unsaved state through the real
                    # current-context callback before invalidating this editor.
                    switch($section){
                        'Tracking' {[void](SettingsAct $section 'DefaultView' 'Diagnostic')}
                        'Detail' {
                            [void](SettingsAct $section 'SelectFamily' 'Admin')
                            [void](SettingsAct $section 'SelectField' 'Reference')
                            [void](SettingsAct $section 'ShowField' 'False')
                        }
                        default {[void](SettingsAct $section 'Select' 'How-To')}
                    }
                    $read=if($section -cin @('Tracking','Detail')){'Request'}else{'Choice'}
                    $stagedBefore=SettingsAct $section $read
                    $before=ActivityPins
                    $otherBefore=@(Get-Slice4beActivityFiles $Other)
                    $configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
                    if($loss -ceq 'OtherTarget'){
                        SelectTarget $Other 'config-admin'
                    }else{
                        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
                        if($loss -ceq 'Reauthenticated'){SelectTarget $Fixture 'config-admin'}
                    }
                    $status=SettingsAct $section $action
                    $label='SettingsActivity.Context.'+$loss+'.'+$section+'.'+$action
                    Check ($label+'.PreservesHeldStaging') ((SettingsAct $section $read) -ceq $stagedBefore)
                    Check ($label+'.RequiresReopen') ($status.Contains('Session or warehouse changed.') -and $status.Contains('Reopen'))
                    Check ($label+'.NoCrossContextActivityOrConfigWrite') ((ActivityPins).Count -eq $before.Count -and (PinsRetained $before) -and
                        @(Get-Slice4beActivityFiles $Other).Count -eq $otherBefore.Count -and
                        (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore -and
                        (Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherConfig)
                }
            }
        }
        Check 'SettingsActivity.OtherWarehouseConfigPreserved' ((Get-FileHash -LiteralPath $Other.Config).Hash -ceq $otherConfig)
    } finally {
        CloseSettingsEditors
        CloseRecordingViewer
        foreach($book in @($excel.Workbooks)){
            if([string]::Equals($book.FullName,$Fixture.Config,[StringComparison]::OrdinalIgnoreCase)){$book.Close($false)}
        }
        [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        SelectTarget $Fixture 'config-admin'
        [void](Run 'invSys.Core.xlam' 'modConfig.Reload')
        if($null -ne $anchor){$anchor.Close($false)}
        $excel.Visible=$entryVisible
        Check 'SettingsActivity.GeneratedConfigRestoredExactly' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $originalConfig)
    }
}
