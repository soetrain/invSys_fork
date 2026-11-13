# Focused restart setup through already accepted packaged controls. It complements
# the preserved full regression gate; it does not replace that evidence.
function Initialize-GuideRestartFixture($Fixture,[switch]$ForPublishedEdit) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    function FixtureControl([string]$Form,[string]$Name,[string]$Action,[string]$Value='') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($Form,$Name,$Action,$Value))
    }
    function FixtureRecording([string]$First,[string]$Second) {
        if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Actual fixture recording cannot start.'}
        $firstAction=SaveRecordedSetting $First;$secondAction=SaveRecordedSetting $Second
        if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED'){throw 'Actual fixture recording cannot stop.'}
        $entries=@(RecordingJournal $firstAction.Attempt.SequenceId|Sort-Object Version)
        $last=if($entries.Count){$entries[-1]}else{$null}
        $observations=if($last){@($last.Observations)}else{@()}
        $chainValid=JournalChain $firstAction.Attempt.SequenceId 14
        $firstValid=HasSequence $firstAction 3
        $secondValid=HasSequence $secondAction 6
        if($entries.Count -ne 14 -or -not $chainValid -or -not $firstValid -or -not $secondValid -or
           -not $last -or $last.Lifecycle -cne 'Stopped' -or $last.ActionCount -ne 6 -or $observations.Count -ne 12){
            if($reportRoot -and (Test-Path -LiteralPath $reportRoot)){
                # Structural facts only: no IDs, entered values, paths or event payloads.
                $facts=[ordered]@{EntryCount=$entries.Count;ChainValid=$chainValid;
                    FirstSaveOrdinal=$firstAction.Attempt.Ordinal;SecondSaveOrdinal=$secondAction.Attempt.Ordinal;
                    FirstSequenceValid=$firstValid;SecondSequenceValid=$secondValid;
                    Lifecycle=if($last){$last.Lifecycle}else{''};
                    ActionCount=if($last){$last.ActionCount}else{$null};ObservationCount=$observations.Count;
                    ControlOutcomes=@($observations|Group-Object ControlId,OutcomeCode|ForEach-Object {
                        [ordered]@{ControlId=$_.Group[0].ControlId;OutcomeCode=$_.Group[0].OutcomeCode;Count=$_.Count}
                    })}
                [IO.File]::WriteAllText((Join-Path $reportRoot 'guide-restart-journal-shape.json'),
                    ($facts|ConvertTo-Json -Depth 4),[Text.UTF8Encoding]::new($false))
            }
            throw 'Actual restart journal fixture is invalid; not behavioral RED.'
        }
        return $last
    }
    SelectTarget $Fixture 'config-admin'
    $published=Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest'
    $publication=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $valid=Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($publication,$Fixture.Warehouse)
    if($published -isnot [bool] -or -not $published -or $valid -isnot [bool] -or -not $valid){throw 'Actual publication fixture unavailable.'}
    Check 'GuidePresentation.RestartFixture.ActualPublicationValid' $true
    SetRecordingPolicy $true;OpenRecordingViewer
    $source=FixtureRecording '701' '702'
    Check 'GuidePresentation.RestartFixture.ActualSourceJournalValid' $true
    if((Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Open','')) -cne 'DELIVERED' -or
       (Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Select',$source.ActionPathId)) -cne 'SELECTED'){
        throw 'Actual source library selection unavailable.'
    }
    if((FixtureControl 'frmActionPaths' 'btnCreateGuide' 'Click') -cne 'DELIVERED'){throw 'Actual Create guide unavailable.'}
    [void](FixtureControl 'frmActionPathGuide' 'txtGuideName' 'Write' 'Preference restart guide')
    [void](FixtureControl 'frmActionPathGuide' 'txtGuideInstructions' 'Write' 'Authored instructions stay separate from observed actions.')
    if($ForPublishedEdit){
        [void](FixtureControl 'frmActionPathGuide' 'btnGuideExpectedConclusion' 'Click')
        foreach($retry in @('True','False')){
            [void](FixtureControl 'frmActionPathExpectation' 'cboExpectedControl' 'Write' 'ADMIN_SETTINGS_SAVE_VALUE')
            [void](FixtureControl 'frmActionPathExpectation' 'cboExpectedOutcome' 'Write' 'COMPLETED')
            [void](FixtureControl 'frmActionPathExpectation' 'chkExpectedRetry' 'Check' $retry)
            if((FixtureControl 'frmActionPathExpectation' 'btnAddExpectedStep' 'Click') -cne 'DELIVERED'){throw 'Actual expected-step fixture unavailable.'}
        }
        [void](FixtureControl 'frmActionPathExpectation' 'cboTerminalStep' 'Select' '1')
        [void](FixtureControl 'frmActionPathExpectation' 'cboTerminalKind' 'Write' 'CommandCompleted')
        if((FixtureControl 'frmActionPathExpectation' 'btnUseExpectation' 'Click') -cne 'DELIVERED'){throw 'Actual guide expectation use unavailable.'}
    }
    $guideRoot=Join-Path $journalRoot 'Guides'
    $before=@();if(Test-Path -LiteralPath $guideRoot){$before=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json'|ForEach-Object FullName)}
    if((FixtureControl 'frmActionPathGuide' 'btnSaveGuide' 'Click') -cne 'DELIVERED'){throw 'Actual Save guide unavailable.'}
    $fresh=@();if(Test-Path -LiteralPath $guideRoot){$fresh=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json'|Where-Object {$_.FullName -cnotin $before})}
    if($fresh.Count -ne 1){throw 'Actual Save guide did not publish exactly one version.'}
    $text=[IO.File]::ReadAllText($fresh[0].FullName);$guide=$text|ConvertFrom-Json
    $match=[regex]::Match($text,',"ContentSha256":"([0-9a-f]{64})"\}$')
    if(-not $match.Success -or $fresh[0].Length -gt 1048576 -or $fresh[0].Length -ne $text.Length -or $text -match '[^\x00-\x7f]'){throw 'Guide wire fixture invalid.'}
    $body=$text.Substring(0,$match.Index)+'}';$sha=[Security.Cryptography.SHA256]::Create()
    try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
    if($guide.ContentSha256 -cne $hash -or $guide.RecordKind -cne 'Guide' -or $guide.Version -ne 1 -or
       $guide.SourceRun.ActionPathId -cne $source.ActionPathId -or $guide.SourceRun.ContentSha256 -cne $source.ContentSha256 -or
       $guide.ActionPathId -ceq $source.ActionPathId -or @($guide.Steps).Count -ne 6 -or
       @($guide.Observations).Count -ne 12){
        if($reportRoot -and (Test-Path -LiteralPath $reportRoot)){
            $facts=[ordered]@{GuideSchema=$guide.SchemaVersion;Version=$guide.Version;
                RecordKind=$guide.RecordKind;HashValid=($guide.ContentSha256 -ceq $hash);
                SourceRunIdMatches=($guide.SourceRun.ActionPathId -ceq $source.ActionPathId);
                SourceRunHashMatches=($guide.SourceRun.ContentSha256 -ceq $source.ContentSha256);
                DistinctGuideId=($guide.ActionPathId -cne $source.ActionPathId);
                Steps=@($guide.Steps).Count;Observations=@($guide.Observations).Count}
            [IO.File]::WriteAllText((Join-Path $reportRoot 'guide-restart-provenance-shape.json'),
                ($facts|ConvertTo-Json -Depth 3),[Text.UTF8Encoding]::new($false))
        }
        throw 'Actual saved guide provenance fixture invalid.'
    }
    $expectedKind=if($ForPublishedEdit){'CommandCompleted'}else{'None'}
    if($guide.ExpectedConclusion.TerminalKind -cne $expectedKind){throw 'Actual guide expectation fixture invalid.'}
    Check 'GuidePresentation.RestartFixture.ActualGuideSaveRetainsExactSource' $true
    $firstGuide=$guide
    if($ForPublishedEdit){
        . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
        if(@($guide.ExpectedConclusion.Steps).Count -ne 2){throw 'Actual first guide must contain two authored expected steps.'}
        [void](FixtureControl 'frmActionPathGuide' 'btnGuideExpectedConclusion' 'Click')
        [void](FixtureControl 'frmActionPathExpectation' 'lstExpectedSteps' 'Select' '0')
        [void](FixtureControl 'frmActionPathExpectation' 'btnRemoveExpectedStep' 'Click')
        [void](FixtureControl 'frmActionPathExpectation' 'cboTerminalStep' 'Select' '0')
        [void](FixtureControl 'frmActionPathExpectation' 'cboTerminalKind' 'Write' 'CommandCompleted')
        [void](FixtureControl 'frmActionPathExpectation' 'btnUseExpectation' 'Click')
        [void](FixtureControl 'frmActionPathGuide' 'btnSaveGuide' 'Click')
        $guide=ReadGuideExpectationRecord (Join-Path $guideRoot ($firstGuide.ActionPathId+'.2.json'))
        if($null -eq $guide -or $guide.PreviousRecordId -cne $firstGuide.RecordId -or $guide.PreviousSha256 -cne $firstGuide.ContentSha256 -or
           @($guide.ExpectedConclusion.Steps).Count -ne 1 -or $guide.ExpectedConclusion.Steps[0].StepId -cne $firstGuide.ExpectedConclusion.Steps[1].StepId -or @($guide.Steps).Count -ne 6){throw 'Actual second guide fixture invalid.'}
        Check 'GuideEditFixture.ActualSecondVersionRetainsPredecessorAndExplicitIntent' $true
    }
    [void](FixtureControl 'frmActionPathGuide' 'btnCancelGuide' 'Click')
    CloseRecordingViewer;OpenRecordingViewer
    $observed=FixtureRecording '703' '704'
    if($observed.ActionPathId -ceq $source.ActionPathId){throw 'Restart observed run must be separate from the guide source.'}
    Check 'GuidePresentation.RestartFixture.ActualDifferentObservedJournalValid' $true
    CloseRecordingViewer
    $script:guidePresentationRestartFixture=[pscustomobject]@{Fixture=$Fixture;FirstGuide=$firstGuide;Guide=$guide;Observed=$observed}
}
