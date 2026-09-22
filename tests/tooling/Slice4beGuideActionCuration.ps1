# D18 direct curation: original unrecorded Admin actions, real publication,
# actual selection/Create/Save handlers. No synthetic guide or recording source.
function Test-GuideActionCuration($Fixture,$Other) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    function Pick([string]$Name,[string]$Action,[string]$Value='') {BoundControl $Name $Action $Value 'frmGuideActionPicker'}
    function Draft([string]$Name,[string]$Action,[string]$Value='') {BoundControl $Name $Action $Value 'frmActionPathGuide'}
    function OpenPicker {[void](BoundControl 'btnChooseGuideActions' 'Click' '' 'frmActionPaths')}
    function ToggleAction([string]$Id) {
        $ids=@((Pick 'lstGuideActions' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        $index=[Array]::IndexOf($ids,$Id)
        if($index -lt 0){return $false}
        return (Pick 'lstGuideActions' 'Toggle' ([string]$index)) -ceq 'SELECTED'
    }
    function OpenChosenDraft {
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
        [void](BoundLibrary 'Open');OpenPicker
        [void](ToggleAction $first.Attempt.ActivityId);[void](ToggleAction $second.Attempt.ActivityId)
        [void](Pick 'btnCreateSelectedGuide' 'Click')
    }
    function CurateCapture([string]$Stage,[string]$Caption='Choose tracked actions') {
        if($CaptureGuideEvidence){CaptureOwnedFormByCaptionEvidence $Caption ('guide-curation-'+$Stage+'.png')}
    }
    CloseRecordingViewer;SelectTarget $Fixture 'config-admin';SetRecordingPolicy $false
    $first=SaveRecordedSetting '801';$second=SaveRecordedSetting '802'
    if($first.Attempt.SequenceId -cne '' -or $second.Attempt.SequenceId -cne '' -or $first.Attempt.Ordinal -ne 0 -or $second.Attempt.Ordinal -ne 0){throw 'Actual unrecorded activity fixture failed; not curation RED.'}
    $published=Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest'
    $publication=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $valid=Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($publication,$Fixture.Warehouse)
    if($published -isnot [bool] -or -not $published -or $valid -isnot [bool] -or -not $valid){throw 'Actual curation publication fixture unavailable.'}
    $model=Get-Content -LiteralPath $publication -Raw|ConvertFrom-Json
    $selectedIds=@($first.Attempt.ActivityId,$second.Attempt.ActivityId)
    $groups=@($model.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -cin $selectedIds})
    if($groups.Count -ne 2 -or @($groups|ForEach-Object Lines).Count -ne 4){throw 'Actual publication omitted the selected source bodies.'}
    $observations=@(foreach($line in ($groups|ForEach-Object Lines)){
        $raw=[IO.File]::ReadAllText((Join-Path $activityRoot ($line.RecordId+'.json')))
        $marker=[regex]::Match($raw,',"ContentSha256":"([0-9a-f]{64})"\}$')
        if(-not $marker.Success){throw 'Original activity envelope lacks its existing digest.'}
        $body=$raw.Substring(0,$marker.Index)+'}';$sha=[Security.Cryptography.SHA256]::Create()
        try{$digest=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
        $original=$raw|ConvertFrom-Json;$record=$body|ConvertFrom-Json
        if($digest -cne $marker.Groups[1].Value -or @($record.PSObject.Properties).Count -ne 26 -or
           (ConvertTo-Json -InputObject $original -Depth 25 -Compress) -cne (ConvertTo-Json -InputObject $line -Depth 25 -Compress)){
            throw 'Published activity differs from its verified original body/envelope.'
        }
        $record
    })
    Check 'GuideCuration.Fixture.ActualUnrecordedActionsAndPublication' $true
    $sourcePins=BoundPins $Fixture.Root
    $activityBefore=ActivityPins;$trainingBefore=BoundPins $journalRoot
    $configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $publishBefore=Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest'
    $authorityBefore=Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest'
    $guideRoot=Join-Path $journalRoot 'Guides'
    try {
        OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Accepted Action Paths entry unavailable.'}
        Check 'GuideCuration.EntryIndependentOfRecording' ((BoundControl 'btnChooseGuideActions' 'State' '' 'frmActionPaths') -ceq 'True|True' -and (BoundControl 'lstActionPaths' 'Rows' '' 'frmActionPaths') -ceq '0')
        Check 'GuideCuration.EntryCaption' ((BoundControl 'btnChooseGuideActions' 'Label' '' 'frmActionPaths') -ceq 'Choose tracked actions')
        foreach($size in @('Minimum','Default','Larger','Restored')){
            Check ('GuideCuration.LibraryLayout.'+$size) ((BoundControl '' 'Fit' $size 'frmActionPaths') -ceq 'True')
            CurateCapture ('library-'+$size.ToLowerInvariant()) 'Action Paths'
        }
        OpenPicker
        $opened=(Pick '' 'Count') -ceq '1'
        Check 'GuideCuration.ActualEntryOpensOnePicker' $opened
        Check 'GuideCuration.EmptySelectionDisablesCreate' ($opened -and (Pick 'btnCreateSelectedGuide' 'State') -ceq 'True|False')
        $source=Pick 'lblGuideActionSource' 'Label'
        Check 'GuideCuration.SourceNamesExactLoadedPublication' ($opened -and $source.Contains([string]$model.PublicationId) -and $source.Contains([string]$model.ContentSha256))
        $ids=@((Pick 'lstGuideActions' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        $permittedIds=@($model.Groups|Where-Object Source -CEQ 'Activity'|ForEach-Object SourceId)
        Check 'GuideCuration.CandidatesArePublishedActionIdentities' ($opened -and $ids.Count -gt 1 -and @($ids|Where-Object {$_ -cnotin $permittedIds}).Count -eq 0 -and $first.Attempt.ActivityId -cin $ids -and $second.Attempt.ActivityId -cin $ids)
        Check 'GuideCuration.SelectFirstActualAction' (ToggleAction $first.Attempt.ActivityId)
        [void](Pick 'txtGuideActionSearch' 'Write' 'no-match-curation-filter')
        Check 'GuideCuration.SearchFiltersWithoutLosingSelection' ($opened -and (Pick 'lstGuideActions' 'Rows') -ceq '0' -and (Pick 'lblGuideActionStatus' 'Label') -match '(?i)1 selected' -and (Pick 'btnCreateSelectedGuide' 'State') -ceq 'True|True')
        [void](Pick 'txtGuideActionSearch' 'Write' 'ADMIN_SETTINGS_SAVE_VALUE')
        Check 'GuideCuration.SearchByControlRestoresSelectedIdentity' ($opened -and (Pick 'lstGuideActions' 'SelectedValues') -ceq ($first.Attempt.ActivityId+"`n"))
        Check 'GuideCuration.SelectSecondActualAction' (ToggleAction $second.Attempt.ActivityId)
        OpenPicker
        Check 'GuideCuration.RepeatedEntryReusesSelection' ($opened -and (Pick '' 'Count') -ceq '1' -and (Pick 'lblGuideActionStatus' 'Label') -match '(?i)2 selected')
        foreach($size in @('Minimum','Default','Larger','Restored')){
            Check ('GuideCuration.PickerLayout.'+$size) ($opened -and (Pick '' 'Fit' $size) -ceq 'True')
            if($opened){CurateCapture ('picker-'+$size.ToLowerInvariant())}
        }
        [void](Pick 'btnCreateSelectedGuide' 'Click')
        $drafted=(Draft '' 'Count') -ceq '1'
        Check 'GuideCuration.CreateUsesExistingGuideEditor' ($opened -and $drafted)
        Check 'GuideCuration.NoRecordedSequenceClaim' ($drafted -and (Draft 'lblGuideSource' 'Label') -match 'No recorded sequence; authored order is not execution evidence')
        $stepIds=@((Draft 'lstGuideSteps' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        Check 'GuideCuration.DistinctAuthoredStepIdentities' ($drafted -and $stepIds.Count -eq 2 -and @($stepIds|Select-Object -Unique).Count -eq 2 -and @($stepIds|Where-Object {$_ -cin $selectedIds}).Count -eq 0)
        Check 'GuideCuration.NoAutomaticExpectation' ($drafted -and (Draft 'lblGuideExpectationSummary' 'Label') -ceq 'Guide expectation: None')
        $evidence=Draft 'txtGuideEvidence' 'Text'
        Check 'GuideCuration.OriginalActionEvidenceDisplayed' ($drafted -and $evidence.Contains([string]$first.Attempt.ActivityId) -and $evidence.Contains([string]$second.Attempt.ActivityId) -and $evidence -match 'REQUESTED' -and $evidence -match 'COMPLETED')
        [void](Draft 'txtGuideName' 'Write' 'Direct tracked-action guide')
        [void](Draft 'txtGuideTags' 'Write' 'direct, unrecorded')
        [void](Draft 'txtGuideInstructions' 'Write' 'Authored order is not execution evidence.')
        [void](Pick 'btnCreateSelectedGuide' 'Click')
        Check 'GuideCuration.RepeatedCreateRetainsStagedText' ($drafted -and (Draft '' 'Count') -ceq '1' -and (Draft 'txtGuideName' 'Text') -ceq 'Direct tracked-action guide')
        [void](Draft 'lstGuideSteps' 'Select' '1');[void](Draft 'btnGuideStepUp' 'Click')
        [void](Draft 'txtGuideStepInstruction' 'Write' 'First authored instruction.')
        Check 'GuideCuration.ReorderRetainsStableStepIdsAndOriginalEvidence' ($drafted -and (Draft 'lstGuideSteps' 'Values') -ceq ($stepIds[1]+"`n"+$stepIds[0]+"`n") -and (Draft 'txtGuideEvidence' 'Text') -ceq $evidence)
        Check 'GuideCuration.NonSaveActionsPreserveEverySourceByte' (BoundSame $sourcePins (BoundPins $Fixture.Root))
        foreach($size in @('Minimum','Default','Larger','Restored')){
            Check ('GuideCuration.EditorLayout.'+$size) ($drafted -and (Draft '' 'Fit' $size) -ceq 'True')
            if($drafted){CurateCapture ('editor-'+$size.ToLowerInvariant()) 'Action Path guide'}
        }
        [void](Draft 'btnSaveGuide' 'Click')
        $files=@();if(Test-Path -LiteralPath $guideRoot){$files=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.1.json')}
        $guide=$null;if($files.Count -eq 1){$guide=ReadGuideExpectationRecord $files[0].FullName}
        $saved=$null -ne $guide
        Check 'GuideCuration.SavePublishesValidatedSchemaOne' ($drafted -and $saved -and $guide.Version -eq 1 -and $guide.Name -ceq 'Direct tracked-action guide')
        Check 'GuideCuration.SaveHasNoFabricatedRunOrExpectation' ($saved -and @($guide.SourceRun.PSObject.Properties).Count -eq 0 -and $guide.ExpectedConclusion.TerminalKind -ceq 'None' -and @($guide.ExpectedConclusion.Steps).Count -eq 0)
        Check 'GuideCuration.SavePreservesEverySelectedOriginalBody' ($saved -and (ConvertTo-Json -InputObject @($guide.Observations) -Depth 25 -Compress) -ceq (ConvertTo-Json -InputObject $observations -Depth 25 -Compress))
        Check 'GuideCuration.SavePreservesAuthoredOrderAndSourceIdentity' ($saved -and $guide.Steps[0].StepId -ceq $stepIds[1] -and $guide.Steps[1].StepId -ceq $stepIds[0] -and $guide.Steps[0].SourceActivityId -ceq $groups[1].SourceId -and $guide.Steps[1].SourceActivityId -ceq $groups[0].SourceId -and $guide.Steps[0].Instruction -ceq 'First authored instruction.')
        [void](Draft 'btnSaveGuide' 'Click')
        $revision=$null;if($saved){$revision=ReadGuideExpectationRecord (Join-Path $guideRoot ($guide.ActionPathId+'.2.json'))}
        Check 'GuideCuration.RepeatedSaveAppendsImmutableRevision' ($saved -and $null -ne $revision -and $revision.PreviousRecordId -ceq $guide.RecordId -and $revision.PreviousSha256 -ceq $guide.ContentSha256 -and @($revision.SourceRun.PSObject.Properties).Count -eq 0)
        $savedPins=BoundPins $journalRoot
        [void](Draft 'btnGuideExpectedConclusion' 'Click');[void](ToggleAction $first.Attempt.ActivityId)
        Check 'GuideCuration.SelectionChangeClosesDraftAndExpectation' ($drafted -and (Draft '' 'Count') -ceq '0' -and (BoundControl '' 'Count' '' 'frmActionPathExpectation') -ceq '0')
        [void](Pick 'btnCreateSelectedGuide' 'Click');[void](Pick 'btnCancelGuideActions' 'Click')
        Check 'GuideCuration.CancelClosesOwnedStagingPreservesSavedGuides' ($opened -and (Pick '' 'Count') -ceq '0' -and (Draft '' 'Count') -ceq '0' -and (BoundSame $savedPins (BoundPins $journalRoot)))
        Check 'GuideCuration.DoesNotCreateRecordingOrEvaluation' (@(Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json' -ErrorAction SilentlyContinue).Count -eq 0 -and (BoundPins (Join-Path $journalRoot 'Evaluations')).Count -eq 0)
        Check 'GuideCuration.SaveCancelPreserveActivityAndConfig' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore -and (PinsRetained $activityBefore) -and (ActivityPins).Count -eq $activityBefore.Count)
        OpenChosenDraft
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('Refresh',''))
        [void](Draft 'btnSaveGuide' 'Click')
        Check 'GuideCuration.ReloadedPublicationInvalidatesCapturedDraft' ($drafted -and (Draft 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (Draft 'btnSaveGuide' 'State') -cne 'True|True' -and (BoundSame $savedPins (BoundPins $journalRoot)))
        OpenChosenDraft;[void](Draft 'btnGuideExpectedConclusion' 'Click')
        [void](BoundControl 'btnClose' 'Click' '' 'frmActionPaths')
        Check 'GuideCuration.LibraryCloseDisposesPickerDraftAndExpectation' ($opened -and (Pick '' 'Count') -ceq '0' -and (Draft '' 'Count') -ceq '0' -and (BoundControl '' 'Count' '' 'frmActionPathExpectation') -ceq '0')
        OpenChosenDraft;CloseRecordingViewer
        Check 'GuideCuration.ViewerCloseDisposesOwnedStaging' ($drafted -and (Pick '' 'Count') -ceq '0' -and (Draft '' 'Count') -ceq '0' -and (BoundSame $savedPins (BoundPins $journalRoot)))
        OpenChosenDraft
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut');[void](Draft 'btnSaveGuide' 'Click')
        Check 'GuideCuration.SignOutClearsEvidenceAndCannotSave' ($drafted -and (Draft 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $savedPins (BoundPins $journalRoot)))
        OpenChosenDraft
        $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
        $authBytes=[IO.File]::ReadAllBytes($authPath)
        try {
            $auth=$excel.Workbooks.Open($authPath,0,$false)
            try {
                $caps=Table $auth 'tblCapabilities';$revoked=0
                foreach($row in $caps.ListRows){
                    if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-admin' -and
                       $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'ACTION_PATH_MAINT'){
                        $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$revoked++
                    }
                }
                if($revoked -ne 1){throw 'Guide maintenance fixture is not unique.'};$auth.Save()
            } finally {$auth.Close($false)}
            [void](Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($Fixture.Warehouse))
            $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
            $signed=Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'
            if($allowed -isnot [bool] -or $allowed -or $signed -isnot [bool] -or -not $signed){throw 'Permission loss fixture unavailable.'}
            [void](Draft 'btnSaveGuide' 'Click')
            Check 'GuideCuration.MaintenanceRevocationInvalidatesStagedDraft' ($drafted -and (Draft 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $savedPins (BoundPins $journalRoot)))
        } finally {CloseRecordingViewer;[IO.File]::WriteAllBytes($authPath,$authBytes);SelectTarget $Fixture 'config-admin'}
        OpenChosenDraft
        try {
            SaveGuideExpectationVisibility $false
            [void](Draft 'btnSaveGuide' 'Click')
            Check 'GuideCuration.PolicyLossInvalidatesWithoutDroppingSteps' ($drafted -and (Draft 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $savedPins (BoundPins $journalRoot)))
        } finally {CloseRecordingViewer;SaveGuideExpectationVisibility $true}
        OpenChosenDraft;SelectTarget $Other 'config-admin';[void](Draft 'btnSaveGuide' 'Click')
        Check 'GuideCuration.ContextLossCannotRetargetSave' ($drafted -and (Draft 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $savedPins (BoundPins $journalRoot)))
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer;[void](BoundLibrary 'Open')
        Check 'GuideCuration.ReaderCannotOpenPicker' ((BoundControl 'btnChooseGuideActions' 'State' '' 'frmActionPaths') -ceq 'True|False' -and (Pick '' 'Count') -ceq '0')
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin'
        $publicationBytes=[IO.File]::ReadAllBytes($publication)
        try {
            $altered=[Text.Encoding]::UTF8.GetString($publicationBytes)|ConvertFrom-Json
            $badGroup=@($altered.Groups|Where-Object SourceId -CEQ $first.Attempt.ActivityId)
            $badAttempt=@($badGroup[0].Lines|Where-Object OutcomeCode -CEQ 'REQUESTED')
            if($badGroup.Count -ne 1 -or $badAttempt.Count -ne 1){throw 'Invalid-digest fixture source is ambiguous.'}
            $badAttempt[0].ContentSha256='0'*64
            $altered.PSObject.Properties.Remove('ContentSha256')
            $body=ConvertTo-Json -InputObject $altered -Depth 30 -Compress
            if($body -match '[^\x00-\x7f]'){throw 'Diagnostic publication fixture must retain ASCII wire.'}
            $sha=[Security.Cryptography.SHA256]::Create()
            try{$digest=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            [IO.File]::WriteAllText($publication,$body.Substring(0,$body.Length-1)+',"ContentSha256":"'+$digest+'"}',[Text.UTF8Encoding]::new($false))
            $valid=Run 'invSys.Core.xlam' 'modInventoryViewerData.PublishedReadFixtureValidForTest' @($publication,$Fixture.Warehouse)
            if($valid -isnot [bool] -or -not $valid){throw 'Outer publication fixture must remain valid.'}
            OpenRecordingViewer;[void](BoundLibrary 'Open');OpenPicker
            $ids=@((Pick 'lstGuideActions' 'Values') -split "`n"|Where-Object {$_ -cne ''})
            Check 'GuideCuration.InvalidInnerDigestRejectedWithoutSourceMutation' ($opened -and $first.Attempt.ActivityId -cnotin $ids -and (Pick 'lblGuideActionStatus' 'Label') -match '1 unavailable' -and (PinsRetained $activityBefore))
            Check 'GuideCuration.InvalidInnerDigestDoesNotSuppressOtherValidAction' ($opened -and $second.Attempt.ActivityId -cin $ids -and (BoundSame $savedPins (BoundPins $journalRoot)))
        } finally {CloseRecordingViewer;[IO.File]::WriteAllBytes($publication,$publicationBytes)}
        if($saved -and $null -ne $revision){
            OpenRecordingViewer;[void](BoundLibrary 'Open');BoundOpen;BoundSelect $revision
        }
        Check 'GuideCuration.PublishedReaderLoadsCuratedGuideWithoutRun' ($saved -and $null -ne $revision -and (BoundControl 'txtPublishedInstructions' 'Text').Contains([string]$guide.Name) -and (BoundControl 'btnUseGuideForRun' 'State') -ceq 'True|False')
        if($saved -and $null -ne $revision){[void](BoundControl 'btnEditPublishedGuide' 'Click')}
        Check 'GuideCuration.PublishedEditorRestoresCuratedStepIds' ($saved -and $null -ne $revision -and (Draft 'lstGuideSteps' 'Values') -ceq ((@($revision.Steps.StepId) -join "`n")+"`n"))
        if($saved -and $null -ne $revision){
            [void](Draft 'txtGuideName' 'Write' 'Reopened directly curated guide');[void](Draft 'btnSaveGuide' 'Click')
            $third=ReadGuideExpectationRecord (Join-Path $guideRoot ($guide.ActionPathId+'.3.json'))
        } else {$third=$null}
        Check 'GuideCuration.PublishedRevisionRetainsEmptyRunAndOriginalBodies' ($saved -and $null -ne $third -and $third.PreviousRecordId -ceq $revision.RecordId -and $third.PreviousSha256 -ceq $revision.ContentSha256 -and @($third.SourceRun.PSObject.Properties).Count -eq 0 -and (ConvertTo-Json -InputObject @($third.Observations) -Depth 25 -Compress) -ceq (ConvertTo-Json -InputObject $observations -Depth 25 -Compress))
        Check 'GuideCuration.NoPublicationOrAuthorityRead' ((Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publishBefore -and (Run 'invSys.Operations.xlam' 'modTS_Shipments.PublishedReadAuthorityCallsForTest') -eq $authorityBefore)
        foreach($path in $sourcePins.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $sourcePins[$path] -and $path -cne $Fixture.Config){throw 'Curation changed a pre-existing source file.'}}
    } finally {CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
