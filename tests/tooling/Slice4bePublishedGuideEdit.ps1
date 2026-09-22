# D18 exact-version published guide editing through the actual packaged controls.
# Called inside Test-GuideEvaluation; its exact-reader helpers remain in scope.
function Test-PublishedGuideEditFocused($Fixture,$Other,$State) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    $guideRoot=Join-Path $journalRoot 'Guides'
    Test-PublishedGuideEdit $Fixture $Other $State.FirstGuide $State.Guide $State.Observed
}

function Test-PublishedGuideEdit($Fixture,$Other,$FirstGuide,$SecondGuide,$Observed) {
    function EditControl([string]$Name,[string]$Action,[string]$Value='') {
        BoundControl $Name $Action $Value 'frmActionPathGuide'
    }
    function EditExpected([string]$Name,[string]$Action,[string]$Value='') {
        BoundControl $Name $Action $Value 'frmActionPathExpectation'
    }
    function EditSelected {BoundControl 'btnEditPublishedGuide' 'Click'}
    function EditExact($Guide) {BoundSelect $Guide;[void](EditSelected)}
    function EditCapture([string]$Stage,[string]$Caption='Action Path guide') {
        if($CaptureGuideEvidence){CaptureOwnedFormByCaptionEvidence $Caption ('guide-edit-'+$Stage+'.png')}
    }
    function PriorFilesPreserved($Pins) {
        foreach($path in $Pins.Keys){if(-not (Test-Path -LiteralPath $path) -or (Get-FileHash -LiteralPath $path).Hash -cne $Pins[$path]){return $false}}
        return $true
    }
    CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
    if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Accepted recording library is unavailable.'}
    BoundOpen
    $before=BoundPins $journalRoot;$activityBefore=ActivityPins
    $configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $publicationBefore=Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest'
    if($publicationBefore -isnot [int]){throw 'Accepted publication counter is unavailable.'}
    $resultsBefore=BoundPins (Join-Path $journalRoot 'Evaluations')
    try {
        Check 'GuideEdit.NoSelectionDisablesEntry' ((BoundControl 'btnEditPublishedGuide' 'State') -ceq 'True|False')
        BoundSelect $SecondGuide
        $present=(BoundControl 'btnEditPublishedGuide' 'State') -ceq 'True|True'
        Check 'GuideEdit.PermittedEntryWithoutObservedRun' ($present -and (BoundControl 'btnUseGuideForRun' 'State') -ceq 'True|False')
        Check 'GuideEdit.ApprovedCaption' ((BoundControl 'btnEditPublishedGuide' 'Label') -ceq 'Edit guide')
        foreach($layout in @('Minimum','Default','Larger','Restored')){
            Check ('GuideEdit.ReaderLayout.'+$layout) ((BoundControl '' 'Fit' $layout) -ceq 'True')
            EditCapture ('reader-'+$layout.ToLowerInvariant()) 'Published guides'
        }
        $opened=(EditSelected) -ceq 'DELIVERED' -and (EditControl '' 'Count') -ceq '1'
        Check 'GuideEdit.ActualHandlerOpensExistingGuideEditor' $opened
        Check 'GuideEdit.ReopenRestoresSavedName' ($opened -and (EditControl 'txtGuideName' 'Text') -ceq $SecondGuide.Name)
        Check 'GuideEdit.ReopenRetainsDraftNotice' ($opened -and (EditControl 'lblGuideStatus' 'Label') -match '^Draft only\.')
        $ids=(@($SecondGuide.Steps.StepId) -join "`n")+"`n"
        Check 'GuideEdit.ReopenRestoresOriginalAuthoredStepIds' ($opened -and (EditControl 'lstGuideSteps' 'Values') -ceq $ids)
        $source=EditControl 'lblGuideSource' 'Label'
        Check 'GuideEdit.SourceNamesExactGuideVersion' ($opened -and $source.Contains([string]$SecondGuide.ActionPathId) -and $source -match '(?i)version\s*:?\s*2\b')
        [void](EditControl 'btnGuideExpectedConclusion' 'Click')
        $expectedIds=(@($SecondGuide.ExpectedConclusion.Steps.StepId) -join "`n")+"`n"
        Check 'GuideEdit.ReopenRestoresExplicitExpectationIdsAndConclusion' ($opened -and (EditExpected 'lstExpectedSteps' 'Values') -ceq $expectedIds -and (EditExpected 'cboTerminalKind' 'Selected') -ceq $SecondGuide.ExpectedConclusion.TerminalKind)
        [void](EditExpected 'btnCancelExpectation' 'Click')
        Check 'GuideEdit.OpenDoesNotWriteTrainingOrConfig' ((BoundSame $before (BoundPins $journalRoot)) -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore)
        [void](EditControl 'txtGuideName' 'Write' 'Reopened authored guide')
        [void](EditControl 'txtGuideTags' 'Write' 'reopened, retained order')
        [void](EditControl 'txtGuideInstructions' 'Write' 'Authored revision; original observations remain separate.')
        [void](EditSelected)
        Check 'GuideEdit.RepeatedEntryReusesStagedEditor' ($opened -and (EditControl '' 'Count') -ceq '1' -and (EditControl 'txtGuideName' 'Text') -ceq 'Reopened authored guide')
        if((BoundLibrary 'Select' $Observed.ActionPathId) -cne 'SELECTED'){throw 'Separate observed-run fixture unavailable.'}
        Check 'GuideEdit.ObservedRunSelectionDoesNotReplaceAuthoredDraft' ($opened -and (EditControl 'txtGuideName' 'Text') -ceq 'Reopened authored guide' -and (EditControl 'lstGuideSteps' 'Values') -ceq $ids -and (EditControl 'lblGuideSource' 'Label') -ceq $source)
        if(@($SecondGuide.Steps).Count -lt 2){throw 'Existing guide fixture needs two authored steps.'}
        [void](EditControl 'lstGuideSteps' 'Select' '1')
        [void](EditControl 'txtGuideStepInstruction' 'Write' 'Retained instruction after reopening.')
        [void](EditControl 'btnGuideStepUp' 'Click')
        $reordered=@($SecondGuide.Steps.StepId);$swap=$reordered[0];$reordered[0]=$reordered[1];$reordered[1]=$swap
        Check 'GuideEdit.ReorderPreservesStepIds' ($opened -and (EditControl 'lstGuideSteps' 'Values') -ceq (($reordered -join "`n")+"`n"))
        foreach($layout in @('Minimum','Default','Larger','Restored')){
            Check ('GuideEdit.Layout.'+$layout) ($opened -and (EditControl '' 'Fit' $layout) -ceq 'True')
            if($opened){EditCapture $layout.ToLowerInvariant()}
        }
        [void](EditControl 'btnSaveGuide' 'Click')
        $third=ReadGuideExpectationRecord (Join-Path $guideRoot ($SecondGuide.ActionPathId+'.3.json'))
        $saved=$null -ne $third
        Check 'GuideEdit.SaveAppendsSameGuideVersionThree' ($opened -and $saved -and $third.ActionPathId -ceq $SecondGuide.ActionPathId -and $third.Version -eq 3 -and $third.RecordId -cne $SecondGuide.RecordId)
        Check 'GuideEdit.RevisionLinksExactPredecessor' ($opened -and $saved -and $third.PreviousRecordId -ceq $SecondGuide.RecordId -and $third.PreviousSha256 -ceq $SecondGuide.ContentSha256)
        Check 'GuideEdit.RevisionPreservesSourceObservationsAndProvenance' ($opened -and $saved -and ($third.Observations|ConvertTo-Json -Depth 20 -Compress) -ceq ($SecondGuide.Observations|ConvertTo-Json -Depth 20 -Compress) -and ($third.SourceRun|ConvertTo-Json -Depth 12 -Compress) -ceq ($SecondGuide.SourceRun|ConvertTo-Json -Depth 12 -Compress))
        Check 'GuideEdit.AuthoredReorderDoesNotRewriteExpectation' ($opened -and $saved -and ($third.ExpectedConclusion|ConvertTo-Json -Depth 12 -Compress) -ceq ($SecondGuide.ExpectedConclusion|ConvertTo-Json -Depth 12 -Compress))
        Check 'GuideEdit.RevisionRetainsAuthoredIdsOrderAndText' ($opened -and $saved -and (@($third.Steps.StepId) -join '|') -ceq ($reordered -join '|') -and $third.Steps[0].Instruction -ceq 'Retained instruction after reopening.')
        $key=[string]$SecondGuide.ActionPathId+'|2|'+[string]$SecondGuide.ContentSha256
        Check 'GuideEdit.SaveDoesNotRetargetReaderSelection' ($opened -and (BoundControl 'lstPublishedGuides' 'Selected') -ceq $key)
        [void](EditControl 'btnSaveGuide' 'Click')
        $fourth=ReadGuideExpectationRecord (Join-Path $guideRoot ($SecondGuide.ActionPathId+'.4.json'))
        Check 'GuideEdit.SubsequentSaveAppendsFromOwnLastVersion' ($opened -and $saved -and $null -ne $fourth -and $fourth.PreviousRecordId -ceq $third.RecordId -and $fourth.PreviousSha256 -ceq $third.ContentSha256)
        [void](EditControl 'btnCancelGuide' 'Click')
        Check 'GuideEdit.CancelClosesEditor' ($opened -and (EditControl '' 'Count') -ceq '0')
        [void](BoundControl 'btnRefreshGuides' 'Click')
        $reopen=$SecondGuide;if($saved){$reopen=$third}
        EditExact $reopen
        Check 'GuideEdit.ClosedEditorRestoresAuthoredFields' ($opened -and $saved -and (EditControl 'txtGuideName' 'Text') -ceq 'Reopened authored guide' -and (EditControl 'txtGuideTags' 'Text') -ceq 'reopened, retained order' -and (EditControl 'txtGuideInstructions' 'Text') -ceq 'Authored revision; original observations remain separate.')
        [void](EditControl 'lstGuideSteps' 'Select' '0')
        Check 'GuideEdit.ClosedEditorRestoresStepInstruction' ($opened -and $saved -and (EditControl 'txtGuideStepInstruction' 'Text') -ceq 'Retained instruction after reopening.')
        if($opened){EditCapture 'reopened'}
        $published=BoundPins $journalRoot
        [void](EditControl 'txtGuideName' 'Write' 'Unsaved change discarded')
        [void](EditControl 'btnCancelGuide' 'Click')
        EditExact $reopen
        Check 'GuideEdit.CancelDiscardsUnsavedTextWithoutDeletingVersions' ($opened -and $saved -and (EditControl 'txtGuideName' 'Text') -ceq 'Reopened authored guide' -and (BoundSame $published (BoundPins $journalRoot)))
        [void](EditControl 'btnCancelGuide' 'Click')
        EditExact $FirstGuide
        [void](EditControl 'txtGuideName' 'Write' 'Conflicting draft remains staged')
        [void](EditControl 'btnSaveGuide' 'Click')
        Check 'GuideEdit.OlderVersionConflictRetainsDraftAndAllVersions' ($opened -and (EditControl 'lblGuideStatus' 'Label') -match '(?i)version conflict' -and (EditControl 'txtGuideName' 'Text') -ceq 'Conflicting draft remains staged' -and (BoundSame $published (BoundPins $journalRoot)))
        if($opened){EditCapture 'conflict'}
        [void](EditControl 'btnGuideExpectedConclusion' 'Click')
        BoundSelect $SecondGuide
        Check 'GuideEdit.ChangedSelectionClosesDraftAndExpectationEditor' ($opened -and (EditControl '' 'Count') -ceq '0' -and (EditExpected 'lstExpectedSteps' 'Rows') -cin @('0','MISSING'))
        EditExact $FirstGuide
        [void](EditControl 'btnGuideExpectedConclusion' 'Click')
        [void](BoundControl 'btnCloseGuides' 'Click')
        Check 'GuideEdit.ReaderCloseDisposesOwnedDraftAndExpectationEditor' ($opened -and (EditControl '' 'Count') -ceq '0' -and (EditExpected 'lstExpectedSteps' 'Rows') -cin @('0','MISSING'))
        BoundOpen;EditExact $FirstGuide
        $path=Join-Path $guideRoot ($FirstGuide.ActionPathId+'.1.json');$bytes=[IO.File]::ReadAllBytes($path)
        try {
            [IO.File]::WriteAllText($path,'{}',[Text.UTF8Encoding]::new($false))
            [void](EditControl 'btnSaveGuide' 'Click')
            Check 'GuideEdit.ChangedGuideBytesInvalidateDraftAndEvidence' ($opened -and (EditControl 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (EditControl 'btnSaveGuide' 'State') -cne 'True|True')
        } finally {[IO.File]::WriteAllBytes($path,$bytes)}
        Check 'GuideEdit.EditSaveCancelPreserveConfigAndActivity' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore -and (PinsRetained $activityBefore) -and (ActivityPins).Count -eq $activityBefore.Count)
        CloseRecordingViewer;OpenRecordingViewer
        [void](BoundLibrary 'Open');BoundOpen;EditExact $FirstGuide
        try {
            SaveGuideExpectationVisibility $false
            [void](EditControl 'btnSaveGuide' 'Click')
            Check 'GuideEdit.CurrentPolicyLossClearsDraftWithoutDroppingHiddenContent' ($opened -and (EditControl 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $published (BoundPins $journalRoot)))
            [void](BoundControl 'btnRefreshGuides' 'Click')
            Check 'GuideEdit.RestrictedEntryHasReason' ((BoundControl 'btnEditPublishedGuide' 'State') -ceq 'True|False' -and (BoundControl 'lblPublishedGuideStatus' 'Label') -match '(?i)Hidden by policy: editing')
        } finally {CloseRecordingViewer;SaveGuideExpectationVisibility $true}
        OpenRecordingViewer;[void](BoundLibrary 'Open');BoundOpen;EditExact $FirstGuide
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
                if($revoked -ne 1){throw 'Guide maintenance fixture is not unique.'}
                $auth.Save()
            } finally {$auth.Close($false)}
            [void](Run 'invSys.Core.xlam' 'modAuth.LoadAuth' @($Fixture.Warehouse))
            $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
            $signed=Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'
            if($allowed -isnot [bool] -or $allowed -or $signed -isnot [bool] -or -not $signed){throw 'Permission loss fixture unavailable.'}
            [void](EditControl 'btnSaveGuide' 'Click')
            Check 'GuideEdit.MaintenanceRevocationInvalidatesPublishedDraft' ($opened -and (EditControl 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $published (BoundPins $journalRoot)))
        } finally {
            CloseRecordingViewer;[IO.File]::WriteAllBytes($authPath,$authBytes);SelectTarget $Fixture 'config-admin'
        }
        OpenRecordingViewer;[void](BoundLibrary 'Open');BoundOpen;EditExact $FirstGuide
        SelectTarget $Other 'config-admin'
        [void](EditControl 'btnSaveGuide' 'Click')
        Check 'GuideEdit.ContextLossCannotRetargetSave' ($opened -and (EditControl 'txtGuideEvidence' 'Text') -cin @('','MISSING') -and (BoundSame $published (BoundPins $journalRoot)))
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        [void](BoundLibrary 'Open');[void](BoundLibrary 'Select' $Observed.ActionPathId);BoundOpen;BoundSelect $FirstGuide
        Check 'GuideEdit.ReaderCannotOpenEdit' ((BoundControl 'btnEditPublishedGuide' 'State') -ceq 'True|False' -and (EditSelected) -ceq 'DISABLED' -and (EditControl '' 'Count') -ceq '0')
        Check 'GuideEdit.DeniedEntryHasReason' ((BoundControl 'lblPublishedGuideStatus' 'Label') -match 'ACTION_PATH_MAINT')
        Check 'GuideEdit.ReaderRetainsPublishedReadAndUseAccess' ((BoundControl 'txtPublishedInstructions' 'Text').Contains([string]$FirstGuide.Name) -and (BoundControl 'btnUseGuideForRun' 'Click') -ceq 'DELIVERED')
        Check 'GuideEdit.AllPriorSourceAndGuideFilesRemainIdentical' (PriorFilesPreserved $before)
        Check 'GuideEdit.NoEvaluationOrBusinessPublication' ((BoundSame $resultsBefore (BoundPins (Join-Path $journalRoot 'Evaluations'))) -and (Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publicationBefore)
    } finally {CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
