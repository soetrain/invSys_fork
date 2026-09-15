# D18 explicit guide intent through the shared real expectation editor.
function Test-GuideExpectation($Fixture,$Other) {
    function GuideExpectationControl([string]$Name,[string]$Action,[string]$Value='', [string]$Form='frmActionPathExpectation') {
        [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($Form,$Name,$Action,$Value))
    }
    function GuideAuthor([string]$Name,[string]$Action,[string]$Value='') {
        GuideExpectationControl $Name $Action $Value 'frmActionPathGuide'
    }
    function OpenGuideExpectation {GuideAuthor 'btnGuideExpectedConclusion' 'Click'}
    function CaptureGuideExpectation([string]$Stage,[string]$Title='Expected steps and conclusion') {
        if(-not ($CaptureEvidence -or $CaptureGuideEvidence)){return}
        Initialize-SettingsCapture
        $handle=[InvSysSettingsCapture]::OwnedVisibleForm($Title,[IntPtr]$excel.Hwnd).ToInt64()
        CaptureOwnedFormEvidence $Title ('guide-expectation-'+$Stage.ToLowerInvariant()+'.png') $handle
        Check ('GuideExpectation.VisibleCapture.'+$Stage) $true
    }
    function AddGuideExpectedStep([string]$Retry='True') {
        [void](GuideExpectationControl 'cboExpectedControl' 'Write' 'ADMIN_SETTINGS_SAVE_VALUE')
        [void](GuideExpectationControl 'cboExpectedOutcome' 'Write' 'COMPLETED')
        [void](GuideExpectationControl 'chkExpectedRetry' 'Check' $Retry)
        GuideExpectationControl 'btnAddExpectedStep' 'Click'
    }
    function UseGuideExpectedSteps([string]$TerminalIndex) {
        [void](GuideExpectationControl 'cboTerminalStep' 'Select' $TerminalIndex)
        [void](GuideExpectationControl 'cboTerminalKind' 'Write' 'CommandCompleted')
        GuideExpectationControl 'btnUseExpectation' 'Click'
    }
    function GuideExpectationPins([string]$Root) {
        $pins=@{}
        if(Test-Path -LiteralPath $Root){foreach($file in Get-ChildItem -LiteralPath $Root -Recurse -File){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}}
        return $pins
    }
    function GuideExpectationSamePins($Before,$After) {
        if($Before.Count -ne $After.Count){return $false}
        foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $Before[$path] -cne $After[$path]){return $false}}
        return $true
    }
    function ReadGuideExpectationRecord([string]$Path) {
        if(-not (Test-Path -LiteralPath $Path -PathType Leaf)){return $null}
        try {
            $text=[IO.File]::ReadAllText($Path);$length=(Get-Item -LiteralPath $Path).Length
            if($text -match '[^\x00-\x7f]' -or $length -ne $text.Length -or $length -gt 1048576){return $null}
            $value=$text|ConvertFrom-Json
            $match=[regex]::Match($text,',"ContentSha256":"([0-9a-f]{64})"\}$')
            if(-not $match.Success -or $value.RecordKind -cne 'Guide' -or $value.SchemaVersion -ne 1 -or @($value.PSObject.Properties).Count -ne 24){return $null}
            $body=$text.Substring(0,$match.Index)+'}';$sha=[Security.Cryptography.SHA256]::Create()
            try{$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($body))).Replace('-','').ToLowerInvariant()}finally{$sha.Dispose()}
            if($value.ContentSha256 -cne $hash){return $null}
            return $value
        } catch {return $null}
    }
    function OpenAuthoredGuide {
        if([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Open','')) -cne 'DELIVERED'){throw 'Existing recording library cannot open; not guide-expectation RED.'}
        if([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Select',$source.ActionPathId)) -cne 'SELECTED'){throw 'Existing source cannot be selected; not guide-expectation RED.'}
        if((GuideExpectationControl 'btnCreateGuide' 'Click' '' 'frmActionPaths') -cne 'DELIVERED' -or (GuideAuthor '' 'Count') -cne '1'){throw 'Existing author fixture cannot open; not guide-expectation RED.'}
        [void](GuideAuthor 'txtGuideName' 'Write' 'Guide with an authored conclusion')
    }
    function SaveGuideExpectationVisibility([bool]$Visible) {
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            $result=Run 'invSys.Admin.xlam' 'TestD5Commands.PublishedReadVisibilityForTest' @($Visible)
            if($result -isnot [bool] -or -not $result){throw 'Actual visibility command fixture failed; not expectation RED.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    }
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    SetRecordingPolicy $true
    OpenRecordingViewer
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Recording fixture unavailable.'}
    $first=SaveRecordedSetting '660';$second=SaveRecordedSetting '661'
    # A real captured expectation must not become the new guide's definition.
    if((GuideExpectationControl 'btnRecordingExpectation' 'Click' '' 'frmInventoryViewer') -cne 'DELIVERED'){throw 'Accepted recording expectation entry unavailable.'}
    if((AddGuideExpectedStep) -cne 'DELIVERED' -or (GuideExpectationControl 'lstExpectedSteps' 'Rows') -cne '1'){throw 'Accepted recording expectation authoring unavailable.'}
    if((UseGuideExpectedSteps '0') -cne 'DELIVERED'){throw 'Accepted recording expectation cannot be staged.'}
    if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED'){throw 'Recording fixture cannot stop.'}
    $journal=@(RecordingJournal $first.Attempt.SequenceId|Sort-Object Version)
    if($journal.Count -ne 6 -or -not (JournalChain $first.Attempt.SequenceId 6)){throw 'Actual recording fixture is invalid.'}
    $source=$journal[-1]
    if($source.ExpectedConclusion.TerminalKind -cne 'CommandCompleted' -or @($source.ExpectedConclusion.Steps).Count -ne 1){throw 'Actual captured expectation fixture is missing.'}
    $guideRoot=Join-Path $journalRoot 'Guides'
    $priorGuides=GuideExpectationPins $guideRoot
    $evaluationBefore=GuideExpectationPins (Join-Path $journalRoot 'Evaluations')
    $sourcePins=@{};foreach($file in Get-ChildItem -LiteralPath $journalRoot -File){$sourcePins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
    $activityBefore=ActivityPins;$configBefore=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $publicationBefore=[long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest')
    OpenAuthoredGuide
    $observed=GuideAuthor 'txtGuideEvidence' 'Text'
    $analysisBefore=GuideExpectationControl 'lblExpectationSummary' 'Label' '' 'frmActionPaths'
    $authoredIds=@((GuideAuthor 'lstGuideSteps' 'Values') -split "`n"|Where-Object {$_ -cne ''})
    try {
        $available=(GuideAuthor 'btnGuideExpectedConclusion' 'State') -ceq 'True|True'
        Check 'GuideExpectation.ActualGuideEntryAvailable' $available
        Check 'GuideExpectation.InitialSummaryIsNoneDespiteCapturedIntent' ($available -and (GuideAuthor 'lblGuideExpectationSummary' 'Label') -ceq 'Guide expectation: None')
        $opened=(OpenGuideExpectation) -ceq 'DELIVERED'
        Check 'GuideExpectation.SharedEditorOpensInGuideScope' ($opened -and (GuideExpectationControl '' 'Count') -ceq '1' -and (GuideExpectationControl 'btnUseExpectation' 'Label') -ceq 'Use for this guide')
        Check 'GuideExpectation.HelpIdentifiesAuthoredGuideIntent' ($opened -and (GuideExpectationControl 'lblExpectationHelp' 'Label') -match '(?i)guide' -and (GuideExpectationControl 'lblExpectationHelp' 'Label') -notmatch '(?i)this recording|this evaluation')
        Check 'GuideExpectation.DoesNotInheritCapturedSteps' ($opened -and (GuideExpectationControl 'lstExpectedSteps' 'Rows') -ceq '0' -and (GuideExpectationControl 'cboTerminalKind' 'Selected') -ceq 'None')
        [void](OpenGuideExpectation)
        Check 'GuideExpectation.RepeatedEntryReusesOneEditor' ($opened -and (GuideExpectationControl '' 'Count') -ceq '1')
        [void](AddGuideExpectedStep)
        [void](GuideExpectationControl 'btnCancelExpectation' 'Click')
        [void](OpenGuideExpectation)
        Check 'GuideExpectation.CancelDiscardsOnlyPendingExpectationEdits' ($opened -and (GuideExpectationControl 'lstExpectedSteps' 'Rows') -ceq '0' -and (GuideAuthor 'lblGuideExpectationSummary' 'Label') -ceq 'Guide expectation: None')
        $retryDefault=GuideExpectationControl 'chkExpectedRetry' 'Text'
        [void](AddGuideExpectedStep)
        [void](AddGuideExpectedStep 'False')
        $stepIds=@((GuideExpectationControl 'lstExpectedSteps' 'Values') -split "`n"|Where-Object {$_ -cne ''})
        Check 'GuideExpectation.RegisteredRepeatedStepsCanBeAuthored' ($opened -and (GuideExpectationControl 'lstExpectedSteps' 'Rows') -ceq '2' -and $retryDefault -ceq 'True')
        foreach($size in @('Minimum','Default','Larger','Restored')){
            Check ('GuideExpectation.EditorLayout.'+$size) ($opened -and (GuideExpectationControl '' 'Fit' $size) -ceq 'True')
            if($opened){CaptureGuideExpectation $size}
        }
        $used=(UseGuideExpectedSteps '1') -ceq 'DELIVERED'
        Check 'GuideExpectation.UseUpdatesGuideSummary' ($opened -and $used -and (GuideAuthor 'lblGuideExpectationSummary' 'Label') -match '(?i)guide expectation.*2')
        if($opened -and $used){CaptureGuideExpectation 'Staged' 'Action Path guide'}
        [void](GuideExpectationControl 'btnPathRefresh' 'Click' '' 'frmActionPaths')
        Check 'GuideExpectation.StagingDoesNotChangeSelectedRunAnalysis' ($opened -and
            (GuideExpectationControl 'lblExpectationSummary' 'Label' '' 'frmActionPaths') -ceq $analysisBefore)
        Check 'GuideExpectation.StagingDoesNotPublishOrRewriteSources' ((GuideExpectationSamePins $priorGuides (GuideExpectationPins $guideRoot)) -and
            (GuideAuthor 'txtGuideEvidence' 'Text') -ceq $observed -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configBefore)
        [void](GuideAuthor 'btnSaveGuide' 'Click')
        $fresh=@(Get-ChildItem -LiteralPath $guideRoot -File -Filter '*.json'|Where-Object {-not $priorGuides.ContainsKey($_.FullName)})
        if($fresh.Count -ne 1){throw 'Existing immutable Save fixture failed; not guide-expectation RED.'}
        $saved=ReadGuideExpectationRecord $fresh[0].FullName
        if($null -eq $saved){throw 'Existing Save record integrity is invalid; not expectation RED.'}
        $firstHash=(Get-FileHash -LiteralPath $fresh[0].FullName).Hash
        $definition=$saved.ExpectedConclusion;$steps=@($definition.Steps)
        Check 'GuideExpectation.SavePublishesExplicitDefinition' ($opened -and $used -and $definition.SchemaVersion -eq 1 -and $definition.TerminalKind -ceq 'CommandCompleted' -and $steps.Count -eq 2)
        $identityValid=$opened -and $steps.Count -eq 2 -and $stepIds.Count -eq 2
        if($identityValid){
            $identityValid=$steps[0].StepId -ceq $stepIds[0] -and $steps[1].StepId -ceq $stepIds[1] -and $steps[0].StepId -cne $steps[1].StepId
            foreach($step in $steps){if($step.StepId -cin $authoredIds -or $step.StepId -ceq $source.ExpectedConclusion.Steps[0].StepId){$identityValid=$false}}
        }
        Check 'GuideExpectation.ExpectedIdsAreStableAndDistinctFromInstructionAndCapturedIds' $identityValid
        Check 'GuideExpectation.RetryChoicesAndExactTerminalPersist' ($steps.Count -eq 2 -and $steps[0].ControlId -ceq 'ADMIN_SETTINGS_SAVE_VALUE' -and $steps[1].ControlId -ceq 'ADMIN_SETTINGS_SAVE_VALUE' -and
            $steps[0].RequiredOutcome -ceq 'COMPLETED' -and $steps[1].RequiredOutcome -ceq 'COMPLETED' -and $steps[0].RetryAllowed -is [bool] -and $steps[0].RetryAllowed -and
            $steps[1].RetryAllowed -is [bool] -and -not $steps[1].RetryAllowed -and $definition.TerminalStepId -ceq $steps[1].StepId)
        Check 'GuideExpectation.SavePreservesOriginalObservationsAndInstructionIds' (($saved.Observations|ConvertTo-Json -Depth 12 -Compress) -ceq ($source.Observations|ConvertTo-Json -Depth 12 -Compress) -and
            ($saved.Steps.StepId -join '|') -ceq ($authoredIds -join '|'))
        [void](OpenGuideExpectation)
        $stagedIds=($stepIds -join "`n")+"`n"
        Check 'GuideExpectation.ReopeningLoadsStagedDefinition' ($opened -and (GuideExpectationControl 'lstExpectedSteps' 'Values') -ceq $stagedIds)
        [void](AddGuideExpectedStep)
        $pendingValues=GuideExpectationControl 'lstExpectedSteps' 'Values'
        $pendingThree=(GuideExpectationControl 'lstExpectedSteps' 'Rows') -ceq '3'
        $startedAnother=(RecordingControl 'Start Recording' 'Click') -ceq 'DELIVERED'
        Check 'GuideExpectation.StartingRecordingRetainsPendingGuideScope' ($opened -and $pendingThree -and $startedAnother -and
            (GuideExpectationControl 'lstExpectedSteps' 'Values') -ceq $pendingValues)
        $stoppedAnother=(RecordingControl 'Stop Recording' 'Click') -ceq 'DELIVERED'
        Check 'GuideExpectation.StoppingRecordingRetainsPendingGuideScope' ($opened -and $pendingThree -and $stoppedAnother -and
            (GuideExpectationControl 'lstExpectedSteps' 'Values') -ceq $pendingValues)
        $runOpened=GuideExpectationControl 'btnExpectedConclusion' 'Click' '' 'frmActionPaths'
        $capturedIds=[string]$source.ExpectedConclusion.Steps[0].StepId+"`n"
        Check 'GuideExpectation.SwitchToRunUsesSeparateCapturedScope' ($opened -and $runOpened -ceq 'DELIVERED' -and
            (GuideExpectationControl 'btnUseExpectation' 'Label') -ceq 'Use for this evaluation' -and
            (GuideExpectationControl 'lstExpectedSteps' 'Values') -ceq $capturedIds)
        [void](GuideExpectationControl 'btnCancelExpectation' 'Click')
        [void](OpenGuideExpectation)
        Check 'GuideExpectation.ScopeSwitchDropsPendingEditsOnly' ($opened -and (GuideExpectationControl 'lstExpectedSteps' 'Values') -ceq $stagedIds)
        [void](GuideExpectationControl 'lstExpectedSteps' 'Select' '0')
        [void](GuideExpectationControl 'btnRemoveExpectedStep' 'Click')
        [void](GuideExpectationControl 'btnCancelExpectation' 'Click')
        [void](OpenGuideExpectation)
        Check 'GuideExpectation.CancelRetainsPreviouslyStagedDefinition' ($opened -and (GuideExpectationControl 'lstExpectedSteps' 'Values') -ceq $stagedIds)
        [void](GuideExpectationControl 'lstExpectedSteps' 'Select' '0')
        [void](GuideExpectationControl 'btnRemoveExpectedStep' 'Click')
        [void](UseGuideExpectedSteps '0')
        [void](GuideAuthor 'btnSaveGuide' 'Click')
        $nextPath=Join-Path $guideRoot ($saved.ActionPathId+'.2.json')
        $next=ReadGuideExpectationRecord $nextPath
        if($null -eq $next){throw 'Existing second immutable Save fixture failed; not expectation RED.'}
        $nextSteps=@($next.ExpectedConclusion.Steps)
        Check 'GuideExpectation.NextVersionRetainsPriorDefinitionAndLinks' ($opened -and $nextSteps.Count -eq 1 -and $steps.Count -eq 2 -and $nextSteps[0].StepId -ceq $steps[1].StepId -and
            $next.PreviousRecordId -ceq $saved.RecordId -and $next.PreviousSha256 -ceq $saved.ContentSha256 -and (Get-FileHash -LiteralPath $fresh[0].FullName).Hash -ceq $firstHash)
        if($opened){CaptureGuideExpectation 'PublishedVersion2' 'Action Path guide'}
        $savedPins=GuideExpectationPins $guideRoot
        [void](OpenGuideExpectation);[void](AddGuideExpectedStep);[void](UseGuideExpectedSteps '1')
        [void](OpenGuideExpectation);[void](AddGuideExpectedStep)
        $pendingAtCancel=(GuideExpectationControl 'lstExpectedSteps' 'Rows') -ceq '3'
        [void](GuideAuthor 'btnCancelGuide' 'Click')
        Check 'GuideExpectation.GuideCancelPreservesPublishedDefinitions' ($opened -and (GuideExpectationSamePins $savedPins (GuideExpectationPins $guideRoot)))
        Check 'GuideExpectation.GuideCancelClearsPendingExpectationEditor' ($opened -and $pendingAtCancel -and (GuideExpectationControl 'lstExpectedSteps' 'Rows') -cin @('0','MISSING'))
        OpenAuthoredGuide
        [void](OpenGuideExpectation)
        Check 'GuideExpectation.FreshGuideDoesNotInheritAnotherDraft' ($opened -and (GuideExpectationControl 'lstExpectedSteps' 'Rows') -ceq '0' -and (GuideAuthor 'lblGuideExpectationSummary' 'Label') -ceq 'Guide expectation: None')
        [void](AddGuideExpectedStep)
        $hadPending=(GuideExpectationControl 'lstExpectedSteps' 'Rows') -ceq '1'
        try {
            SaveGuideExpectationVisibility $false
            [void](UseGuideExpectedSteps '0')
            Check 'GuideExpectation.CurrentPolicyInvalidatesPendingGuideIntent' ($opened -and $hadPending -and
                (GuideExpectationControl 'lstExpectedSteps' 'Rows') -cin @('0','MISSING') -and (GuideExpectationSamePins $savedPins (GuideExpectationPins $guideRoot)))
        } finally {
            CloseRecordingViewer
            SaveGuideExpectationVisibility $true
        }
        OpenRecordingViewer
        OpenAuthoredGuide
        [void](OpenGuideExpectation);[void](AddGuideExpectedStep)
        $hadPending=(GuideExpectationControl 'lstExpectedSteps' 'Rows') -ceq '1'
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
            $signedIn=Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'
            if($allowed -isnot [bool] -or $allowed -or $signedIn -isnot [bool] -or -not $signedIn){throw 'Maintenance loss is not isolated from sign-out.'}
            [void](UseGuideExpectedSteps '0')
            Check 'GuideExpectation.MaintenanceRevocationInvalidatesPendingIntent' ($opened -and $hadPending -and
                (GuideExpectationControl 'lstExpectedSteps' 'Rows') -cin @('0','MISSING') -and (GuideExpectationSamePins $savedPins (GuideExpectationPins $guideRoot)))
        } finally {
            CloseRecordingViewer
            [IO.File]::WriteAllBytes($authPath,$authBytes)
            SelectTarget $Fixture 'config-admin'
        }
        if([Convert]::ToBase64String([IO.File]::ReadAllBytes($authPath)) -cne [Convert]::ToBase64String($authBytes)){throw 'Auth fixture bytes were not restored.'}
        $restored=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
        if($restored -isnot [bool] -or -not $restored){throw 'Guide capability restoration failed.'}
        OpenRecordingViewer
        OpenAuthoredGuide
        [void](OpenGuideExpectation);[void](AddGuideExpectedStep)
        $hadPending=(GuideExpectationControl 'lstExpectedSteps' 'Rows') -ceq '1'
        SelectTarget $Other 'config-admin'
        [void](UseGuideExpectedSteps '0')
        Check 'GuideExpectation.TargetChangeInvalidatesPendingGuideIntent' ($opened -and $hadPending -and (GuideExpectationControl 'lstExpectedSteps' 'Rows') -cin @('0','MISSING') -and
            (GuideExpectationSamePins $savedPins (GuideExpectationPins $guideRoot)))
        $sourcesUnchanged=$true;foreach($path in $sourcePins.Keys){if((Get-FileHash -LiteralPath $path).Hash -cne $sourcePins[$path]){$sourcesUnchanged=$false}}
        Check 'GuideExpectation.AllActionsPreserveRecordedExpectationsAndActivity' ($sourcesUnchanged -and (PinsRetained $activityBefore) -and (ActivityPins).Count -eq $activityBefore.Count)
        Check 'GuideExpectation.NoBusinessPublicationOrEvaluationResult' ([long](Run 'invSys.Core.xlam' 'modWarehouseSync.PublishedReadPublishCallsForTest') -eq $publicationBefore -and
            (GuideExpectationSamePins $evaluationBefore (GuideExpectationPins (Join-Path $journalRoot 'Evaluations'))))
    } finally {
        CloseRecordingViewer
        SelectTarget $Fixture 'config-admin'
    }
}
