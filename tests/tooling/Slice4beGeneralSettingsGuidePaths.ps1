# Authored intent compared with independent recordings through packaged controls.
function Test-GeneralSettingsGuidePaths($Fixture,$Runs,$Steps,[switch]$CloseActions,[switch]$WorksheetScale,[switch]$SettingsNavigation,[switch]$Reconcile,[switch]$VerifyAddins,[switch]$AdminConsole,[switch]$Users,[switch]$Seed,[switch]$DirectoryActions,[switch]$RootActions,[switch]$CreateActions,[switch]$TesterActions,[switch]$AggregationActions) {
    if(@($CloseActions,$WorksheetScale,$SettingsNavigation,$Reconcile,$VerifyAddins,$AdminConsole,$Users,$Seed,$DirectoryActions,$RootActions,$CreateActions,$TesterActions,$AggregationActions|Where-Object{$_}).Count -gt 1){throw 'Choose one independent guide fixture.'}
    function Delivered([string]$Value){if($Value -cne 'DELIVERED'){throw 'Actual General guide control unavailable; not product RED.'}}
    function Author([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathGuide'}
    function View([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathView'}
    function Retained($Before,$After){foreach($path in $Before.Keys){if(-not $After.ContainsKey($path) -or $After[$path] -cne $Before[$path]){return $false}};return $true}
    $prefix=if($AggregationActions){'AggregationGuide'}elseif($TesterActions){'TesterGuide'}elseif($CreateActions){'CreateGuide'}elseif($RootActions){'RootGuide'}elseif($DirectoryActions){'DirectoryGuide'}elseif($Seed){'SeedGuide'}elseif($Users){'UsersGuide'}elseif($AdminConsole){'AdminConsoleGuide'}elseif($VerifyAddins){'AdminVerifyGuide'}elseif($Reconcile){'ReconcileGuide'}elseif($SettingsNavigation){'SettingsNavigationGuide'}elseif($WorksheetScale){'WorksheetScaleGuide'}elseif($CloseActions){'AdminCloseGuide'}else{'GeneralGuide'}
    $capturePrefix=if($AggregationActions){'aggregation-guide-'}elseif($TesterActions){'tester-guide-'}elseif($CreateActions){'create-guide-'}elseif($RootActions){'root-guide-'}elseif($DirectoryActions){'directory-guide-'}elseif($Seed){'seed-guide-'}elseif($Users){'users-guide-'}elseif($AdminConsole){'admin-console-guide-'}elseif($VerifyAddins){'admin-verify-guide-'}elseif($Reconcile){'reconcile-guide-'}elseif($SettingsNavigation){'settings-navigation-guide-'}elseif($WorksheetScale){'worksheet-scale-guide-'}elseif($CloseActions){'admin-close-guide-'}else{'general-guide-'}
    function Pass([string]$Name,[bool]$Value){Check ($prefix+'.'+$Name) $Value|Out-Host}
    $source=@($Runs|Where-Object Kind -CEQ 'Source')[0]
    $observed=@($Runs|Where-Object Kind -CEQ 'Observed')[0]
    $eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $publication=Get-Content $eventsPath -Raw|ConvertFrom-Json;$publicationHash=Hash $eventsPath
    $originalPins=BoundPins $journalRoot;$activityPins=ActivityPins
    $instruction="Use General Settings to inspect configuration and UOM choices, manage carriers, and save the connection option.`r`nCarrier and connection changes apply to this Windows user; they do not prove inventory changes.`r`nRead the observed result. Interrupted or denied actions cannot borrow completion from this guide."
    if($CloseActions){$instruction="Dismiss Settings with Close, reopen it, then use the window close button.`r`nUnsaved form edits are discarded. Dismissal proves neither a settings save nor Domain application.`r`nBoth gestures record the same Close control; the recording does not identify the gesture. Interrupted actions cannot borrow completion from this guide."}
    if($WorksheetScale){$instruction="With Designs disabled and the local worksheet recipe available, select the recipe and enter 200 percent. Click Apply Scale and inspect the prepared rows.`r`nClick Apply Scale again. The recipe quantity remains four, not eight; existing System_Key and custom fields are preserved.`r`nThis concludes two local staging commands, not inventory application. An interrupted recording cannot borrow completion from this guide."}
    if($SettingsNavigation){$instruction="Open Settings, select Event Tracking, Event Detail, Action Paths, Tracking, and General, then Close.`r`nKeep Settings open throughout the tab changes. These observations prove presentation and dismissal, not a settings save or Domain application.`r`nAn interrupted recording cannot borrow completion from this guide."}
    if($Reconcile){$instruction="Open Shipment Reconcile in the training warehouse and inspect the selected exact inventory entity's shipment.`r`nRead the current quantity, balance after shipment and proposed correction. Supply a repair narrative and confirm only the intended correction.`r`nA queued correction proves command completion, not inventory application. An interrupted recording cannot borrow completion from this guide."}
    if($VerifyAddins){$instruction="Use Verify Add-ins Published twice and read each result.`r`nSuccess means the five required files are present and nonempty. It does not validate hashes, promote a deployment or apply inventory events.`r`nAn interrupted recording cannot borrow completion from this guide."}
    if($AdminConsole){$instruction="Open Admin Console twice and review its status before continuing.`r`nThese observations prove console presentation. They do not prove inventory processing or Domain application.`r`nAn interrupted recording cannot borrow completion from this guide."}
    if($DirectoryActions){$instruction="Use View Warehouses twice and review the directory after each action.`r`nOpening the directory does not create a warehouse, change the selected target, or prove complete discovery.`r`nAn interrupted recording cannot borrow completion from this guide."}
    if($RootActions){$instruction="Use Add Warehouse Root twice with generated training folders and read each result.`r`nA confirmed preference and refreshed directory do not create a warehouse, select a target or apply inventory events.`r`nAn interrupted recording cannot borrow completion from this guide."}
    if($Users){$instruction="In a generated training warehouse, use Users & Roles to update a training account, deactivate it, then delete it.`r`nRead each owner result before continuing. SAVED proves Auth persistence, not inventory application.`r`nAn interrupted recording cannot borrow completion from this guide."}
    if($Seed){$instruction="In a generated training warehouse, seed built-in demo inventory, then seed the selected training CSV.`r`nRead each owner result and inspect the created inventory. Command completion does not itself assert published Domain evidence.`r`nAn interrupted recording cannot borrow completion from this guide."}
    if($TesterActions){$instruction="Use Setup to create two separate disposable environments in empty folders. Read each owning command's result before continuing.`r`nSupply credentials interactively; no entered values or destination paths are retained in this guide.`r`nCompletion proves the Tester command only, not inventory application or Training/replay eligibility. A refused or interrupted attempt cannot borrow completion from the source recording."}
    if($AggregationActions){$instruction="Aggregate the explicitly selected published snapshots twice. Read each owner result before continuing.`r`nComplete output is advisory; it does not apply inventory or grant authority over selected warehouses. Connection inputs and paths are not retained in the guide.`r`nA refusal, partial or indeterminate attempt cannot borrow completion from the source recording. Retained attempts require explicit review."}
    if($CreateActions){$instruction="Create two separate generated Training warehouses with initial publication enabled. Read each result before continuing.`r`nBoth local creation and publication must complete for each step. A partial publication does not undo local creation.`r`nThese are command results, not inventory application. A fresh failed or interrupted run cannot borrow completion from this guide."}
    try {
        Pass 'IndependentSourceAndObservedRun' ($source.Journal.ActionPathId -cne $observed.Journal.ActionPathId -and @($observed.Records|Where-Object {$_.ActivityId -cin @($source.Records.ActivityId)}).Count -eq 0)
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
        Delivered (BoundLibrary 'Open')
        if((BoundLibrary 'Select' $source.Journal.ActionPathId) -cne 'SELECTED'){throw 'Exact General guide source unavailable.'}
        Delivered (BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths')
        Delivered (Author 'txtGuideName' 'Write' $(if($AggregationActions){'Aggregation source-set guide'}elseif($TesterActions){'Test Environment Setup guide'}elseif($CreateActions){'Create Warehouse training'}elseif($RootActions){'Add Warehouse Root training'}elseif($DirectoryActions){'View Warehouses training'}elseif($Seed){'Seed Demo Inventory training'}elseif($Users){'Users and Roles training'}elseif($AdminConsole){'Admin Console training'}elseif($VerifyAddins){'Verify Add-ins Published training'}elseif($Reconcile){'Shipment Reconcile training'}elseif($SettingsNavigation){'Settings navigation training'}elseif($WorksheetScale){'Worksheet Scale training'}elseif($CloseActions){'Settings Close training'}else{'General Settings training'}))
        Delivered (Author 'txtGuideInstructions' 'Write' $instruction)
        Pass $(if($Users){'ThreeObservedGuideSteps'}elseif($SettingsNavigation){'SevenObservedGuideSteps'}elseif($CloseActions -or $WorksheetScale -or $Reconcile -or $VerifyAddins -or $AdminConsole -or $Users -or $Seed -or $DirectoryActions -or $RootActions -or $CreateActions -or $TesterActions -or $AggregationActions){'TwoObservedGuideSteps'}else{'NineObservedGuideSteps'}) ((Author 'lstGuideSteps' 'Rows') -ceq [string]$Steps.Count)
        for($index=0;$index -lt $Steps.Count;$index++){
            if((Author 'lstGuideSteps' 'Select' ([string]$index)) -cne 'SELECTED'){throw 'General guide step unavailable.'}
            Delivered (Author 'txtGuideStepInstruction' 'Write' ('Use '+$Steps[$index][0]+'. Read the owner result before continuing.'))
        }
        Delivered (Author 'btnGuideExpectedConclusion' 'Click')
        $count=BoundControl 'lstExpectedSteps' 'Rows' '' 'frmActionPathExpectation'
        for($index=[int]$count-1;$index -ge 0;$index--){
            if((BoundControl 'lstExpectedSteps' 'Select' ([string]$index) 'frmActionPathExpectation') -cne 'SELECTED'){throw 'Expectation removal selection unavailable.'}
            Delivered (BoundControl 'btnRemoveExpectedStep' 'Click' '' 'frmActionPathExpectation')
        }
        foreach($step in $Steps){
            Delivered (BoundControl 'cboExpectedControl' 'Write' $step[1] 'frmActionPathExpectation')
            if($AggregationActions){
                $choices=@((BoundControl 'cboExpectedOutcome' 'Values' '' 'frmActionPathExpectation') -split "`n"|Where-Object {$_ -cne ''})
                $expected=@('REQUESTED','COMPLETED','REJECTED','DENIED','FAILED','CANCELLED','INDETERMINATE','COMPLETED_STOPPED')
                $offered=@($expected|Where-Object {$_ -cnotin $choices}).Count -eq 0 -and @($choices|Where-Object {$_ -cnotin $expected}).Count -eq 0
                Pass ('ExactAggregationOutcomesOffered.'+[array]::IndexOf($Steps,$step)) $offered
                if(-not $offered){return}
            }
            if($TesterActions){
                $choices=@((BoundControl 'cboExpectedOutcome' 'Values' '' 'frmActionPathExpectation') -split "`n"|Where-Object {$_ -cne ''})
                $expected=@('REQUESTED','COMPLETED','REJECTED','DENIED','FAILED')
                $offered=@($expected|Where-Object {$_ -cnotin $choices}).Count -eq 0 -and @($choices|Where-Object {$_ -cnotin $expected}).Count -eq 0
                Pass ('ExactTesterOutcomesOffered.'+[array]::IndexOf($Steps,$step)) $offered
                if(-not $offered){return}
            }
            if($CreateActions){
                $choices=@((BoundControl 'cboExpectedOutcome' 'Values' '' 'frmActionPathExpectation') -split "`n"|Where-Object {$_ -cne ''})
                $expected=@('REQUESTED','REJECTED','DENIED','FAILED','CREATED_LOCAL','CREATED_PUBLISHED','CREATED_PUBLISH_FAILED','CREATED_PUBLISH_DENIED')
                $offered=@($expected|Where-Object {$_ -cnotin $choices}).Count -eq 0 -and @($choices|Where-Object {$_ -cnotin $expected}).Count -eq 0
                Pass ('ExactCreateOutcomesOffered.'+[array]::IndexOf($Steps,$step)) $offered
                if(-not $offered){return}
            }
            if($RootActions){
                $choices=@((BoundControl 'cboExpectedOutcome' 'Values' '' 'frmActionPathExpectation') -split "`n"|Where-Object {$_ -cne ''})
                $offered=@('REMEMBERED_REFRESHED','REMEMBERED_REFRESH_FAILED','UNCONFIRMED_REFRESHED','UNCONFIRMED_REFRESH_FAILED'|Where-Object {$_ -cnotin $choices}).Count -eq 0
                Pass ('AllRootOutcomesOffered.'+[array]::IndexOf($Steps,$step)) $offered
                if(-not $offered){return}
            }
            if($VerifyAddins -or $Users -or $Seed){
                $choices=@((BoundControl 'cboExpectedOutcome' 'Values' '' 'frmActionPathExpectation') -split "`n"|Where-Object {$_ -cne ''})
                $offered=$step[2] -cin $choices
                Pass ($(if($Seed){'AppliedOutcomeOffered.'}elseif($Users){'SavedOutcomeOffered.'}else{'VerifiedOutcomeOffered.'})+[array]::IndexOf($Steps,$step)) $offered
                if(-not $offered){return}
            }
            Delivered (BoundControl 'cboExpectedOutcome' 'Write' $step[2] 'frmActionPathExpectation')
            Delivered (BoundControl 'btnAddExpectedStep' 'Click' '' 'frmActionPathExpectation')
        }
        if((BoundControl 'cboTerminalStep' 'Select' ([string]($Steps.Count-1)) 'frmActionPathExpectation') -cne 'SELECTED'){throw 'Last expected terminal step unavailable.'}
        Delivered (BoundControl 'cboTerminalKind' 'Write' 'CommandCompleted' 'frmActionPathExpectation')
        Delivered (BoundControl 'btnUseExpectation' 'Click' '' 'frmActionPathExpectation')
        $directory=Join-Path $journalRoot 'Guides';$prior=BoundPins $directory
        Delivered (Author 'btnSaveGuide' 'Click')
        $fresh=@(Get-ChildItem -LiteralPath $directory -Filter '*.json' -File|Where-Object {-not $prior.ContainsKey($_.FullName)})
        if($fresh.Count -ne 1){throw 'Exactly one newly authored General guide required.'}
        $guide=ReadGuideExpectationRecord $fresh[0].FullName
        if($null -eq $guide){throw 'General guide integrity unavailable.'}
        $exact=$guide.SourceRun.RecordId -ceq $source.Journal.RecordId -and $guide.SourceRun.ContentSha256 -ceq $source.Journal.ContentSha256 -and @($guide.Steps).Count -eq $Steps.Count -and @($guide.ExpectedConclusion.Steps).Count -eq $Steps.Count -and $guide.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted'
        for($index=0;$index -lt $Steps.Count;$index++){$exact=$exact -and $guide.ExpectedConclusion.Steps[$index].ControlId -ceq $Steps[$index][1] -and $guide.ExpectedConclusion.Steps[$index].RequiredOutcome -ceq $Steps[$index][2]}
        Pass 'ExactSourceAndAuthoredIntent' $exact
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        Pass 'ReaderCannotAuthor' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-reader',$Fixture.Warehouse,'S1')))
        $kinds=if($RootActions -or $CreateActions -or $TesterActions -or $AggregationActions){@('Observed','Interrupted','Unconfirmed')}elseif($CloseActions -or $WorksheetScale -or $SettingsNavigation -or $Reconcile -or $VerifyAddins -or $AdminConsole -or $Users -or $Seed -or $DirectoryActions -or $RootActions -or $CreateActions -or $TesterActions -or $AggregationActions){@('Observed','Interrupted')}else{@('Observed','Interrupted','Denied')}
        foreach($kind in $kinds){
            $run=@($Runs|Where-Object Kind -CEQ $kind)[0];$journal=$run.Journal
            Delivered (BoundLibrary 'Open')
            if((BoundLibrary 'Select' $journal.ActionPathId) -cne 'SELECTED'){throw 'Exact General comparison recording unavailable.'}
            BoundOpen;BoundSelect $guide
            Delivered (BoundControl 'btnUseGuideForRun' 'Click');Delivered (BoundControl 'btnCloseGuides' 'Click')
            $before=@(EvaluationFiles|ForEach-Object FullName)
            Delivered (BoundControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
            if($fresh.Count -ne 1){throw 'Paired Evaluate must append exactly one result.'}
            $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
            $state=if($run.Incomplete){'Incomplete'}elseif($run.Success){'Concluded'}else{'Failed'}
            $reason=if($run.Incomplete){'CAPTURE_INCOMPLETE'}elseif($run.Success){'COMMAND_COMPLETED'}else{'REQUIRED_OUTCOME_MISMATCH'}
            Pass ($kind+'.ExactDiagnosticState') ($result.ResultState -ceq $state -and $reason -cin @($result.ReasonCodes))
            Pass ($kind+'.ExactGuideAndObservedJournal') ($result.JournalRecordId -ceq $journal.RecordId -and $result.JournalSha256 -ceq $journal.ContentSha256 -and $result.Guide.ContentSha256 -ceq $guide.ContentSha256 -and $result.Guide.Version -eq $guide.Version)
            Pass ($kind+'.ExactPublication') ($result.Publication.PublicationId -ceq $publication.PublicationId -and $result.Publication.ContentSha256 -ceq $publication.ContentSha256 -and $result.Publication.WarehouseId -ceq $Fixture.Warehouse)
            Pass ($kind+'.NoDomainApplicationClaim') (@($result.TerminalSources).Count -eq 0 -and $result.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted')
            # Incomplete/failed conclusions retain the actions actually completed.
            # Interruption follows Reload; denial still permits four selections.
            $ordinals=@(if($run.Success){1..$Steps.Count}elseif($run.Incomplete -and $Users){1,2}elseif($run.Incomplete -or $RootActions -or $CreateActions -or $TesterActions -or $AggregationActions){1}else{2,3,6,8})
            $exactMatches=@($result.Matches).Count -eq $ordinals.Count
            foreach($ordinal in $ordinals){
                $expected=@($run.Records|Where-Object {$_.Ordinal -eq $ordinal -and $_.OutcomeCode -ceq $Steps[$ordinal-1][2]})
                $matches=@($result.Matches|Where-Object Ordinal -EQ $ordinal)
                $exactMatches=$exactMatches -and $expected.Count -eq 1 -and $matches.Count -eq 1
                if($expected.Count -eq 1 -and $matches.Count -eq 1){$exactMatches=$exactMatches -and $matches[0].ActivityId -ceq $expected[0].ActivityId -and $matches[0].StepId -ceq $guide.ExpectedConclusion.Steps[$ordinal-1].StepId}
            }
            Pass ($kind+'.OnlyObservedRunCanMatch') $exactMatches
            # Reconcile loses CONFIRMED after REQUESTED was saved; the attempted
            # terminal step has an unmatched outcome. Capture remains Incomplete.
            $missing=if($run.Incomplete -and -not ($Reconcile -or $VerifyAddins -or $AdminConsole -or $Users -or $Seed -or $DirectoryActions -or $RootActions -or $CreateActions -or $TesterActions -or $AggregationActions)){$Steps.Count-1}else{0}
            $failed=if(($Reconcile -or $VerifyAddins -or $AdminConsole -or $Users -or $Seed -or $DirectoryActions -or $RootActions -or $CreateActions -or $TesterActions -or $AggregationActions) -and $run.Incomplete){1}elseif(($RootActions -or $CreateActions -or $TesterActions -or $AggregationActions) -and -not $run.Success){1}elseif(-not $run.Success -and -not $run.Incomplete){5}else{0}
            $unmatched=@($result.MissingSteps).Count -eq $missing -and @($result.FailedSteps).Count -eq $failed -and @($result.UnavailableSteps).Count -eq 0
            if(($Reconcile -or $VerifyAddins -or $AdminConsole -or $Users -or $Seed -or $DirectoryActions -or $RootActions -or $CreateActions -or $TesterActions -or $AggregationActions) -and $run.Incomplete){$unmatched=$unmatched -and $result.FailedSteps[0] -ceq $guide.ExpectedConclusion.TerminalStepId -and 'REQUIRED_OUTCOME_MISMATCH' -cin @($result.ReasonCodes)}
            Pass ($kind+'.ExactUnmatchedSteps') $unmatched
            $pins=BoundPins $journalRoot
            Delivered (BoundControl 'btnViewActionPath' 'Click' '' 'frmActionPaths')
            $pair=View 'lblActionPathPair' 'Label';$how=View 'txtActionPathHowTo' 'Text';$diagnostic=View 'txtActionPathDiagnostic' 'Text'
            Pass ($kind+'.ExactPairAndMultilineInstructions') ($pair.Contains($guide.ContentSha256) -and $pair.Contains($journal.ActionPathId) -and $how.Contains($instruction))
            $noBorrowing=$diagnostic.Contains($result.EvaluationId) -and $diagnostic.Contains($run.Records[0].ActivityId)
            foreach($id in @($source.Records.ActivityId|Select-Object -Unique)){$noBorrowing=$noBorrowing -and -not $diagnostic.Contains($id)}
            Pass ($kind+'.DiagnosticCannotBorrowGuideSourceSuccess') $noBorrowing
            $caption=if($run.Incomplete){'Incomplete evidence'}elseif($run.Success){'Conclusion observed. Command completed; Domain application not asserted'}else{'Failed'}
            Pass ($kind+'.DisplayedSavedConclusion') ($diagnostic.Contains($caption+"`r`nSaved diagnostic result: "+$result.EvaluationId))
            foreach($method in @('How-To','Diagnostic','Compare both')){
                Delivered (View 'cboActionPathView' 'Write' $method)
                Pass ($kind+'.View.'+$method) ((View 'cboActionPathView' 'Selected') -ceq $method -and ((View 'txtActionPathHowTo' 'State') -split '\|')[0] -ceq $(if($method -ceq 'Diagnostic'){'False'}else{'True'}) -and ((View 'txtActionPathDiagnostic' 'State') -split '\|')[0] -ceq $(if($method -ceq 'How-To'){'False'}else{'True'}))
                Pass ($kind+'.SameEvidence.'+$method) ((View 'lblActionPathPair' 'Label') -ceq $pair -and (View 'txtActionPathHowTo' 'Text') -ceq $how -and (View 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
                Pass ($kind+'.Layout.'+$method) ((BoundControl '' 'Fit' 'Current' 'frmActionPathView') -ceq 'True')
                if($kind -ceq 'Observed' -or $method -ceq 'Compare both'){CaptureOwnedFormByCaptionEvidence 'Action Path view' ($capturePrefix+$kind.ToLowerInvariant()+'-'+$method.Replace(' ','-').ToLowerInvariant()+'.png')}
            }
            Delivered (View 'txtActionPathDiagnostic' 'ViewportText' ($caption+"`r`nSaved diagnostic result: "+$result.EvaluationId))
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ($capturePrefix+$kind.ToLowerInvariant()+'-conclusion.png')
            Pass ($kind+'.ViewsPreserveTrainingEvidence') (BoundSame $pins (BoundPins $journalRoot))
            Delivered (View 'btnCloseActionPathView' 'Click')
        }
        Pass 'SourceJournalsAndActivityPreserved' ((Retained $originalPins (BoundPins $journalRoot)) -and (PinsRetained $activityPins) -and (Hash $eventsPath) -ceq $publicationHash)
    }finally{CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
