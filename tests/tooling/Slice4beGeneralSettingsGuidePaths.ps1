# One authored nine-control guide compared with three independent recordings.
function Test-GeneralSettingsGuidePaths($Fixture,$Runs,$Steps) {
    function Delivered([string]$Value){if($Value -cne 'DELIVERED'){throw 'Actual General guide control unavailable; not product RED.'}}
    function Author([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathGuide'}
    function View([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathView'}
    function Pass([string]$Name,[bool]$Value){Check ('GeneralGuide.'+$Name) $Value|Out-Host}
    $source=@($Runs|Where-Object Kind -CEQ 'Source')[0]
    $observed=@($Runs|Where-Object Kind -CEQ 'Observed')[0]
    $eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $publication=Get-Content $eventsPath -Raw|ConvertFrom-Json;$publicationHash=Hash $eventsPath
    $originalPins=BoundPins $journalRoot;$activityPins=ActivityPins
    $instruction="Use General Settings to inspect configuration and UOM choices, manage carriers, and save the connection option.`r`nCarrier and connection changes apply to this Windows user; they do not prove inventory changes.`r`nRead the observed result. Interrupted or denied actions cannot borrow completion from this guide."
    try {
        Pass 'IndependentSourceAndObservedRun' ($source.Journal.ActionPathId -cne $observed.Journal.ActionPathId -and @($observed.Records|Where-Object {$_.ActivityId -cin @($source.Records.ActivityId)}).Count -eq 0)
        CloseRecordingViewer;SelectTarget $Fixture 'config-admin';OpenRecordingViewer
        Delivered (BoundLibrary 'Open')
        if((BoundLibrary 'Select' $source.Journal.ActionPathId) -cne 'SELECTED'){throw 'Exact General guide source unavailable.'}
        Delivered (BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths')
        Delivered (Author 'txtGuideName' 'Write' 'General Settings training')
        Delivered (Author 'txtGuideInstructions' 'Write' $instruction)
        Pass 'NineObservedGuideSteps' ((Author 'lstGuideSteps' 'Rows') -ceq '9')
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
            Delivered (BoundControl 'cboExpectedOutcome' 'Write' $step[2] 'frmActionPathExpectation')
            Delivered (BoundControl 'btnAddExpectedStep' 'Click' '' 'frmActionPathExpectation')
        }
        if((BoundControl 'cboTerminalStep' 'Select' '8' 'frmActionPathExpectation') -cne 'SELECTED'){throw 'Save Connection terminal step unavailable.'}
        Delivered (BoundControl 'cboTerminalKind' 'Write' 'CommandCompleted' 'frmActionPathExpectation')
        Delivered (BoundControl 'btnUseExpectation' 'Click' '' 'frmActionPathExpectation')
        $directory=Join-Path $journalRoot 'Guides';$prior=BoundPins $directory
        Delivered (Author 'btnSaveGuide' 'Click')
        $fresh=@(Get-ChildItem -LiteralPath $directory -Filter '*.json' -File|Where-Object {-not $prior.ContainsKey($_.FullName)})
        if($fresh.Count -ne 1){throw 'Exactly one newly authored General guide required.'}
        $guide=ReadGuideExpectationRecord $fresh[0].FullName
        if($null -eq $guide){throw 'General guide integrity unavailable.'}
        $exact=$guide.SourceRun.RecordId -ceq $source.Journal.RecordId -and $guide.SourceRun.ContentSha256 -ceq $source.Journal.ContentSha256 -and @($guide.Steps).Count -eq 9 -and @($guide.ExpectedConclusion.Steps).Count -eq 9 -and $guide.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted'
        for($index=0;$index -lt $Steps.Count;$index++){$exact=$exact -and $guide.ExpectedConclusion.Steps[$index].ControlId -ceq $Steps[$index][1] -and $guide.ExpectedConclusion.Steps[$index].RequiredOutcome -ceq $Steps[$index][2]}
        Pass 'ExactSourceAndAuthoredIntent' $exact
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        Pass 'ReaderCannotAuthor' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-reader',$Fixture.Warehouse,'S1')))
        foreach($kind in @('Observed','Interrupted','Denied')){
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
            Pass ($kind+'.OnlyObservedRunCanMatch') (@($result.Matches).Count -eq $(if($run.Success){9}else{0}) -and @($result.Matches|Where-Object {$_.ActivityId -cnotin @($run.Records.ActivityId)}).Count -eq 0)
            $pins=BoundPins $journalRoot
            Delivered (BoundControl 'btnViewActionPath' 'Click' '' 'frmActionPaths')
            $pair=View 'lblActionPathPair' 'Label';$how=View 'txtActionPathHowTo' 'Text';$diagnostic=View 'txtActionPathDiagnostic' 'Text'
            Pass ($kind+'.ExactPairAndMultilineInstructions') ($pair.Contains($guide.ContentSha256) -and $pair.Contains($journal.ActionPathId) -and $how.Contains($instruction))
            $noBorrowing=$diagnostic.Contains($result.EvaluationId) -and $diagnostic.Contains($run.Records[0].ActivityId)
            foreach($id in @($source.Records.ActivityId|Select-Object -Unique)){$noBorrowing=$noBorrowing -and -not $diagnostic.Contains($id)}
            Pass ($kind+'.DiagnosticCannotBorrowGuideSourceSuccess') $noBorrowing
            foreach($method in @('How-To','Diagnostic','Compare both')){
                Delivered (View 'cboActionPathView' 'Write' $method)
                Pass ($kind+'.View.'+$method) ((View 'cboActionPathView' 'Selected') -ceq $method -and ((View 'txtActionPathHowTo' 'State') -split '\|')[0] -ceq $(if($method -ceq 'Diagnostic'){'False'}else{'True'}) -and ((View 'txtActionPathDiagnostic' 'State') -split '\|')[0] -ceq $(if($method -ceq 'How-To'){'False'}else{'True'}))
                Pass ($kind+'.SameEvidence.'+$method) ((View 'lblActionPathPair' 'Label') -ceq $pair -and (View 'txtActionPathHowTo' 'Text') -ceq $how -and (View 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
                if($kind -ceq 'Observed' -or $method -ceq 'Compare both'){CaptureOwnedFormByCaptionEvidence 'Action Path view' ('general-guide-'+$kind.ToLowerInvariant()+'-'+$method.Replace(' ','-').ToLowerInvariant()+'.png')}
            }
            Delivered (View 'txtActionPathDiagnostic' 'ViewportBottom')
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ('general-guide-'+$kind.ToLowerInvariant()+'-conclusion.png')
            Pass ($kind+'.ViewsPreserveTrainingEvidence') (BoundSame $pins (BoundPins $journalRoot))
            Delivered (View 'btnCloseActionPathView' 'Click')
        }
        Pass 'SourceJournalsAndActivityPreserved' ((Retained $originalPins (BoundPins $journalRoot)) -and (PinsRetained $activityPins) -and (Hash $eventsPath) -ceq $publicationHash)
    }finally{CloseRecordingViewer;SelectTarget $Fixture 'config-admin'}
}
