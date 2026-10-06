# One authored Print guide, independently recorded success and interrupted runs.
function Test-ProductionPrintPaths($Fixture,$Runs){
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideResourceTrace.ps1')
    function Pass([string]$Name,[bool]$Value){Check ('PrintPaths.'+$Name) $Value|Out-Host}
    function Delivered([string]$Result){if($Result -cne 'DELIVERED'){throw 'Actual Print guide handler unavailable; not product RED.'}}
    function Author([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathGuide'}
    function View([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathView'}
    $source=@($Runs|Where-Object Label -CEQ 'PrintRecorded.Recovery')
    $observed=@($Runs|Where-Object Label -CEQ 'PrintRecorded.GuideRun')
    if($source.Count -ne 1 -or $observed.Count -ne 1){throw 'Two actual successful recordings required.'}
    $source=$source[0];$observed=$observed[0]
    Pass 'IndependentGuideSourceAndObservedRun' ($source.Success -and $observed.Success -and $source.Journal.ActionPathId -cne $observed.Journal.ActionPathId -and $source.Original[0].ActivityId -cne $observed.Original[0].ActivityId)
    $activityPins=ActivityPins;$originalPins=BoundPins $journalRoot
    $eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
    $publication=Get-Content $eventsPath -Raw|ConvertFrom-Json;$publishedPin=Hash $eventsPath
    $instruction="In Production Run - List, choose Print Recall.`r`nReview the preview, then close it and read the current status.`r`nPreview closure proves command completion, not a physically printed page. If tracking is unavailable, inspect the incomplete recording before using it for training."
    try{
        CloseRecordingViewer;SelectTarget $Fixture;OpenRecordingViewer
        Delivered (BoundLibrary 'Open')
        if((BoundLibrary 'Select' $source.Journal.ActionPathId) -cne 'SELECTED'){throw 'Print guide source unavailable.'}
        Delivered (BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths')
        Delivered (Author 'txtGuideName' 'Write' 'Review a Production recall preview')
        Delivered (Author 'txtGuideInstructions' 'Write' $instruction)
        Pass 'OneObservedGuideStep' ((Author 'lstGuideSteps' 'Rows') -ceq '1')
        if((Author 'lstGuideSteps' 'Select' '0') -cne 'SELECTED'){throw 'Print guide step unavailable.'}
        Delivered (Author 'txtGuideStepInstruction' 'Write' $instruction)
        Delivered (Author 'btnGuideExpectedConclusion' 'Click')
        foreach($pair in @(@('cboExpectedControl','PRODUCTION_RUN_PRINT'),@('cboExpectedOutcome','PREVIEW_RETURNED'))){Delivered (BoundControl $pair[0] 'Write' $pair[1] 'frmActionPathExpectation')}
        Delivered (BoundControl 'btnAddExpectedStep' 'Click' '' 'frmActionPathExpectation')
        if((BoundControl 'cboTerminalStep' 'Select' '0' 'frmActionPathExpectation') -cne 'SELECTED'){throw 'Print terminal step unavailable.'}
        Delivered (BoundControl 'cboTerminalKind' 'Write' 'CommandCompleted' 'frmActionPathExpectation')
        Delivered (BoundControl 'btnUseExpectation' 'Click' '' 'frmActionPathExpectation')
        $guides=Join-Path $journalRoot 'Guides';$prior=BoundPins $guides
        Delivered (Author 'btnSaveGuide' 'Click')
        $fresh=@(Get-ChildItem $guides -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)})
        if($fresh.Count -ne 1){throw 'Exactly one authored guide required.'}
        $guide=ReadGuideExpectationRecord $fresh[0].FullName
        if($null -eq $guide){throw 'Saved Print guide integrity unavailable.'}
        Pass 'ExactGuideSourceAndCommandExpectation' ($guide.SourceRun.RecordId -ceq $source.Journal.RecordId -and $guide.SourceRun.ContentSha256 -ceq $source.Journal.ContentSha256 -and @($guide.Steps).Count -eq 1 -and $guide.Steps[0].SourceActivityId -ceq $source.Original[0].ActivityId -and $guide.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted' -and $guide.ExpectedConclusion.Steps[0].RequiredOutcome -ceq 'PREVIEW_RETURNED')
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        Pass 'ReaderCannotAuthor' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-reader',$Fixture.Warehouse,'S1')))
        foreach($case in @('GuideRun','PolicyChanged','Permission')){
            $selected=@($Runs|Where-Object Label -CEQ ('PrintRecorded.'+$case))
            if($selected.Count -ne 1){throw 'Exact Print comparison recording required.'}
            $run=$selected[0];$journal=$run.Journal;$activity=$run.Original[0].ActivityId
            Delivered (BoundLibrary 'Open')
            if((BoundLibrary 'Select' $journal.ActionPathId) -cne 'SELECTED'){throw 'Observed Print run unavailable.'}
            BoundOpen;BoundSelect $guide
            Delivered (BoundControl 'btnUseGuideForRun' 'Click');Delivered (BoundControl 'btnCloseGuides' 'Click')
            $before=@(EvaluationFiles|ForEach-Object FullName)
            Delivered (BoundControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before})
            if($fresh.Count -ne 1){throw 'Actual paired Evaluate must append one result.'}
            $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
            $state=if($run.Incomplete){'Incomplete'}elseif($run.Success){'Concluded'}else{'Failed'}
            $reason=if($run.Incomplete){'CAPTURE_INCOMPLETE'}elseif($run.Success){'COMMAND_COMPLETED'}else{'REQUIRED_OUTCOME_MISMATCH'}
            Pass ($case+'.ExactDiagnosticState') ($result.ResultState -ceq $state -and $reason -cin @($result.ReasonCodes))
            Pass ($case+'.ExactGuideAndObservedJournal') ($result.JournalRecordId -ceq $journal.RecordId -and $result.JournalSha256 -ceq $journal.ContentSha256 -and $result.Guide.ContentSha256 -ceq $guide.ContentSha256 -and $result.Guide.Version -eq $guide.Version)
            Pass ($case+'.ExactPublishedProvenance') ($result.Publication.PublicationId -ceq $publication.PublicationId -and $result.Publication.ContentSha256 -ceq $publication.ContentSha256 -and $result.Publication.WarehouseId -ceq $Fixture.Warehouse)
            Pass ($case+'.NoInventedBusinessApplication') (@($result.TerminalSources).Count -eq 0 -and $result.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted')
            Pass ($case+'.OnlyObservedRunCanMatch') (@($result.Matches).Count -eq $(if($run.Success){1}else{0}) -and @($result.Matches|Where-Object ActivityId -CNE $activity).Count -eq 0)
            $trainingPins=BoundPins $journalRoot
            Get-GuideResourceSample ($case+'.BeforeView')|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'print-path-view-resources.jsonl')
            Delivered (BoundControl 'btnViewActionPath' 'Click' '' 'frmActionPaths')
            Get-GuideResourceSample ($case+'.AfterView')|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'print-path-view-resources.jsonl')
            $pair=View 'lblActionPathPair' 'Label';$how=View 'txtActionPathHowTo' 'Text';$diagnostic=View 'txtActionPathDiagnostic' 'Text'
            Pass ($case+'.ExactPairAndMultilineInstruction') ($pair.Contains($guide.ContentSha256) -and $pair.Contains($journal.ActionPathId) -and $how.Contains($instruction))
            Pass ($case+'.DiagnosticShowsObservedNotSourceAction') ($diagnostic.Contains($activity) -and -not $diagnostic.Contains($source.Original[0].ActivityId) -and $diagnostic.Contains($result.EvaluationId))
            foreach($method in @('How-To','Diagnostic','Compare both')){
                Delivered (View 'cboActionPathView' 'Write' $method)
                Pass ($case+'.View.'+$method) ((View 'cboActionPathView' 'Selected') -ceq $method -and ((View 'txtActionPathHowTo' 'State') -split '\|')[0] -ceq $(if($method -ceq 'Diagnostic'){'False'}else{'True'}) -and ((View 'txtActionPathDiagnostic' 'State') -split '\|')[0] -ceq $(if($method -ceq 'How-To'){'False'}else{'True'}))
                Pass ($case+'.SameEvidence.'+$method) ((View 'lblActionPathPair' 'Label') -ceq $pair -and (View 'txtActionPathHowTo' 'Text') -ceq $how -and (View 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
                CaptureOwnedFormByCaptionEvidence 'Action Path view' ('print-path-'+$case.ToLowerInvariant()+'-'+$method.Replace(' ','-').ToLowerInvariant()+'.png')
            }
            Delivered (View 'txtActionPathDiagnostic' 'ViewportBottom')
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ('print-path-'+$case.ToLowerInvariant()+'-conclusion.png')
            Pass ($case+'.ViewPreservesAllTrainingEvidence') (BoundSame $trainingPins (BoundPins $journalRoot))
            Delivered (View 'btnCloseActionPathView' 'Click')
        }
        $retained=$true;foreach($file in $originalPins.Keys){$retained=$retained -and (Hash $file) -ceq $originalPins[$file]}
        Pass 'OriginalJournalsAndEvaluationsImmutable' $retained
        Pass 'ActivityAndPublicationImmutable' ((PinsRetained $activityPins) -and (Hash $eventsPath) -ceq $publishedPin)
    }finally{CloseRecordingViewer;SelectTarget $Fixture}
}
