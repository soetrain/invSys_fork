# Real lifecycle recordings supply distinct guide provenance and observed evidence.
function Test-ProductionLifecyclePresentation($Fixture) {
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    function Probe([string]$Method,[object[]]$Values=@()) {Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Author([string]$Name,[string]$Action,[string]$Value='') {BoundControl $Name $Action $Value 'frmActionPathGuide'}
    function View([string]$Name,[string]$Action,[string]$Value='') {BoundControl $Name $Action $Value 'frmActionPathView'}
    function Expected([string]$Name,[string]$Action,[string]$Value='') {BoundControl $Name $Action $Value 'frmActionPathExpectation'}
    function SavedHash([string]$Path) {
        $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
    }
    function RequireDelivery([string]$Result) {
        if($Result -ceq 'DELIVERED'){return}
        $kind=if($Result -cin @('DISABLED','MISSING','OUT OF RANGE','NOT FOUND','SELECTED')){$Result}else{'OTHER'}
        [pscustomobject]@{CallerLine=$MyInvocation.ScriptLineNumber;ResultClass=$kind}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $reportRoot 'presentation-fixture-failure.json')
        throw 'Existing packaged action fixture is unavailable; inspect prerequisites before claiming RED.'
    }
    function RecordSeries([string]$Label,[ref]$Recorded) {
        $ready=[string](Probe 'LifecyclePrepare' @($canary))
        if($ready -cnotmatch '^READY\|([^|]+)\|([^|]+)$'){throw 'Owning Process prerequisite unavailable.'}
        $processId=$Matches[1];$version=$Matches[2]
        $prior=@(if(Test-Path $journalRoot){Get-ChildItem $journalRoot -File -Filter '*.json'|ForEach-Object FullName})
        RequireDelivery (RecordingControl 'Start Recording' 'Click')
        $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {$_.FullName -cnotin $prior}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
        if($starts.Count -ne 1){throw 'Actual recording start unavailable.'}
        $terminals=@()
        foreach($action in $actions){
            if($action[0] -ceq 'Recipe' -and $action[1] -ceq 'Save'){
                if(-not [bool](Probe 'ReleasedRecipe' @($processId,$version,$canary))){throw 'Released Process prerequisite unavailable.'}
            }
            $before=@(Get-Slice4beActivityFiles $Fixture)
            $delivered=[string](Probe 'LifecycleAct' @($action[0],$action[1]))
            $rows=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
            $terminal=@($rows|Where-Object OutcomeCode -CEQ 'CONFIRMED')
            $valid=$delivered -ceq 'RETURNED' -and [string](Probe 'LifecycleStatus' @($action[0])) -ceq $action[2] -and $rows.Count -eq 2 -and $terminal.Count -eq 1
            Check ('LifecyclePresentation.'+$Label+'.'+$action[0]+'.'+$action[1]+'.ActualConfirmedHandler') $valid
            if(-not $valid){throw 'Confirmed owning series prerequisite unavailable.'}
            $terminals+=$terminal[0]
        }
        RequireDelivery (RecordingControl 'Stop Recording' 'Click')
        $closed=@(RecordingJournal $starts[0].SequenceId|Where-Object RecordType -CEQ 'Close')
        $valid=$closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq 6 -and (JournalChain $starts[0].SequenceId 14)
        if($valid){$valid=(@($closed[0].Observations|Where-Object OutcomeCode -CEQ 'CONFIRMED'|ForEach-Object RecordId) -join '|') -ceq ($terminals.RecordId -join '|')}
        Check ('LifecyclePresentation.'+$Label+'.ExactSixActionJournal') $valid
        if(-not $valid){throw 'Original six-action journal unavailable.'}
        $Recorded.Value=[pscustomobject]@{Journal=$closed[0];Terminals=$terminals}
    }
    $actions=@(@('Process','Save','DRAFT'),@('Process','Release','RELEASED'),@('Recipe','Save','DRAFT'),@('Recipe','Release','RELEASED'),@('Recipe','Obsolete','OBSOLETE'),@('Process','Obsolete','OBSOLETE'))
    $canary='PAIR'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null
    try {
        SelectTarget $Fixture
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.SaveSettings' @('DesignsEnabled','TRUE'))){throw 'Explicit Designs setup failed.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
        SetRecordingPolicy $true;SelectTarget $Fixture 'config-producer';OpenRecordingViewer
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $book.SaveAs((Join-Path $runRoot 'lifecycle-presentation-operator.xlsb'),50)
        $decoy=$excel.Workbooks.Add();[void](Probe 'OpenDesigner' @($book.Name));[void](Probe 'RefreshDesigners')
        $source=$null;$observed=$null
        RecordSeries 'GuideSource' ([ref]$source)
        RecordSeries 'ObservedRun' ([ref]$observed)
        Check 'LifecyclePresentation.DistinctSourceAndObservedRuns' ($source.Journal.ActionPathId -cne $observed.Journal.ActionPathId -and @($source.Terminals|Where-Object {$_.ActivityId -cin $observed.Terminals.ActivityId}).Count -eq 0)
        Check 'LifecyclePresentation.UnknownOperatorColumnPreserved' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary)
        [void](Probe 'CloseDesigner');CloseRecordingViewer;SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Actual Admin publication unavailable.'}
        $authorAllowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
        Check 'LifecyclePresentation.Setup.AuthorHasExplicitMaintenanceCapability' ($authorAllowed -is [bool] -and $authorAllowed)
        if($authorAllowed -isnot [bool] -or -not $authorAllowed){throw 'Explicit guide-author fixture capability unavailable; not behavioral RED.'}
        OpenRecordingViewer
        RequireDelivery (BoundLibrary 'Open')
        if((BoundLibrary 'Select' $source.Journal.ActionPathId) -cne 'SELECTED'){throw 'Guide source selection unavailable.'}
        RequireDelivery (BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths')
        Check 'LifecyclePresentation.SixObservedControlsBecomeGuideSteps' ((Author 'lstGuideSteps' 'Rows') -ceq '6')
        RequireDelivery (Author 'txtGuideName' 'Write' 'Process and Recipe lifecycle')
        RequireDelivery (Author 'txtGuideInstructions' 'Write' 'Save and release a Process, save and release its Recipe, then obsolete the Recipe and Process.')
        for($i=0;$i -lt 6;$i++){
            if((Author 'lstGuideSteps' 'Select' ([string]$i)) -cne 'SELECTED'){throw 'Guide step selection unavailable.'}
            RequireDelivery (Author 'txtGuideStepInstruction' 'Write' ($actions[$i][1]+' the '+$actions[$i][0]+' and verify the recorded outcome.'))
        }
        RequireDelivery (Author 'btnGuideExpectedConclusion' 'Click')
        foreach($action in $actions){
            RequireDelivery (Expected 'cboExpectedControl' 'Write' ('PRODUCTION_'+$action[0].ToUpperInvariant()+'_'+$action[1].ToUpperInvariant()))
            RequireDelivery (Expected 'cboExpectedOutcome' 'Write' 'CONFIRMED')
            RequireDelivery (Expected 'btnAddExpectedStep' 'Click')
        }
        if((Expected 'cboTerminalStep' 'Select' '5') -cne 'SELECTED'){throw 'Sixth terminal step unavailable.'}
        RequireDelivery (Expected 'cboTerminalKind' 'Write' 'SourceEventsApplied')
        RequireDelivery (Expected 'btnUseExpectation' 'Click')
        $guideRoot=Join-Path $journalRoot 'Guides';$guideBefore=BoundPins $guideRoot
        RequireDelivery (Author 'btnSaveGuide' 'Click')
        $fresh=@(Get-ChildItem $guideRoot -File -Filter '*.json'|Where-Object {-not $guideBefore.ContainsKey($_.FullName)})
        if($fresh.Count -ne 1){throw 'Actual guide save did not append one immutable version.'}
        $guide=ReadGuideExpectationRecord $fresh[0].FullName
        if($null -eq $guide){throw 'Saved guide integrity unavailable.'}
        Check 'LifecyclePresentation.SavedGuideHasExplicitSixStepIntent' ($guide.ExpectedConclusion.TerminalKind -ceq 'SourceEventsApplied' -and @($guide.ExpectedConclusion.Steps).Count -eq 6)
        Check 'LifecyclePresentation.GuideRetainsExactDifferentSource' ($guide.SourceRun.ActionPathId -ceq $source.Journal.ActionPathId -and $guide.SourceRun.RecordId -ceq $source.Journal.RecordId -and $guide.SourceRun.ContentSha256 -ceq $source.Journal.ContentSha256 -and ($guide.Steps.SourceActivityId -join '|') -ceq ($source.Terminals.ActivityId -join '|'))
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        $readerAllowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-reader',$Fixture.Warehouse,'S1')
        Check 'LifecyclePresentation.ReaderHasNoMaintenanceCapability' ($readerAllowed -is [bool] -and -not $readerAllowed)
        RequireDelivery (BoundLibrary 'Open')
        if((BoundLibrary 'Select' $observed.Journal.ActionPathId) -cne 'SELECTED'){throw 'Separate observed run unavailable.'}
        BoundOpen;BoundSelect $guide
        RequireDelivery (BoundControl 'btnUseGuideForRun' 'Click')
        RequireDelivery (BoundControl 'btnCloseGuides' 'Click')
        $evaluationRoot=Join-Path $journalRoot 'Evaluations';$before=BoundPins $evaluationRoot
        RequireDelivery (BoundControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
        $fresh=@(Get-ChildItem $evaluationRoot -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)})
        if($fresh.Count -ne 1){throw 'Actual paired Evaluate did not append one result.'}
        $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
        Check 'LifecyclePresentation.ConcludesOnlyExactObservedRun' ($result.ResultState -ceq 'Concluded' -and $result.JournalRecordId -ceq $observed.Journal.RecordId -and $result.JournalSha256 -ceq $observed.Journal.ContentSha256 -and $result.Guide.ContentSha256 -ceq $guide.ContentSha256)
        Check 'LifecyclePresentation.MatchesAllSixOriginalObservedActions' (($result.Matches.ActivityId -join '|') -ceq ($observed.Terminals.ActivityId -join '|') -and @($result.ExtraActivityIds).Count -eq 0)
        $trainingPins=BoundPins $journalRoot;$activityPins=ActivityPins
        $configHash=(Get-FileHash $Fixture.Config).Hash
        $authorityPins=@{}
        foreach($file in Get-ChildItem $Fixture.Root -File -Filter '*.xlsb'|Where-Object {$_.Name -notlike '*.Snapshot.*' -and -not $_.Name.StartsWith('~$')}){$authorityPins[$file.FullName]=SavedHash $file.FullName}
        $publication=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json');$publishedHash=(Get-FileHash $publication).Hash
        RequireDelivery (BoundControl 'btnViewActionPath' 'Click' '' 'frmActionPaths')
        $pair=View 'lblActionPathPair' 'Label';$how=View 'txtActionPathHowTo' 'Text';$diagnostic=View 'txtActionPathDiagnostic' 'Text'
        Check 'LifecyclePresentation.ExactGuideAndObservedPairNamed' ($pair.Contains([string]$guide.ContentSha256) -and $pair.Contains([string]$observed.Journal.ActionPathId))
        $instructions=$how.Contains('Authored instruction') -and $how.Contains([string]$guide.Name)
        foreach($action in $actions){$instructions=$instructions -and $how.Contains($action[1]+' the '+$action[0]+' and verify the recorded outcome.')}
        Check 'LifecyclePresentation.AllSixAuthoredInstructionsRetainedInHowTo' $instructions
        $ordered=$true;$last=-1
        foreach($record in $observed.Terminals){$index=$diagnostic.IndexOf([string]$record.ActivityId,[StringComparison]::Ordinal);$ordered=$ordered -and $index -gt $last;$last=$index}
        foreach($record in $source.Terminals){$ordered=$ordered -and -not $diagnostic.Contains([string]$record.ActivityId)}
        Check 'LifecyclePresentation.DiagnosticUsesObservedOrderNotGuideSource' $ordered
        Check 'LifecyclePresentation.DiagnosticRetainsSavedConclusion' ($diagnostic.Contains([string]$result.EvaluationId) -and $diagnostic.Contains('Conclusion observed') -and $diagnostic.Contains('Additional observed actions: 0'))
        Check 'LifecyclePresentation.BothPanesLocked' ((View 'txtActionPathHowTo' 'Locked') -ceq 'True' -and (View 'txtActionPathDiagnostic' 'Locked') -ceq 'True')
        foreach($method in @('How-To','Diagnostic','Compare both')){
            RequireDelivery (View 'cboActionPathView' 'Write' $method)
            Check ('LifecyclePresentation.Method.'+$method) ((View 'cboActionPathView' 'Selected') -ceq $method -and ((View 'txtActionPathHowTo' 'State') -split '\|')[0] -ceq $(if($method -ceq 'Diagnostic'){'False'}else{'True'}) -and ((View 'txtActionPathDiagnostic' 'State') -split '\|')[0] -ceq $(if($method -ceq 'How-To'){'False'}else{'True'}))
            Check ('LifecyclePresentation.SameEvidence.'+$method) ((View 'lblActionPathPair' 'Label') -ceq $pair -and (View 'txtActionPathHowTo' 'Text') -ceq $how -and (View 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ('lifecycle-'+$method.Replace(' ','-').ToLowerInvariant()+'.png')
            Check ('LifecyclePresentation.Captured.'+$method) $true
            if($method -ceq 'How-To'){
                RequireDelivery (View 'txtActionPathHowTo' 'ViewportBottom')
                CaptureOwnedFormByCaptionEvidence 'Action Path view' 'lifecycle-how-to-last-instruction.png'
                Check 'LifecyclePresentation.CapturedLastAuthoredInstruction' $true
                RequireDelivery (View 'txtActionPathHowTo' 'ViewportTop')
            }
        }
        RequireDelivery (View 'txtActionPathDiagnostic' 'ViewportBottom')
        CaptureOwnedFormByCaptionEvidence 'Action Path view' 'lifecycle-conclusion-bottom.png'
        Check 'LifecyclePresentation.CapturedSavedConclusion' $true
        Check 'LifecyclePresentation.ViewsPreserveTrainingActivityConfigAndPublication' ((BoundSame $trainingPins (BoundPins $journalRoot)) -and (PinsRetained $activityPins) -and (ActivityPins).Count -eq $activityPins.Count -and (Get-FileHash $Fixture.Config).Hash -ceq $configHash -and (Get-FileHash $publication).Hash -ceq $publishedHash)
        $same=$true;foreach($file in $authorityPins.Keys){$same=$same -and (SavedHash $file) -ceq $authorityPins[$file]}
        Check 'LifecyclePresentation.ViewsPreserveSavedAuthority' $same
    } finally {
        [void](Probe 'CloseDesigner');CloseRecordingViewer
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
