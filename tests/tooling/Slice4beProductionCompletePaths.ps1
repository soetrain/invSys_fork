# Real Complete Run recordings, owning publication and independent guide evaluation.
function Install-CompletePathsProbe {
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function CompletePathWorksheetKeysForTest() As Variant
    Dim output As ListObject
    Set output = ProductionTable(TABLE_MANAGER_OUTPUT)
    CompletePathWorksheetKeysForTest = Array(mRunSheetKey, CellByHeader(output, 1, "System_Key"))
End Function
'@)
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function CompletePathWorksheetKeys() As Variant
    CompletePathWorksheetKeys = mForm.CompletePathWorksheetKeysForTest()
End Function
'@)
}

function Test-ProductionCompletePaths($Fixture,$Book,$Decoy,[string]$Canary,[string]$Mode){
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationEvidence.ps1')
    $control='PRODUCTION_RUN_COMPLETE';$prefix='CompletePaths.'+$Mode+'.'
    function Pass([string]$Name,[bool]$Value){Check ($prefix+$Name) $Value|Out-Host}
    function Delivered([string]$Value){
        if($Value -cne 'DELIVERED'){
            $kind=if($Value -cin @('MISSING','DISABLED','HIDDEN','CHOICE_UNAVAILABLE','SELECTED','NOT FOUND')){$Value}else{'OTHER'}
            [pscustomobject]@{CallerLine=$MyInvocation.ScriptLineNumber;ResultClass=$kind}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'complete-path-prerequisite.json')
            throw 'Packaged guide prerequisite unavailable; not product RED.'
        }
    }
    function Author([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathGuide'}
    function View([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathView'}
    function Evaluate([string]$Kind,[string]$Stage){
        $ready=Set-EvaluationDraft @(,@($control,'CONFIRMED','True')) 0 $Kind -StopAtMissingChoice
        Pass ('Choice.'+$Stage+'.'+$Kind) $ready
        if(-not $ready){return $null}
        $prior=@(EvaluationFiles|ForEach-Object FullName)
        Delivered (ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
        $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $prior})
        if($fresh.Count -ne 1){throw 'Actual Evaluate result unavailable.'}
        Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
    }
    function Record([string]$Label){
        $work=$null
        try{
            $work=$excel.Workbooks.Add();$sheet=$work.Worksheets.Item(1)
            $sheet.Cells.Item(2,1).Value2=$Canary;$sheet.Cells.Item(2,2).Formula='=1+2'
            [void](Probe 'RunLocalReopen' @($work.Name))
            if($Mode -ceq 'Reusable'){$ready=[bool](Probe 'CompleteBaselinePrepare' @($true))}
            else{
                $ready=[bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CompleteWorksheetInventoryForTest' @($work.Name))
                if($ready){$ready=[bool](Probe 'CompleteWorksheetPrepare' @($Canary))}
            }
            if(-not $ready){throw 'Actual Complete Run prerequisite unavailable; not product RED.'}
            [void](Probe 'RunLocalShowAndCapture' @($work.Name,'CHECK_IN'))
            $prior=BoundPins $journalRoot
            Delivered (RecordingControl 'Start Recording' 'Click')
            $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object{Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
            if($starts.Count -ne 1){throw 'Actual recording start unavailable.'}
            $before=@(Get-Slice4beActivityFiles $Fixture);$Decoy.Activate()
            [void](Probe 'CompleteSubmissionArm' @('','','ObserveOnly'))
            Pass ($Label+'.ActualHandlerReturned') ([bool](Probe 'CompleteEntryAct' @('')))
            Pass ($Label+'.GuardsRestored') ([bool](Probe 'CompleteEntryFact' @('GuardsRestored')))
            if($Mode -ceq 'Reusable'){
                $ids=@([string](Probe 'CompleteSubmissionEvent'),[string](Probe 'CompleteSubmissionOutputEvent'))
                $selected=Owner 'CompleteSubmissionInputForTest';$output=Owner 'CompleteSubmissionOutputSelectionForTest'
                $keys=@([string]$selected[0],[string]$output[0]);$deltas=@(-[double]$selected[1],[double]$output[1])
                Pass ($Label+'.OwnerCompleted') ([bool](Probe 'CompleteBaselineCompleted'))
                Pass ($Label+'.ExactInputAndFreshOutput') ([bool](Owner 'CompleteBaselineBalancesForTest' @($true)) -and [bool](Owner 'CompleteBaselineOutputForTest'))
            }else{
                $ids=([string](Probe 'CompleteWorksheetEvents')).Split("`n")
                $keys=Probe 'CompletePathWorksheetKeys';$deltas=@(-10.0,1.0)
                foreach($fact in @('ExactInput','FreshOutput','Log','OutputCustom','CheckCustom')){Pass ($Label+'.'+$fact) ([bool](Probe 'CompleteWorksheetFact' @($fact,$Canary)))}
            }
            if($ids.Count -ne 2 -or @($ids|Where-Object {$_ -ceq ''}).Count -or $ids[0] -ceq $ids[1] -or $keys.Count -ne 2 -or $keys[0] -ceq $keys[1]){throw 'Distinct real owner identities unavailable.'}
            Delivered (RecordingControl 'Stop Recording' 'Click')
            $rows=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object{Get-Content $_ -Raw|ConvertFrom-Json})
            $attempt=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$terminal=@($rows|Where-Object OutcomeCode -CEQ 'CONFIRMED')
            $valid=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $terminal.Count -eq 1
            if($valid){$valid=$terminal[0].ControlId -ceq $control -and $terminal[0].OwnerId -ceq 'PRODUCTION_RUN_COMPLETION' -and $terminal[0].CatalogVersion -eq 28 -and $terminal[0].DataEffect -ceq 'Unknown' -and $attempt[0].ActivityId -ceq $terminal[0].ActivityId -and $terminal[0].SequenceId -ceq $starts[0].SequenceId -and $terminal[0].Ordinal -eq 1 -and ($terminal[0].SourceEventRefs.EventId -join '|') -ceq ($ids -join '|')}
            Pass ($Label+'.ExactRecordedCommandAndReferences') $valid
            if(-not $valid){return $null}
            $closed=@(RecordingJournal $starts[0].SequenceId|Where-Object RecordType -CEQ 'Close')
            $valid=$closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq 1 -and (JournalChain $starts[0].SequenceId 4)
            if($valid){$valid=($closed[0].Observations.RecordId -join '|') -ceq (@($attempt[0].RecordId,$terminal[0].RecordId) -join '|')}
            Pass ($Label+'.ExactStoppedJournal') $valid
            if(-not $valid){return $null}
            Test-CompletePathAuthority $Fixture $ids $keys $deltas ($prefix+$Label)|Out-Host
            Pass ($Label+'.CustomValueFormulaAndDecoy') ($sheet.Cells.Item(2,1).Value2 -ceq $Canary -and $sheet.Cells.Item(2,2).Formula -ceq '=1+2' -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
            [pscustomobject]@{Journal=$closed[0];Original=@($attempt[0],$terminal[0]);Terminal=$terminal[0];Ids=$ids;Keys=$keys;Deltas=$deltas}
        }finally{[void](Probe 'CompleteSubmissionReset');[void](Probe 'CloseDesigner');if($null -ne $work){$work.Close($false)}}
    }
    function SourceProof($Result,$Observed,$Publication,[string]$Label){
        $valid=$Result.ResultState -ceq 'Concluded' -and 'SOURCE_APPLIED' -cin @($Result.ReasonCodes) -and @($Result.TerminalSources).Count -eq 2
        foreach($id in $Observed.Ids){
            $i=[array]::IndexOf($Observed.Ids,$id)
            $source=@($Result.TerminalSources|Where-Object EventId -CEQ $id)
            $group=@($Publication.Groups|Where-Object {$_.Source -ceq 'Inventory' -and $_.SourceId -ceq $id})
            if($source.Count -ne 1 -or $group.Count -ne 1){$valid=$false;continue}
            $line=@($group[0].Lines);$body=[ordered]@{Lines=$line}|ConvertTo-Json -Depth 16 -Compress
            $valid=$valid -and $source[0].OwnerStatus -ceq 'Applied' -and $source[0].WarehouseId -ceq $Fixture.Warehouse -and $source[0].SourceKind -ceq 'Inventory' -and $source[0].SubmissionState -ceq 'Submitted' -and $source[0].LineCount -eq 1 -and ($source[0].SystemKeys -join '|') -ceq $Observed.Keys[$i] -and $source[0].LinesSha256 -ceq (EvaluationSha $body) -and $line.Count -eq 1 -and $line[0].System_Key -ceq $Observed.Keys[$i] -and [double]$line[0].QtyDelta -eq $Observed.Deltas[$i]
        }
        Pass ($Label+'.EveryFreshExactSourceApplied') $valid
        Pass ($Label+'.ExactPublicationProvenance') ($Result.Publication.PublicationId -ceq $Publication.PublicationId -and $Result.Publication.ContentSha256 -ceq $Publication.ContentSha256 -and $Result.Publication.WarehouseId -ceq $Fixture.Warehouse)
    }
    try{
        SelectTarget $Fixture
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($true))){throw 'Complete recording policy unavailable.'}
        SetRecordingPolicy $true
        if(-not [bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')){throw 'Initial owning publication unavailable.'}
        $eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json')
        $initial=Get-Content $eventsPath -Raw|ConvertFrom-Json
        SelectTarget $Fixture 'config-producer';OpenRecordingViewer
        $source=Record 'GuideSource';if($null -eq $source){return}
        $observed=Record 'ObservedRun';if($null -eq $observed){return}
        Pass 'FreshDistinctRunsAndBusinessEvents' ($source.Journal.ActionPathId -cne $observed.Journal.ActionPathId -and $source.Terminal.ActivityId -cne $observed.Terminal.ActivityId -and @($source.Ids|Where-Object {$_ -cin $observed.Ids}).Count -eq 0 -and $source.Keys[1] -cne $observed.Keys[1])
        $activityPins=ActivityPins;$journalPins=BoundPins $journalRoot
        # Keep the operator's originally loaded publication. Reopening would
        # legitimately load the newer snapshot emitted by the completion owner.
        Delivered (BoundLibrary 'Open')
        if((BoundLibrary 'Select' $observed.Journal.ActionPathId) -cne 'SELECTED'){throw 'Observed run unavailable.'}
        $stale=Evaluate 'SourceEventsApplied' 'BeforePublication';if($null -eq $stale){return}
        Pass 'OlderPublicationCannotProveFreshSources' ($stale.ResultState -ceq 'Awaiting' -and $stale.Publication.PublicationId -ceq $initial.PublicationId -and @($stale.TerminalSources).Count -eq 2 -and @($stale.TerminalSources|Where-Object OwnerStatus -CNE 'Awaiting').Count -eq 0)
        CloseRecordingViewer;SelectTarget $Fixture
        Pass 'ActualAdminPublication' ([bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest'))
        $publication=Get-Content $eventsPath -Raw|ConvertFrom-Json
        $authorityPins=@{};foreach($file in Get-ChildItem $Fixture.Root -Recurse -File|Where-Object {$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '*.Snapshot.*' -and $_.Name -notlike '~$*'}){$authorityPins[$file.FullName]=Hash $file.FullName}
        OpenRecordingViewer
        foreach($record in @($source.Original)+@($observed.Original)){
            $lines=@($publication.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -ceq $record.ActivityId}|ForEach-Object Lines|Where-Object RecordId -CEQ $record.RecordId)
            Pass ('Publication.'+$(if($record.SequenceId -ceq $source.Journal.SequenceId){'Source'}else{'Observed'})+'.'+$record.OutcomeCode) ($lines.Count -eq 1 -and ($lines[0]|ConvertTo-Json -Depth 20 -Compress) -ceq ($record|ConvertTo-Json -Depth 20 -Compress))
        }
        Delivered (BoundLibrary 'Open');[void](BoundLibrary 'Select' $observed.Journal.ActionPathId)
        $command=Evaluate 'CommandCompleted' 'Published';$applied=Evaluate 'SourceEventsApplied' 'Published'
        if($null -eq $command -or $null -eq $applied){return}
        Pass 'SeparateCommandConclusion' ($command.ResultState -ceq 'Concluded' -and 'COMMAND_COMPLETED' -cin @($command.ReasonCodes) -and @($command.TerminalSources).Count -eq 0)
        SourceProof $applied $observed $publication 'Evaluation'
        Pass 'ExactObservedMatch' ($applied.JournalRecordId -ceq $observed.Journal.RecordId -and @($applied.Matches).Count -eq 1 -and $applied.Matches[0].ActivityId -ceq $observed.Terminal.ActivityId -and @($applied.ExtraActivityIds).Count -eq 0)
        if((BoundLibrary 'Select' $source.Journal.ActionPathId) -cne 'SELECTED'){throw 'Guide source unavailable.'}
        Delivered (BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths')
        $instruction='Prepare valid '+$Mode.ToLowerInvariant()+' Production allocations, Check In, select the output and enter its actual quantity. Click Complete Run. Inspect the result before continuing; partial work is not automatically retried or rolled back.'
        Delivered (Author 'txtGuideName' 'Write' ('Complete '+$Mode.ToLowerInvariant()+' Production'))
        Delivered (Author 'txtGuideInstructions' 'Write' ($instruction+' Command completion and published Inventory application are separate conclusions.'))
        Pass 'OneObservedGuideStep' ((Author 'lstGuideSteps' 'Rows') -ceq '1')
        if((Author 'lstGuideSteps' 'Select' '0') -cne 'SELECTED'){throw 'Guide step unavailable.'}
        Delivered (Author 'txtGuideStepInstruction' 'Write' $instruction)
        Delivered (Author 'btnGuideExpectedConclusion' 'Click')
        foreach($pair in @(@('cboExpectedControl',$control),@('cboExpectedOutcome','CONFIRMED'))){Delivered (BoundControl $pair[0] 'Write' $pair[1] 'frmActionPathExpectation')}
        Delivered (BoundControl 'btnAddExpectedStep' 'Click' '' 'frmActionPathExpectation')
        if((BoundControl 'cboTerminalStep' 'Select' '0' 'frmActionPathExpectation') -cne 'SELECTED'){throw 'Guide terminal unavailable.'}
        Delivered (BoundControl 'cboTerminalKind' 'Write' 'SourceEventsApplied' 'frmActionPathExpectation');Delivered (BoundControl 'btnUseExpectation' 'Click' '' 'frmActionPathExpectation')
        $guides=Join-Path $journalRoot 'Guides';$prior=BoundPins $guides;Delivered (Author 'btnSaveGuide' 'Click')
        $fresh=@(Get-ChildItem $guides -File -Filter '*.json'|Where-Object {-not $prior.ContainsKey($_.FullName)})
        if($fresh.Count -ne 1){throw 'Saved guide unavailable.'}
        $guide=ReadGuideExpectationRecord $fresh[0].FullName;if($null -eq $guide){throw 'Guide integrity unavailable.'}
        Pass 'ImmutableGuideExactSourceAndExplicitConclusion' ($guide.SourceRun.RecordId -ceq $source.Journal.RecordId -and $guide.SourceRun.ContentSha256 -ceq $source.Journal.ContentSha256 -and $guide.Steps[0].SourceActivityId -ceq $source.Terminal.ActivityId -and $guide.ExpectedConclusion.TerminalKind -ceq 'SourceEventsApplied')
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        Pass 'ReaderCannotAuthor' (-not [bool](Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-reader',$Fixture.Warehouse,'S1')))
        Delivered (BoundLibrary 'Open');if((BoundLibrary 'Select' $observed.Journal.ActionPathId) -cne 'SELECTED'){throw 'Separate observed run unavailable.'}
        BoundOpen;BoundSelect $guide;Delivered (BoundControl 'btnUseGuideForRun' 'Click');Delivered (BoundControl 'btnCloseGuides' 'Click')
        $prior=@(EvaluationFiles|ForEach-Object FullName);Delivered (BoundControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
        $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $prior});if($fresh.Count -ne 1){throw 'Paired evaluation unavailable.'}
        $paired=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
        SourceProof $paired $observed $publication 'GuidePair'
        Pass 'PairUsesObservedRunAndExactGuide' ($paired.JournalRecordId -ceq $observed.Journal.RecordId -and $paired.Guide.ContentSha256 -ceq $guide.ContentSha256 -and @($paired.Matches).Count -eq 1 -and $paired.Matches[0].ActivityId -ceq $observed.Terminal.ActivityId)
        $trainingPins=BoundPins $journalRoot;$publishedPin=Hash $eventsPath
        Delivered (BoundControl 'btnViewActionPath' 'Click' '' 'frmActionPaths')
        $pair=View 'lblActionPathPair' 'Label';$how=View 'txtActionPathHowTo' 'Text';$diagnostic=View 'txtActionPathDiagnostic' 'Text'
        Pass 'ExactPairAndAuthoredInstructionVisible' ($pair.Contains($guide.ContentSha256) -and $pair.Contains($observed.Journal.ActionPathId) -and $how.Contains($instruction))
        Pass 'DiagnosticUsesObservedEvidence' ($diagnostic.Contains($observed.Terminal.ActivityId) -and -not $diagnostic.Contains($source.Terminal.ActivityId) -and $diagnostic.Contains($paired.EvaluationId))
        foreach($method in @('How-To','Diagnostic','Compare both')){
            Delivered (View 'cboActionPathView' 'Write' $method)
            Pass ('View.'+$method) ((View 'cboActionPathView' 'Selected') -ceq $method -and ((View 'txtActionPathHowTo' 'State') -split '\|')[0] -ceq $(if($method -ceq 'Diagnostic'){'False'}else{'True'}) -and ((View 'txtActionPathDiagnostic' 'State') -split '\|')[0] -ceq $(if($method -ceq 'How-To'){'False'}else{'True'}))
            Pass ('SameEvidence.'+$method) ((View 'lblActionPathPair' 'Label') -ceq $pair -and (View 'txtActionPathHowTo' 'Text') -ceq $how -and (View 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ('complete-'+$Mode.ToLowerInvariant()+'-'+$method.Replace(' ','-').ToLowerInvariant()+'.png')
        }
        Delivered (View 'txtActionPathDiagnostic' 'ViewportBottom');CaptureOwnedFormByCaptionEvidence 'Action Path view' ('complete-'+$Mode.ToLowerInvariant()+'-conclusion.png')
        Pass 'ViewingPreservesEvidence' ((BoundSame $trainingPins (BoundPins $journalRoot)) -and (PinsRetained $activityPins) -and (Hash $eventsPath) -ceq $publishedPin)
        $same=$true;foreach($file in $journalPins.Keys){$same=$same -and (Hash $file) -ceq $journalPins[$file]};Pass 'OriginalJournalsImmutable' $same
        $same=$true;foreach($file in $authorityPins.Keys){$same=$same -and (Hash $file) -ceq $authorityPins[$file]};Pass 'EvaluationAndViewingPreserveAuthority' $same
    }finally{[void](Probe 'CloseDesigner');CloseRecordingViewer;SelectTarget $Fixture}
}

function Test-CompletePathAuthority($Fixture,[string[]]$Ids,$Keys,$Deltas,[string]$Prefix){
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $before=@($excel.Workbooks|Where-Object {$_.FullName -ieq $path});$pin=Hash $path
    if($before.Count -gt 1){throw 'Duplicate Inventory authority.'}
    $opened=$before.Count -eq 0
    if($opened){$authority=$excel.Workbooks.Open($path,0,$true)}else{$authority=$before[0]}
    try{
        $applied=Table $authority 'tblAppliedEvents';$audit=Table $authority 'tblInventoryLog';$exact=$true
        for($i=0;$i -lt $Ids.Count;$i++){
            $events=@($applied.ListRows|Where-Object {$_.Range.Cells.Item(1,$applied.ListColumns.Item('EventID').Index).Value2 -ceq $Ids[$i]})
            $rows=@($audit.ListRows|Where-Object {$_.Range.Cells.Item(1,$audit.ListColumns.Item('EventID').Index).Value2 -ceq $Ids[$i]})
            $exact=$exact -and $events.Count -eq 1 -and $rows.Count -eq 1
            if($rows.Count -eq 1){$exact=$exact -and $rows[0].Range.Cells.Item(1,$audit.ListColumns.Item('System_Key').Index).Value2 -ceq $Keys[$i] -and [double]$rows[0].Range.Cells.Item(1,$audit.ListColumns.Item('QtyDelta').Index).Value2 -eq $Deltas[$i]}
        }
        Check ($Prefix+'.IndependentExactAppliedEventsAndAudit') $exact
    }finally{if($opened){$authority.Close($false)}}
    Check ($Prefix+'.ReadOnlyAuthorityInspection') ((Hash $path) -ceq $pin -and @($excel.Workbooks|Where-Object {$_.FullName -ieq $path}).Count -eq $before.Count)
}
