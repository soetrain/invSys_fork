# Original instruction handlers supply separate guide provenance and observed runs.
function Test-ProductionInstructionPaths($Fixture,[switch]$Uom,[switch]$Components,[switch]$RecipeOrder,[switch]$RecipeStructure) {
    if(([int][bool]$Uom+[int][bool]$Components+[int][bool]$RecipeOrder+[int][bool]$RecipeStructure) -gt 1){throw 'Select one Production path family.'}
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Author([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathGuide'}
    function View([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathView'}
    function Expected([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathExpectation'}
    function PathCheck([string]$Name,[bool]$Passed){
        if($Uom){$Name=$Name.Replace('InstructionPaths.','UomPaths.').Replace('Five','Two')}
        if($Components){$Name=$Name.Replace('InstructionPaths.','ComponentPaths.').Replace('Five','Ten')}
        if($RecipeOrder){$Name=$Name.Replace('InstructionPaths.','RecipeOrderPaths.').Replace('Five','Three')}
        if($RecipeStructure){$Name=$Name.Replace('InstructionPaths.','RecipeStructurePaths.')}
        Check $Name $Passed
    }
    function ActionId([string]$Action){if($Uom){'PRODUCTION_UOM_EDIT'}elseif($Components){'PRODUCTION_PROCESS_'+$Action}elseif($RecipeStructure){'PRODUCTION_RECIPE_'+$Action}elseif($RecipeOrder){if($Action -ceq 'AUTO'){'PRODUCTION_RECIPE_AUTO_ORDER'}else{'PRODUCTION_RECIPE_MOVE_'+$Action}}else{'PRODUCTION_PROCESS_INSTRUCTION_'+$Action}}
    function ActionOutcome([string]$Action){if($Uom){if($Action -ceq 'OPEN'){'OPENED'}else{'REUSED'}}else{'STAGED'}}
    function Instruction([string]$Action){
        if($Uom){$Action+' the UOM workbench; retain existing edits and verify that the saved catalog is unchanged.'}
        elseif($Components){$Action.Replace('_',' ')+' in the local Process draft; for UPDATE, select the added row first. Verify the staged outcome.'}
        elseif($RecipeStructure){
            switch($Action){
                'DISCONNECT'{'Select the connected source output and click Disconnect. Verify the local connection is removed.'}
                'CONNECT'{'Select the source Process output and its compatible Feeds Process target, then click Connect.'}
                'UPDATE_CONNECTION'{'With that connection selected, edit Required Qty and Required %, click Update and verify both values.'}
                'ADD_PROCESS'{'Select a released Process and click Add Process. Verify the new local Recipe node is selected.'}
                'REMOVE_PROCESS'{'Click Remove Process to remove the newly added node. Inspect the remaining draft before validation and Save.'}
            }
        }
        elseif($RecipeOrder){
            switch($Action){
                'UP'{'Select a recipe node and click Move Up. Verify its local order; saved definitions remain unchanged.'}
                'DOWN'{'Keep that node selected and click Move Down. Verify its local order; saved definitions remain unchanged.'}
                'AUTO'{'Click Auto Order and inspect the local recipe order before validation and Save.'}
            }
        }
        else{$Action+' the local instruction and verify the staged outcome.'}
    }
    function SavedHash([string]$Path){
        $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
        try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
    }
    function Delivered([string]$Result){
        if($Result -ceq 'DELIVERED'){return}
        $kind=if($Result -cin @('DISABLED','MISSING','OUT OF RANGE','NOT FOUND','SELECTED')){$Result}else{'OTHER'}
        [pscustomobject]@{CallerLine=$MyInvocation.ScriptLineNumber;ResultClass=$kind}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'instruction-path-prerequisite.json')
        throw 'Existing packaged UI prerequisite unavailable; not behavioral RED.'
    }
    function RecordSeries([string]$Label,[ref]$Recorded){
        if($Uom){
            # Reset only this disposable workbench between independent recordings.
            if(-not [string]::Equals([string]$book.FullName,[IO.Path]::GetFullPath($path),[StringComparison]::OrdinalIgnoreCase)){throw 'Unexpected UOM path fixture workbook.'}
            $priorAlerts=$excel.DisplayAlerts
            try{
                $excel.DisplayAlerts=$false
                foreach($stageSheet in @($book.Worksheets)){if($stageSheet.Name -ceq 'invSys UOM Catalog'){$stageSheet.Delete()}}
            }finally{$excel.DisplayAlerts=$priorAlerts}
        }elseif($Components){[void](Probe 'ComponentStage' @('REQUIREMENT',$canary,'Valid'))}
        elseif($RecipeOrder){[void](Probe 'OrderStage' @($canary,'Unordered'))}
        elseif($RecipeStructure){[void](Probe 'StructurePathReset')}
        else{[void](Probe 'InstructionStage' @($canary,1,(' '+$canary+'4 ')))}
        $prior=@(if(Test-Path $journalRoot){Get-ChildItem $journalRoot -File -Filter '*.json'|ForEach-Object FullName})
        Delivered (RecordingControl 'Start Recording' 'Click')
        $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {$_.FullName -cnotin $prior}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
        if($starts.Count -ne 1){throw 'Actual recording Start unavailable.'}
        $original=@();$terminals=@();$ordinal=0
        foreach($action in $actions){
            $ordinal++;$before=@(Get-Slice4beActivityFiles $Fixture);$decoy.Activate()
            if($Uom){[void](Probe 'SendUom')}
            elseif($Components){
                $parts=$action.Split('_')
                if($parts[1] -ceq 'UPDATE'){[void](Probe 'ComponentSelectAdded' @($parts[0],$canary))}
                [void](Probe 'ComponentAct' @($parts[0],$parts[1]))
            }elseif($RecipeOrder){[void](Probe 'OrderAct' @($action))}
            elseif($RecipeStructure){[void](Probe 'StructurePathInput' @($action));[void](Probe 'StructureAct' @($action))}
            else{[void](Probe 'InstructionAct' @($action))}
            $rows=@(Get-Slice4beActivityFiles $Fixture|Where-Object {$_ -cnotin $before}|ForEach-Object {Get-Content $_ -Raw|ConvertFrom-Json})
            $attempt=@($rows|Where-Object OutcomeCode -CEQ 'REQUESTED');$terminal=@($rows|Where-Object OutcomeCode -CEQ (ActionOutcome $action))
            $valid=$rows.Count -eq 2 -and $attempt.Count -eq 1 -and $terminal.Count -eq 1
            if($valid){$valid=$terminal[0].ControlId -ceq (ActionId $action) -and $attempt[0].ActivityId -ceq $terminal[0].ActivityId -and $terminal[0].SequenceId -ceq $starts[0].SequenceId -and $terminal[0].Ordinal -eq $ordinal -and @($terminal[0].SourceEventRefs).Count -eq 0}
            PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.OriginalHandlerPair') $valid
            if(-not $valid){throw 'Original instruction recording incomplete; retain behavioral failures.'}
            $original+=@($attempt[0],$terminal[0]);$terminals+=$terminal[0]
        }
        Delivered (RecordingControl 'Stop Recording' 'Click')
        if($RecipeStructure){PathCheck ('InstructionPaths.'+$Label+'.SeriesReachesExpectedLocalDraft') ([bool](Probe 'StructurePathResult'))}
        $closed=@(RecordingJournal $starts[0].SequenceId|Where-Object RecordType -CEQ 'Close')
        $valid=$closed.Count -eq 1 -and $closed[0].Lifecycle -ceq 'Stopped' -and $closed[0].ActionCount -eq $actions.Count -and (JournalChain $starts[0].SequenceId ($actions.Count*2+2))
        if($valid){$valid=($closed[0].Observations.RecordId -join '|') -ceq ($original.RecordId -join '|')}
        PathCheck ('InstructionPaths.'+$Label+'.ExactFiveActionJournal') $valid
        if(-not $valid){throw 'Original five-action journal unavailable.'}
        $Recorded.Value=[pscustomobject]@{Journal=$closed[0];Terminals=$terminals;Original=$original}
    }
    $actions=if($Uom){@('OPEN','REOPEN')}else{@('ADD','UPDATE','UP','DOWN','REMOVE')}
    if($Components){$actions=@('REQUIREMENT_ADD','REQUIREMENT_UPDATE','REQUIREMENT_UP','REQUIREMENT_DOWN','REQUIREMENT_REMOVE','OUTPUT_ADD','OUTPUT_UPDATE','OUTPUT_UP','OUTPUT_DOWN','OUTPUT_REMOVE')}
    if($RecipeOrder){$actions=@('UP','DOWN','AUTO')}
    if($RecipeStructure){$actions=@('DISCONNECT','CONNECT','UPDATE_CONNECTION','ADD_PROCESS','REMOVE_PROCESS')}
    $canary='PATH'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null
    $imagePrefix=if($Uom){'uom'}elseif($Components){'component'}elseif($RecipeOrder){'recipe-order'}elseif($RecipeStructure){'recipe-structure'}else{'instruction'}
    try {
        SelectTarget $Fixture
        $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
        PathCheck 'InstructionPaths.Setup.AuthorHasExplicitCapability' ($allowed -is [bool] -and $allowed)
        if($allowed -isnot [bool] -or -not $allowed){throw 'Explicit guide author capability fixture unavailable; not behavioral RED.'}
        SetRecordingPolicy $true;SelectTarget $Fixture 'config-producer';OpenRecordingViewer
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $path=Join-Path $runRoot 'instruction-path-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $workbookPin=SavedHash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();[void](Probe 'OpenDesigner' @($book.Name))
        if($RecipeStructure){
            if([string](Probe 'StructureReleasedSetup') -cne 'READY' -or -not [bool](Probe 'StructureReleasedNodes')){throw 'Released Process path fixture unavailable; not behavioral RED.'}
            [void](Probe 'StructurePathRemember')
        }
        $authorityPins=@{}
        # Publication owns snapshot projections; canonical files remain byte-pinned.
        foreach($file in Get-ChildItem $Fixture.Root -Recurse -File|Where-Object {$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '*.Snapshot.*' -and -not $_.Name.StartsWith('~$')}){$authorityPins[$file.FullName]=SavedHash $file.FullName}
        $source=$null;$observed=$null;RecordSeries 'GuideSource' ([ref]$source);RecordSeries 'ObservedRun' ([ref]$observed)
        PathCheck 'InstructionPaths.DistinctGuideAndObservedRuns' ($source.Journal.ActionPathId -cne $observed.Journal.ActionPathId -and @($source.Terminals|Where-Object {$_.ActivityId -cin $observed.Terminals.ActivityId}).Count -eq 0)
        $window=[long](Probe 'Present' @($(if($Uom){5}elseif($RecipeOrder -or $RecipeStructure){1}else{0})));CaptureOwnedFormEvidence 'Production' ($imagePrefix+'-editor.png') $window
        $activityPins=ActivityPins;$journalPins=BoundPins $journalRoot
        [void](Probe 'CloseDesigner');CloseRecordingViewer;SelectTarget $Fixture
        $published=[bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')
        PathCheck 'InstructionPaths.ActualAdminPublication' $published
        if(-not $published){throw 'Owning publication unavailable.'}
        $eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json');$publication=Get-Content $eventsPath -Raw|ConvertFrom-Json
        OpenRecordingViewer
        foreach($record in @($source.Original)+@($observed.Original)){
            $group=@($publication.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -ceq $record.ActivityId})
            $lines=@($group|ForEach-Object Lines|Where-Object RecordId -CEQ $record.RecordId)
            PathCheck ('InstructionPaths.Publication.'+$(if($record.SequenceId -ceq $source.Journal.SequenceId){'GuideSource'}else{'ObservedRun'})+'.'+$record.Ordinal+'.'+$record.OutcomeCode) ($lines.Count -eq 1 -and ($lines[0]|ConvertTo-Json -Depth 20 -Compress) -ceq ($record|ConvertTo-Json -Depth 20 -Compress))
        }
        foreach($record in $observed.Terminals){
            $id=[string]$record.ActivityId
            PathCheck ('InstructionPaths.Detail.'+$record.ControlId+$(if($Uom){'.'+$record.Ordinal}else{''})) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('SelectSource',$id)) -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadDetailForTest' @('Source event / activity ID',$id)))
        }
        if($Uom -or $Components -or $RecipeOrder -or $RecipeStructure){
            # Published contributing lines are not ordered by outcome. Select
            # the exact terminal record, rather than assuming it is row two.
            $detailGroup=@($publication.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -ceq $observed.Terminals[-1].ActivityId})
            if($detailGroup.Count -ne 1){throw 'Exact terminal detail group unavailable.'}
            $terminalIndex=-1
            for($lineIndex=0;$lineIndex -lt $detailGroup[0].Lines.Count;$lineIndex++){
                if($detailGroup[0].Lines[$lineIndex].RecordId -ceq $observed.Terminals[-1].RecordId){$terminalIndex=$lineIndex;break}
            }
            if($terminalIndex -lt 0){throw 'Exact terminal detail line unavailable.'}
            if($Components -or $RecipeOrder -or $RecipeStructure){
                $requestedIndex=-1
                for($lineIndex=0;$lineIndex -lt $detailGroup[0].Lines.Count;$lineIndex++){
                    if($detailGroup[0].Lines[$lineIndex].OutcomeCode -ceq 'REQUESTED'){$requestedIndex=$lineIndex;break}
                }
                if($requestedIndex -lt 0){throw 'Exact requested detail line unavailable.'}
                Delivered (ExpectationControl 'lstEventLines' 'Index' ([string]$requestedIndex) 'frmEventDetail')
                PathCheck 'InstructionPaths.Detail.RequestedObservationMatchesVisibleSelection' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedSelectedFieldForTest' @('Outcome','REQUESTED')) -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedSelectedFieldForTest' @('Data effect','Unknown')) -and -not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedSelectedFieldForTest' @('Outcome','STAGED')))
            }
            Delivered (ExpectationControl 'lstEventLines' 'Index' ([string]$terminalIndex) 'frmEventDetail')
            if($Components -or $RecipeOrder -or $RecipeStructure){
                PathCheck 'InstructionPaths.Detail.ExactTerminalOutcomeAndEffect' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedSelectedFieldForTest' @('Outcome','STAGED')) -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedSelectedFieldForTest' @('Data effect','Unchanged')))
            }
        }
        CaptureOwnedFormEvidence 'Event Detail' ($imagePrefix+'-event-detail.png') ([long](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedDetailLabelForTest' @('Window','')))
        Delivered (BoundLibrary 'Open')
        if((BoundLibrary 'Select' $observed.Journal.ActionPathId) -cne 'SELECTED'){throw 'Observed run unavailable.'}
        foreach($kind in @('CommandCompleted','SourceEventsApplied')){
            $ready=Set-EvaluationDraft @(,@((ActionId $actions[-1]),(ActionOutcome $actions[-1]),'True')) 0 $kind -StopAtMissingChoice
            PathCheck ('InstructionPaths.TerminalChoice.'+$kind) $ready
            if(-not $ready){continue}
            $before=@(EvaluationFiles|ForEach-Object FullName);Delivered (ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before});if($fresh.Count -ne 1){throw 'Actual evaluation unavailable.'}
            $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
            $valid=if($kind -ceq 'CommandCompleted'){$result.ResultState -ceq 'Concluded' -and 'COMMAND_COMPLETED' -cin @($result.ReasonCodes)}else{$result.ResultState -ceq 'Incomplete' -and 'SOURCE_UNAVAILABLE' -cin @($result.ReasonCodes)}
            PathCheck ('InstructionPaths.ExactTerminal.'+$kind) ($valid -and $result.JournalRecordId -ceq $observed.Journal.RecordId -and $result.Matches[0].ActivityId -ceq $observed.Terminals[-1].ActivityId)
        }
        if((BoundLibrary 'Select' $source.Journal.ActionPathId) -cne 'SELECTED'){throw 'Guide source selection unavailable.'}
        Delivered (BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths')
        PathCheck 'InstructionPaths.FiveObservedGuideSteps' ((Author 'lstGuideSteps' 'Rows') -ceq [string]$actions.Count)
        Delivered (Author 'txtGuideName' 'Write' $(if($Uom){'Open and reuse the UOM draft'}elseif($Components){'Edit Process requirements and outputs'}elseif($RecipeOrder){'Order a local Recipe draft'}elseif($RecipeStructure){'Edit local Recipe connections and nodes'}else{'Edit Process instructions'}))
        Delivered (Author 'txtGuideInstructions' 'Write' $(if($Uom){'Open the UOM workbench and reuse the existing local draft without publishing the catalog.'}elseif($Components){'Add, select and update, move up/down, then remove a requirement and an output. Validate before saving the Process.'}elseif($RecipeOrder){'Move the selected recipe node up, down, then apply Auto Order. Inspect the local draft before validation and Save.'}elseif($RecipeStructure){'Disconnect and reconnect the source output, update its required quantities, then add and remove a local Recipe node. Inspect and validate before Save.'}else{'Edit the local instruction draft, then validate before saving the Process.'}))
        for($i=0;$i -lt $actions.Count;$i++){
            if((Author 'lstGuideSteps' 'Select' ([string]$i)) -cne 'SELECTED'){throw 'Guide step unavailable.'}
            Delivered (Author 'txtGuideStepInstruction' 'Write' (Instruction $actions[$i]))
        }
        Delivered (Author 'btnGuideExpectedConclusion' 'Click')
        foreach($action in $actions){Delivered (Expected 'cboExpectedControl' 'Write' (ActionId $action));Delivered (Expected 'cboExpectedOutcome' 'Write' (ActionOutcome $action));Delivered (Expected 'btnAddExpectedStep' 'Click')}
        if((Expected 'cboTerminalStep' 'Select' ([string]($actions.Count-1))) -cne 'SELECTED'){throw 'Final terminal step unavailable.'}
        Delivered (Expected 'cboTerminalKind' 'Write' 'CommandCompleted');Delivered (Expected 'btnUseExpectation' 'Click')
        $guideRoot=Join-Path $journalRoot 'Guides';$before=BoundPins $guideRoot;Delivered (Author 'btnSaveGuide' 'Click')
        $fresh=@(Get-ChildItem $guideRoot -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)})
        if($fresh.Count -ne 1){throw 'Immutable guide Save unavailable.'}
        $guide=ReadGuideExpectationRecord $fresh[0].FullName;if($null -eq $guide){throw 'Saved guide integrity unavailable.'}
        PathCheck 'InstructionPaths.ExactSourceAndExplicitIntent' ($guide.SourceRun.RecordId -ceq $source.Journal.RecordId -and $guide.SourceRun.ContentSha256 -ceq $source.Journal.ContentSha256 -and ($guide.Steps.SourceActivityId -join '|') -ceq ($source.Terminals.ActivityId -join '|') -and $guide.ExpectedConclusion.TerminalKind -ceq 'CommandCompleted' -and @($guide.ExpectedConclusion.Steps).Count -eq $actions.Count)
        CloseRecordingViewer;SelectTarget $Fixture 'config-reader';OpenRecordingViewer
        $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-reader',$Fixture.Warehouse,'S1')
        PathCheck 'InstructionPaths.ReaderHasNoAuthorCapability' ($allowed -is [bool] -and -not $allowed)
        Delivered (BoundLibrary 'Open')
        if((BoundLibrary 'Select' $observed.Journal.ActionPathId) -cne 'SELECTED'){throw 'Separate observed run unavailable.'}
        BoundOpen;BoundSelect $guide;Delivered (BoundControl 'btnUseGuideForRun' 'Click');Delivered (BoundControl 'btnCloseGuides' 'Click')
        $evaluationRoot=Join-Path $journalRoot 'Evaluations';$before=BoundPins $evaluationRoot;Delivered (BoundControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
        $fresh=@(Get-ChildItem $evaluationRoot -File -Filter '*.json'|Where-Object {-not $before.ContainsKey($_.FullName)});if($fresh.Count -ne 1){throw 'Paired Evaluate unavailable.'}
        $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
        PathCheck 'InstructionPaths.PairedConclusionUsesExactObservedRun' ($result.ResultState -ceq 'Concluded' -and $result.JournalRecordId -ceq $observed.Journal.RecordId -and $result.Guide.ContentSha256 -ceq $guide.ContentSha256 -and ($result.Matches.ActivityId -join '|') -ceq ($observed.Terminals.ActivityId -join '|') -and @($result.ExtraActivityIds).Count -eq 0)
        $trainingPins=BoundPins $journalRoot;$publishedHash=(Get-FileHash $eventsPath).Hash
        Delivered (BoundControl 'btnViewActionPath' 'Click' '' 'frmActionPaths')
        $pair=View 'lblActionPathPair' 'Label';$how=View 'txtActionPathHowTo' 'Text';$diagnostic=View 'txtActionPathDiagnostic' 'Text'
        PathCheck 'InstructionPaths.ExactPairNamed' ($pair.Contains([string]$guide.ContentSha256) -and $pair.Contains([string]$observed.Journal.ActionPathId))
        $instructions=$true;foreach($action in $actions){$instructions=$instructions -and $how.Contains((Instruction $action))}
        PathCheck 'InstructionPaths.AllAuthoredInstructionsInHowTo' $instructions
        $ordered=$true;$last=-1;foreach($record in $observed.Terminals){$index=$diagnostic.IndexOf([string]$record.ActivityId,[StringComparison]::Ordinal);$ordered=$ordered -and $index -gt $last;$last=$index}
        foreach($record in $source.Terminals){$ordered=$ordered -and -not $diagnostic.Contains([string]$record.ActivityId)}
        PathCheck 'InstructionPaths.DiagnosticUsesObservedOrder' $ordered
        PathCheck 'InstructionPaths.DiagnosticPreservesLocalOnlyConclusion' ($diagnostic.Contains([string]$result.EvaluationId) -and $diagnostic.Contains('Command completed; Domain application not asserted'))
        foreach($method in @('How-To','Diagnostic','Compare both')){
            Delivered (View 'cboActionPathView' 'Write' $method)
            PathCheck ('InstructionPaths.Method.'+$method) ((View 'cboActionPathView' 'Selected') -ceq $method -and ((View 'txtActionPathHowTo' 'State') -split '\|')[0] -ceq $(if($method -ceq 'Diagnostic'){'False'}else{'True'}) -and ((View 'txtActionPathDiagnostic' 'State') -split '\|')[0] -ceq $(if($method -ceq 'How-To'){'False'}else{'True'}))
            PathCheck ('InstructionPaths.SameEvidence.'+$method) ((View 'lblActionPathPair' 'Label') -ceq $pair -and (View 'txtActionPathHowTo' 'Text') -ceq $how -and (View 'txtActionPathDiagnostic' 'Text') -ceq $diagnostic)
            CaptureOwnedFormByCaptionEvidence 'Action Path view' ($imagePrefix+'-'+$method.Replace(' ','-').ToLowerInvariant()+'.png')
        }
        Delivered (View 'txtActionPathDiagnostic' 'ViewportBottom');CaptureOwnedFormByCaptionEvidence 'Action Path view' ($imagePrefix+'-conclusion.png')
        PathCheck 'InstructionPaths.ViewingPreservesEvidence' ((BoundSame $trainingPins (BoundPins $journalRoot)) -and (PinsRetained $activityPins) -and (Get-FileHash $eventsPath).Hash -ceq $publishedHash)
        $same=$true;foreach($file in $journalPins.Keys){$same=$same -and (Get-FileHash $file).Hash -ceq $journalPins[$file]}
        PathCheck 'InstructionPaths.OriginalJournalsImmutable' $same
        $same=$true;foreach($file in $authorityPins.Keys){$same=$same -and (SavedHash $file) -ceq $authorityPins[$file]}
        PathCheck 'InstructionPaths.SavedAuthorityPreserved' $same
        PathCheck 'InstructionPaths.UnknownColumnAndWorkbookBytes' ($sheet.Cells.Item(1,1).Value2 -ceq 'Operator Extra' -and $sheet.Cells.Item(2,1).Value2 -ceq $canary -and (SavedHash $path) -ceq $workbookPin)
    }finally{
        [void](Probe 'CloseDesigner');CloseRecordingViewer
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
