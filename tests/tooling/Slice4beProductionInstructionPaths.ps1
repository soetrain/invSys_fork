# Original instruction handlers supply separate guide provenance and observed runs.
function Test-ProductionInstructionPaths($Fixture,[switch]$Uom,[switch]$Components,[switch]$RecipeOrder,[switch]$RecipeStructure,[switch]$DesignReads,[switch]$Regulation,[switch]$CloseActions,[switch]$Worksheet,[switch]$Assignment,[switch]$RunPresentation,[switch]$RunClear,[switch]$RunLoad,[switch]$RunRefresh) {
    if(([int][bool]$Uom+[int][bool]$Components+[int][bool]$RecipeOrder+[int][bool]$RecipeStructure+[int][bool]$DesignReads+[int][bool]$Regulation+[int][bool]$CloseActions+[int][bool]$Worksheet+[int][bool]$Assignment+[int][bool]$RunPresentation+[int][bool]$RunClear+[int][bool]$RunLoad+[int][bool]$RunRefresh) -gt 1){throw 'Select one Production path family.'}
    . (Join-Path $PSScriptRoot 'Slice4beRecordingFixture.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beGuideTestActions.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beEvaluationContracts.ps1')
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Author([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathGuide'}
    function View([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathView'}
    function Expected([string]$Name,[string]$Action,[string]$Value=''){BoundControl $Name $Action $Value 'frmActionPathExpectation'}
    function PathCheck([string]$Name,[bool]$Passed){
        if($RunRefresh){$Name=$Name.Replace('InstructionPaths.','RunRefreshPaths.')}
        if($RunLoad){$Name=$Name.Replace('InstructionPaths.','RunLoadPaths.').Replace('Five','One')}
        if($RunClear){$Name=$Name.Replace('InstructionPaths.','RunClearPaths.').Replace('Five','Two')}
        if($RunPresentation){$Name=$Name.Replace('InstructionPaths.','RunPresentationPaths.').Replace('Five','Two')}
        if($Uom){$Name=$Name.Replace('InstructionPaths.','UomPaths.').Replace('Five','Two')}
        if($Components){$Name=$Name.Replace('InstructionPaths.','ComponentPaths.').Replace('Five','Ten')}
        if($RecipeOrder){$Name=$Name.Replace('InstructionPaths.','RecipeOrderPaths.').Replace('Five','Three')}
        if($RecipeStructure){$Name=$Name.Replace('InstructionPaths.','RecipeStructurePaths.')}
        if($DesignReads){$Name=$Name.Replace('InstructionPaths.','DesignReadPaths.')}
        if($Regulation){$Name=$Name.Replace('InstructionPaths.','RegulationPaths.').Replace('Five','Four')}
        if($CloseActions){$Name=$Name.Replace('InstructionPaths.','ClosePaths.').Replace('Five','Two')}
        if($Assignment){$Name=$Name.Replace('InstructionPaths.','AssignmentPaths.').Replace('Five','Twelve')}
        if($Worksheet){$Name=$Name.Replace('InstructionPaths.','WorksheetPaths.').Replace('Five','Four')}
        Check $Name $Passed
    }
    function ActionId([string]$Action){if($RunRefresh){'PRODUCTION_RUN_'+($Action -replace '^(Reusable|Worksheet)_','')}elseif($RunLoad){'PRODUCTION_RUN_LOAD'}elseif($RunClear){'PRODUCTION_RUN_CLEAR'}elseif($RunPresentation){'PRODUCTION_RUN_'+$Action}elseif($Assignment){'PRODUCTION_ASSIGNMENT_'+$Action}elseif($Worksheet){'PRODUCTION_PROCESS_WORKSHEET_'+$Action}elseif($CloseActions){'PRODUCTION_CLOSE'}elseif($Regulation){'PRODUCTION_OUTPUT_REGULATION_'+$Action.Split('_')[1]}elseif($DesignReads){'PRODUCTION_'+$Action}elseif($Uom){'PRODUCTION_UOM_EDIT'}elseif($Components){'PRODUCTION_PROCESS_'+$Action}elseif($RecipeStructure){'PRODUCTION_RECIPE_'+$Action}elseif($RecipeOrder){if($Action -ceq 'AUTO'){'PRODUCTION_RECIPE_AUTO_ORDER'}else{'PRODUCTION_RECIPE_MOVE_'+$Action}}else{'PRODUCTION_PROCESS_INSTRUCTION_'+$Action}}
    function ActionOutcome([string]$Action){if($RunRefresh){if($Action -ceq 'CLEAR'){'STAGED'}else{'REFRESHED'}}elseif($RunClear){'STAGED'}elseif($RunPresentation){'PRESENTED'}elseif($Assignment){switch($Action){'REFRESH'{'REFRESHED'}{$_ -cin @('PROCESS','PROCESS_SELECT')}{'PRESENTED'}{$_ -cin @('REQUIREMENT','REQUIREMENT_SELECT')}{'SELECTED'}'SAVE'{'CONFIRMED'}default{'STAGED'}}}elseif($Worksheet){if($Action -ceq 'RETRIEVE'){'CONFIRMED'}else{'STAGED'}}elseif($CloseActions){'CLOSED'}elseif($DesignReads){if($Action.EndsWith('REFRESH')){'REFRESHED'}elseif($Action -ceq 'PROCESS_REUSE'){'STAGED'}else{'PRESENTED'}}elseif($Uom){if($Action -ceq 'OPEN'){'OPENED'}else{'REUSED'}}else{'STAGED'}}
    function Instruction([string]$Action){
        if($RunRefresh){
            if($Action -ceq 'CLEAR'){'Click Clear Run to remove the loaded reusable run. Inspect the remaining worksheet staging before continuing.'}
            else{
                $button=if($Action.EndsWith('LOADER_REFRESH')){'Recipe Loader Refresh'}else{'Run Manager Refresh'}
                $scope=if($Action.StartsWith('Reusable')){'loaded reusable run'}else{'remaining worksheet staging'}
                'Click '+$button+' and inspect the '+$scope+'. A completed local refresh does not prove source availability, freshness or inventory application.'
            }
            return
        }
        if($RunLoad){
            'With a released Recipe version already selected and a valid batch scale entered, click Load Recipe. Inspect the loaded Process, ingredient and output rows before allocation or check-in. Loading replaces prior local run preparation; it does not apply inventory changes.'
            return
        }
        if($RunClear){
            if($Action -ceq 'Reusable'){'With a reusable run and separate worksheet staging already prepared, click Clear Run. Verify the reusable run is cleared; this does not apply inventory changes.'}
            else{'Click Clear Run again to clear the remaining worksheet staging. Dismiss the completion message and inspect the local staging. The two observations do not identify which branch ran or prove inventory application.'}
            return
        }
        if($RunPresentation){
            if($Action -ceq 'TREE_COLLAPSE'){'In Production Run - Tree, click Collapse to hide ingredient choices. This changes the local view only.'}
            else{'Click Expand to show ingredient choices again. Inspect the local view; this does not allocate or apply inventory.'}
            return
        }
        if($Assignment){Get-AssignmentPathInstruction $Action}elseif($Worksheet){Get-WorksheetPathInstruction $Action}
        elseif($CloseActions){
            if($Action -ceq 'Button'){'Open Production and click Close. Verify dismissal; this does not save, post or apply inventory changes.'}
            else{'Reopen Production and click its window X. Verify dismissal; workbook staging remains unchanged.'}
        }
        elseif($Regulation){
            $scope=$Action.Split('_')[0];$command=if($Action.EndsWith('APPLY')){'Apply Regulation'}else{'Clear Override'}
            'In Production Settings, choose '+$scope+' scope and select the output. '+$(if($Action.EndsWith('APPLY')){'Enable regulation and enter a floor of 2 and ceiling of 8. '}else{''})+'Click '+$command+' and inspect the local draft; saved definitions remain unchanged.'
        }
        elseif($DesignReads){
            switch($Action){
                'PROCESS_REFRESH'{'Click Refresh in Process Designer and inspect the lists. A completed refresh does not prove source availability.'}
                'PROCESS_LOAD'{'Select the saved Process and click View Process. Inspect the displayed definition; loading does not validate it.'}
                'PROCESS_REUSE'{'Keep that Process selected and click Edit as New Version. Inspect the editable draft and proposed version; no version is reserved or saved.'}
                'RECIPE_REFRESH'{'Click Refresh in Recipe Designer and inspect the lists. A completed refresh does not prove source availability.'}
                'RECIPE_LOAD'{'Select the saved Recipe and click Load. Inspect the displayed graph; loading does not establish validity or suitability for production.'}
            }
        }
        elseif($Uom){$Action+' the UOM workbench; retain existing edits and verify that the saved catalog is unchanged.'}
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
        if($RunLoad){
            if([string](Probe 'RunLocalStage' @('LOAD','Normal')) -cne 'READY'){throw 'Released Load selection, scale and prior run unavailable; not behavioral RED.'}
        }elseif($RunClear -or $RunRefresh){
            if(-not [bool](Probe 'RunSheetOwnerRecreate')){throw 'Clear path local surface unavailable; not behavioral RED.'}
            if([string](Probe 'RunSheetOwnerStage' @('CLEAR','Normal',$canary)) -cne 'READY'){throw 'Clear worksheet staging unavailable; not behavioral RED.'}
            if([string](Probe 'RunLocalStage' @('CLEAR','Normal')) -cne 'READY'){throw 'Clear released Run staging unavailable; not behavioral RED.'}
            if(-not [bool](Probe 'RunSheetOwnerPending' @($canary))){throw 'Both Clear branches must be staged before recording; not behavioral RED.'}
        }elseif($RunPresentation){
            if([string](Probe 'RunPresentationStage' @('Expanded',$canary)) -cne 'READY'){throw 'Tree presentation path fixture unavailable; not behavioral RED.'}
            $palette=[string](Probe 'RunPresentationState' @($true))
        }elseif($Assignment){
            [void](Probe 'AssignmentStage' @('Normal'))
            $queue=Join-Path $runRoot ('assignment-path-'+$Label)
            [void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @($queue,'Observe',$Fixture.Warehouse))
            [void](Probe 'LifecycleFaultMode' @(''))
        }elseif($Uom){
            # Reset only this disposable workbench between independent recordings.
            if(-not [string]::Equals([string]$book.FullName,[IO.Path]::GetFullPath($path),[StringComparison]::OrdinalIgnoreCase)){throw 'Unexpected UOM path fixture workbook.'}
            $priorAlerts=$excel.DisplayAlerts
            try{
                $excel.DisplayAlerts=$false
                foreach($stageSheet in @($book.Worksheets)){if($stageSheet.Name -ceq 'invSys UOM Catalog'){$stageSheet.Delete()}}
            }finally{$excel.DisplayAlerts=$priorAlerts}
        }elseif($Regulation){[void](Probe 'RegulationStage' @('Process','Normal',$canary,$regProcessId,$regProcessVersion))}
        elseif($DesignReads){[void](Probe 'ReadStage' @('Normal'))}
        elseif($Components){[void](Probe 'ComponentStage' @('REQUIREMENT',$canary,'Valid'))}
        elseif($RecipeOrder){[void](Probe 'OrderStage' @($canary,'Unordered'))}
        elseif($RecipeStructure){[void](Probe 'StructurePathReset')}
        elseif(-not $CloseActions -and -not $Worksheet){[void](Probe 'InstructionStage' @($canary,1,(' '+$canary+'4 ')))}
        $prior=@(if(Test-Path $journalRoot){Get-ChildItem $journalRoot -File -Filter '*.json'|ForEach-Object FullName})
        Delivered (RecordingControl 'Start Recording' 'Click')
        $starts=@(Get-ChildItem $journalRoot -File -Filter '*.json'|Where-Object {$_.FullName -cnotin $prior}|ForEach-Object {Get-Content $_.FullName -Raw|ConvertFrom-Json}|Where-Object RecordType -CEQ 'Start')
        if($starts.Count -ne 1){throw 'Actual recording Start unavailable.'}
        $original=@();$terminals=@();$ordinal=0
        foreach($action in $actions){
            $ordinal++;$before=@(Get-Slice4beActivityFiles $Fixture);$decoy.Activate()
            if($RunRefresh){
                $command=$action -replace '^(Reusable|Worksheet)_',''
                $notice=[string](Probe 'RunFaultAct' @($command))
                $reads=if($action.StartsWith('Worksheet')){if($command -ceq 'LOADER_REFRESH'){1}else{2}}else{0}
                $owner=if($action -ceq 'CLEAR'){[bool](Probe 'RunLocalPreserved' @('CLEAR','Normal'))}
                    elseif($action.StartsWith('Reusable')){[bool](Probe 'RunLocalPreserved' @($command,'Normal'))}
                    else{[bool](Probe 'RunSheetPathReadReceipt' @($reads))}
                $message=if($action -ceq 'CLEAR'){$notice -ceq 'Reusable Production Run cleared.'}
                    elseif($action.StartsWith('Reusable')){$notice -ceq 'Reusable Production Run inventory refreshed from the exact entity projection.'}
                    else{$notice.StartsWith('Production Run inventory refreshed. ')}
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.ActualOwnerAndMessage') ($owner -and $message)
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.CumulativeCapturedReadReceipt') ([bool](Probe 'RunSheetPathReadReceipt' @($reads)))
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.WorksheetStagingRetained') ([bool](Probe 'RunSheetOwnerPending' @($canary)))
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.GuardsRestoredWithoutAdapterReset') ([bool](Probe 'RunFaultGuards'))
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.ExactKeyCustomColumnsAndCanonicalSource') ([bool](Probe 'RunSheetOwnerPreserved') -and [bool](Probe 'RunWorksheetSource'))
            }elseif($RunLoad){
                $notice=[string](Probe 'RunFaultAct' @('LOAD'))
                PathCheck ('InstructionPaths.'+$Label+'.LOAD.ActualOwner') ([bool](Probe 'RunLocalPreserved' @('LOAD','Normal')))
                PathCheck ('InstructionPaths.'+$Label+'.LOAD.ExistingMessage') ($notice.StartsWith('Loaded released Recipe ') -and $notice -cmatch ' with 1 Process\(es\) at 100(?:\.0*)?% scale\.$')
                PathCheck ('InstructionPaths.'+$Label+'.LOAD.GuardsRestoredWithoutAdapterReset') ([bool](Probe 'RunFaultGuards'))
            }elseif($RunClear){
                $notice=[string](Probe 'RunFaultAct' @('CLEAR'))
                $owner=if($action -ceq 'Reusable'){
                    [bool](Probe 'RunLocalPreserved' @('CLEAR','Normal')) -and [bool](Probe 'RunSheetOwnerPending' @($canary)) -and [string](Probe 'RunSheetClearTarget') -ceq 'NotEntered'
                }else{[bool](Probe 'RunSheetOwnerResult' @('CLEAR','Normal')) -and [string](Probe 'RunSheetClearTarget') -ceq 'Captured'}
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.ActualOwnerAndMessage') ($owner -and $notice -ceq $(if($action -ceq 'Reusable'){'Reusable Production Run cleared.'}else{'Production Run cleared.'}))
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.GuardsRestoredWithoutAdapterReset') ([bool](Probe 'RunFaultGuards'))
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.ExactKeyCustomColumnsAndCanonicalSource') ([bool](Probe 'RunSheetOwnerPreserved') -and [bool](Probe 'RunWorksheetSource'))
            }elseif($RunPresentation){
                $notice=[string](Probe 'RunPresentationAct' @($action,''))
                $shape=if($action -ceq 'TREE_COLLAPSE'){'2|2|1'}else{'2|4|0'}
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.OwnerPresentationAndPalette') (-not $notice.StartsWith('HANDLER_ERROR|') -and [string](Probe 'RunPresentationShape') -ceq $shape -and [string](Probe 'RunPresentationState' @($true)) -ceq $palette)
            }elseif($Assignment){[void](Probe 'AssignmentPathInput' @($action));[void](Probe 'AssignmentAct' @($action))}
            elseif($Worksheet){Invoke-WorksheetPathAction $action $book $decoy $canary}
            elseif($CloseActions){
                [void](Probe 'OpenDesigner' @($book.Name));[void](Probe 'CloseShowForTest')
                $decoy.Activate();$dismissed=$true
                if($action -ceq 'Native'){$dismissed=Close-ProductionNativeFixture}else{[void](Probe 'CloseButtonForTest')}
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.ActualDismissal') ($dismissed -and [long](Probe 'CloseLoadedFormsForTest') -eq 0)
                [void](Probe 'CloseForgetForTest')
            }
            elseif($Regulation){[void](Probe 'RegulationPathInput' @($action));[void](Probe 'RegulationAct' @($action.Split('_')[1]))}
            elseif($DesignReads){[void](Probe 'ReadPathInput' @($action));[void](Probe 'ReadAct' @($action))}
            elseif($Uom){[void](Probe 'SendUom')}
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
            if($valid){$valid=$terminal[0].ControlId -ceq (ActionId $action) -and $attempt[0].ActivityId -ceq $terminal[0].ActivityId -and $terminal[0].SequenceId -ceq $starts[0].SequenceId -and $terminal[0].Ordinal -eq $ordinal -and ($Worksheet -or $Assignment -or @($terminal[0].SourceEventRefs).Count -eq 0)}
            PathCheck ('InstructionPaths.'+$Label+'.'+$action+$(if($Worksheet -or $Assignment){'.'+$ordinal}else{''})+'.OriginalHandlerPair') $valid
            if(-not $valid){
                if($CloseActions -or $RunRefresh -or $RunClear -or $RunLoad){Delivered (RecordingControl 'Stop Recording' 'Click');return}
                throw 'Original instruction recording incomplete; retain behavioral failures.'
            }
            if($Assignment){Test-AssignmentPathAction $terminal[0] $attempt[0] $Fixture $Label $ordinal}
            if($RunRefresh -or $RunLoad -or $RunClear -or $RunPresentation){
                PathCheck ('InstructionPaths.'+$Label+'.'+$action+'.ExactLocalOwnerFacts') ($attempt[0].ControlId -ceq (ActionId $action) -and $attempt[0].SequenceId -ceq $starts[0].SequenceId -and $attempt[0].Ordinal -eq $ordinal -and @($attempt[0].SourceEventRefs).Count -eq 0 -and $terminal[0].OwnerId -ceq 'PRODUCTION_RUN_LOCAL' -and $terminal[0].CatalogVersion -eq 24 -and $terminal[0].DataEffect -ceq 'Unchanged' -and $terminal[0].WarehouseId -ceq $Fixture.Warehouse -and $terminal[0].UserId -ceq 'config-producer')
            }
            if($Worksheet){Test-WorksheetPathAction $action $book $decoy $terminal[0] $attempt[0] $Fixture $Label $ordinal}
            $original+=@($attempt[0],$terminal[0]);$terminals+=$terminal[0]
        }
        Delivered (RecordingControl 'Stop Recording' 'Click')
        if($Assignment){PathCheck ('InstructionPaths.'+$Label+'.SeriesReachesSavedDraft') ([bool](Probe 'AssignmentPathResult'))}
        if($RecipeStructure){PathCheck ('InstructionPaths.'+$Label+'.SeriesReachesExpectedLocalDraft') ([bool](Probe 'StructurePathResult'))}
        if($DesignReads){PathCheck ('InstructionPaths.'+$Label+'.SeriesReachesExpectedLocalDraft') ([bool](Probe 'ReadPathResult'))}
        if($Regulation){PathCheck ('InstructionPaths.'+$Label+'.SeriesReachesExpectedLocalDraft') ([bool](Probe 'RegulationPathResult'))}
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
    if($DesignReads){$actions=@('PROCESS_REFRESH','PROCESS_LOAD','PROCESS_REUSE','RECIPE_REFRESH','RECIPE_LOAD')}
    if($Regulation){$actions=@('PROCESS_APPLY','PROCESS_CLEAR','RECIPE_APPLY','RECIPE_CLEAR')}
    if($CloseActions){$actions=@('Button','Native')}
    if($Assignment){$actions=@('REFRESH','PROCESS_SELECT','PROCESS','REQUIREMENT_SELECT','REQUIREMENT','ADD','REMOVE','CLEAR','PROCESS','REQUIREMENT','ADD','SAVE')}
    if($Worksheet){$actions=@('SEND','ADD_ITEM','SEND','RETRIEVE')}
    if($RunPresentation){$actions=@('TREE_COLLAPSE','TREE_EXPAND')}
    if($RunClear){$actions=@('Reusable','Worksheet')}
    if($RunLoad){$actions=@('LOAD')}
    if($RunRefresh){$actions=@('Reusable_LOADER_REFRESH','Reusable_MANAGER_REFRESH','CLEAR','Worksheet_LOADER_REFRESH','Worksheet_MANAGER_REFRESH')}
    $canary='PATH'+[guid]::NewGuid().ToString('N');$book=$null;$decoy=$null
    $imagePrefix=if($RunRefresh){'run-refresh'}elseif($RunLoad){'run-load'}elseif($RunClear){'run-clear'}elseif($RunPresentation){'run-presentation'}elseif($Assignment){'assignment'}elseif($Worksheet){'worksheet'}elseif($CloseActions){'close'}elseif($Regulation){'regulation'}elseif($DesignReads){'design-read'}elseif($Uom){'uom'}elseif($Components){'component'}elseif($RecipeOrder){'recipe-order'}elseif($RecipeStructure){'recipe-structure'}else{'instruction'}
    try {
        SelectTarget $Fixture
        if($RunRefresh -or $RunClear -or $RunLoad){
            $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
            if(-not $seed.StartsWith('OK|')){throw 'Admin Clear path inventory fixture unavailable; not behavioral RED.'}
        }
        $allowed=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-admin',$Fixture.Warehouse,'S1')
        PathCheck 'InstructionPaths.Setup.AuthorHasExplicitCapability' ($allowed -is [bool] -and $allowed)
        if($allowed -isnot [bool] -or -not $allowed){throw 'Explicit guide author capability fixture unavailable; not behavioral RED.'}
        if($Assignment -and -not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.AssignmentPolicyForTest' @($true,$true))){throw 'Assignment navigation policy unavailable.'}
        if($RunPresentation -and -not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RunPresentationPolicy' @($true))){throw 'Tree navigation policy unavailable; not behavioral RED.'}
        if(($RunRefresh -or $RunClear -or $RunLoad) -and -not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.RunPresentationPolicy' @($false))){throw 'Clear command policy unavailable; not behavioral RED.'}
        SetRecordingPolicy $true
        SelectTarget $Fixture 'config-producer';OpenRecordingViewer
        $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1)
        $sheet.Cells.Item(1,1).Value2='Operator Extra';$sheet.Cells.Item(2,1).Value2=$canary
        $path=Join-Path $runRoot 'instruction-path-operator.xlsb';$book.SaveAs($path,50);$book.Close($false)
        $workbookPin=SavedHash $path;$book=$excel.Workbooks.Open($path,0,$false);$sheet=$book.Worksheets.Item(1)
        $decoy=$excel.Workbooks.Add();[void](Probe 'OpenDesigner' @($book.Name))
        if($RunLoad -and -not [bool](Probe 'ReadPrepare' @($canary))){throw 'Released Load definitions unavailable; not behavioral RED.'}
        if($RunClear -or $RunRefresh){
            if(-not [bool](Probe 'ReadPrepare' @($canary)) -or -not [bool](Probe 'RunSheetOwnerPrepare')){throw 'Released definitions and supported worksheet surfaces unavailable; not behavioral RED.'}
        }
        if($RecipeStructure){
            if([string](Probe 'StructureReleasedSetup') -cne 'READY' -or -not [bool](Probe 'StructureReleasedNodes')){throw 'Released Process path fixture unavailable; not behavioral RED.'}
            [void](Probe 'StructurePathRemember')
        }
        if($Assignment -and -not [bool](Probe 'AssignmentPrepare' @($canary))){throw 'Released Assignment fixture unavailable; not behavioral RED.'}
        if($DesignReads -and -not [bool](Probe 'ReadPrepare' @($canary))){throw 'Released designer read fixture unavailable; not behavioral RED.'}
        if($Regulation){
            $ready=([string](Probe 'ReleasedProcess' @($canary))).Split('|')
            if($ready.Count -ne 3 -or $ready[0] -cne 'READY'){throw 'Released regulation fixture unavailable; not behavioral RED.'}
            $regProcessId=$ready[1];$regProcessVersion=$ready[2]
        }
        $authorityPins=@{}
        # Publication owns snapshot projections; canonical files remain byte-pinned.
        foreach($file in Get-ChildItem $Fixture.Root -Recurse -File|Where-Object {$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '*.Snapshot.*' -and -not $_.Name.StartsWith('~$')}){$authorityPins[$file.FullName]=SavedHash $file.FullName}
        if($Worksheet -or $Assignment){$worksheetBefore=Get-WorksheetPathBusinessState $Fixture}
        $source=$null;$observed=$null;RecordSeries 'GuideSource' ([ref]$source)
        if(($CloseActions -or $RunRefresh -or $RunClear -or $RunLoad) -and $null -eq $source){return}
        RecordSeries 'ObservedRun' ([ref]$observed)
        if(($CloseActions -or $RunRefresh -or $RunClear -or $RunLoad) -and $null -eq $observed){return}
        PathCheck 'InstructionPaths.DistinctGuideAndObservedRuns' ($source.Journal.ActionPathId -cne $observed.Journal.ActionPathId -and @($source.Terminals|Where-Object {$_.ActivityId -cin $observed.Terminals.ActivityId}).Count -eq 0)
        if($Worksheet -or $Assignment){
            PathCheck 'InstructionPaths.RecordingPreservesAuthConfigInventory' ((Get-WorksheetPathBusinessState $Fixture) -ceq $worksheetBefore)
            if($Worksheet){$workbookPin=SavedHash $path}
        }
        if($CloseActions){[void](Probe 'OpenDesigner' @($book.Name))}
        $window=[long](Probe 'Present' @($(if($RunRefresh -or $RunLoad -or $RunClear){3}elseif($RunPresentation){4}elseif($Assignment){2}elseif($Uom -or $Regulation){5}elseif($RecipeOrder -or $RecipeStructure -or $DesignReads -or $Regulation){1}else{0})));CaptureOwnedFormEvidence 'Production' ($imagePrefix+'-editor.png') $window
        $activityPins=ActivityPins;$journalPins=BoundPins $journalRoot
        [void](Probe 'CloseDesigner');CloseRecordingViewer;SelectTarget $Fixture
        $published=[bool](Run 'invSys.Admin.xlam' 'modAdminConsole.PublishReadFixtureForTest')
        PathCheck 'InstructionPaths.ActualAdminPublication' $published
        if(-not $published){throw 'Owning publication unavailable.'}
        $eventsPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Snapshot.Events.json');$publication=Get-Content $eventsPath -Raw|ConvertFrom-Json
        if($Worksheet -or $Assignment){
            $authorityPins=@{};foreach($file in Get-ChildItem $Fixture.Root -Recurse -File|Where-Object {$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '*.Snapshot.*' -and -not $_.Name.StartsWith('~$')}){$authorityPins[$file.FullName]=SavedHash $file.FullName}
            if($Assignment){Test-AssignmentPathSources $source $observed $publication $Fixture}else{Test-WorksheetPathPublishedSources $source $observed $publication $Fixture}
        }
        OpenRecordingViewer
        foreach($record in @($source.Original)+@($observed.Original)){
            $group=@($publication.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -ceq $record.ActivityId})
            $lines=@($group|ForEach-Object Lines|Where-Object RecordId -CEQ $record.RecordId)
            PathCheck ('InstructionPaths.Publication.'+$(if($record.SequenceId -ceq $source.Journal.SequenceId){'GuideSource'}else{'ObservedRun'})+'.'+$record.Ordinal+'.'+$record.OutcomeCode) ($lines.Count -eq 1 -and ($lines[0]|ConvertTo-Json -Depth 20 -Compress) -ceq ($record|ConvertTo-Json -Depth 20 -Compress))
        }
        foreach($record in $observed.Terminals){
            $id=[string]$record.ActivityId
            PathCheck ('InstructionPaths.Detail.'+$record.ControlId+$(if($RunRefresh -or $RunClear -or $Uom -or $Regulation -or $CloseActions -or $Worksheet -or $Assignment){'.'+$record.Ordinal}else{''})) ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadActionForTest' @('SelectSource',$id)) -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedReadDetailForTest' @('Source event / activity ID',$id)))
        }
        if($RunRefresh -or $RunLoad -or $RunClear -or $RunPresentation -or $Uom -or $Components -or $RecipeOrder -or $RecipeStructure -or $DesignReads -or $Regulation -or $CloseActions -or $Worksheet -or $Assignment){
            # Published contributing lines are not ordered by outcome. Select
            # the exact terminal record, rather than assuming it is row two.
            $detailGroup=@($publication.Groups|Where-Object {$_.Source -ceq 'Activity' -and $_.SourceId -ceq $observed.Terminals[-1].ActivityId})
            if($detailGroup.Count -ne 1){throw 'Exact terminal detail group unavailable.'}
            $terminalIndex=-1
            for($lineIndex=0;$lineIndex -lt $detailGroup[0].Lines.Count;$lineIndex++){
                if($detailGroup[0].Lines[$lineIndex].RecordId -ceq $observed.Terminals[-1].RecordId){$terminalIndex=$lineIndex;break}
            }
            if($terminalIndex -lt 0){throw 'Exact terminal detail line unavailable.'}
            if($RunRefresh -or $RunLoad -or $RunClear -or $RunPresentation -or $Components -or $RecipeOrder -or $RecipeStructure -or $DesignReads -or $Regulation -or $CloseActions -or $Worksheet -or $Assignment){
                $requestedIndex=-1
                for($lineIndex=0;$lineIndex -lt $detailGroup[0].Lines.Count;$lineIndex++){
                    if($detailGroup[0].Lines[$lineIndex].OutcomeCode -ceq 'REQUESTED'){$requestedIndex=$lineIndex;break}
                }
                if($requestedIndex -lt 0){throw 'Exact requested detail line unavailable.'}
                Delivered (ExpectationControl 'lstEventLines' 'Index' ([string]$requestedIndex) 'frmEventDetail')
                PathCheck 'InstructionPaths.Detail.RequestedObservationMatchesVisibleSelection' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedSelectedFieldForTest' @('Outcome','REQUESTED')) -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedSelectedFieldForTest' @('Data effect','Unknown')) -and -not [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedSelectedFieldForTest' @('Outcome',(ActionOutcome $actions[-1]))))
            }
            Delivered (ExpectationControl 'lstEventLines' 'Index' ([string]$terminalIndex) 'frmEventDetail')
            if($RunRefresh -or $RunLoad -or $RunClear -or $RunPresentation -or $Components -or $RecipeOrder -or $RecipeStructure -or $DesignReads -or $Regulation -or $CloseActions -or $Worksheet -or $Assignment){
                PathCheck 'InstructionPaths.Detail.ExactTerminalOutcomeAndEffect' ([bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedSelectedFieldForTest' @('Outcome',(ActionOutcome $actions[-1]))) -and [bool](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedSelectedFieldForTest' @('Data effect',$(if($Worksheet -or $Assignment){'Unknown'}else{'Unchanged'}))))
            }
        }
        CaptureOwnedFormEvidence 'Event Detail' ($imagePrefix+'-event-detail.png') ([long](Run 'invSys.Operations.xlam' 'modInventoryViewer.PublishedDetailLabelForTest' @('Window','')))
        Delivered (BoundLibrary 'Open')
        if((BoundLibrary 'Select' $observed.Journal.ActionPathId) -cne 'SELECTED'){throw 'Observed run unavailable.'}
        if($DesignReads){
            $allChoices=$true
            foreach($action in $actions){
                $choice=Set-EvaluationDraft @(,@((ActionId $action),(ActionOutcome $action),'True')) 0 'CommandCompleted' -StopAtMissingChoice
                PathCheck ('InstructionPaths.ExpectedOutcome.'+$action) $choice
                $allChoices=$allChoices -and $choice
            }
            # Retain real missing editor choices as behavioral RED, without
            # disguising the later unavailable guide prerequisite as a harness defect.
            if(-not $allChoices){return}
        }
        if($Worksheet){Test-WorksheetPathLocalTerminals $observed}
        if($Assignment){Test-AssignmentPathLocalTerminals $observed}
        foreach($kind in @('CommandCompleted','SourceEventsApplied')){
            $steps=@(,@((ActionId $actions[-1]),(ActionOutcome $actions[-1]),'True'));$terminal=0
            if($RunRefresh -or $RunLoad -or $RunClear -or $RunPresentation -or $Regulation -or $CloseActions -or $Worksheet -or $Assignment){
                # Repeated controls carry no scope or button/native distinction.
                # Explicit ordered intent identifies
                # the final occurrence without changing first-match semantics.
                $steps=@(foreach($action in $actions){,@((ActionId $action),(ActionOutcome $action),'True')})
                $terminal=$actions.Count-1
            }
            $ready=Set-EvaluationDraft $steps $terminal $kind -StopAtMissingChoice
            PathCheck ('InstructionPaths.TerminalChoice.'+$kind) $ready
            if(-not $ready){continue}
            $before=@(EvaluationFiles|ForEach-Object FullName);Delivered (ExpectationControl 'btnEvaluatePath' 'Click' '' 'frmActionPaths')
            $fresh=@(EvaluationFiles|Where-Object {$_.FullName -cnotin $before});if($fresh.Count -ne 1){throw 'Actual evaluation unavailable.'}
            $result=Get-Content $fresh[0].FullName -Raw|ConvertFrom-Json
            $valid=if($kind -ceq 'CommandCompleted'){$result.ResultState -ceq 'Concluded' -and 'COMMAND_COMPLETED' -cin @($result.ReasonCodes)}elseif($Worksheet -or $Assignment){$result.ResultState -ceq 'Concluded' -and 'SOURCE_APPLIED' -cin @($result.ReasonCodes)}else{$result.ResultState -ceq 'Incomplete' -and 'SOURCE_UNAVAILABLE' -cin @($result.ReasonCodes)}
            if($RunRefresh -or $RunLoad -or $RunClear -or $RunPresentation -or $Regulation -or $CloseActions -or $Worksheet -or $Assignment){$valid=$valid -and @($result.Matches).Count -eq $actions.Count -and @($result.ExtraActivityIds).Count -eq 0 -and ($result.Matches.ActivityId -join '|') -ceq ($observed.Terminals.ActivityId -join '|')}
            if($Assignment -and $kind -ceq 'SourceEventsApplied'){Test-AssignmentPathTerminalSource $result $observed.Terminals[-1] $publication}
            if($Worksheet -and $kind -ceq 'SourceEventsApplied'){Test-WorksheetPathTerminalSources $result $observed.Terminals[-1] $publication}
            PathCheck ('InstructionPaths.ExactTerminal.'+$kind) ($valid -and $result.JournalRecordId -ceq $observed.Journal.RecordId -and $result.Matches[-1].ActivityId -ceq $observed.Terminals[-1].ActivityId)
        }
        if((BoundLibrary 'Select' $source.Journal.ActionPathId) -cne 'SELECTED'){throw 'Guide source selection unavailable.'}
        Delivered (BoundControl 'btnCreateGuide' 'Click' '' 'frmActionPaths')
        PathCheck 'InstructionPaths.FiveObservedGuideSteps' ((Author 'lstGuideSteps' 'Rows') -ceq [string]$actions.Count)
        Delivered (Author 'txtGuideName' 'Write' $(if($RunRefresh){'Refresh reusable and worksheet Run staging'}elseif($RunLoad){'Load a released Recipe into local Run staging'}elseif($RunClear){'Clear local Production staging'}elseif($RunPresentation){'Inspect ingredient choices in Run Tree'}elseif($Assignment){'Edit and save acceptable item alternatives'}elseif($Worksheet){'Stage and retrieve Process worksheets'}elseif($CloseActions){'Dismiss Production through its Close controls'}elseif($Regulation){'Stage and clear output regulation'}elseif($DesignReads){'Refresh and load local designer views'}elseif($Uom){'Open and reuse the UOM draft'}elseif($Components){'Edit Process requirements and outputs'}elseif($RecipeOrder){'Order a local Recipe draft'}elseif($RecipeStructure){'Edit local Recipe connections and nodes'}else{'Edit Process instructions'}))
        Delivered (Author 'txtGuideInstructions' 'Write' $(if($RunRefresh){'Start with a loaded reusable run and separate worksheet staging already prepared. Click Recipe Loader Refresh and Run Manager Refresh, then Clear Run. Repeat both Refresh buttons on the remaining worksheet staging. Inspect each local result. The observations do not identify the branches or prove source freshness or inventory application.'}elseif($RunLoad){'Start with a released Recipe version selected and a valid batch scale. Click Load Recipe and inspect the local Process, ingredient and output rows before allocation or check-in. Loading replaces prior local preparation. The diagnostic proves local loading completed; the control observation does not identify the selected Recipe or prove inventory application.'}elseif($RunClear){'Start with a loaded reusable run and separate worksheet staging. Click Clear Run twice and inspect the local results. The diagnostic proves two locally completed Clear commands; observations do not distinguish the branches or assert inventory or Domain application.'}elseif($RunPresentation){'Collapse then expand the ingredient choices in the experimental Run Tree. This changes presentation only. The diagnostic proves local presentation completed; it does not prove inventory allocation or Domain application.'}elseif($Assignment){'Select a released Process and requirement, add and remove an acceptable item, clear, reload and save an alternative as a new DRAFT. These item types do not allocate inventory. Command completion and published Designs application are separate conclusions.'}elseif($Worksheet){'Send two local Process tables, add an acceptable-item pair, fill valid definitions and retrieve both as DRAFT. Command completion does not itself prove Domain application.'}elseif($CloseActions){'Demonstrate the Close button and window X in separate openings. The diagnostic proves two dismissals; observations do not distinguish the gesture or assert saved work or Domain application.'}elseif($Regulation){'Apply then clear regulation on the selected Process output and Recipe node output. Inspect the resulting local draft; this sequence does not save, release, validate or apply inventory changes.'}elseif($DesignReads){'Refresh and view a Process, prepare its next editable draft, then refresh and load a Recipe. These local completions do not validate definitions, prove source availability, reserve versions or apply inventory changes.'}elseif($Uom){'Open the UOM workbench and reuse the existing local draft without publishing the catalog.'}elseif($Components){'Add, select and update, move up/down, then remove a requirement and an output. Validate before saving the Process.'}elseif($RecipeOrder){'Move the selected recipe node up, down, then apply Auto Order. Inspect the local draft before validation and Save.'}elseif($RecipeStructure){'Disconnect and reconnect the source output, update its required quantities, then add and remove a local Recipe node. Inspect and validate before Save.'}else{'Edit the local instruction draft, then validate before saving the Process.'}))
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
        if($Assignment){[void](Probe 'LifecycleFaultMode' @(''));[void](Run 'invSys.Core.xlam' 'modRoleEventWriter.LifecycleWriterFixtureForTest' @('','',''))}
        [void](Probe 'CloseDesigner');CloseRecordingViewer
        if($null -ne $book){$book.Close($false)}
        if($null -ne $decoy){$decoy.Close($false)}
        SelectTarget $Fixture
    }
}
