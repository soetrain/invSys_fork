# Sign out after real owning reads return; payloads and original handlers are retained.
function Install-ProductionRunYieldProbe {
    function AfterRead($Module,[string]$Procedure,[string]$Assignment,[string]$Boundary,[string]$Available){
        $start=$Module.ProcStartLine($Procedure,0);$end=$start+$Module.ProcCountLines($Procedure,0)
        $matches=@(for($i=$start;$i -lt $end;$i++){if($Module.Lines($i,1).Trim().StartsWith($Assignment,[StringComparison]::Ordinal)){$i}})
        if($matches.Count -ne 1){throw ('Run read-return fixture anchor changed: '+$Procedure+'; not product RED.')}
        $line=$matches[0]
        while($Module.Lines($line,1).TrimEnd().EndsWith('_')){$line++;if($line -ge $end){throw 'Incomplete read assignment; not product RED.'}}
        $Module.InsertLines($line+1,('    TestProductionDesigner.RunYieldReadReturned "'+$Boundary+'", '+$Available))
    }
    function AfterConversion($Module,[string]$Procedure,[string]$Boundary){
        $start=$Module.ProcStartLine($Procedure,0);$end=$start+$Module.ProcCountLines($Procedure,0)
        $hits=@(for($line=$start;$line -lt $end;$line++){if($Module.Lines($line,1).Trim().StartsWith('If Not modUomSettings.GetUomConversion(',[StringComparison]::Ordinal)){$line}})
        if($hits.Count -ne 1){throw 'Run conversion fixture anchor changed; not product RED.'}
        $line=$hits[0]
        while($line -lt $end -and $Module.Lines($line,1).Trim() -cne 'End If'){$line++}
        if($line -ge $end -or $Module.Lines($line-1,1).Trim() -cne 'Exit Function'){throw 'Run conversion return anchor changed; not product RED.'}
        $native=if($Boundary -ceq 'StockUomConversion'){'CStr(entities(representativeRow, 5))'}else{'nativeUom'}
        $version=if($Boundary -ceq 'StockUomConversion'){'conversionVersion'}else{'catalogVersion'}
        $Module.InsertLines($line+1,('    TestProductionDesigner.RunYieldReadReturned "'+$Boundary+'", (conversionFactor > 0 And Len('+$version+') > 0 And StrComp('+$native+', RunRecordText(requirement, "UOM"), vbTextCompare) <> 0)'))
    }
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('modProductionReusableRun').CodeModule
    AfterConversion $owner 'ApplyReusableRunStockAllocation' 'StockUomConversion'
    AfterConversion $owner 'ApplyReusableRunAllocation' 'ExactUomConversion'
    AfterRead $owner 'LoadReleasedReusableRecipe' 'validation = modOperationsPrimitiveBridge.ValidateReleasedRecipe(' 'ValidateReleasedRecipe' '(Left$(validation, 2) = "1" & vbTab)'
    AfterRead $owner 'LoadReleasedReusableRecipe' 'jsonText = modOperationsPrimitiveBridge.GetRecipeGraph(' 'GetRecipeGraph' 'Len(jsonText) > 2'
    $definitions=$owner
    foreach($component in $project.VBComponents){
        if($component.Name -ceq 'modProductionRunDefinitionLoad'){$definitions=$component.CodeModule;break}
    }
    AfterRead $definitions 'LoadNodeProcessDefinitions' 'jsonText = modOperationsPrimitiveBridge.GetProcessVersion(' 'GetProcessVersion' 'Len(jsonText) > 2'
    $readModules=@($owner)
    foreach($component in $project.VBComponents){
        if($component.Name -ceq 'modProductionRunEntityReads'){$readModules+=,$component.CodeModule}
    }
    $reads=@(foreach($module in $readModules){for($line=1;$line -le $module.CountOfLines;$line++){
        if($module.Lines($line,1).Trim() -cmatch '^(entities|entity|entityRows) = modInventoryDomainBridge\.ListAvailableInventoryEntitiesBridge\(""\)$'){
            [pscustomobject]@{Module=$module;Line=$line;Variable=$Matches[1]}
        }
    }})
    if($reads.Count -ne 8){throw 'Reusable Inventory read-return fixture anchors changed; not product RED.'}
    foreach($read in $reads|Sort-Object Line -Descending){
        $read.Module.InsertLines($read.Line+1,('    TestProductionDesigner.RunYieldReadReturned "InventoryEntities", IsArray('+$read.Variable+')'))
    }
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    # Change the editor before the actual Save/Release fixture, never a loaded definition.
    $start=$form.ProcStartLine('DesignerReleasedProcessForTest',0);$end=$start+$form.ProcCountLines('DesignerReleasedProcessForTest',0)
    $hits=@(for($line=$start;$line -lt $end;$line++){if($form.Lines($line,1).Trim() -ceq 'mLstProcessRequirements.List(0, 5) = "lbs"'){$line}})
    if($hits.Count -ne 1){throw 'Released differing-UOM fixture anchor changed; not product RED.'}
    $form.ReplaceLine($hits[0],'    mLstProcessRequirements.List(0, 5) = "lb"')
    AfterRead $form 'DesignReadListForTest' 'DesignReadListForTest = modOperationsPrimitiveBridge.ListProcesses(' 'ListProcesses' 'IsArray(DesignReadListForTest)'
    AfterRead $form 'DesignReadListForTest' 'DesignReadListForTest = modOperationsPrimitiveBridge.ListRecipes(' 'ListRecipes' 'IsArray(DesignReadListForTest)'
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mRunYieldTarget As String, mRunYieldArmed As Boolean, mRunYieldReached As Boolean
Private mRunYieldAvailable As Boolean, mRunYieldLaterReads As Long
Private mRunYieldOrdinal As Long, mRunYieldSeen As Long
Private mRunYieldOwner As String, mRunYieldProjection As String
'@)
    $adapter.AddFromString(@'
Public Sub RunYieldReset()
    mRunYieldTarget = "": mRunYieldArmed = False: mRunYieldReached = False
    mRunYieldAvailable = False: mRunYieldLaterReads = 0
    mRunYieldOrdinal = 1: mRunYieldSeen = 0
    mRunYieldOwner = "": mRunYieldProjection = ""
End Sub
Public Sub RunYieldArm(ByVal boundary As String, Optional ByVal ordinal As Long = 1)
    RunYieldReset
    mRunYieldTarget = boundary: mRunYieldOrdinal = ordinal: mRunYieldArmed = True
End Sub
Public Sub RunYieldReadReturned(ByVal boundary As String, ByVal available As Boolean)
    If mRunYieldReached Then mRunYieldLaterReads = mRunYieldLaterReads + 1
    If Not mRunYieldArmed Or boundary <> mRunYieldTarget Then Exit Sub
    mRunYieldSeen = mRunYieldSeen + 1
    If mRunYieldSeen <> mRunYieldOrdinal Then Exit Sub
    mRunYieldArmed = False: mRunYieldReached = True: mRunYieldAvailable = available
    mRunYieldOwner = modProductionReusableRun.RunLocalStateForTest()
    mRunYieldProjection = mForm.RunLocalStateForTest()
    modAuth.SignOut
End Sub
Public Function RunYieldEvidence() As String
    RunYieldEvidence = CStr(mRunYieldReached) & "|" & CStr(mRunYieldAvailable) & "|" & CStr(mRunYieldLaterReads)
End Function
Public Function RunYieldOwnerPreserved() As Boolean
    RunYieldOwnerPreserved = (modProductionReusableRun.RunLocalStateForTest() = mRunYieldOwner)
End Function
Public Function RunYieldProjectionPreserved() As Boolean
    RunYieldProjectionPreserved = (mForm.RunLocalStateForTest() = mRunYieldProjection)
End Function
'@)
}

function Test-ProductionRunYield($Fixture,$Other,$Book) {
    # The parent fault fixture supplies Probe, Files, the saved byte pins and canary.
    $cases=@(
        @{Action='LOAD';Read='ValidateReleasedRecipe'},
        @{Action='LOAD';Read='GetRecipeGraph'},
        @{Action='LOAD';Read='GetProcessVersion'},
        @{Action='LOAD';Read='InventoryEntities'},
        @{Action='SCALE';Read='InventoryEntities'},
        @{Action='LOADER_REFRESH';Read='ListProcesses'},
        @{Action='LOADER_REFRESH';Read='ListRecipes'},
        @{Action='LOADER_REFRESH';Read='InventoryEntities'},
        @{Action='MANAGER_REFRESH';Read='InventoryEntities'},
        @{Action='ALLOCATE';Read='InventoryEntities'},
        @{Action='TREE_ALLOCATE';Read='InventoryEntities'}
    )
    foreach($action in @('ALLOCATE','TREE_ALLOCATE')){
        foreach($read in @('StockUomConversion','ExactUomConversion')){$cases+=@{Action=$action;Read=$read;Ordinal=1}}
        foreach($ordinal in @(2,3,4)){$cases+=@{Action=$action;Read='InventoryEntities';Ordinal=$ordinal}}
    }
    try{
        foreach($case in $cases){
            [void](Probe 'RunYieldReset')
            SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
            [void](Probe 'RunFaultResetFixture')
            if([string](Probe 'RunLocalStage' @($case.Action,'Normal')) -cne 'READY'){throw 'Released Run yield fixture unavailable; not product RED.'}
            $ordinal=if($case.ContainsKey('Ordinal')){[int]$case.Ordinal}else{1}
            [void](Probe 'RunYieldArm' @($case.Read,$ordinal));$before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            $notice=[string](Probe 'RunFaultAct' @($case.Action));$evidence=([string](Probe 'RunYieldEvidence')).Split('|')
            $label='RunYield.'+$case.Action+'.'+$case.Read
            if($ordinal -gt 1){$label+='.'+$ordinal}
            $reached=$evidence.Count -eq 3 -and $evidence[0] -ceq 'True' -and $evidence[1] -ceq 'True'
            if(-not $reached){throw ('Real Run read boundary unavailable: '+$case.Read+'; not product RED.')}
            Check ($label+'.RealReadReturnedBeforeSignOut') $reached
            Check ($label+'.NoLaterObservedReads') ([int]$evidence[2] -eq 0)
            Check ($label+'.OwnerStateAtBoundaryPreserved') ([bool](Probe 'RunYieldOwnerPreserved'))
            Check ($label+'.ProjectionAtBoundaryPreserved') ([bool](Probe 'RunYieldProjectionPreserved'))
            Check ($label+'.RefusalVisible') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))
            Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Check ($label+'.GuardsRestoredBeforeAdapterReset') ([bool](Probe 'RunFaultGuards'))
            $rows=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json})
            Check ($label+'.OriginalAttemptWithoutMisattributedOutcome') ($rows.Count -eq 1 -and $rows[0].ControlId -ceq ('PRODUCTION_RUN_'+$case.Action) -and $rows[0].OutcomeCode -ceq 'REQUESTED' -and $rows[0].UserId -ceq 'config-producer' -and $rows[0].WarehouseId -ceq $Fixture.Warehouse -and @($rows[0].SourceEventRefs).Count -eq 0)
            Check ($label+'.NoOtherWarehouseActivity') ((@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
        }
    }finally{
        [void](Probe 'RunYieldReset')
        SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
    }
}
