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
    $project=$packages['invSys.Operations.xlam'].VBProject
    $owner=$project.VBComponents.Item('modProductionReusableRun').CodeModule
    AfterRead $owner 'LoadReleasedReusableRecipe' 'validation = modOperationsPrimitiveBridge.ValidateReleasedRecipe(' 'ValidateReleasedRecipe' '(Left$(validation, 2) = "1" & vbTab)'
    AfterRead $owner 'LoadReleasedReusableRecipe' 'jsonText = modOperationsPrimitiveBridge.GetRecipeGraph(' 'GetRecipeGraph' 'Len(jsonText) > 2'
    $definitions=$owner
    foreach($component in $project.VBComponents){
        if($component.Name -ceq 'modProductionRunDefinitionLoad'){$definitions=$component.CodeModule;break}
    }
    AfterRead $definitions 'LoadNodeProcessDefinitions' 'jsonText = modOperationsPrimitiveBridge.GetProcessVersion(' 'GetProcessVersion' 'Len(jsonText) > 2'
    $reads=@(for($line=1;$line -le $owner.CountOfLines;$line++){
        if($owner.Lines($line,1).Trim() -cmatch '^(entities|entity|entityRows) = modInventoryDomainBridge\.ListAvailableInventoryEntitiesBridge\(""\)$'){
            [pscustomobject]@{Line=$line;Variable=$Matches[1]}
        }
    })
    if($reads.Count -ne 8){throw 'Reusable Inventory read-return fixture anchors changed; not product RED.'}
    foreach($read in $reads|Sort-Object Line -Descending){
        $owner.InsertLines($read.Line+1,('    TestProductionDesigner.RunYieldReadReturned "InventoryEntities", IsArray('+$read.Variable+')'))
    }
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    AfterRead $form 'DesignReadListForTest' 'DesignReadListForTest = modOperationsPrimitiveBridge.ListProcesses(' 'ListProcesses' 'IsArray(DesignReadListForTest)'
    AfterRead $form 'DesignReadListForTest' 'DesignReadListForTest = modOperationsPrimitiveBridge.ListRecipes(' 'ListRecipes' 'IsArray(DesignReadListForTest)'
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mRunYieldTarget As String, mRunYieldArmed As Boolean, mRunYieldReached As Boolean
Private mRunYieldAvailable As Boolean, mRunYieldLaterReads As Long
Private mRunYieldOwner As String, mRunYieldProjection As String
'@)
    $adapter.AddFromString(@'
Public Sub RunYieldReset()
    mRunYieldTarget = "": mRunYieldArmed = False: mRunYieldReached = False
    mRunYieldAvailable = False: mRunYieldLaterReads = 0
    mRunYieldOwner = "": mRunYieldProjection = ""
End Sub
Public Sub RunYieldArm(ByVal boundary As String)
    RunYieldReset
    mRunYieldTarget = boundary: mRunYieldArmed = True
End Sub
Public Sub RunYieldReadReturned(ByVal boundary As String, ByVal available As Boolean)
    If mRunYieldReached Then mRunYieldLaterReads = mRunYieldLaterReads + 1
    If Not mRunYieldArmed Or boundary <> mRunYieldTarget Then Exit Sub
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
    try{
        foreach($case in $cases){
            [void](Probe 'RunYieldReset')
            SelectTarget $Fixture 'config-producer';[void](Probe 'RunLocalReopen' @($Book.Name))
            [void](Probe 'RunFaultResetFixture')
            if([string](Probe 'RunLocalStage' @($case.Action,'Normal')) -cne 'READY'){throw 'Released Run yield fixture unavailable; not product RED.'}
            [void](Probe 'RunYieldArm' @($case.Read));$before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            $notice=[string](Probe 'RunFaultAct' @($case.Action));$evidence=([string](Probe 'RunYieldEvidence')).Split('|')
            $label='RunYield.'+$case.Action+'.'+$case.Read
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
