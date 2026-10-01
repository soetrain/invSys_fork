# Unsaved probes observe real worksheet Clear notification/read returns.
function Install-ProductionRunClearYieldProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    function After-ClearBoundary($Module,[string]$Procedure,[string]$Anchor,[string]$Boundary,[string]$Available){
        $start=$Module.ProcStartLine($Procedure,0);$end=$start+$Module.ProcCountLines($Procedure,0)
        $hits=@(for($line=$start;$line -lt $end;$line++){if($Module.Lines($line,1).Trim() -ceq $Anchor){$line}})
        if($hits.Count -ne 1){throw ('Clear yield fixture anchor changed: '+$Procedure+'; not product RED.')}
        $Module.InsertLines($hits[0]+1,('    TestProductionDesigner.RunClearYieldReturned "'+$Boundary+'", '+$Available))
    }
    $owner=$project.VBComponents.Item('mProduction').CodeModule
    After-ClearBoundary $owner 'LoadProductionRunInventoryPickerItems' 'result = modInventoryDomainBridge.ListInventoryPickerItemsBridge(filterText)' 'InventoryPicker' 'IsArray(result)'
    After-ClearBoundary $owner 'GetProductionRunDefaultLocation' 'GetProductionRunDefaultLocation = Trim$(modConfig.GetString("DefaultLocation", ""))' 'DefaultLocation' 'True'
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    After-ClearBoundary $adapter 'RunSheetNotice' 'mRunSheetNotices = mRunSheetNotices + 1' 'ClearNotification' 'True'
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.AddFromString(@'
Public Function RunClearYieldStageForTest() As Boolean
    Dim control As MSForms.ListBox
    On Error GoTo Failed
    mLoading = True
    mRunInventoryCacheLoaded = False: mRunInventoryRows = Empty
    For Each control In RunClearYieldListsForTest()
        control.Clear
        control.AddItem "CLEAR-YIELD-LOCAL"
    Next control
    mLoading = False
    RunClearYieldStageForTest = True
    Exit Function
Failed:
    mLoading = False
End Function
Public Function RunClearYieldStateForTest() As String
    Dim control As MSForms.ListBox, row As Long, column As Long, value As String
    value = RunLocalStateForTest() & "|" & CStr(mRunInventoryCacheLoaded) & "|" & CStr(IsEmpty(mRunInventoryRows))
    For Each control In RunClearYieldListsForTest()
        value = value & "|" & CStr(control.ListCount) & ":" & CStr(control.ListIndex)
        For row = 0 To control.ListCount - 1
            For column = 0 To control.ColumnCount - 1
                value = value & "|" & NzStr(control.List(row, column))
            Next column
        Next row
    Next control
    RunClearYieldStateForTest = value & "|" & CStr(mCmbRunLocation.ListCount) & "|" & CStr(mCmbTreeRunLocation.ListCount)
End Function
Private Function RunClearYieldListsForTest() As Collection
    Dim lists As New Collection
    lists.Add mLstLoaderLines
    lists.Add mLstRunPalette
    lists.Add mLstManagerCheck
    lists.Add mLstManagerOutput
    lists.Add mLstRunInstructions
    Set RunClearYieldListsForTest = lists
End Function
Public Function RunClearYieldFixtureForTest() As Variant
    RunClearYieldFixtureForTest = Array(mRunSheetKey, mRunSheetLocation, mRunSheetUom, mRunSheetAvailable)
End Function
Public Sub RunClearYieldRestoreForTest(ByVal values As Variant)
    mRunSheetKey = CStr(values(0)): mRunSheetLocation = CStr(values(1))
    mRunSheetUom = CStr(values(2)): mRunSheetAvailable = CDbl(values(3))
End Sub
'@)
    $adapter.InsertLines(1,@'
Private mClearYieldTarget As String, mClearYieldReached As Boolean, mClearYieldAvailable As Boolean
Private mClearYieldLater As Long, mClearYieldState As String
'@)
    $adapter.AddFromString(@'
Public Sub RunClearYieldReset()
    mClearYieldTarget = "": mClearYieldReached = False: mClearYieldAvailable = False
    mClearYieldLater = 0: mClearYieldState = ""
End Sub
Public Function RunClearYieldArm(ByVal boundary As String) As Boolean
    RunClearYieldReset
    If Not mForm.RunClearYieldStageForTest() Then Exit Function
    mClearYieldTarget = boundary
    RunClearYieldArm = True
End Function
Public Sub RunClearYieldReturned(ByVal boundary As String, ByVal available As Boolean)
    If mClearYieldReached Then mClearYieldLater = mClearYieldLater + 1
    If mClearYieldTarget = "" Or boundary <> mClearYieldTarget Then Exit Sub
    mClearYieldTarget = "": mClearYieldReached = True: mClearYieldAvailable = available
    mClearYieldState = mForm.RunClearYieldStateForTest()
    modAuth.SignOut
End Sub
Public Function RunClearYieldEvidence() As String
    RunClearYieldEvidence = CStr(mClearYieldReached) & "|" & CStr(mClearYieldAvailable) & "|" & CStr(mClearYieldLater)
End Function
Public Function RunClearYieldPreserved() As Boolean
    RunClearYieldPreserved = mForm.RunClearYieldStateForTest() = mClearYieldState
End Function
Public Sub RunClearYieldReopen(ByVal workbookName As String)
    Dim values As Variant
    values = mForm.RunClearYieldFixtureForTest()
    RunLocalReopen workbookName
    mForm.RunClearYieldRestoreForTest values
End Sub
'@)
}

function Test-ProductionRunClearYield($Fixture,$Other,$Book){
    [void](Probe 'RunLocalRememberFixture')
    try{
        foreach($boundary in @('ClearNotification','InventoryPicker','DefaultLocation')){
            [void](Probe 'RunClearYieldReset')
            SelectTarget $Fixture 'config-producer';[void](Probe 'RunClearYieldReopen' @($Book.Name))
            [void](Probe 'RunFaultResetFixture')
            if(-not [bool](Probe 'RunSheetOwnerRecreate')){throw 'Clear yield owner surface unavailable; not product RED.'}
            if([string](Probe 'RunSheetOwnerStage' @('CLEAR','Normal',$canary)) -cne 'READY'){throw 'Clear yield staging unavailable; not product RED.'}
            if(-not [bool](Probe 'RunClearYieldArm' @($boundary))){throw 'Clear yield local controls unavailable; not product RED.'}
            $before=@(Files);$otherBefore=@(Get-Slice4beActivityFiles $Other)
            $notice=[string](Probe 'RunFaultAct' @('CLEAR'))
            $evidence=([string](Probe 'RunClearYieldEvidence')).Split('|');$label='RunClearYield.'+$boundary
            $reached=$evidence.Count -eq 3 -and $evidence[0] -ceq 'True' -and $evidence[1] -ceq 'True'
            if(-not $reached){throw ('Real Clear yield boundary unavailable: '+$boundary+'; not product RED.')}
            Check ($label+'.RealBoundaryBeforeSignOut') $reached
            Check ($label+'.NoLaterObservedBoundaries') ([int]$evidence[2] -eq 0)
            Check ($label+'.LocalStateAtBoundaryPreserved') ([bool](Probe 'RunClearYieldPreserved'))
            Check ($label+'.ExistingClearEffectsAndOneNotification') ([bool](Probe 'RunSheetOwnerResult' @('CLEAR','Normal')))
            Check ($label+'.RefusalVisible') ($notice.StartsWith('Session, warehouse, or captured workbook changed.') -and $notice.Contains('Reopen Production'))
            Check ($label+'.NoUnhandledError') (-not $notice.StartsWith('HANDLER_ERROR|'))
            Check ($label+'.GuardsRestoredWithoutAdapterReset') ([bool](Probe 'RunFaultGuards'))
            Check ($label+'.CustomColumnsFormulaAndExactKey') ([bool](Probe 'RunSheetOwnerPreserved'))
            $rows=@(Files|Where-Object{$_ -cnotin $before}|ForEach-Object{[IO.File]::ReadAllText($_)|ConvertFrom-Json})
            Check ($label+'.OriginalAttemptWithoutMisattributedOutcome') ($rows.Count -eq 1 -and $rows[0].ControlId -ceq 'PRODUCTION_RUN_CLEAR' -and $rows[0].OutcomeCode -ceq 'REQUESTED' -and $rows[0].UserId -ceq 'config-producer' -and $rows[0].WarehouseId -ceq $Fixture.Warehouse -and @($rows[0].SourceEventRefs).Count -eq 0)
            Check ($label+'.NoOtherWarehouseActivity') ((@(Get-Slice4beActivityFiles $Other) -join '|') -ceq ($otherBefore -join '|'))
            SelectTarget $Fixture 'config-producer'
            Check ($label+'.CanonicalSourcePreserved') ([bool](Probe 'RunWorksheetSource'))
        }
    }finally{
        [void](Probe 'RunClearYieldReset')
        SelectTarget $Fixture 'config-producer';[void](Probe 'RunClearYieldReopen' @($Book.Name))
    }
}
