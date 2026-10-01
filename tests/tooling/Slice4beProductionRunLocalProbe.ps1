# Real released definitions and Admin-seeded stock; original action handlers stay intact.
# Fault/staging helpers are unsaved adapters, not runtime implementations.
function Install-ProductionRunLocalProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $start=$form.ProcStartLine('DesignerReleasedProcessForTest',0)
    $body=$form.Lines($start,$form.ProcCountLines('DesignerReleasedProcessForTest',0)) -split '\r?\n'
    $anchor=@(for($i=0;$i -lt $body.Count;$i++){if($body[$i].Trim() -ieq 'stage = "SAVE"'){$start+$i}})
    if($anchor.Count -ne 1){throw 'Released Run fixture anchor changed; not product RED.'}
    $form.InsertLines($anchor[0],@'
    Dim runAlternative As Object
    mLstProcessRequirements.AddItem "A02"
    mLstProcessRequirements.List(0, 1) = fixtureName
    mLstProcessRequirements.List(0, 2) = "2"
    mLstProcessRequirements.List(0, 5) = "lbs"
    mLstProcessRequirements.List(0, 6) = "FIXED"
    Set runAlternative = NewReusableRecord("ALTERNATIVE")
    runAlternative("RequirementId") = "A02"
    runAlternative("ITEM_CODE") = "DEMO-RAW-BLACK-TEA"
    mProcessAlternatives.Add runAlternative
'@)
    $form.InsertLines($form.CountOfDeclarationLines+1,@'
Private mRunLocalInitialTreeForTest As Long
Private mRunLocalInitialOwnerForTest As String
Private mRunLocalSelectedKeyForTest As String
'@)
    $form.AddFromString(@'
Public Function RunLocalStageForTest(ByVal action As String, ByVal mode As String) As String
    Dim prior As Boolean, phase As String, row As Long, report As String
    On Error GoTo Failed
    prior = mLoading
    phase = "LoadRealRecipe"
    mReadModeForTest = ""
    mLoading = True
    mCmbRunLocation.Clear
    mCmbTreeRunLocation.Clear
    RefreshRecipeLists
    mLoading = prior
    mTxtBatchScalePercent.Text = "100"
    If Not LoadReusableRecipeIntoRun(mReadRecipeIdForTest, mReadRecipeVersionForTest) Then GoTo Unavailable
    phase = "SeededPalette"
    If mLstRunPalette.ListCount <> 1 Or mLstManagerOutput.ListCount <> 1 Then GoTo Unavailable
    mLoading = True
    mLstRunPalette.ListIndex = 0
    SelectComboText mCmbRunLocation, NzStr(mLstRunPalette.List(0, 9))
    SelectComboText mCmbTreeRunLocation, NzStr(mLstRunPalette.List(0, 9))
    mRunLocalSelectedKeyForTest = NzStr(mLstRunPalette.List(0, 3))
    phase = "PriorAllocation"
    If Not modProductionReusableRun.ApplyReusableRunStockAllocation( _
        NzStr(mLstRunPalette.List(0, 0)), NzStr(mLstRunPalette.List(0, 1)), _
        mRunLocalSelectedKeyForTest, 0.25, report) Then GoTo Unavailable
    If Not modProductionReusableRun.StageReusableRunActualOutput(1, "1", report) Then GoTo Unavailable
    If Not modProductionReusableRun.SetReusableRunBatchNote(mReadCanaryForTest, report) Then GoTo Unavailable
    RefreshReusableRunControls False
    mLoading = True
    row = FindIdentityListRow(mLstLoaderRecipes, mReadRecipeIdForTest, mReadRecipeVersionForTest)
    If row < 0 Then GoTo Unavailable
    mLstLoaderRecipes.ListIndex = row
    mLstRunPalette.ListIndex = 0
    mUpdatingPaletteInputs = True
    mTxtPaletteSplit.Text = "75": mTxtPaletteQty.Text = "0.5"
    mTxtTreePaletteSplit.Text = "75": mTxtTreePaletteQty.Text = "0.5"
    mUpdatingPaletteInputs = False
    If action = "SCALE" Then mTxtBatchScalePercent.Text = "200"
    phase = "CaseState"
    Select Case mode
        Case "Minimum": mTxtBatchScalePercent.Text = "0.001"
        Case "Maximum": mTxtBatchScalePercent.Text = "1000"
        Case "BelowMinimum": mTxtBatchScalePercent.Text = "0"
        Case "AboveMaximum": mTxtBatchScalePercent.Text = "1001"
        Case "BadScale": mTxtBatchScalePercent.Text = "not a number"
        Case "NoSelection"
            If action = "LOAD" Then mLstLoaderRecipes.ListIndex = -1 Else mLstRunPalette.ListIndex = -1
        Case "UnavailableRecipe"
            mLstLoaderRecipes.List(row, 0) = "ZZZ": mLstLoaderRecipes.List(row, 1) = "999999"
        Case "WrongLocation"
            mCmbRunLocation.AddItem "RUN-OTHER": mCmbTreeRunLocation.AddItem "RUN-OTHER"
            SelectComboText mCmbRunLocation, "RUN-OTHER"
            SelectComboText mCmbTreeRunLocation, "RUN-OTHER"
        Case "SyntheticCompleted": modProductionReusableRun.RunLocalCompletedFixtureForTest
        Case "RefillPalette": mLstRunPalette.Clear
    End Select
    mUpdatingPaletteInputs = True
    Select Case mode
        Case "NoQuantity": mTxtPaletteSplit.Text = "": mTxtPaletteQty.Text = ""
        Case "BadQuantity": mTxtPaletteQty.Text = "not a number"
        Case "NegativeQuantity": mTxtPaletteQty.Text = "-1"
        Case "OverRequirement": mTxtPaletteQty.Text = "3"
        Case "PercentOnly": mTxtPaletteSplit.Text = "50": mTxtPaletteQty.Text = ""
        Case "ZeroQuantity": mTxtPaletteQty.Text = "0"
    End Select
    mUpdatingPaletteInputs = False
    mRunLocalInitialTreeForTest = mLstRunTree.ListCount
    mRunLocalInitialOwnerForTest = modProductionReusableRun.RunLocalStateForTest()
    mReadCallsForTest = 0
    mLoading = prior
    RunLocalStageForTest = "READY"
    Exit Function
Unavailable:
    mUpdatingPaletteInputs = False: mLoading = prior
    RunLocalStageForTest = "FIXTURE_FAILED|" & phase
    Exit Function
Failed:
    mUpdatingPaletteInputs = False: mLoading = prior
    RunLocalStageForTest = "FIXTURE_FAILED|" & phase & "|" & CStr(Err.Number)
End Function
Public Function RunLocalActForTest(ByVal action As String, ByVal guard As String) As String
    Dim priorLoading As Boolean, priorBusy As Boolean
    On Error GoTo Failed
    priorLoading = mLoading: priorBusy = mDesignerActionInProgress
    If guard = "Loading" Then mLoading = True
    If guard = "Busy" Then mDesignerActionInProgress = True
    Select Case action
        Case "LOAD": mBtnLoaderLoad_Click
        Case "SCALE": mBtnApplyBatchScale_Click
        Case "CLEAR": mBtnLoaderClear_Click
        Case "LOADER_REFRESH": mBtnLoaderRefresh_Click
        Case "MANAGER_REFRESH": mBtnManagerRefresh_Click
        Case "ALLOCATE": mBtnRunApplyPalette_Click
        Case "TREE_ALLOCATE": mBtnRunTreeApplyPalette_Click
        Case Else: Err.Raise 5
    End Select
    RunLocalActForTest = mTxtStatus.Text
Finished:
    mLoading = priorLoading: mDesignerActionInProgress = priorBusy
    Exit Function
Failed:
    RunLocalActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
    Resume Finished
End Function
Public Function RunLocalStateForTest() As String
    RunLocalStateForTest = modProductionReusableRun.RunLocalStateForTest() & "|" & _
        CStr(mLstLoaderRecipes.ListIndex) & "|" & CStr(mLstLoaderLines.ListCount) & "|" & _
        CStr(mLstRunPalette.ListCount) & "|" & CStr(mLstRunPalette.ListIndex) & "|" & _
        CStr(mLstManagerOutput.ListCount) & "|" & CStr(mLstRunTree.ListCount) & "|" & _
        mTxtPaletteQty.Text & "|" & mTxtPaletteSplit.Text & "|" & ActiveRunLocation()
End Function
Public Function RunLocalShowForTest(ByVal action As String) As Long
    Dim prior As Boolean
    prior = mLoading: mLoading = True
    If action = "TREE_ALLOCATE" Then mPages.Value = 4 Else mPages.Value = 3
    mLoading = prior
    Me.Show vbModeless
    RunLocalShowForTest = mPages.Value
End Function
Public Function RunLocalClosedActForTest(ByVal action As String) As String
    TestProductionDesigner.RunLocalClosedEnteredForTest = True
    RunLocalClosedActForTest = RunLocalActForTest(action, "")
End Function
Public Function RunLocalPreservedForTest(ByVal action As String, ByVal mode As String) As Boolean
    Dim expectedScale As Double, expectedQty As Double, report As String
    If mode = "UnavailableRecipe" Then
        RunLocalPreservedForTest = Not modProductionReusableRun.ReusableRunIsLoaded() And _
            modProductionReusableRun.RunLocalAllocatedForTest() = 0 And _
            mLstRunPalette.ListCount = 1 And mLstManagerOutput.ListCount = 1
        Exit Function
    End If
    Select Case action
        Case "LOAD", "SCALE"
            expectedScale = 100
            If action = "SCALE" Then expectedScale = 200
            If mode = "Minimum" Then expectedScale = 0.001
            If mode = "Maximum" Then expectedScale = 1000
            RunLocalPreservedForTest = modProductionReusableRun.ReusableRunIsLoaded() And _
                modProductionReusableRun.ReusableRunScalePercent() = expectedScale And _
                modProductionReusableRun.RunLocalAllocatedForTest() = 0 And _
                modProductionReusableRun.ReusableRunActualOutput(1) = "" And _
                modProductionReusableRun.ReusableRunBatchNote() = "" And _
                Not modProductionReusableRun.ReusableRunIsCheckedIn() And _
                mLstLoaderLines.ListCount = 2 And mLstRunPalette.ListCount = 1 And mLstManagerOutput.ListCount = 1
        Case "CLEAR"
            RunLocalPreservedForTest = Not modProductionReusableRun.ReusableRunIsLoaded() And _
                modProductionReusableRun.RunLocalAllocatedForTest() = 0 And _
                mLstLoaderLines.ListCount = 0 And mLstRunPalette.ListCount = 0 And _
                mLstManagerCheck.ListCount = 0 And mLstManagerOutput.ListCount = 0 And _
                mLstRunInstructions.ListCount = 0 And mLstRunTree.ListCount = mRunLocalInitialTreeForTest
        Case "LOADER_REFRESH", "MANAGER_REFRESH"
            RunLocalPreservedForTest = modProductionReusableRun.RunLocalStateForTest() = mRunLocalInitialOwnerForTest And _
                mLstLoaderLines.ListCount = 2 And mLstRunPalette.ListCount = 1 And mLstManagerOutput.ListCount = 1
        Case "ALLOCATE", "TREE_ALLOCATE"
            expectedQty = 0.5
            If mode = "PercentOnly" Then expectedQty = 1
            If mode = "ZeroQuantity" Then expectedQty = 0
            If mode = "NoSelection" Or mode = "RefillPalette" Then expectedQty = 0.25
            RunLocalPreservedForTest = modProductionReusableRun.RunLocalAllocatedForTest() = expectedQty And _
                Not modProductionReusableRun.ReusableRunIsCheckedIn() And _
                Not modProductionReusableRun.ReusableRunIsCompleted() And _
                modProductionReusableRun.RunLocalOnlyExactKeyForTest(mRunLocalSelectedKeyForTest)
    End Select
End Function
'@)
    $project.VBComponents.Item('modProductionReusableRun').CodeModule.AddFromString(@'
Public Function RunLocalAllocatedForTest() As Double
    Dim key As Variant
    If mAllocations Is Nothing Then Exit Function
    For Each key In mAllocations.Keys
        RunLocalAllocatedForTest = RunLocalAllocatedForTest + CDbl(mAllocations(key))
    Next key
End Function
Public Function RunLocalOnlyExactKeyForTest(ByVal expectedKey As String) As Boolean
    Dim key As Variant
    RunLocalOnlyExactKeyForTest = True
    For Each key In mAllocations.Keys
        If AllocationSystemKey(CStr(key)) <> expectedKey Then RunLocalOnlyExactKeyForTest = False
    Next key
End Function
Public Sub RunLocalCompletedFixtureForTest()
    mCompletedNodes(RunRecordText(mNodes(1), "ProcessNodeId")) = True
End Sub
Public Function RunLocalStateForTest() As String
    Dim key As Variant, result As String
    result = CStr(mLoaded) & "|" & mRecipeId & "|" & mRecipeVersion & "|" & CStr(mScalePercent) & "|" & _
        CStr(mCheckedIn) & "|" & CStr(mCompleted) & "|" & mBatchNote & "|" & CStr(mBatchNoteFrozen) & "|" & _
        CStr(mNodes.Count) & "|" & CStr(mRequirements.Count) & "|" & CStr(mOutputs.Count) & "|" & _
        CStr(mCompletedNodes.Count) & "|" & mCheckedInNodeId & "|" & mRunId & "|" & mEventIds
    For Each key In mAllocations.Keys
        result = result & "|ALLOC|" & CStr(key) & "|" & CStr(mAllocations(key))
    Next key
    For Each key In mActualOutputQty.Keys
        result = result & "|ACTUAL|" & CStr(key) & "|" & CStr(mActualOutputQty(key))
    Next key
    For Each key In mOutputKeys.Keys
        result = result & "|OUTPUT|" & CStr(key) & "|" & CStr(mOutputKeys(key))
    Next key
    RunLocalStateForTest = result
End Function
'@)
    $adapter=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,@'
Private mRunLocalFixtureForTest As Variant
Private mRunLocalCapturedBookForTest As Workbook, mRunLocalCapturedContextForTest As String
Private mRunLocalClosedErrorForTest As Long
Public RunLocalClosedEnteredForTest As Boolean
'@)
    $adapter.AddFromString(@'
Public Sub RunLocalRememberFixture()
    mRunLocalFixtureForTest = mForm.DesignReadFixtureForTest()
End Sub
Public Sub RunLocalReopen(ByVal workbookName As String)
    OpenDesigner workbookName
    mForm.DesignReadRestoreFixtureForTest mRunLocalFixtureForTest
End Sub
Public Function RunLocalShowAndCapture(ByVal workbookName As String, ByVal action As String) As Long
    Set mRunLocalCapturedBookForTest = Application.Workbooks(workbookName)
    mRunLocalCapturedContextForTest = modActivity.CaptureContext()
    RunLocalShowAndCapture = mForm.RunLocalShowForTest(action)
End Function
Public Function RunLocalClosedAct(ByVal action As String) As String
    On Error GoTo Failed
    RunLocalClosedEnteredForTest = False: mRunLocalClosedErrorForTest = 0
    RunLocalClosedAct = mForm.RunLocalClosedActForTest(action)
    Exit Function
Failed:
    mRunLocalClosedErrorForTest = Err.Number
End Function
Public Function RunLocalClosedStatus() As String
    RunLocalClosedStatus = CStr(RunLocalClosedEnteredForTest) & "|" & CStr(mRunLocalClosedErrorForTest)
End Function
Public Function RunLocalClosedBindingCurrent() As Boolean
    RunLocalClosedBindingCurrent = modProductionDesignerActions.ContextIsCurrent(mRunLocalCapturedContextForTest, mRunLocalCapturedBookForTest)
End Function
Public Function RunLocalOwnerState() As String
    RunLocalOwnerState = modProductionReusableRun.RunLocalStateForTest()
End Function
Public Sub RunLocalSafeClose()
    On Error Resume Next
    Unload mForm: Set mForm = Nothing
    Set mRunLocalCapturedBookForTest = Nothing
End Sub
Public Function RunLocalStage(ByVal action As String, ByVal mode As String) As String
    On Error GoTo Failed
    RunLocalStage = mForm.RunLocalStageForTest(action, mode)
    Exit Function
Failed:
    RunLocalStage = "FIXTURE_FAILED|FormBinding|" & CStr(Err.Number)
End Function
Public Function RunLocalAct(ByVal action As String, ByVal guard As String) As String
    RunLocalAct = mForm.RunLocalActForTest(action, guard)
End Function
Public Function RunLocalState() As String
    RunLocalState = mForm.RunLocalStateForTest()
End Function
Public Function RunLocalPreserved(ByVal action As String, ByVal mode As String) As Boolean
    RunLocalPreserved = mForm.RunLocalPreservedForTest(action, mode)
End Function
'@)
}
