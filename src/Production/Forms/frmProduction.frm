VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProduction
   Caption         =   "Production"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13200
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProduction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

'@FormLayout Strategy=WINDOWS_API_ANCHORS MinWidth=1110 MinHeight=800 DefaultWidth=1110 DefaultHeight=800 ExpandedWidth=1350 ExpandedHeight=980
Private Const RUN_LOADER_RECIPE_WIDTHS As String = "0 pt;120 pt;130 pt"
Private Const RUN_LOADER_LINE_WIDTHS As String = "110 pt;0 pt;45 pt;155 pt;45 pt;45 pt;50 pt;95 pt;0 pt"
Private Const RUN_PALETTE_WIDTHS As String = "0 pt;0 pt;190 pt;0 pt;240 pt;45 pt;45 pt;145 pt;230 pt;115 pt"
Private Const RUN_OUTPUT_WIDTHS As String = "120 pt;190 pt;48 pt;72 pt;45 pt;72 pt;82 pt;100 pt;240 pt"
Private Const RUN_CHECK_WIDTHS As String = "52 pt;150 pt;170 pt;150 pt;72 pt;130 pt;50 pt;72 pt;90 pt"
Private Const RUN_LIST_CHECK_MIN_HEIGHT As Single = 102
Private Const RUN_LIST_INSTRUCTIONS_MIN_HEIGHT As Single = 52
Private Const RUN_LIST_SCROLL_HEIGHT As Single = 680
Private Const RUN_LIST_CHECK_READABLE_HEIGHT As Single = 96
Private Const RUN_LIST_INSTRUCTIONS_READABLE_HEIGHT As Single = 48
Private Const BATCH_SCALE_MIN_PERCENT As Double = 0.001
Private Const BATCH_SCALE_MAX_PERCENT As Double = 1000

Private WithEvents mPages As MSForms.MultiPage
Private WithEvents mLstBuilderRecipes As MSForms.ListBox
Private WithEvents mLstBuilderLines As MSForms.ListBox
Private WithEvents mCmbLineIo As MSForms.ComboBox
Private WithEvents mBtnBuilderRefresh As MSForms.CommandButton
Private WithEvents mBtnBuilderNew As MSForms.CommandButton
Private WithEvents mBtnBuilderLoad As MSForms.CommandButton
Private WithEvents mBtnBuilderSave As MSForms.CommandButton
Private WithEvents mBtnBuilderProcess As MSForms.CommandButton
Private WithEvents mBtnBuilderFormulas As MSForms.CommandButton
Private WithEvents mBtnBuilderClear As MSForms.CommandButton
Private WithEvents mBtnBuilderRelease As MSForms.CommandButton
Private WithEvents mBtnLineAdd As MSForms.CommandButton
Private WithEvents mBtnLineUpdate As MSForms.CommandButton
Private WithEvents mBtnLineRemove As MSForms.CommandButton
Private WithEvents mBtnLineMoveUp As MSForms.CommandButton
Private WithEvents mBtnLineMoveDown As MSForms.CommandButton
Private WithEvents mBtnLineUomAdd As MSForms.CommandButton

Private WithEvents mLstProcesses As MSForms.ListBox
Private WithEvents mBtnProcessRefresh As MSForms.CommandButton
Private WithEvents mBtnProcessNew As MSForms.CommandButton
Private WithEvents mBtnProcessLoad As MSForms.CommandButton
Private WithEvents mBtnProcessReuse As MSForms.CommandButton
Private WithEvents mBtnProcessValidate As MSForms.CommandButton
Private WithEvents mBtnProcessSave As MSForms.CommandButton
Private WithEvents mBtnProcessRelease As MSForms.CommandButton
Private WithEvents mBtnProcessObsolete As MSForms.CommandButton
Private WithEvents mBtnProcessClear As MSForms.CommandButton
Private WithEvents mBtnProcessWorksheetCreate As MSForms.CommandButton
Private WithEvents mBtnProcessWorksheetRetrieve As MSForms.CommandButton
Private WithEvents mBtnProcessWorksheetAddAlternative As MSForms.CommandButton
Private WithEvents mBtnProcessRequirementAdd As MSForms.CommandButton
Private WithEvents mBtnProcessRequirementUpdate As MSForms.CommandButton
Private WithEvents mBtnProcessRequirementRemove As MSForms.CommandButton
Private WithEvents mBtnProcessRequirementUp As MSForms.CommandButton
Private WithEvents mBtnProcessRequirementDown As MSForms.CommandButton
Private WithEvents mBtnProcessOutputAdd As MSForms.CommandButton
Private WithEvents mBtnProcessOutputUpdate As MSForms.CommandButton
Private WithEvents mBtnProcessOutputRemove As MSForms.CommandButton
Private WithEvents mBtnProcessOutputUp As MSForms.CommandButton
Private WithEvents mBtnProcessOutputDown As MSForms.CommandButton
Private WithEvents mBtnProcessInstructionAdd As MSForms.CommandButton
Private WithEvents mBtnProcessInstructionUpdate As MSForms.CommandButton
Private WithEvents mBtnProcessInstructionRemove As MSForms.CommandButton
Private WithEvents mBtnProcessInstructionUp As MSForms.CommandButton
Private WithEvents mBtnProcessInstructionDown As MSForms.CommandButton
Private mTxtProcessName As MSForms.TextBox
Private mTxtProcessId As MSForms.TextBox
Private mTxtProcessVersion As MSForms.TextBox
Private mTxtProcessDescription As MSForms.TextBox
Private WithEvents mLstProcessRequirements As MSForms.ListBox
Private mTxtRequirementId As MSForms.TextBox
Private mTxtRequirementName As MSForms.TextBox
Private mTxtRequirementQty As MSForms.TextBox
Private mTxtRequirementPercent As MSForms.TextBox
Private mTxtRequirementYieldBasis As MSForms.TextBox
Private mTxtRequirementUom As MSForms.TextBox
Private WithEvents mCmbRequirementQtyMode As MSForms.ComboBox
Private WithEvents mLstProcessOutputs As MSForms.ListBox
Private mTxtProcessOutputId As MSForms.TextBox
Private mTxtProcessOutputName As MSForms.TextBox
Private mTxtProcessOutputItemCode As MSForms.TextBox
Private mTxtProcessOutputDesignId As MSForms.TextBox
Private mTxtProcessOutputDesignVersion As MSForms.TextBox
Private mTxtProcessOutputQty As MSForms.TextBox
Private mTxtProcessOutputPercent As MSForms.TextBox
Private mTxtProcessOutputYieldBasis As MSForms.TextBox
Private WithEvents mCmbProcessOutputUom As MSForms.ComboBox
Private WithEvents mCmbProcessOutputQtyMode As MSForms.ComboBox
Private WithEvents mCmbOutputRegulationScope As MSForms.ComboBox
Private WithEvents mCmbOutputRegulationNode As MSForms.ComboBox
Private WithEvents mLstOutputRegulations As MSForms.ListBox
Private WithEvents mChkOutputRegulated As MSForms.CheckBox
Private mTxtOutputRegulationFloor As MSForms.TextBox
Private mTxtOutputRegulationCeiling As MSForms.TextBox
Private mTxtOutputRegulationHelp As MSForms.TextBox
Private WithEvents mBtnOutputRegulationApply As MSForms.CommandButton
Private WithEvents mBtnOutputRegulationClear As MSForms.CommandButton
Private WithEvents mBtnUomCatalogSend As MSForms.CommandButton
Private WithEvents mBtnUomCatalogRetrieve As MSForms.CommandButton
Private WithEvents mLstProcessInstructions As MSForms.ListBox
Private mTxtProcessInstruction As MSForms.TextBox

Private WithEvents mLstRecipes As MSForms.ListBox
Private WithEvents mLstReleasedProcesses As MSForms.ListBox
Private WithEvents mLstRecipeNodes As MSForms.ListBox
Private WithEvents mLstRecipeConnections As MSForms.ListBox
Private WithEvents mLstRecipeConnectionDisplay As MSForms.ListBox
Private WithEvents mLstRecipeValidation As MSForms.ListBox
Private WithEvents mBtnRecipeRefresh As MSForms.CommandButton
Private WithEvents mBtnRecipeNew As MSForms.CommandButton
Private WithEvents mBtnRecipeLoad As MSForms.CommandButton
Private WithEvents mBtnRecipeAddProcess As MSForms.CommandButton
Private WithEvents mBtnRecipeRemoveProcess As MSForms.CommandButton
Private WithEvents mBtnRecipeConnect As MSForms.CommandButton
Private WithEvents mBtnRecipeUpdateConnection As MSForms.CommandButton
Private WithEvents mBtnRecipeDisconnect As MSForms.CommandButton
Private WithEvents mBtnRecipeMoveUp As MSForms.CommandButton
Private WithEvents mBtnRecipeMoveDown As MSForms.CommandButton
Private WithEvents mBtnRecipeAutoOrder As MSForms.CommandButton
Private WithEvents mBtnRecipeValidate As MSForms.CommandButton
Private WithEvents mBtnRecipeSave As MSForms.CommandButton
Private WithEvents mBtnRecipeRelease As MSForms.CommandButton
Private WithEvents mBtnRecipeObsolete As MSForms.CommandButton
Private WithEvents mBtnRecipeClear As MSForms.CommandButton
Private mTxtReusableRecipeName As MSForms.TextBox
Private mTxtReusableRecipeId As MSForms.TextBox
Private mTxtReusableRecipeVersion As MSForms.TextBox
Private mTxtReusableRecipeDescription As MSForms.TextBox
Private WithEvents mCmbConnectionFromNode As MSForms.ComboBox
Private WithEvents mCmbConnectionOutput As MSForms.ComboBox
Private WithEvents mCmbConnectionToNode As MSForms.ComboBox
Private WithEvents mCmbConnectionRequirement As MSForms.ComboBox
Private mTxtConnectionQty As MSForms.TextBox
Private mTxtConnectionPercent As MSForms.TextBox
Private WithEvents mCmbConnectionUom As MSForms.ComboBox
Private mLblRecipeFinishedOutput As MSForms.Label
Private mProcessAlternatives As Collection
Private mProcessOutputRegulations As Object
Private mRecipeOutputRegulations As Collection
Private mSelectedProcessRequirementIndex As Long
Private mSelectedProcessOutputIndex As Long
Private mReusableActionTestInProgress As Boolean
Private mReusableTestSourceId As String
Private mReusableTestSinkId As String
Private mReusableTestRecipeId As String

Private WithEvents mLstAssignRecipes As MSForms.ListBox
Private WithEvents mLstAssignIngredients As MSForms.ListBox
Private WithEvents mTxtInventorySearch As MSForms.TextBox
Private WithEvents mLstAssignInventory As MSForms.ListBox
Private WithEvents mLstAssignAllowed As MSForms.ListBox
Private WithEvents mBtnAssignRefresh As MSForms.CommandButton
Private WithEvents mBtnAssignRecipe As MSForms.CommandButton
Private WithEvents mBtnAssignIngredient As MSForms.CommandButton
Private WithEvents mBtnAssignAdd As MSForms.CommandButton
Private WithEvents mBtnAssignRemove As MSForms.CommandButton
Private WithEvents mBtnAssignSave As MSForms.CommandButton
Private WithEvents mBtnAssignClear As MSForms.CommandButton

Private WithEvents mLstLoaderRecipes As MSForms.ListBox
Private WithEvents mLstLoaderLines As MSForms.ListBox
Private WithEvents mLstLoaderOutput As MSForms.ListBox
Private WithEvents mLstRunPalette As MSForms.ListBox
Private WithEvents mLstRunTree As MSForms.ListBox
Private WithEvents mCmbRunProcess As MSForms.ComboBox
Private WithEvents mCmbTreeRunProcess As MSForms.ComboBox
Private WithEvents mCmbRunLocation As MSForms.ComboBox
Private WithEvents mCmbTreeRunLocation As MSForms.ComboBox
Private WithEvents mBtnLoaderRefresh As MSForms.CommandButton
Private WithEvents mBtnLoaderLoad As MSForms.CommandButton
Private WithEvents mBtnLoaderClear As MSForms.CommandButton
Private WithEvents mBtnRunTreeExpandAll As MSForms.CommandButton
Private WithEvents mBtnRunTreeCollapseAll As MSForms.CommandButton
Private WithEvents mBtnRunTreeApplyPalette As MSForms.CommandButton

Private WithEvents mLstManagerOutput As MSForms.ListBox
Private WithEvents mLstManagerCheck As MSForms.ListBox
Private WithEvents mLstRunInstructions As MSForms.ListBox
Private WithEvents mBtnRunApplyPalette As MSForms.CommandButton
Private WithEvents mBtnApplyBatchScale As MSForms.CommandButton
Private WithEvents mBtnManagerCheckIn As MSForms.CommandButton
Private WithEvents mBtnManagerRefresh As MSForms.CommandButton
Private WithEvents mBtnManagerApplyOutput As MSForms.CommandButton
Private WithEvents mBtnManagerNext As MSForms.CommandButton
Private WithEvents mBtnManagerPrint As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton

Private mTxtRecipeName As MSForms.TextBox
Private mTxtRecipeId As MSForms.TextBox
Private mTxtRecipeDescription As MSForms.TextBox
Private mTxtRecipeRowBudget As MSForms.TextBox
Private mTxtLineProcess As MSForms.TextBox
Private mTxtLineIngredient As MSForms.TextBox
Private mTxtLinePercent As MSForms.TextBox
Private mTxtLineUom As MSForms.ComboBox
Private mTxtLineAmount As MSForms.TextBox
Private WithEvents mTxtPaletteSplit As MSForms.TextBox
Private WithEvents mTxtPaletteQty As MSForms.TextBox
Private WithEvents mTxtTreePaletteSplit As MSForms.TextBox
Private WithEvents mTxtTreePaletteQty As MSForms.TextBox
Private WithEvents mTxtOutputReal As MSForms.TextBox
Private WithEvents mTxtRunBatchNote As MSForms.TextBox
Private WithEvents mTxtBatchScalePercent As MSForms.TextBox
Private mChkRunTargetOutputScale As MSForms.CheckBox
Private mCmbRunTargetOutput As MSForms.ComboBox
Private mTxtRunTargetOutputQty As MSForms.TextBox
Private mTxtStatus As MSForms.TextBox
Private mOperatorWorkbook As Workbook
Private mLayout As cOperationsAnchorManager
Private mOperatorWorkbookCaptured As Boolean
Private mInventoryRows As Variant
Private mInventoryCacheLoaded As Boolean
Private mRunInventoryRows As Variant
Private mRunInventoryCacheLoaded As Boolean
Private mBuilt As Boolean
Private mLoading As Boolean
Private mResizeInitialized As Boolean
Private mResizingLayout As Boolean
Private mRunTreeCollapsed As Object
Private mRunSplitOverrides As Object
Private mRunBaseQtyByKey As Object
Private mUpdatingPaletteInputs As Boolean
Private mPaletteInputSource As String
Private mCheckRowsDiagnostic As String
Private mRunProcessByKey As Object
Private mRunItemCodeByKey As Object
Private mBuilderLineTableRows() As Long
Private mBuilderLineTableRowCount As Long

Private Const ASSIGN_INVENTORY_MAX_VISIBLE As Long = 250
Private Const PRODUCTION_MIN_WIDTH As Double = 1110
Private Const PRODUCTION_MIN_HEIGHT As Double = 800
Private Const PRODUCTION_DEFAULT_WIDTH As Double = 1110
Private Const PRODUCTION_DEFAULT_HEIGHT As Double = 800
Private Const PRODUCTION_LAYOUT_TEST_MAX_WIDTH As Double = 1350
Private Const PRODUCTION_LAYOUT_TEST_MAX_HEIGHT As Double = 980
Private Const PRODUCTION_DEFAULT_ROW_BUDGET As Long = 50
Private Const PRODUCTION_MAX_ROW_BUDGET As Long = 1000
Private Const RUN_TREE_PARENT_MARKER As String = "__RUN_TREE_PARENT__"
Private Const RUN_TREE_OUTPUT_MARKER As String = "__RUN_TREE_OUTPUT__"

Private Const TABLE_BUILDER_HEADER As String = "RB_AddRecipeName"
Private Const TABLE_BUILDER_LINES As String = "RecipeBuilder"
Private Const TABLE_ASSIGN_RECIPE As String = "IP_ChooseRecipe"
Private Const TABLE_ASSIGN_INGREDIENT As String = "IP_ChooseIngredient"
Private Const TABLE_ASSIGN_ITEM As String = "IP_ChooseItem"
Private Const TABLE_LOADER_CHOOSE As String = "RC_RecipeChoose"
Private Const TABLE_LOADER_LINES As String = "RecipeChooser_generated"
Private Const TABLE_MANAGER_PALETTE As String = "InventoryPalette_generated"
Private Const TABLE_MANAGER_OUTPUT As String = "ProductionOutput"
Private Const TABLE_MANAGER_CHECK As String = "Prod_invSys_Check"

Private Sub UserForm_Initialize()
    On Error GoTo FailInitialize
    BuildLayout
    Exit Sub

FailInitialize:
    MsgBox "Production form failed to initialize: " & Err.Description, vbExclamation, "invSys Production"
End Sub

Private Sub UserForm_Activate()
    On Error GoTo FailActivate
    If Not mResizeInitialized Then
        mResizingLayout = True
        On Error Resume Next
        modProductionFormWindow.EnableResizable Me, True, True
        modProductionFormWindow.ApplyDpiLayoutZoom Me
        On Error GoTo FailActivate
        ConfigureProductionAnchors
        mResizingLayout = False
        mResizeInitialized = True
    End If
    ResizeProductionLayout
    Exit Sub

FailActivate:
    mResizingLayout = False
    On Error Resume Next
    ShowStatus "Production form activation skipped: " & Err.Description
    On Error GoTo 0
End Sub

Private Sub UserForm_Resize()
    On Error Resume Next
    ResizeProductionLayout
    On Error GoTo 0
End Sub

Private Sub UserForm_Terminate()
    mProduction.ClearProductionOperatorWorkbookBinding mOperatorWorkbook
    Set mOperatorWorkbook = Nothing
    Set mRunTreeCollapsed = Nothing
    Set mRunSplitOverrides = Nothing
    Set mRunBaseQtyByKey = Nothing
    Set mRunProcessByKey = Nothing
    Set mRunItemCodeByKey = Nothing
End Sub

Public Sub SetOperatorWorkbook(ByVal wb As Workbook)
    If IsUsableWorkbook(wb) Then
        Set mOperatorWorkbook = wb
        mOperatorWorkbookCaptured = True
        mProduction.BindProductionOperatorWorkbook wb
    End If
End Sub

Public Sub InitializeFromProduction()
    On Error GoTo ErrHandler

    Dim wb As Workbook
    If Not mBuilt Then BuildLayout
    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then
        If mOperatorWorkbookCaptured Then
            ShowStatus "The captured Production operator workbook is closed or no longer valid. Reopen the form from the intended operator workbook."
        Else
            ShowStatus "Open a Production operator workbook before using the Production form."
        End If
        Exit Sub
    End If

    mLoading = True
    BindOperatorWorkbookForRun
    mProduction.InitializeProductionUiForWorkbook wb
    ResetInventoryCache
    RefreshAllViews
    EnsureRecipeDraftIdentity
    mLoading = False
    ShowStatus "Production form loaded for " & wb.Name & ". " & _
               NzStr(mProduction.GetProductionInventoryModeStatus()) & " " & _
               NzStr(mProduction.GetProductionDesignsModeStatus())
    Exit Sub

ErrHandler:
    mLoading = False
    ShowStatus "Production form load failed: " & Err.Description
End Sub

Public Function TestInitializeForWorkbook(ByVal wb As Workbook) As Long
    SetOperatorWorkbook wb
    InitializeFromProduction
    TestInitializeForWorkbook = TestPageCount()
End Function

Public Function TestPageCount() As Long
    If mPages Is Nothing Then BuildLayout
    TestPageCount = mPages.Pages.Count
End Function

Public Function TestLayoutGeometryReportForSize(ByVal requestedWidth As Double, _
                                                ByVal requestedHeight As Double, _
                                                Optional ByVal pageIndex As Long = 2) As String
    If Not mBuilt Then BuildLayout
    Me.Width = requestedWidth
    Me.Height = requestedHeight
    ResizeProductionLayout
    TestLayoutGeometryReportForSize = _
        BuildLayoutGeometryReport(requestedWidth, requestedHeight, pageIndex)
End Function

Public Function TestPrepareLayoutForScreenshot(ByVal requestedWidth As Double, _
                                               ByVal requestedHeight As Double, _
                                               Optional ByVal pageIndex As Long = 2) As String
    Me.Caption = "Production Layout Validation"
    TestPrepareLayoutForScreenshot = _
        TestLayoutGeometryReportForSize(requestedWidth, requestedHeight, pageIndex)
End Function

Public Function TestCurrentLayoutGeometryReport(Optional ByVal pageIndex As Long = 2) As String
    TestCurrentLayoutGeometryReport = _
        BuildLayoutGeometryReport(Me.Width, Me.Height, pageIndex)
End Function

Public Function TestRunListResponsiveLayoutReportForSize() As String
    Dim originalWidth As Double
    Dim originalHeight As Double
    Dim baselineRecipeHeight As Double
    Dim baselinePlanHeight As Double
    Dim baselinePaletteHeight As Double
    Dim baselineCheckHeight As Double
    Dim baselineInstructionHeight As Double
    Dim baselineOutputHeight As Double
    Dim readableRows As Boolean
    Dim listsGrew As Boolean
    Dim instructionLeftStable As Boolean
    Dim headersAligned As Boolean
    Dim runPlanQtyHeaderAligned As Boolean
    Dim verticalScrollbarVisible As Boolean
    Dim geometryHealthy As Boolean

    If Not mBuilt Then BuildLayout
    originalWidth = Me.Width
    originalHeight = Me.Height

    Me.Width = PRODUCTION_DEFAULT_WIDTH
    Me.Height = PRODUCTION_DEFAULT_HEIGHT
    ResizeProductionLayout
    baselineRecipeHeight = mLstLoaderRecipes.Height
    baselinePlanHeight = mLstLoaderLines.Height
    baselinePaletteHeight = mLstRunPalette.Height
    baselineCheckHeight = mLstManagerCheck.Height
    baselineInstructionHeight = mLstRunInstructions.Height
    baselineOutputHeight = mLstManagerOutput.Height
    readableRows = baselineCheckHeight >= RUN_LIST_CHECK_READABLE_HEIGHT And _
                   baselineInstructionHeight >= RUN_LIST_INSTRUCTIONS_READABLE_HEIGHT
    headersAligned = RunListHeaderBandAligned("ManagerCheck", mLstManagerCheck) And _
                     RunListHeaderBandAligned("RunInstructions", mLstRunInstructions) And _
                     RunListHeaderBandAligned("ManagerOutput", mLstManagerOutput)
    runPlanQtyHeaderAligned = RunListHeaderCellAligned("LoaderLines", 7, mLstLoaderLines)
    verticalScrollbarVisible = mPages.Pages(3).ScrollBars = fmScrollBarsVertical And _
                              mPages.Pages(3).KeepScrollBarsVisible = fmScrollBarsVertical

    Me.Width = PRODUCTION_LAYOUT_TEST_MAX_WIDTH
    Me.Height = PRODUCTION_LAYOUT_TEST_MAX_HEIGHT
    ResizeProductionLayout
    listsGrew = mLstLoaderRecipes.Height > baselineRecipeHeight And _
                mLstLoaderLines.Height > baselinePlanHeight And _
                mLstRunPalette.Height > baselinePaletteHeight And _
                mLstManagerCheck.Height > baselineCheckHeight And _
                mLstRunInstructions.Height > baselineInstructionHeight And _
                mLstManagerOutput.Height > baselineOutputHeight
    instructionLeftStable = Abs(mLstRunInstructions.Left - mLstManagerCheck.Left) < 0.1
    headersAligned = headersAligned And _
                     RunListHeaderBandAligned("ManagerCheck", mLstManagerCheck) And _
                     RunListHeaderBandAligned("RunInstructions", mLstRunInstructions) And _
                     RunListHeaderBandAligned("ManagerOutput", mLstManagerOutput)
    runPlanQtyHeaderAligned = runPlanQtyHeaderAligned And _
                               RunListHeaderCellAligned("LoaderLines", 7, mLstLoaderLines)
    geometryHealthy = Left$(BuildLayoutGeometryReport(Me.Width, Me.Height, 3), 3) = "OK|"

    Me.Width = originalWidth
    Me.Height = originalHeight
    ResizeProductionLayout

    TestRunListResponsiveLayoutReportForSize = "OK" & _
        "|CheckEightRows=" & CStr(readableRows) & _
        "|InstructionsFourRows=" & CStr(readableRows) & _
        "|CheckHeight=" & Format$(baselineCheckHeight, "0.0") & _
        "|InstructionHeight=" & Format$(baselineInstructionHeight, "0.0") & _
        "|AllListsGrew=" & CStr(listsGrew) & _
        "|InstructionLeftStable=" & CStr(instructionLeftStable) & _
        "|HeadersAligned=" & CStr(headersAligned) & _
        "|RunListVerticalScrollbar=" & CStr(verticalScrollbarVisible) & _
        "|RunPlanQtyHeaderAligned=" & CStr(runPlanQtyHeaderAligned) & _
        "|GeometryHealthy=" & CStr(geometryHealthy)
End Function

Private Function BuildLayoutGeometryReport(ByVal requestedWidth As Double, _
                                           ByVal requestedHeight As Double, _
                                           ByVal pageIndex As Long) As String
    Dim issueDetail As String
    Dim outOfBoundsCount As Long
    Dim overlapCount As Long
    Dim i As Long
    Dim windowStyle As String
    Dim resultState As String

    If pageIndex < 0 Then pageIndex = 0
    If pageIndex >= mPages.Pages.Count Then pageIndex = mPages.Pages.Count - 1
    mPages.Value = pageIndex
    DoEvents

    outOfBoundsCount = CountOutOfBoundsControls(Me, issueDetail)
    overlapCount = CountOverlappingInteractiveControls(Me, issueDetail)
    For i = 0 To mPages.Pages.Count - 1
        outOfBoundsCount = outOfBoundsCount + _
                           CountOutOfBoundsControls(mPages.Pages(i), issueDetail)
        overlapCount = overlapCount + _
                       CountOverlappingInteractiveControls(mPages.Pages(i), issueDetail)
    Next i

    Call modProductionFormWindow.EnableResizable(Me, True, True)
    windowStyle = modProductionFormWindow.DiagnoseWindowStyle(Me)
    resultState = "OK"
    If Me.Width < PRODUCTION_MIN_WIDTH Or Me.Height < PRODUCTION_MIN_HEIGHT Then resultState = "FAIL"
    If outOfBoundsCount <> 0 Or overlapCount <> 0 Then resultState = "FAIL"
    If InStr(1, windowStyle, "Resizable=True", vbTextCompare) = 0 Then resultState = "FAIL"
    If InStr(1, windowStyle, "Minimize=True", vbTextCompare) = 0 Then resultState = "FAIL"
    If InStr(1, windowStyle, "Maximize=True", vbTextCompare) = 0 Then resultState = "FAIL"

    BuildLayoutGeometryReport = resultState & _
        "|Requested=" & Format$(requestedWidth, "0.0") & "x" & Format$(requestedHeight, "0.0") & _
        "|Actual=" & Format$(Me.Width, "0.0") & "x" & Format$(Me.Height, "0.0") & _
        "|Page=" & CStr(pageIndex) & _
        "|Zoom=" & CStr(Me.Zoom) & _
        "|Anchors=" & CStr(mLayout.RegisteredControlCount) & _
        "|OutOfBounds=" & CStr(outOfBoundsCount) & _
        "|Overlap=" & CStr(overlapCount) & _
        "|WindowStyle=" & windowStyle & _
        "|Detail=" & issueDetail
End Function

Public Function TestStatusText() As String
    If mTxtStatus Is Nothing Then BuildLayout
    TestStatusText = mTxtStatus.Text
End Function

Public Function TestRunTwoConsecutiveBatchesForWorkbook(ByVal operatorWb As Workbook, _
                                                        ByVal inputItemCode As String, _
                                                        ByVal inputItemName As String, _
                                                        ByVal inputQty As Double, _
                                                        ByVal inputUom As String, _
                                                        ByVal inputLocation As String, _
                                                        ByVal outputQty As Double, _
                                                        Optional ByVal activatedWb As Workbook = Nothing) As String
    Dim batchNumber As Long
    Dim batchStatus As String
    Dim activeChoices As MSForms.ListBox
    Dim splitInput As MSForms.TextBox
    Dim qtyInput As MSForms.TextBox
    Dim choiceIndex As Long
    Dim actionStage As String

    On Error GoTo FailAction
    actionStage = "BuildLayout"
    If Not mBuilt Then BuildLayout
    actionStage = "BindWorkbook"
    SetOperatorWorkbook operatorWb

    For batchNumber = 1 To 2
        actionStage = "PrepareChoice"
        If Not PrepareRunChoiceForActionTest(inputItemCode, inputItemName, inputQty, _
                                             inputUom, inputLocation) Then
            TestRunTwoConsecutiveBatchesForWorkbook = _
                "FAIL|Batch=" & CStr(batchNumber) & "|Prepare|" & TestStatusText()
            Exit Function
        End If
        actionStage = "ActivateAlternateWorkbook"
        If Not activatedWb Is Nothing Then activatedWb.Activate
        actionStage = "SetLocation"
        SetRunLocationForActionTest inputLocation
        actionStage = "BuildRunTree"
        BuildRunTreeFromPaletteList
        actionStage = "ResolvePalette"
        Set activeChoices = ActiveRunPaletteList()
        actionStage = "ResolveChoice"
        choiceIndex = FirstSelectableRunChoice(activeChoices)
        If choiceIndex < 0 Then
            TestRunTwoConsecutiveBatchesForWorkbook = _
                "FAIL|Batch=" & CStr(batchNumber) & "|Prepare|No selectable inventory choice."
            Exit Function
        End If
        actionStage = "SelectChoice"
        activeChoices.ListIndex = choiceIndex
        If activeChoices Is mLstRunTree Then
            mLstRunTree_Click
        Else
            mLstRunPalette_Click
        End If
        actionStage = "ResolveInputs"
        Set splitInput = ActiveRunSplitTextBox()
        Set qtyInput = ActiveRunQtyTextBox()
        splitInput.Text = "100"
        qtyInput.Text = CStr(inputQty)
        actionStage = "ApplyPalette"
        mBtnRunApplyPalette_Click
        actionStage = "CheckIn"
        mBtnManagerCheckIn_Click
        batchStatus = TestStatusText()
        If InStr(1, batchStatus, "Checked in ", vbTextCompare) = 0 Then
            TestRunTwoConsecutiveBatchesForWorkbook = _
                "FAIL|Batch=" & CStr(batchNumber) & "|CheckIn|" & batchStatus
            Exit Function
        End If

        actionStage = "RefreshManager"
        RefreshManagerState
        If mLstManagerOutput.ListCount = 0 Then
            TestRunTwoConsecutiveBatchesForWorkbook = _
                "FAIL|Batch=" & CStr(batchNumber) & "|OutputMissing|" & TestStatusText()
            Exit Function
        End If
        actionStage = "SelectOutput"
        If modProductionReusableRun.ReusableRunIsLoaded() Then
            mLstManagerOutput.ListIndex = FirstReusableActiveOutputRow()
        Else
            mLstManagerOutput.ListIndex = 0
        End If
        mLstManagerOutput_Click
        mTxtOutputReal.Text = CStr(outputQty)
        actionStage = "CompleteRun"
        mBtnManagerApplyOutput_Click
        batchStatus = TestStatusText()
        If InStr(1, batchStatus, "Production run completed.", vbTextCompare) = 0 Then
            TestRunTwoConsecutiveBatchesForWorkbook = _
                "FAIL|Batch=" & CStr(batchNumber) & "|Complete|" & batchStatus
            Exit Function
        End If
        If batchNumber = 1 Then
            actionStage = "NextBatch"
            mBtnManagerNext_Click
        End If
    Next batchNumber

    TestRunTwoConsecutiveBatchesForWorkbook = _
        "OK|Batches=2|BoundWorkbook=" & mOperatorWorkbook.Name
    Exit Function

FailAction:
    TestRunTwoConsecutiveBatchesForWorkbook = "FAIL|Batch=" & CStr(batchNumber) & _
        "|Stage=" & actionStage & "|Error=" & CStr(Err.Number) & _
        "|Source=" & Err.Source & "|Description=" & Err.Description
End Function

Private Sub SetRunLocationForActionTest(ByVal locationValue As String)
    locationValue = Trim$(locationValue)
    If locationValue = "" Then Exit Sub
    EnsureComboValueForActionTest mCmbRunLocation, locationValue
    EnsureComboValueForActionTest mCmbTreeRunLocation, locationValue
End Sub

Private Sub EnsureComboValueForActionTest(ByVal combo As MSForms.ComboBox, ByVal valueText As String)
    Dim i As Long
    Dim found As Boolean

    If combo Is Nothing Then Exit Sub
    For i = 0 To combo.ListCount - 1
        If StrComp(Trim$(NzStr(combo.List(i))), valueText, vbTextCompare) = 0 Then
            found = True
            Exit For
        End If
    Next i
    If Not found Then combo.AddItem valueText
    combo.Value = valueText
End Sub

Private Function FirstSelectableRunChoice(ByVal choices As MSForms.ListBox) As Long
    Dim i As Long

    FirstSelectableRunChoice = -1
    If choices Is Nothing Then Exit Function
    For i = 0 To choices.ListCount - 1
        If Not IsRunTreeParentRow(choices, i) And Not IsRunTreeOutputRow(choices, i) Then
            FirstSelectableRunChoice = i
            Exit Function
        End If
    Next i
End Function

Private Function PrepareRunChoiceForActionTest(ByVal itemCode As String, _
                                               ByVal itemName As String, _
                                               ByVal qtyValue As Double, _
                                               ByVal uomValue As String, _
                                               ByVal locationValue As String) As Boolean
    Dim values(1 To 1, 1 To 11) As Variant

    mLstRunPalette.Clear
    Set mRunItemCodeByKey = Nothing
    Set mRunProcessByKey = Nothing
    Set mRunBaseQtyByKey = Nothing
    values(1, 1) = "TEST"
    values(1, 2) = "TEST-INPUT"
    values(1, 3) = "Test input"
    values(1, 4) = ""
    values(1, 5) = itemName
    values(1, 6) = ""
    values(1, 7) = ""
    values(1, 8) = uomValue
    values(1, 9) = locationValue
    values(1, 10) = qtyValue
    values(1, 11) = itemCode
    AddRunChoiceRows values
    If mLstRunPalette.ListCount <> 1 Then Exit Function
    mLstRunPalette.ListIndex = 0
    PrepareRunChoiceForActionTest = True
End Function

Public Function TestSelectedProductionOutputTableRow(ByVal wb As Workbook, ByVal listIndex As Long) As Long
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    RefreshProductionOutputList mLstManagerOutput
    If listIndex < 0 Or listIndex >= mLstManagerOutput.ListCount Then Exit Function
    mLstManagerOutput.ListIndex = listIndex
    TestSelectedProductionOutputTableRow = SelectedProductionOutputTableRow()
End Function

Public Function TestProductionOutputDisplayedBatch(ByVal wb As Workbook, ByVal listIndex As Long) As String
    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    RefreshProductionOutputList mLstManagerOutput
    If listIndex < 0 Or listIndex >= mLstManagerOutput.ListCount Then Exit Function
    TestProductionOutputDisplayedBatch = NzStr(mLstManagerOutput.List(listIndex, 4))
End Function

Public Function TestFillAssignmentIoCount(ByVal values As Variant, ByVal ioValue As String) As Long
    Dim i As Long
    Dim wanted As String

    If Not mBuilt Then BuildLayout
    FillIngredientListFromArray mLstAssignIngredients, values
    wanted = UCase$(Trim$(ioValue))
    For i = 0 To mLstAssignIngredients.ListCount - 1
        If UCase$(Trim$(NzStr(mLstAssignIngredients.List(i, 4)))) = wanted Then
            TestFillAssignmentIoCount = TestFillAssignmentIoCount + 1
        End If
    Next i
End Function

Public Function TestFilterAssignmentInventoryCount(ByVal values As Variant, ByVal filterText As String) As Long
    If Not mBuilt Then BuildLayout
    FillInventoryListFromArray values, filterText
    TestFilterAssignmentInventoryCount = mLstAssignInventory.ListCount
End Function

Public Function TestProductionRunLocationAllowed(ByVal runLocation As String, ByVal inventoryLocation As String, ByVal qtyValue As Double) As Long
    TestProductionRunLocationAllowed = IIf(RunChoiceLocationAllowed(runLocation, inventoryLocation, qtyValue), 1, 0)
End Function

Public Function TestRunPaletteCanonicalItemCodeStorage() As Long
    Dim values(1 To 1, 1 To 11) As Variant

    If Not mBuilt Then BuildLayout
    mLstRunPalette.Clear
    mRunInventoryRows = Empty
    mRunInventoryCacheLoaded = True
    Set mRunItemCodeByKey = Nothing

    values(1, 1) = "BREW"
    values(1, 2) = "ING-WATER"
    values(1, 3) = "Filtered Water"
    values(1, 4) = ""
    values(1, 5) = "Filtered Water"
    values(1, 7) = 500
    values(1, 8) = "LB"
    values(1, 10) = 500
    values(1, 11) = "ITEM-0061"
    AddRunChoiceRows values

    If mLstRunPalette.ColumnCount = 10 _
       And mLstRunPalette.ListCount = 1 _
       And StrComp(RunItemCodeFromList(mLstRunPalette, 0), "ITEM-0061", vbTextCompare) = 0 Then
        TestRunPaletteCanonicalItemCodeStorage = 1
    End If
End Function

Public Function TestAssignmentItemRowsGrowWithoutTableCollision(ByVal wb As Workbook) As Long
    Dim lo As ListObject
    Dim lr As ListRow
    Dim i As Long

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    Set lo = ProductionTable(TABLE_ASSIGN_ITEM)
    If lo Is Nothing Then Exit Function
    ClearTableContentsKeepBlank lo

    For i = 1 To 3
        Set lr = WritableAssignmentItemRow(lo)
        If lr Is Nothing Then Exit Function
        SetRowValue lr, lo, "ITEMS", "Test acceptable " & CStr(i)
        SetRowValue lr, lo, "ITEM_CODE", "TEST-SKU-" & CStr(i)
    Next i

    If lo.ListRows.Count >= 3 _
       And StrComp(CellByHeader(lo, 3, "ITEM_CODE"), "TEST-SKU-3", vbTextCompare) = 0 Then
        TestAssignmentItemRowsGrowWithoutTableCollision = 1
    End If
End Function

Public Function TestProductionCheckRowsRecognizeSkuIdentity(ByVal wb As Workbook) As Long
    Dim lo As ListObject

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If lo Is Nothing Then Exit Function
    ClearTableContentsKeepBlank lo
    SetCellByHeader lo, 1, "System_Key", "SYS-CHECK-ONLY"
    SetCellByHeader lo, 1, "ITEM_CODE", "SKU-CHECK-ONLY"
    SetCellByHeader lo, 1, "ITEM", "SKU-only checked input"
    SetCellByHeader lo, 1, "USED", 12
    If HasProductionCheckRows() Then TestProductionCheckRowsRecognizeSkuIdentity = 1
End Function

Public Function TestRecipeBuilderSelectedLineProcessUpdate(ByVal wb As Workbook, _
                                                           ByVal newProcess As String, _
                                                           Optional ByVal selectedIndex As Long = 0) As Long
    Dim lo As ListObject
    Dim tableRow As Long

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    RefreshBuilderLines
    If mLstBuilderLines Is Nothing Then Exit Function
    If selectedIndex < 0 Or selectedIndex >= mLstBuilderLines.ListCount Then Exit Function
    mLstBuilderLines.ListIndex = selectedIndex
    tableRow = BuilderTableRowForListIndex(selectedIndex)
    If tableRow <= 0 Then Exit Function
    LoadSelectedBuilderLine
    mTxtLineProcess.Text = newProcess
    If Not WriteSelectedRecipeBuilderLineFromForm(False) Then Exit Function
    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If Not lo Is Nothing Then
        If StrComp(CellByHeader(lo, tableRow, "PROCESS"), newProcess, vbTextCompare) = 0 Then _
            TestRecipeBuilderSelectedLineProcessUpdate = 1
    End If
End Function

Public Function TestRecipeBuilderLineMove(ByVal wb As Workbook, _
                                          ByVal selectedIndex As Long, _
                                          ByVal moveDirection As Long) As String
    Dim lo As ListObject

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    RefreshBuilderLines
    If mLstBuilderLines Is Nothing Then Exit Function
    If selectedIndex < 0 Or selectedIndex >= mLstBuilderLines.ListCount Then Exit Function
    mLstBuilderLines.ListIndex = selectedIndex
    If Not MoveSelectedRecipeBuilderLine(moveDirection, False) Then Exit Function

    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    TestRecipeBuilderLineMove = CellByHeader(lo, 1, "INGREDIENT") & "|" & _
                                CellByHeader(lo, 2, "INGREDIENT") & "|" & _
                                CellByHeader(lo, 1, "RECIPE_LIST_ROW") & "|" & _
                                CellByHeader(lo, 2, "RECIPE_LIST_ROW")
End Function

Public Function TestRecipeBuilderHasInstructionIo() As Long
    Dim i As Long

    If Not mBuilt Then BuildLayout
    For i = 0 To mCmbLineIo.ListCount - 1
        If StrComp(NzStr(mCmbLineIo.List(i)), "INSTRUCTION", vbTextCompare) = 0 Then
            TestRecipeBuilderHasInstructionIo = 1
            Exit Function
        End If
    Next i
End Function

Public Function TestRecipeBuilderLineActionsFitLayout() As Long
    If Not mBuilt Then BuildLayout

    If mBtnLineAdd.Left + mBtnLineAdd.Width + 10 > mBtnLineUpdate.Left Then Exit Function
    If mBtnLineUpdate.Left + mBtnLineUpdate.Width + 10 > mBtnLineRemove.Left Then Exit Function
    If mBtnLineRemove.Left + mBtnLineRemove.Width + 10 > mBtnLineMoveUp.Left Then Exit Function
    If mBtnLineMoveUp.Left + mBtnLineMoveUp.Width + 10 > mBtnLineMoveDown.Left Then Exit Function
    If mBtnLineMoveDown.Left + mBtnLineMoveDown.Width + 10 > mBtnBuilderClear.Left Then Exit Function
    If mBtnLineAdd.Top + mBtnLineAdd.Height + 10 > mLstBuilderLines.Top Then Exit Function

    TestRecipeBuilderLineActionsFitLayout = 1
End Function

Public Function TestRecipeBuilderLifecycleAndHeadersReady() As Long
    If Not mBuilt Then BuildLayout
    If mBtnBuilderRelease Is Nothing Then Exit Function
    If StrComp(mBtnBuilderRelease.Caption, "Release for Production", vbTextCompare) <> 0 Then Exit Function
    If Not ControlExistsByName(mPages.Pages(0), "hdrBuilderLines1") Then Exit Function
    If Not ControlExistsByName(mPages.Pages(0), "hdrBuilderLines8") Then Exit Function
    TestRecipeBuilderLifecycleAndHeadersReady = 1
End Function

Public Function TestRecipeBuilderUomCatalogContains(ByVal uomName As String) As Long
    Dim idx As Long

    If Not mBuilt Then BuildLayout
    RefreshRecipeUomCatalog
    For idx = 0 To mTxtLineUom.ListCount - 1
        If StrComp(CStr(mTxtLineUom.List(idx)), Trim$(uomName), vbTextCompare) = 0 Then
            TestRecipeBuilderUomCatalogContains = 1
            Exit Function
        End If
    Next idx
End Function

Public Function TestRecipeBuilderSelectUom(ByVal uomName As String) As String
    If Not mBuilt Then BuildLayout
    RefreshRecipeUomCatalog uomName
    If mTxtLineUom.ListIndex >= 0 Then _
        TestRecipeBuilderSelectUom = CStr(mTxtLineUom.List(mTxtLineUom.ListIndex))
End Function

Public Function TestWriteRecipeHeaderToBoundWorkbook(ByVal wb As Workbook, _
                                                     ByVal recipeName As String, _
                                                     ByVal recipeId As String) As String
    Dim lo As ListObject

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb
    mTxtRecipeName.Text = recipeName
    mTxtRecipeId.Text = recipeId
    mTxtRecipeDescription.Text = "bound workbook test"
    mTxtRecipeRowBudget.Text = CStr(PRODUCTION_DEFAULT_ROW_BUDGET)
    WriteRecipeHeaderFromForm

    Set lo = ProductionTable(TABLE_BUILDER_HEADER)
    If lo Is Nothing Then Exit Function
    TestWriteRecipeHeaderToBoundWorkbook = FirstRowValue(lo, "RECIPE_ID") & "|" & _
                                           FirstRowValue(lo, "RECIPE_NAME")
End Function

Public Function TestAssignmentIngredientProcessForRecipe(ByVal wb As Workbook, ByVal recipeId As String, ByVal ingredientName As String) As String
    Dim lo As ListObject
    Dim i As Long

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb

    Set lo = ProductionTable(TABLE_ASSIGN_RECIPE)
    If lo Is Nothing Then Exit Function
    EnsureTableRow lo
    SetFirstRowValue lo, "RECIPE_ID", recipeId
    RefreshAssignmentState

    For i = 0 To mLstAssignIngredients.ListCount - 1
        If StrComp(NzStr(mLstAssignIngredients.List(i, 1)), ingredientName, vbTextCompare) = 0 Then
            TestAssignmentIngredientProcessForRecipe = NzStr(mLstAssignIngredients.List(i, 3))
            Exit Function
        End If
    Next i
End Function

Public Function TestAssignmentOutputSelectionClearsStaging(ByVal wb As Workbook, ByVal recipeId As String, ByVal outputName As String) As Long
    Dim loRecipe As ListObject
    Dim loIng As ListObject
    Dim loItems As ListObject
    Dim lr As ListRow
    Dim i As Long

    If Not mBuilt Then BuildLayout
    SetOperatorWorkbook wb

    Set loRecipe = ProductionTable(TABLE_ASSIGN_RECIPE)
    Set loIng = ProductionTable(TABLE_ASSIGN_INGREDIENT)
    Set loItems = ProductionTable(TABLE_ASSIGN_ITEM)
    If loRecipe Is Nothing Or loIng Is Nothing Or loItems Is Nothing Then Exit Function

    EnsureTableRow loRecipe
    SetFirstRowValue loRecipe, "RECIPE_ID", recipeId
    SetFirstRowValue loRecipe, "RECIPE_NAME", "Recipe"

    EnsureTableRow loIng
    SetFirstRowValue loIng, "RECIPE_ID", recipeId
    SetFirstRowValue loIng, "INGREDIENT_ID", "OLD-ING"
    SetFirstRowValue loIng, "INGREDIENT", "Old ingredient"

    Set lr = loItems.ListRows.Add(AlwaysInsert:=False)
    SetRowValue lr, loItems, "RECIPE_ID", recipeId
    SetRowValue lr, loItems, "INGREDIENT_ID", "OLD-ING"
    SetRowValue lr, loItems, "ITEMS", "Old acceptable"

    RefreshAssignmentState
    For i = 0 To mLstAssignIngredients.ListCount - 1
        If StrComp(NzStr(mLstAssignIngredients.List(i, 1)), outputName, vbTextCompare) = 0 Then
            mLstAssignIngredients.ListIndex = i
            SelectAssignmentIngredientFromList
            Exit For
        End If
    Next i
    If i >= mLstAssignIngredients.ListCount Then Exit Function

    If FirstRowValue(loIng, "INGREDIENT_ID") <> "" Then Exit Function
    If Not loItems.DataBodyRange Is Nothing Then
        For i = 1 To loItems.DataBodyRange.Rows.Count
            If CellByHeader(loItems, i, "ITEMS") <> "" Then Exit Function
            If CellByHeader(loItems, i, "ITEM") <> "" Then Exit Function
            If CellByHeader(loItems, i, "INGREDIENT_ID") <> "" Then Exit Function
        Next i
    End If
    If InStr(1, TestStatusText(), "Outputs do not accept ingredients", vbTextCompare) > 0 Then
        TestAssignmentOutputSelectionClearsStaging = 1
    End If
End Function

Private Sub BuildLayout()
    If mBuilt Then Exit Sub

    Me.Caption = "Production"
    Me.Width = PRODUCTION_DEFAULT_WIDTH
    Me.Height = PRODUCTION_DEFAULT_HEIGHT
    Me.ScrollBars = fmScrollBarsBoth
    Me.KeepScrollBarsVisible = fmScrollBarsNone
    Me.ScrollWidth = PRODUCTION_MIN_WIDTH - 20
    Me.ScrollHeight = PRODUCTION_MIN_HEIGHT - 35

    Set mPages = Me.Controls.Add("Forms.MultiPage.1", "mpProduction", True)
    With mPages
        .Left = 12
        .Top = 10
        .Width = 1070
        .Height = PRODUCTION_MIN_HEIGHT - 115
    End With
    Do While mPages.Pages.Count < 6
        mPages.Pages.Add
    Loop
    mPages.Pages(0).Caption = "Process Designer"
    mPages.Pages(1).Caption = "Recipe Designer"
    mPages.Pages(2).Caption = "Ingredients Assignment"
    mPages.Pages(3).Caption = "Production Run - List"
    mPages.Pages(4).Caption = "Production Run - Tree"
    mPages.Pages(5).Caption = "Production Settings"

    BuildProcessDesignerPage mPages.Pages(0)
    BuildRecipeDesignerPage mPages.Pages(1)
    BuildAssignmentPage mPages.Pages(2)
    BuildLoaderPage mPages.Pages(3)
    BuildRunTreePage mPages.Pages(4)
    BuildOutputRegulationSettingsPage mPages.Pages(5)

    Set mTxtStatus = Me.Controls.Add("Forms.TextBox.1", "txtProductionStatus", True)
    With mTxtStatus
        .Left = 12
        .Top = PRODUCTION_MIN_HEIGHT - 94
        .Width = 900
        .Height = 42
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Locked = True
        .Text = ""
    End With
    Set mBtnClose = AddButton(Me, "btnProductionClose", "Close", 930, _
        PRODUCTION_MIN_HEIGHT - 94, 150, 42)

    ConfigureProductionAnchors
    mBuilt = True
    ResizeProductionLayout
End Sub

Private Sub BuildProcessDesignerPage(ByVal pg As MSForms.Page)
    AddLabel pg, "Saved Processes", 12, 8, 130, 16
    AddColumnHeaders pg, "SavedProcesses", Array("ID", "Ver", "Process", "", "Status", ""), _
        12, 26, "70 pt;35 pt;95 pt;0 pt;55 pt;0 pt"
    Set mLstProcesses = AddList(pg, "lstProcesses", 12, 44, 260, 92, 6, _
        "70 pt;35 pt;95 pt;0 pt;55 pt;0 pt")
    Set mBtnProcessRefresh = AddButton(pg, "btnProcessRefresh", "Refresh", 12, 144, 60, 22)
    Set mBtnProcessNew = AddButton(pg, "btnProcessNew", "New Process", 76, 144, 82, 22)
    Set mBtnProcessLoad = AddButton(pg, "btnProcessLoad", "View Process", 162, 144, 76, 22)
    Set mBtnProcessReuse = AddButton(pg, "btnProcessReuse", "Edit as New Version", 242, 144, 116, 22)

    AddLabel pg, "Process Name", 290, 8, 85, 16
    Set mTxtProcessName = AddText(pg, "txtProcessName", 290, 26, 205, 22)
    AddLabel pg, "Process ID", 505, 8, 70, 16
    Set mTxtProcessId = AddText(pg, "txtProcessId", 505, 26, 145, 22)
    mTxtProcessId.Locked = True
    AddLabel pg, "Version", 660, 8, 55, 16
    Set mTxtProcessVersion = AddText(pg, "txtProcessVersion", 660, 26, 55, 22)
    mTxtProcessVersion.Locked = True
    AddLabel pg, "Description", 290, 54, 80, 16
    Set mTxtProcessDescription = AddText(pg, "txtProcessDescription", 290, 72, 425, 64)
    mTxtProcessDescription.MultiLine = True

    Set mBtnProcessValidate = AddButton(pg, "btnProcessValidate", "Validate", 740, 12, 86, 24)
    Set mBtnProcessSave = AddButton(pg, "btnProcessSave", "Save Draft", 834, 12, 86, 24)
    Set mBtnProcessRelease = AddButton(pg, "btnProcessRelease", "Release", 928, 12, 86, 24)
    Set mBtnProcessObsolete = AddButton(pg, "btnProcessObsolete", "Obsolete", 740, 44, 86, 24)
    Set mBtnProcessClear = AddButton(pg, "btnProcessClear", "Clear", 834, 44, 86, 24)
    Set mBtnProcessWorksheetCreate = AddButton(pg, "btnProcessWorksheetCreate", _
        "Send Process to Sheet", 922, 44, 108, 24)
    Set mBtnProcessWorksheetRetrieve = AddButton(pg, "btnProcessWorksheetRetrieve", _
        "Retrieve Selected Process", 860, 140, 170, 24)
    Set mBtnProcessWorksheetAddAlternative = AddButton(pg, _
        "btnProcessWorksheetAddAlternative", "Add Acceptable Item", 680, 140, 170, 24)

    AddLabel pg, "Requirements", 12, 174, 100, 16
    AddColumnHeaders pg, "ProcessRequirements", _
        Array("ID", "Requirement", "Qty", "%", "Batch basis", "UOM", "", ""), _
        12, 192, "50 pt;78 pt;38 pt;38 pt;52 pt;42 pt;0 pt;0 pt"
    Set mLstProcessRequirements = AddList(pg, "lstProcessRequirements", 12, 208, 315, 112, 8, _
        "50 pt;78 pt;38 pt;38 pt;52 pt;42 pt;0 pt;0 pt")
    AddLabel pg, "ID / Name / Qty / % / Batch basis qty / UOM", 12, 324, 300, 16
    Set mTxtRequirementId = AddText(pg, "txtRequirementId", 12, 342, 50, 22)
    mTxtRequirementId.Locked = True
    Set mTxtRequirementName = AddText(pg, "txtRequirementName", 66, 342, 91, 22)
    Set mTxtRequirementQty = AddText(pg, "txtRequirementQty", 161, 342, 38, 22)
    Set mTxtRequirementPercent = AddText(pg, "txtRequirementPercent", 203, 342, 38, 22)
    Set mTxtRequirementYieldBasis = AddText(pg, "txtRequirementYieldBasis", 245, 342, 42, 22)
    Set mTxtRequirementUom = AddText(pg, "txtRequirementUom", 291, 342, 36, 22)
    Set mBtnProcessRequirementAdd = AddButton(pg, "btnProcessRequirementAdd", "Add", 12, 372, 52, 22)
    Set mBtnProcessRequirementUpdate = AddButton(pg, "btnProcessRequirementUpdate", "Update", 68, 372, 58, 22)
    Set mBtnProcessRequirementRemove = AddButton(pg, "btnProcessRequirementRemove", "Remove", 130, 372, 58, 22)
    Set mBtnProcessRequirementUp = AddButton(pg, "btnProcessRequirementUp", "Up", 192, 372, 42, 22)
    Set mBtnProcessRequirementDown = AddButton(pg, "btnProcessRequirementDown", "Down", 238, 372, 52, 22)
    Set mCmbRequirementQtyMode = AddCombo(pg, "cmbRequirementQtyMode", 12, 400, 220, 22)
    mCmbRequirementQtyMode.AddItem "Enter a number"
    mCmbRequirementQtyMode.AddItem "Variable -- determined at Check In"
    mCmbRequirementQtyMode.ListIndex = 0

    AddLabel pg, "Outputs (at least one)", 340, 174, 150, 16
    AddColumnHeaders pg, "ProcessOutputs", _
        Array("ID", "Output", "", "Design", "Ver", "Qty", "Yield %", "Yield basis", "UOM"), _
        340, 192, "48 pt;68 pt;0 pt;70 pt;35 pt;38 pt;38 pt;50 pt;38 pt"
    Set mLstProcessOutputs = AddList(pg, "lstProcessOutputs", 340, 208, 390, 112, 12, _
        "48 pt;68 pt;0 pt;70 pt;35 pt;38 pt;38 pt;50 pt;38 pt;0 pt;0 pt;0 pt")
    AddLabel pg, "ID / Name / Design / Ver / Qty / % / Yield basis qty / UOM", 340, 324, 390, 16
    Set mTxtProcessOutputId = AddText(pg, "txtProcessOutputId", 340, 342, 48, 22)
    mTxtProcessOutputId.Locked = True
    Set mTxtProcessOutputName = AddText(pg, "txtProcessOutputName", 392, 342, 67, 22)
    Set mTxtProcessOutputItemCode = AddText(pg, "txtProcessOutputItemCode", 463, 342, 1, 1)
    mTxtProcessOutputItemCode.Visible = False
    Set mTxtProcessOutputDesignId = AddText(pg, "txtProcessOutputDesignId", 463, 342, 48, 22)
    mTxtProcessOutputDesignId.Locked = True
    Set mTxtProcessOutputDesignVersion = AddText(pg, "txtProcessOutputDesignVersion", 515, 342, 31, 22)
    mTxtProcessOutputDesignVersion.Locked = True
    Set mTxtProcessOutputQty = AddText(pg, "txtProcessOutputQty", 550, 342, 36, 22)
    Set mTxtProcessOutputPercent = AddText(pg, "txtProcessOutputPercent", 590, 342, 36, 22)
    Set mTxtProcessOutputYieldBasis = AddText(pg, "txtProcessOutputYieldBasis", 630, 342, 40, 22)
    Set mCmbProcessOutputUom = AddCombo(pg, "cmbProcessOutputUom", 674, 342, 56, 22)
    RefreshProcessOutputUomCatalog
    Set mBtnProcessOutputAdd = AddButton(pg, "btnProcessOutputAdd", "Add", 340, 372, 52, 22)
    Set mBtnProcessOutputUpdate = AddButton(pg, "btnProcessOutputUpdate", "Update", 396, 372, 58, 22)
    Set mBtnProcessOutputRemove = AddButton(pg, "btnProcessOutputRemove", "Remove", 458, 372, 58, 22)
    Set mBtnProcessOutputUp = AddButton(pg, "btnProcessOutputUp", "Up", 520, 372, 42, 22)
    Set mBtnProcessOutputDown = AddButton(pg, "btnProcessOutputDown", "Down", 566, 372, 52, 22)
    Set mCmbProcessOutputQtyMode = AddCombo(pg, "cmbProcessOutputQtyMode", 340, 400, 250, 22)
    mCmbProcessOutputQtyMode.AddItem "Enter a number"
    mCmbProcessOutputQtyMode.AddItem "Variable -- determined by Actual Output"
    mCmbProcessOutputQtyMode.ListIndex = 0

    AddLabel pg, "Instructions", 744, 174, 100, 16
    AddColumnHeaders pg, "ProcessInstructions", Array("Step", "Instruction"), _
        744, 192, "35 pt;235 pt"
    Set mLstProcessInstructions = AddList(pg, "lstProcessInstructions", 744, 208, 286, 112, 2, "35 pt;235 pt")
    Set mTxtProcessInstruction = AddText(pg, "txtProcessInstruction", 744, 342, 286, 50)
    mTxtProcessInstruction.MultiLine = True
    Set mBtnProcessInstructionAdd = AddButton(pg, "btnProcessInstructionAdd", "Add", 744, 400, 52, 22)
    Set mBtnProcessInstructionUpdate = AddButton(pg, "btnProcessInstructionUpdate", "Update", 800, 400, 58, 22)
    Set mBtnProcessInstructionRemove = AddButton(pg, "btnProcessInstructionRemove", "Remove", 862, 400, 58, 22)
    Set mBtnProcessInstructionUp = AddButton(pg, "btnProcessInstructionUp", "Up", 924, 400, 42, 22)
    Set mBtnProcessInstructionDown = AddButton(pg, "btnProcessInstructionDown", "Down", 970, 400, 52, 22)
    Set mProcessAlternatives = New Collection
    Set mProcessOutputRegulations = CreateObject("Scripting.Dictionary")
    mProcessOutputRegulations.CompareMode = vbTextCompare
End Sub

Private Sub BuildOutputRegulationSettingsPage(ByVal pg As MSForms.Page)
    AddLabel pg, "Output Regulation Settings", 12, 10, 230, 18
    AddLabel pg, "Saved Process defaults and Recipe-output overrides are versioned. Actual Output is authoritative; planned yield is comparison only.", 12, 32, 900, 18
    AddLabel pg, "Scope", 12, 64, 45, 16
    Set mCmbOutputRegulationScope = AddCombo(pg, "cmbOutputRegulationScope", 62, 60, 150, 22)
    mCmbOutputRegulationScope.AddItem "Process default"
    mCmbOutputRegulationScope.AddItem "Recipe override"
    mCmbOutputRegulationScope.ListIndex = 0
    AddLabel pg, "Recipe Process", 230, 64, 82, 16
    Set mCmbOutputRegulationNode = AddCombo(pg, "cmbOutputRegulationNode", 315, 60, 300, 22)
    With mCmbOutputRegulationNode
        .ColumnCount = 4
        .ColumnWidths = "0 pt;0 pt;0 pt;297 pt"
        .Style = fmStyleDropDownList
    End With
    AddColumnHeaders pg, "OutputRegulation", _
        Array("Output ID", "Output", "UOM", "Regulated", "Floor", "Ceiling", "", "", ""), _
        12, 94, "70 pt;230 pt;55 pt;65 pt;65 pt;65 pt;0 pt;0 pt;0 pt"
    Set mLstOutputRegulations = AddList(pg, "lstOutputRegulations", 12, 112, 720, 230, 9, _
        "70 pt;230 pt;55 pt;65 pt;65 pt;65 pt;0 pt;0 pt;0 pt")
    Set mChkOutputRegulated = pg.Controls.Add("Forms.CheckBox.1", "chkOutputRegulated", True)
    With mChkOutputRegulated
        .Caption = "Regulated actual output"
        .Left = 12
        .Top = 362
        .Width = 165
        .Height = 20
    End With
    AddLabel pg, "Floor", 194, 364, 36, 16
    Set mTxtOutputRegulationFloor = AddText(pg, "txtOutputRegulationFloor", 232, 360, 85, 22)
    AddLabel pg, "Ceiling", 335, 364, 48, 16
    Set mTxtOutputRegulationCeiling = AddText(pg, "txtOutputRegulationCeiling", 385, 360, 85, 22)
    Set mBtnOutputRegulationApply = AddButton(pg, "btnOutputRegulationApply", "Apply Regulation", 490, 360, 105, 22)
    Set mBtnOutputRegulationClear = AddButton(pg, "btnOutputRegulationClear", "Clear Override", 605, 360, 100, 22)
    AddLabel pg, "Enabled bounds scale with batch. Completion must meet the larger of its floor and routed commitment; EA values are whole units.", 12, 398, 900, 18
    Set mTxtOutputRegulationHelp = AddText(pg, "txtOutputRegulationHelp", 12, 420, 900, 84)
    With mTxtOutputRegulationHelp
        .Locked = True
        .MultiLine = True
        .ScrollBars = fmScrollBarsVertical
        .Text = "How to use:" & vbCrLf & _
            "1. Process default: on Process Designer, load or create a DRAFT Process and add its outputs. Here choose Process default, select the output, set Regulated/Floor/Ceiling, then Save Draft and Release that Process version." & vbCrLf & _
            "2. Recipe override: on Recipe Designer, load or create a DRAFT Recipe and add its Process nodes. Here choose Recipe override, select the Recipe Process and output, set the bounds, then Save Draft and Release that Recipe version." & vbCrLf & _
            "3. Production Run only reads the released Recipe version; loading or selecting a Process there cannot change its regulation or an active run. Actual Output is the measured quantity and must meet every routed commitment."
    End With
    AddLabel pg, "UOM Catalog", 12, 516, 100, 16
    Set mBtnUomCatalogSend = AddButton(pg, "btnUomCatalogSend", "Edit UOM Catalog on Sheet", 110, 512, 155, 22)
    Set mBtnUomCatalogRetrieve = AddButton(pg, "btnUomCatalogRetrieve", "Retrieve UOM Catalog", 275, 512, 135, 22)
    AddLabel pg, "Export to the captured workbook, edit the table, select it, then retrieve. EA and CS are not convertible.", 420, 516, 600, 16
End Sub

Private Function OutputRegulationRecipeScope() As Boolean
    OutputRegulationRecipeScope = (StrComp(ComboText(mCmbOutputRegulationScope), _
        "Recipe override", vbTextCompare) = 0)
End Function

Private Sub RefreshOutputRegulationSettings()
    Dim i As Long, rowIndex As Long, records As Collection, record As Object
    Dim parseReport As String, nodeId As String, processId As String, processVersion As String
    Dim selectedNodeId As String
    Dim overrideRecord As Object
    If mLstOutputRegulations Is Nothing Then Exit Sub
    selectedNodeId = ComboText(mCmbOutputRegulationNode)
    mLoading = True
    mLstOutputRegulations.Clear
    mCmbOutputRegulationNode.Clear
    If OutputRegulationRecipeScope() Then
        For i = 0 To mLstRecipeNodes.ListCount - 1
            mCmbOutputRegulationNode.AddItem NzStr(mLstRecipeNodes.List(i, 0))
            rowIndex = mCmbOutputRegulationNode.ListCount - 1
            mCmbOutputRegulationNode.List(rowIndex, 1) = NzStr(mLstRecipeNodes.List(i, 1))
            mCmbOutputRegulationNode.List(rowIndex, 2) = NzStr(mLstRecipeNodes.List(i, 2))
            mCmbOutputRegulationNode.List(rowIndex, 3) = NzStr(mLstRecipeNodes.List(i, 3))
        Next i
        For i = 0 To mCmbOutputRegulationNode.ListCount - 1
            If StrComp(NzStr(mCmbOutputRegulationNode.List(i, 0)), selectedNodeId, vbTextCompare) = 0 Then
                mCmbOutputRegulationNode.ListIndex = i
                Exit For
            End If
        Next i
        If mCmbOutputRegulationNode.ListIndex < 0 And mCmbOutputRegulationNode.ListCount > 0 Then _
            mCmbOutputRegulationNode.ListIndex = 0
        nodeId = ComboText(mCmbOutputRegulationNode)
        If nodeId <> "" Then
            processId = NzStr(mCmbOutputRegulationNode.List(mCmbOutputRegulationNode.ListIndex, 1))
            processVersion = NzStr(mCmbOutputRegulationNode.List(mCmbOutputRegulationNode.ListIndex, 2))
            Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords( _
                modOperationsPrimitiveBridge.GetProcessVersion(processId, processVersion), parseReport)
            If Not records Is Nothing Then
                For Each record In records
                    If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"), "OUTPUT", vbTextCompare) = 0 Then
                        Set overrideRecord = FindRecipeOutputRegulation(nodeId, modProductionReusableDesigns.ReusableRecordText(record, "OutputId"))
                        AddOutputRegulationListRow modProductionReusableDesigns.ReusableRecordText(record, "OutputId"), _
                            modProductionReusableDesigns.ReusableRecordText(record, "OutputName"), _
                            modProductionReusableDesigns.ReusableRecordText(record, "UOM"), overrideRecord, _
                            nodeId, processId, processVersion
                    End If
                Next record
            End If
        End If
    Else
        For i = 0 To mLstProcessOutputs.ListCount - 1
            AddOutputRegulationListRow NzStr(mLstProcessOutputs.List(i, 0)), _
                NzStr(mLstProcessOutputs.List(i, 1)), NzStr(mLstProcessOutputs.List(i, 8)), _
                Nothing, "", Trim$(mTxtProcessId.Text), Trim$(mTxtProcessVersion.Text), i
        Next i
    End If
    mChkOutputRegulated.Value = False
    mTxtOutputRegulationFloor.Text = ""
    mTxtOutputRegulationCeiling.Text = ""
    mLoading = False
End Sub

Private Sub AddOutputRegulationListRow(ByVal outputId As String, ByVal outputName As String, _
                                       ByVal uom As String, ByVal regulation As Object, _
                                       ByVal nodeId As String, ByVal processId As String, _
                                       ByVal processVersion As String, Optional ByVal sourceRow As Long = -1)
    Dim rowIndex As Long
    mLstOutputRegulations.AddItem outputId
    rowIndex = mLstOutputRegulations.ListCount - 1
    With mLstOutputRegulations
        .List(rowIndex, 1) = outputName
        .List(rowIndex, 2) = uom
        .List(rowIndex, 6) = nodeId
        .List(rowIndex, 7) = processId
        .List(rowIndex, 8) = processVersion
        If sourceRow >= 0 Then
            .List(rowIndex, 3) = ProcessOutputRegulationValue(NzStr(mLstProcessOutputs.List(sourceRow, 0)), 0)
            .List(rowIndex, 4) = ProcessOutputRegulationValue(NzStr(mLstProcessOutputs.List(sourceRow, 0)), 1)
            .List(rowIndex, 5) = ProcessOutputRegulationValue(NzStr(mLstProcessOutputs.List(sourceRow, 0)), 2)
            .List(rowIndex, 6) = CStr(sourceRow)
        ElseIf Not regulation Is Nothing Then
            .List(rowIndex, 3) = CStr(modProductionReusableDesigns.ReusableRecordValue(regulation, "OutputRegulationEnabled"))
            .List(rowIndex, 4) = NzStr(modProductionReusableDesigns.ReusableRecordValue(regulation, "OutputFloorQty"))
            .List(rowIndex, 5) = NzStr(modProductionReusableDesigns.ReusableRecordValue(regulation, "OutputCeilingQty"))
        End If
    End With
End Sub

Private Function FindRecipeOutputRegulation(ByVal nodeId As String, ByVal outputId As String) As Object
    Dim rawRecord As Variant, record As Object
    If mRecipeOutputRegulations Is Nothing Then Exit Function
    For Each rawRecord In mRecipeOutputRegulations
        Set record = rawRecord
        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(record, "OutputId"), outputId, vbTextCompare) = 0 Then
            Set FindRecipeOutputRegulation = record
            Exit Function
        End If
    Next rawRecord
End Function

Private Function ProcessOutputRegulationValue(ByVal outputId As String, ByVal valueIndex As Long) As String
    Dim values As Variant
    If mProcessOutputRegulations Is Nothing Then Exit Function
    If mProcessOutputRegulations.Exists(outputId) Then
        values = mProcessOutputRegulations(outputId)
        ProcessOutputRegulationValue = NzStr(values(valueIndex))
    End If
End Function

Private Sub SetProcessOutputRegulation(ByVal outputId As String, ByVal enabled As Boolean, _
                                       ByVal floorText As String, ByVal ceilingText As String)
    If mProcessOutputRegulations Is Nothing Then
        Set mProcessOutputRegulations = CreateObject("Scripting.Dictionary")
        mProcessOutputRegulations.CompareMode = vbTextCompare
    End If
    mProcessOutputRegulations(outputId) = Array(CStr(enabled), Trim$(floorText), Trim$(ceilingText))
End Sub

Private Function ReusableRecordBoolean(ByVal record As Object, ByVal keyName As String) As Boolean
    Dim valueIn As Variant
    valueIn = modProductionReusableDesigns.ReusableRecordValue(record, keyName)
    If VarType(valueIn) = vbBoolean Then
        ReusableRecordBoolean = CBool(valueIn)
    Else
        ReusableRecordBoolean = (StrComp(NzStr(valueIn), "True", vbTextCompare) = 0 _
            Or NzStr(valueIn) = "1")
    End If
End Function

Private Sub BuildRecipeDesignerPage(ByVal pg As MSForms.Page)
    AddLabel pg, "Saved Recipes", 12, 8, 120, 16
    Set mLstRecipes = AddList(pg, "lstRecipes", 12, 26, 260, 105, 6, "70 pt;35 pt;95 pt;0 pt;55 pt;0 pt")
    Set mBtnRecipeRefresh = AddButton(pg, "btnRecipeRefresh", "Refresh", 12, 136, 60, 22)
    Set mBtnRecipeNew = AddButton(pg, "btnRecipeNew", "New Recipe", 76, 136, 78, 22)
    Set mBtnRecipeLoad = AddButton(pg, "btnRecipeLoad", "Load", 158, 136, 50, 22)
    AddLabel pg, "Recipe Name", 290, 8, 80, 16
    Set mTxtReusableRecipeName = AddText(pg, "txtReusableRecipeName", 290, 26, 205, 22)
    AddLabel pg, "Recipe ID", 505, 8, 65, 16
    Set mTxtReusableRecipeId = AddText(pg, "txtReusableRecipeId", 505, 26, 145, 22)
    mTxtReusableRecipeId.Locked = True
    AddLabel pg, "Version", 660, 8, 55, 16
    Set mTxtReusableRecipeVersion = AddText(pg, "txtReusableRecipeVersion", 660, 26, 55, 22)
    mTxtReusableRecipeVersion.Locked = False
    AddLabel pg, "Description", 290, 54, 80, 16
    Set mTxtReusableRecipeDescription = AddText(pg, "txtReusableRecipeDescription", 290, 72, 425, 60)
    mTxtReusableRecipeDescription.MultiLine = True
    Set mBtnRecipeValidate = AddButton(pg, "btnRecipeValidate", "Validate Recipe", 740, 12, 96, 24)
    Set mBtnRecipeSave = AddButton(pg, "btnRecipeSave", "Save Draft", 844, 12, 82, 24)
    Set mBtnRecipeRelease = AddButton(pg, "btnRecipeRelease", "Release", 934, 12, 82, 24)
    Set mBtnRecipeObsolete = AddButton(pg, "btnRecipeObsolete", "Obsolete", 740, 44, 82, 24)
    Set mBtnRecipeClear = AddButton(pg, "btnRecipeClear", "Clear", 830, 44, 82, 24)
    AddLabel pg, "Validation", 740, 76, 90, 16
    Set mLstRecipeValidation = AddList(pg, "lstRecipeValidation", 740, 94, 290, 64, 2, "58 pt;224 pt")

    AddLabel pg, "Released Processes", 12, 174, 130, 16
    AddColumnHeaders pg, "ReleasedProcesses", Array("", "Version", "Process", "", "Status", ""), _
        12, 192, "0 pt;55 pt;335 pt;0 pt;90 pt;0 pt"
    Set mLstReleasedProcesses = AddList(pg, "lstReleasedProcesses", 12, 208, 500, 80, 6, _
        "0 pt;55 pt;335 pt;0 pt;90 pt;0 pt")
    Set mBtnRecipeAddProcess = AddButton(pg, "btnRecipeAddProcess", "Add Process", 12, 294, 82, 22)

    AddLabel pg, "Recipe Process Nodes", 525, 174, 150, 16
    AddColumnHeaders pg, "RecipeNodes", Array("Step", "", "Version", "Process", "Order"), _
        525, 192, "55 pt;0 pt;55 pt;320 pt;55 pt"
    Set mLstRecipeNodes = AddList(pg, "lstRecipeNodes", 525, 208, 505, 80, 5, _
        "55 pt;0 pt;55 pt;320 pt;55 pt")
    Set mBtnRecipeRemoveProcess = AddButton(pg, "btnRecipeRemoveProcess", "Remove Process", 525, 294, 92, 22)
    Set mBtnRecipeMoveUp = AddButton(pg, "btnRecipeMoveUp", "Move Up", 621, 294, 65, 22)
    Set mBtnRecipeMoveDown = AddButton(pg, "btnRecipeMoveDown", "Move Down", 690, 294, 72, 22)
    Set mBtnRecipeAutoOrder = AddButton(pg, "btnRecipeAutoOrder", "Auto Order", 766, 294, 72, 22)

    AddLabel pg, "Output Flow", 12, 321, 90, 16
    AddColumnHeaders pg, "RecipeConnections", Array("Stage", "Produced by", "Output", "Feeds Process", "Output Qty", "Yield %", "UOM", "", "", ""), _
        12, 339, "70 pt;245 pt;220 pt;245 pt;80 pt;70 pt;85 pt;0 pt;0 pt;0 pt"
    Set mLstRecipeConnections = AddList(pg, "lstRecipeConnections", 12, 355, 1018, 72, 7, _
        "200 pt;180 pt;200 pt;200 pt;80 pt;70 pt;85 pt")
    mLstRecipeConnections.Visible = False
    Set mLstRecipeConnectionDisplay = AddList(pg, "lstRecipeConnectionDisplay", 12, 355, 1018, 72, 10, _
        "70 pt;245 pt;220 pt;245 pt;80 pt;70 pt;85 pt;0 pt;0 pt;0 pt")

    AddLabel pg, "Produced by / Output / Feeds Process / Required Qty / Required % / UOM", 12, 431, 650, 16
    mLoading = True
    Set mCmbConnectionFromNode = AddCombo(pg, "cmbConnectionFromNode", 12, 449, 220, 22)
    ConfigureNamedIdCombo mCmbConnectionFromNode, 217
    Set mCmbConnectionOutput = AddCombo(pg, "cmbConnectionOutput", 238, 449, 220, 22)
    ConfigureNamedIdCombo mCmbConnectionOutput, 217
    Set mCmbConnectionToNode = AddCombo(pg, "cmbConnectionToNode", 464, 449, 260, 22)
    ConfigureCompatibleTargetCombo mCmbConnectionToNode, 257
    Set mCmbConnectionRequirement = AddCombo(pg, "cmbConnectionRequirement", 464, 449, 1, 1)
    mCmbConnectionRequirement.Visible = False
    mLoading = False
    Set mTxtConnectionQty = AddText(pg, "txtConnectionQty", 730, 449, 80, 22)
    Set mTxtConnectionPercent = AddText(pg, "txtConnectionPercent", 816, 449, 70, 22)
    Set mCmbConnectionUom = AddCombo(pg, "cmbConnectionUom", 892, 449, 96, 22)
    RefreshConnectionUomCatalog
    Set mBtnRecipeConnect = AddButton(pg, "btnRecipeConnect", "Connect", 12, 479, 62, 22)
    Set mBtnRecipeUpdateConnection = AddButton(pg, "btnRecipeUpdateConnection", "Update", 78, 479, 62, 22)
    Set mBtnRecipeDisconnect = AddButton(pg, "btnRecipeDisconnect", "Disconnect", 144, 479, 70, 22)
    Set mLblRecipeFinishedOutput = pg.Controls.Add("Forms.Label.1", "lblRecipeFinishedOutputGuidance", True)
    With mLblRecipeFinishedOutput
        .Caption = "Final output: leave unconnected; Production creates it as finished inventory."
        .Left = 235
        .Top = 480
        .Width = 610
        .Height = 20
        .Font.Bold = True
    End With
End Sub

Private Sub BuildRecipeBuilderPage(ByVal pg As MSForms.Page)
    AddLabel pg, "Saved Recipes", 12, 12, 180, 16
    Set mLstBuilderRecipes = AddList(pg, "lstBuilderRecipes", 12, 32, 320, 230, 3, "0 pt;130 pt;170 pt")
    AddLabel pg, "Recipe Name", 350, 12, 100, 16
    Set mTxtRecipeName = AddText(pg, "txtRecipeName", 350, 32, 240, 22)
    AddLabel pg, "Recipe ID", 610, 12, 80, 16
    Set mTxtRecipeId = AddText(pg, "txtRecipeId", 610, 32, 230, 22)
    AddLabel pg, "Row Budget", 610, 62, 90, 16
    Set mTxtRecipeRowBudget = AddText(pg, "txtRecipeRowBudget", 700, 58, 140, 22)
    AddLabel pg, "Description", 350, 62, 120, 16
    Set mTxtRecipeDescription = AddText(pg, "txtRecipeDescription", 350, 96, 490, 30)
    mTxtRecipeDescription.MultiLine = True

    Set mBtnBuilderRefresh = AddButton(pg, "btnBuilderRefresh", "Refresh", 860, 32, 170, 24)
    Set mBtnBuilderNew = AddButton(pg, "btnBuilderNew", "New Recipe", 860, 64, 170, 24)
    Set mBtnBuilderLoad = AddButton(pg, "btnBuilderLoad", "Load Selected", 860, 96, 170, 24)
    Set mBtnBuilderSave = AddButton(pg, "btnBuilderSave", "Save Recipe", 860, 128, 170, 24)
    Set mBtnBuilderProcess = AddButton(pg, "btnBuilderProcess", "Add Process Table", 860, 160, 170, 24)
    Set mBtnBuilderFormulas = AddButton(pg, "btnBuilderFormulas", "Save Formulas", 860, 192, 170, 24)
    Set mBtnBuilderClear = AddButton(pg, "btnBuilderClear", "Clear Builder", 860, 224, 170, 24)
    Set mBtnBuilderRelease = AddButton(pg, "btnBuilderRelease", "Release for Production", 860, 256, 170, 24)

    AddLabel pg, "Process", 350, 145, 70, 16
    Set mTxtLineProcess = AddText(pg, "txtLineProcess", 350, 165, 120, 22)
    AddLabel pg, "In/Out", 485, 145, 70, 16
    Set mCmbLineIo = AddCombo(pg, "cmbLineIo", 485, 165, 95, 22)
    mCmbLineIo.AddItem "USED"
    mCmbLineIo.AddItem "OUTPUT"
    mCmbLineIo.AddItem "INSTRUCTION"
    mCmbLineIo.ListIndex = 0
    AddLabel pg, "Ingredient / Output / Instruction", 595, 145, 220, 16
    Set mTxtLineIngredient = AddText(pg, "txtLineIngredient", 595, 165, 220, 22)
    AddLabel pg, "Percent", 350, 197, 70, 16
    Set mTxtLinePercent = AddText(pg, "txtLinePercent", 350, 217, 70, 22)
    AddLabel pg, "UOM", 435, 197, 55, 16
    Set mTxtLineUom = AddCombo(pg, "cmbLineUom", 435, 217, 90, 22)
    RefreshRecipeUomCatalog
    Set mBtnLineUomAdd = AddButton(pg, "btnLineUomAdd", "Add UOM", 535, 216, 75, 24)
    AddLabel pg, "Amount", 625, 197, 70, 16
    Set mTxtLineAmount = AddText(pg, "txtLineAmount", 625, 217, 90, 22)
    Set mBtnLineAdd = AddButton(pg, "btnLineAdd", "Add Line", 350, 252, 90, 24)
    Set mBtnLineUpdate = AddButton(pg, "btnLineUpdate", "Update Line", 450, 252, 90, 24)
    Set mBtnLineRemove = AddButton(pg, "btnLineRemove", "Remove Line", 550, 252, 90, 24)
    Set mBtnLineMoveUp = AddButton(pg, "btnLineMoveUp", "Move Up", 650, 252, 90, 24)
    Set mBtnLineMoveDown = AddButton(pg, "btnLineMoveDown", "Move Down", 750, 252, 90, 24)

    AddLabel pg, "Recipe Builder Lines", 12, 290, 220, 16
    AddColumnHeaders pg, "BuilderLines", _
        Array("Process", "", "I/O", "Ingredient / Output / Instruction", "%", "UOM", "Amount", "Ingredient ID"), _
        12, 310, "90 pt;55 pt;70 pt;210 pt;55 pt;55 pt;55 pt;70 pt"
    Set mLstBuilderLines = AddList(pg, "lstBuilderLines", 12, 328, 1018, 192, 8, "90 pt;55 pt;70 pt;210 pt;55 pt;55 pt;55 pt;70 pt")
End Sub

Private Sub BuildAssignmentPage(ByVal pg As MSForms.Page)
    AddLabel pg, "Processes", 12, 12, 160, 16
    AddColumnHeaders pg, "AssignmentProcesses", Array("", "Version", "Process"), _
        12, 32, "0 pt;130 pt;150 pt"
    Set mLstAssignRecipes = AddList(pg, "lstAssignRecipes", 12, 50, 300, 162, 3, "0 pt;130 pt;150 pt")
    Set mBtnAssignRecipe = AddButton(pg, "btnAssignRecipe", "Select Process", 12, 220, 140, 24)
    Set mBtnAssignRefresh = AddButton(pg, "btnAssignRefresh", "Refresh", 172, 220, 140, 24)

    AddLabel pg, "Ingredient Requirements", 330, 12, 180, 16
    AddColumnHeaders pg, "AssignmentRequirements", _
        Array("", "Requirement", "UOM", "Process", "Type", "Qty", "%"), _
        330, 32, "0 pt;135 pt;45 pt;70 pt;55 pt;45 pt;45 pt"
    Set mLstAssignIngredients = AddList(pg, "lstAssignIngredients", 330, 50, 340, 194, 7, "0 pt;135 pt;45 pt;70 pt;55 pt;45 pt;45 pt")
    Set mBtnAssignIngredient = AddButton(pg, "btnAssignIngredient", "Select Requirement", 690, 32, 150, 24)
    Set mBtnAssignSave = AddButton(pg, "btnAssignSave", "Save Alternatives", 690, 64, 150, 24)
    Set mBtnAssignClear = AddButton(pg, "btnAssignClear", "Clear", 690, 96, 150, 24)

    AddLabel pg, "Search Inventory", 12, 262, 130, 16
    Set mTxtInventorySearch = AddText(pg, "txtInventorySearch", 130, 258, 230, 22)
    Set mBtnAssignAdd = AddButton(pg, "btnAssignAdd", "Add Acceptable", 380, 258, 150, 24)
    Set mBtnAssignRemove = AddButton(pg, "btnAssignRemove", "Remove Row", 548, 258, 122, 24)
    AddLabel pg, "Managed Items", 12, 292, 120, 16
    AddColumnHeaders pg, "AssignmentInventory", _
        Array("System_Key", "Managed Item", "UOM", "Inv", "Location", "Description", ""), _
        12, 312, "190 pt;105 pt;35 pt;45 pt;70 pt;65 pt;0 pt"
    Set mLstAssignInventory = AddList(pg, "lstAssignInventory", 12, 330, 510, 190, 7, "190 pt;105 pt;35 pt;45 pt;70 pt;65 pt;0 pt")
    AddLabel pg, "Acceptable Items", 540, 292, 150, 16
    AddColumnHeaders pg, "AssignmentAllowed", _
        Array("", "Managed Item", "UOM", "Item Code", "", "", ""), _
        540, 312, "0 pt;210 pt;55 pt;210 pt;0 pt;0 pt;0 pt"
    Set mLstAssignAllowed = AddList(pg, "lstAssignAllowed", 540, 330, 490, 190, 7, "0 pt;210 pt;55 pt;210 pt;0 pt;0 pt;0 pt")
End Sub

Private Sub BuildLoaderPage(ByVal pg As MSForms.Page)
    With pg
        .ScrollBars = fmScrollBarsVertical
        .KeepScrollBarsVisible = fmScrollBarsVertical
        .ScrollHeight = RUN_LIST_SCROLL_HEIGHT
    End With

    AddLabel pg, "Recipes", 12, 12, 140, 16
    AddColumnHeaders pg, "LoaderRecipes", Array("", "Recipe", "Description"), 12, 32, RUN_LOADER_RECIPE_WIDTHS
    Set mLstLoaderRecipes = AddList(pg, "lstLoaderRecipes", 12, 50, 270, 70, 3, RUN_LOADER_RECIPE_WIDTHS)
    Set mBtnLoaderRefresh = AddButton(pg, "btnLoaderRefresh", "Refresh", 300, 32, 130, 24)
    Set mBtnLoaderLoad = AddButton(pg, "btnLoaderLoad", "Load Recipe", 300, 64, 130, 24)
    Set mBtnLoaderClear = AddButton(pg, "btnLoaderClear", "Clear Run", 300, 96, 130, 24)
    AddLabel pg, "Batch scale %", 12, 132, 88, 16
    Set mTxtBatchScalePercent = AddText(pg, "txtBatchScalePercent", 104, 128, 70, 22)
    mTxtBatchScalePercent.Text = "100"
    Set mBtnApplyBatchScale = AddButton(pg, "btnApplyBatchScale", "Apply Scale", 184, 127, 100, 24)

    Set mChkRunTargetOutputScale = pg.Controls.Add("Forms.CheckBox.1", "chkRunTargetOutputScale", True)
    With mChkRunTargetOutputScale
        .Caption = "Scale from target output Qty (coming later)"
        .Left = 12
        .Top = 154
        .Width = 220
        .Height = 18
        .Enabled = False
        .ControlTipText = "Reserved for a later approved scaling solver."
    End With
    Set mCmbRunTargetOutput = AddCombo(pg, "cmbRunTargetOutput", 238, 152, 135, 22)
    mCmbRunTargetOutput.Enabled = False
    Set mTxtRunTargetOutputQty = AddText(pg, "txtRunTargetOutputQty", 379, 152, 51, 22)
    mTxtRunTargetOutputQty.Enabled = False

    AddLabel pg, "Multi-Process Run Plan", 455, 12, 180, 16
    AddColumnHeaders pg, "LoaderLines", Array("Process", "", "Type", "Requirement / Output", "%", "UOM", "Qty", "Status", ""), 455, 32, RUN_LOADER_LINE_WIDTHS
    Set mLstLoaderLines = AddList(pg, "lstLoaderLines", 455, 50, 575, 106, 9, RUN_LOADER_LINE_WIDTHS)

    AddLabel pg, "Process filter", 12, 192, 80, 16
    Set mCmbRunProcess = AddCombo(pg, "cmbRunProcess", 95, 188, 160, 22)
    AddLabel pg, "Run Location", 270, 192, 90, 16
    Set mCmbRunLocation = AddCombo(pg, "cmbRunLocation", 365, 188, 160, 22)
    AddLabel pg, "% of Requirement", 545, 192, 110, 16
    Set mTxtPaletteSplit = AddText(pg, "txtPaletteSplit", 660, 188, 90, 22)
    AddLabel pg, "Qty", 770, 192, 45, 16
    Set mTxtPaletteQty = AddText(pg, "txtPaletteQty", 810, 188, 90, 22)
    Set mBtnRunApplyPalette = AddButton(pg, "btnRunApplyPalette", "Apply", 920, 187, 90, 24)

    AddLabel pg, "Acceptable Inventory For Run", 12, 214, 230, 16
    AddColumnHeaders pg, "RunPalette", Array("", "", "Process / Ingredient", "", "Inventory Stock", "% Req", "Qty", "Stock / Requirement UOM", "Native / Requirement Available", "Location"), 12, 232, RUN_PALETTE_WIDTHS
    Set mLstRunPalette = AddList(pg, "lstRunPalette", 12, 250, 1018, 96, 10, RUN_PALETTE_WIDTHS)

    AddLabel pg, "Inventory Check", 12, 352, 150, 16
    AddColumnHeaders pg, "ManagerCheck", Array("Type", "Process / Requirement", "Source Process / Output", "System_Key", "Code", "Item", "UOM", "Committed / Used", "Remaining Balance"), 12, 370, RUN_CHECK_WIDTHS
    Set mLstManagerCheck = AddList(pg, "lstManagerCheck", 12, 386, 1018, RUN_LIST_CHECK_MIN_HEIGHT, 9, RUN_CHECK_WIDTHS)
    AddLabel pg, "Selected Process Instructions", 12, 492, 220, 16
    AddColumnHeaders pg, "RunInstructions", Array("Step", "Instruction"), _
        12, 510, "45 pt;973 pt"
    Set mLstRunInstructions = AddList(pg, "lstRunInstructions", 12, 526, 1018, RUN_LIST_INSTRUCTIONS_MIN_HEIGHT, 2, _
        "45 pt;973 pt")

    AddLabel pg, "Production Output", 12, 584, 170, 16
    AddColumnHeaders pg, "ManagerOutput", Array("Process", "Output", "UOM", "Last Actual", "Batch", "Used Goods", "Process Total", "Recall", "System_Key"), 12, 602, RUN_OUTPUT_WIDTHS
    Set mLstManagerOutput = AddList(pg, "lstManagerOutput", 12, 618, 1018, 18, 9, RUN_OUTPUT_WIDTHS)

    AddLabel pg, "Actual Output", 12, 640, 80, 16
    Set mTxtOutputReal = AddText(pg, "txtOutputReal", 100, 638, 105, 22)
    Set mBtnManagerCheckIn = AddButton(pg, "btnManagerCheckIn", "Check In", 230, 638, 95, 24)
    Set mBtnManagerApplyOutput = AddButton(pg, "btnManagerApplyOutput", "Complete Run", 340, 638, 120, 24)
    Set mBtnManagerRefresh = AddButton(pg, "btnManagerRefresh", "Refresh", 480, 638, 95, 24)
    Set mBtnManagerNext = AddButton(pg, "btnManagerNext", "Next Batch", 595, 638, 120, 24)
    Set mBtnManagerPrint = AddButton(pg, "btnManagerPrint", "Print Recall", 735, 638, 120, 24)
    AddLabel pg, "Batch Note", 866, 640, 60, 16
    Set mTxtRunBatchNote = AddText(pg, "txtRunBatchNote", 928, 638, 102, 22)
    mTxtRunBatchNote.MaxLength = 500
End Sub

Private Sub BuildRunTreePage(ByVal pg As MSForms.Page)
    AddLabel pg, "Recipe Ingredient / Inventory Choices", 12, 12, 260, 16
    AddLabel pg, "% of Requirement", 300, 10, 110, 16
    Set mTxtTreePaletteSplit = AddText(pg, "txtTreePaletteSplit", 410, 8, 70, 22)
    AddLabel pg, "Qty", 492, 10, 35, 16
    Set mTxtTreePaletteQty = AddText(pg, "txtTreePaletteQty", 530, 8, 80, 22)
    Set mBtnRunTreeApplyPalette = AddButton(pg, "btnRunTreeApplyPalette", "Apply", 622, 7, 70, 24)
    AddLabel pg, "Process", 704, 10, 55, 16
    Set mCmbTreeRunProcess = AddCombo(pg, "cmbTreeRunProcess", 760, 8, 100, 22)
    AddLabel pg, "Location", 870, 10, 58, 16
    Set mCmbTreeRunLocation = AddCombo(pg, "cmbTreeRunLocation", 930, 8, 100, 22)
    Set mBtnRunTreeExpandAll = AddButton(pg, "btnRunTreeExpandAll", "Expand", 850, 528, 70, 24)
    Set mBtnRunTreeCollapseAll = AddButton(pg, "btnRunTreeCollapseAll", "Collapse", 930, 528, 80, 24)
    AddRunPaletteHeader pg, 12, 42
    Set mLstRunTree = AddList(pg, "lstRunTree", 12, 60, 1018, 460, 10, "0 pt;0 pt;300 pt;42 pt;165 pt;58 pt;68 pt;48 pt;80 pt;110 pt")
    mLstRunTree.Font.Size = 11
End Sub

Private Sub AddRunPaletteHeader(ByVal pg As MSForms.Page, ByVal leftVal As Single, ByVal topVal As Single)
    AddLabel pg, "Ingredient", leftVal + 2, topVal, 135, 14
    AddLabel pg, "System_Key", leftVal + 130, topVal, 95, 14
    AddLabel pg, "Inventory Item", leftVal + 168, topVal, 145, 14
    AddLabel pg, "% Req", leftVal + 315, topVal, 46, 14
    AddLabel pg, "Qty", leftVal + 365, topVal, 50, 14
    AddLabel pg, "UOM", leftVal + 425, topVal, 40, 14
    AddLabel pg, "Inv", leftVal + 465, topVal, 45, 14
    AddLabel pg, "Location", leftVal + 540, topVal, 70, 14
End Sub

Private Sub AddColumnHeaders(ByVal parent As Object, ByVal groupName As String, ByVal labels As Variant, _
                             ByVal leftVal As Single, ByVal topVal As Single, ByVal widths As String)
    Dim i As Long
    Dim idx As Long
    Dim x As Single
    Dim colWidth As Single
    Dim caption As String
    Dim lbl As MSForms.Label

    x = leftVal
    For i = LBound(labels) To UBound(labels)
        idx = i - LBound(labels)
        colWidth = ColumnWidthPoints(widths, idx, 60)
        caption = Trim$(CStr(labels(i)))
        If colWidth > 0 And caption <> "" Then
            Set lbl = parent.Controls.Add("Forms.Label.1", "hdr" & CleanControlName(groupName) & CStr(idx + 1), True)
            With lbl
                .Caption = caption
                .Left = x + 2
                .Top = topVal
                .Width = MaxDoubleForm(1, colWidth - 4)
                .Height = 14
                .Font.Bold = True
            End With
        End If
        x = x + colWidth
    Next i
End Sub

Private Sub PositionColumnHeaders(ByVal parent As Object, ByVal groupName As String, _
                                  ByVal leftVal As Single, ByVal topVal As Single, ByVal widths As String)
    Dim parts As Variant
    Dim i As Long
    Dim x As Single
    Dim colWidth As Single
    Dim lbl As MSForms.Label

    parts = Split(widths, ";")
    x = leftVal
    For i = LBound(parts) To UBound(parts)
        colWidth = ColumnWidthPoints(widths, i, 60)
        Set lbl = Nothing
        On Error Resume Next
        Set lbl = parent.Controls("hdr" & CleanControlName(groupName) & CStr(i + 1))
        On Error GoTo 0
        If Not lbl Is Nothing Then
            lbl.Left = x + 2
            lbl.Top = topVal
            lbl.Width = MaxDoubleForm(1, colWidth - 4)
        End If
        x = x + colWidth
    Next i
End Sub

Private Function ColumnWidthPoints(ByVal widths As String, ByVal zeroBasedIndex As Long, _
                                   Optional ByVal defaultWidth As Single = 60) As Single
    Dim parts As Variant
    Dim rawValue As String

    parts = Split(widths, ";")
    If zeroBasedIndex < LBound(parts) Or zeroBasedIndex > UBound(parts) Then
        ColumnWidthPoints = defaultWidth
        Exit Function
    End If
    rawValue = Replace$(LCase$(Trim$(CStr(parts(zeroBasedIndex)))), "pt", "")
    ColumnWidthPoints = CSng(Val(rawValue))
End Function

Private Sub ResizeProductionLayout()
    On Error GoTo CleanExit
    If Not mBuilt Then Exit Sub
    If mResizingLayout Then Exit Sub
    mResizingLayout = True

    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout

CleanExit:
    mResizingLayout = False
End Sub

Private Sub ConfigureProductionAnchors()
    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, PRODUCTION_MIN_WIDTH, PRODUCTION_MIN_HEIGHT

    mLayout.RegisterControl mPages, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or _
                                    OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mTxtStatus, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_RIGHT Or _
                                        OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mBtnClose, OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM

    ConfigureReusableDesignerAnchors
    ConfigureAssignmentAnchors
    ConfigureRunListAnchors
    ConfigureRunTreeAnchors
End Sub

Private Sub ConfigureReusableDesignerAnchors()
    mLayout.RegisterControl mLstProcesses, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mTxtProcessDescription, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLstProcessRequirements, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mLstProcessOutputs, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLstProcessInstructions, OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mTxtProcessInstruction, OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mLstRecipes, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mTxtReusableRecipeDescription, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLstReleasedProcesses, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mLstRecipeNodes, OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mLstRecipeConnectionDisplay, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mLstRecipeValidation, OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_TOP
End Sub

Private Sub ConfigureRecipeBuilderAnchors()
    Dim rightTop As Long
    rightTop = OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_TOP

    mLayout.RegisterControl mLstBuilderRecipes, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mTxtRecipeName, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or _
                                           OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mTxtRecipeId, rightTop
    mLayout.RegisterControl mTxtRecipeRowBudget, rightTop
    mLayout.RegisterControl mTxtRecipeDescription, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or _
                                                  OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mBtnBuilderRefresh, rightTop
    mLayout.RegisterControl mBtnBuilderNew, rightTop
    mLayout.RegisterControl mBtnBuilderLoad, rightTop
    mLayout.RegisterControl mBtnBuilderSave, rightTop
    mLayout.RegisterControl mBtnBuilderProcess, rightTop
    mLayout.RegisterControl mBtnBuilderFormulas, rightTop
    mLayout.RegisterControl mBtnBuilderClear, rightTop
    mLayout.RegisterControl mBtnBuilderRelease, rightTop
    mLayout.RegisterControl mLstBuilderLines, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or _
                                             OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
End Sub

Private Sub ConfigureAssignmentAnchors()
    Dim rightTop As Long
    rightTop = OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_TOP

    mLayout.RegisterControl mLstAssignRecipes, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    mLayout.RegisterControl mLstAssignIngredients, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or _
                                                  OPERATIONS_ANCHOR_RIGHT
    mLayout.RegisterControl mBtnAssignIngredient, rightTop
    mLayout.RegisterControl mBtnAssignSave, rightTop
    mLayout.RegisterControl mBtnAssignClear, rightTop
    mLayout.RegisterControl mLstAssignInventory, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or _
                                                OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mLstAssignAllowed, OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_TOP Or _
                                              OPERATIONS_ANCHOR_BOTTOM
End Sub

Private Sub ConfigureRunListAnchors()
    Dim leftRightTop As Long

    leftRightTop = OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_TOP
    ConfigureRunListProportionalLayout leftRightTop
End Sub

Private Sub ConfigureRunListProportionalLayout(ByVal leftRightTop As Long)
    Dim leftTop As Long
    Dim rightTop As Long

    leftTop = OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    rightTop = OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_TOP

    ' The five vertical bands share added/removed page height.  This keeps the
    ' operator's recipes, plan, palette, audit check, instructions, and output
    ' readable together instead of assigning all growth to Production Output.
    mLayout.RegisterProportionalVerticalControl mLstLoaderRecipes, leftTop, 0, 0.2
    mLayout.RegisterProportionalVerticalControl mLstLoaderLines, leftRightTop, 0, 0.2
    RegisterRunListHeaderBand "LoaderLines", 0

    RegisterRunListCaptionBand "Batch scale %", leftTop, 0.2
    RegisterRunListCaptionBand "Process filter", leftTop, 0.2
    RegisterRunListCaptionBand "Run Location", leftTop, 0.2
    RegisterRunListCaptionBand "% of Requirement", leftTop, 0.2
    RegisterRunListCaptionBand "Qty", leftTop, 0.2
    RegisterRunListCaptionBand "Acceptable Inventory For Run", leftTop, 0.2
    mLayout.RegisterProportionalVerticalControl mTxtBatchScalePercent, leftTop, 0.2, 0
    mLayout.RegisterProportionalVerticalControl mBtnApplyBatchScale, leftTop, 0.2, 0
    mLayout.RegisterProportionalVerticalControl mChkRunTargetOutputScale, leftTop, 0.2, 0
    mLayout.RegisterProportionalVerticalControl mCmbRunTargetOutput, leftTop, 0.2, 0
    mLayout.RegisterProportionalVerticalControl mTxtRunTargetOutputQty, leftTop, 0.2, 0
    mLayout.RegisterProportionalVerticalControl mCmbRunProcess, leftTop, 0.2, 0
    mLayout.RegisterProportionalVerticalControl mCmbRunLocation, leftTop, 0.2, 0
    mLayout.RegisterProportionalVerticalControl mTxtPaletteSplit, leftTop, 0.2, 0
    mLayout.RegisterProportionalVerticalControl mTxtPaletteQty, leftTop, 0.2, 0
    mLayout.RegisterProportionalVerticalControl mBtnRunApplyPalette, rightTop, 0.2, 0
    RegisterRunListHeaderBand "RunPalette", 0.2
    mLayout.RegisterProportionalVerticalControl mLstRunPalette, leftRightTop, 0.2, 0.2

    RegisterRunListCaptionBand "Inventory Check", leftTop, 0.4
    RegisterRunListHeaderBand "ManagerCheck", 0.4
    mLayout.RegisterProportionalVerticalControl mLstManagerCheck, leftRightTop, 0.4, 0.2

    RegisterRunListCaptionBand "Selected Process Instructions", leftTop, 0.6
    RegisterRunListHeaderBand "RunInstructions", 0.6
    mLayout.RegisterProportionalVerticalControl mLstRunInstructions, leftRightTop, 0.6, 0.2

    RegisterRunListCaptionBand "Production Output", leftTop, 0.8
    RegisterRunListHeaderBand "ManagerOutput", 0.8
    mLayout.RegisterProportionalVerticalControl mLstManagerOutput, leftRightTop, 0.8, 0.2

    RegisterRunListCaptionBand "Actual Output", leftTop, 1
    mLayout.RegisterProportionalVerticalControl mTxtOutputReal, leftTop, 1, 0
    mLayout.RegisterProportionalVerticalControl mBtnManagerCheckIn, leftTop, 1, 0
    mLayout.RegisterProportionalVerticalControl mBtnManagerApplyOutput, leftTop, 1, 0
    mLayout.RegisterProportionalVerticalControl mBtnManagerRefresh, leftTop, 1, 0
    mLayout.RegisterProportionalVerticalControl mBtnManagerNext, leftTop, 1, 0
    mLayout.RegisterProportionalVerticalControl mBtnManagerPrint, rightTop, 1, 0
    RegisterRunListCaptionBand "Batch Note", rightTop, 1
    mLayout.RegisterProportionalVerticalControl mTxtRunBatchNote, rightTop, 1, 0
End Sub

Private Sub RegisterRunListCaptionBand(ByVal caption As String, ByVal horizontalAnchorMask As Long, _
                                       ByVal verticalTopFactor As Double)
    Dim ctl As MSForms.Control

    For Each ctl In mPages.Pages(3).Controls
        If TypeName(ctl) = "Label" And _
           StrComp(Left$(ctl.Name, 3), "hdr", vbTextCompare) <> 0 Then
            If StrComp(CStr(ctl.Caption), caption, vbTextCompare) = 0 Then
                mLayout.RegisterProportionalVerticalControl ctl, horizontalAnchorMask, _
                                                           verticalTopFactor, 0
            End If
        End If
    Next ctl
End Sub

Private Sub RegisterRunListHeaderBand(ByVal groupName As String, ByVal verticalTopFactor As Double)
    Dim ctl As MSForms.Control
    Dim prefix As String
    Dim leftTop As Long

    prefix = "hdr" & CleanControlName(groupName)
    leftTop = OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP
    For Each ctl In mPages.Pages(3).Controls
        If TypeName(ctl) = "Label" Then
            If StrComp(Left$(ctl.Name, Len(prefix)), prefix, vbTextCompare) = 0 Then
                mLayout.RegisterProportionalVerticalControl ctl, leftTop, verticalTopFactor, 0
            End If
        End If
    Next ctl
End Sub

Private Sub ConfigureRunTreeAnchors()
    Dim rightTop As Long
    rightTop = OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_TOP

    mLayout.RegisterControl mTxtTreePaletteSplit, rightTop
    mLayout.RegisterControl mTxtTreePaletteQty, rightTop
    mLayout.RegisterControl mBtnRunTreeApplyPalette, rightTop
    mLayout.RegisterControl mCmbTreeRunProcess, rightTop
    mLayout.RegisterControl mCmbTreeRunLocation, rightTop
    mLayout.RegisterControl mLstRunTree, OPERATIONS_ANCHOR_LEFT Or OPERATIONS_ANCHOR_TOP Or _
                                         OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mBtnRunTreeExpandAll, OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
    mLayout.RegisterControl mBtnRunTreeCollapseAll, OPERATIONS_ANCHOR_RIGHT Or OPERATIONS_ANCHOR_BOTTOM
End Sub

Private Function RunListHeaderBandAligned(ByVal groupName As String, _
                                           ByVal targetList As MSForms.ListBox) As Boolean
    Dim firstHeader As MSForms.Label
    Dim expectedHeaderTop As Double

    On Error GoTo NotAligned
    Set firstHeader = mPages.Pages(3).Controls("hdr" & CleanControlName(groupName) & "1")
    expectedHeaderTop = targetList.Top - 16
    RunListHeaderBandAligned = Abs(firstHeader.Top - expectedHeaderTop) < 0.1 And _
                               Abs(firstHeader.Left - (targetList.Left + 2)) < 0.1
    Exit Function

NotAligned:
    RunListHeaderBandAligned = False
End Function

Private Function RunListHeaderCellAligned(ByVal groupName As String, _
                                           ByVal headerIndex As Long, _
                                           ByVal targetList As MSForms.ListBox) As Boolean
    Dim headerCell As MSForms.Label
    Dim firstHeader As MSForms.Label

    On Error GoTo NotAligned
    Set headerCell = mPages.Pages(3).Controls("hdr" & CleanControlName(groupName) & CStr(headerIndex))
    Set firstHeader = mPages.Pages(3).Controls("hdr" & CleanControlName(groupName) & "1")
    RunListHeaderCellAligned = Abs(headerCell.Top - firstHeader.Top) < 0.1 And _
                                headerCell.Visible And firstHeader.Visible
    Exit Function

NotAligned:
    RunListHeaderCellAligned = False
End Function

Private Function CountOutOfBoundsControls(ByVal parent As Object, ByRef issueDetail As String) As Long
    Dim ctl As Object
    Dim parentWidth As Double
    Dim parentHeight As Double

    parentWidth = LayoutParentExtent(parent, True)
    parentHeight = LayoutParentExtent(parent, False)
    For Each ctl In parent.Controls
        If IsDirectLayoutChild(ctl, parent) And ctl.Visible Then
            If CDbl(ctl.Left) < -0.5 Or CDbl(ctl.Top) < -0.5 Or _
               CDbl(ctl.Left) + CDbl(ctl.Width) > parentWidth + 1 Or _
               CDbl(ctl.Top) + CDbl(ctl.Height) > parentHeight + 1 Then
                CountOutOfBoundsControls = CountOutOfBoundsControls + 1
                AppendLayoutIssue issueDetail, _
                    "BOUNDS:" & CStr(ctl.Name) & "@" & LayoutControlRectangleText(ctl) & _
                    "/P=" & Format$(parentWidth, "0") & "x" & Format$(parentHeight, "0")
            End If
        End If
    Next ctl
End Function

Private Function CountOverlappingInteractiveControls(ByVal parent As Object, _
                                                     ByRef issueDetail As String) As Long
    Dim leftControl As Object
    Dim rightControl As Object
    Dim leftIndex As Long
    Dim rightIndex As Long

    For leftIndex = 0 To parent.Controls.Count - 2
        Set leftControl = parent.Controls.Item(leftIndex)
        If IsDirectLayoutChild(leftControl, parent) And _
           IsInteractiveLayoutControl(leftControl) Then
            For rightIndex = leftIndex + 1 To parent.Controls.Count - 1
                Set rightControl = parent.Controls.Item(rightIndex)
                If IsDirectLayoutChild(rightControl, parent) And _
                   IsInteractiveLayoutControl(rightControl) Then
                    If LayoutRectanglesOverlap(leftControl, rightControl) Then
                        CountOverlappingInteractiveControls = _
                            CountOverlappingInteractiveControls + 1
                        AppendLayoutIssue issueDetail, _
                            "OVERLAP:" & CStr(leftControl.Name) & "@" & _
                            LayoutControlRectangleText(leftControl) & "+" & _
                            CStr(rightControl.Name) & "@" & _
                            LayoutControlRectangleText(rightControl)
                    End If
                End If
            Next rightIndex
        End If
    Next leftIndex
End Function

Private Function LayoutControlRectangleText(ByVal ctl As Object) As String
    LayoutControlRectangleText = _
        Format$(CDbl(ctl.Left), "0") & "," & Format$(CDbl(ctl.Top), "0") & "," & _
        Format$(CDbl(ctl.Width), "0") & "," & Format$(CDbl(ctl.Height), "0")
End Function

Private Function IsDirectLayoutChild(ByVal ctl As Object, ByVal expectedParent As Object) As Boolean
    Dim actualParent As Object

    On Error Resume Next
    Set actualParent = CallByName(ctl, "Container", VbGet)
    If actualParent Is Nothing Then
        Err.Clear
        Set actualParent = ctl.Parent
    End If
    IsDirectLayoutChild = (actualParent Is expectedParent)
    On Error GoTo 0
End Function

Private Function IsInteractiveLayoutControl(ByVal ctl As Object) As Boolean
    If Not ctl.Visible Then Exit Function
    Select Case TypeName(ctl)
        Case "CommandButton", "ComboBox", "ListBox", "MultiPage", "TextBox"
            IsInteractiveLayoutControl = True
    End Select
End Function

Private Function LayoutRectanglesOverlap(ByVal leftControl As Object, _
                                         ByVal rightControl As Object) As Boolean
    LayoutRectanglesOverlap = _
        (CDbl(leftControl.Left) < CDbl(rightControl.Left) + CDbl(rightControl.Width) - 0.5) And _
        (CDbl(rightControl.Left) < CDbl(leftControl.Left) + CDbl(leftControl.Width) - 0.5) And _
        (CDbl(leftControl.Top) < CDbl(rightControl.Top) + CDbl(rightControl.Height) - 0.5) And _
        (CDbl(rightControl.Top) < CDbl(leftControl.Top) + CDbl(leftControl.Height) - 0.5)
End Function

Private Function LayoutParentExtent(ByVal parent As Object, _
                                    ByVal horizontal As Boolean) As Double
    Dim pageHost As Object
    Dim primaryProperty As String
    Dim fallbackProperty As String

    On Error Resume Next
    If horizontal Then
        primaryProperty = "InsideWidth"
        fallbackProperty = "Width"
    Else
        primaryProperty = "InsideHeight"
        fallbackProperty = "Height"
    End If
    If TypeName(parent) = "Page" Then
        Set pageHost = parent.Parent
        LayoutParentExtent = CDbl(CallByName(pageHost, fallbackProperty, VbGet)) - 20
        On Error GoTo 0
        Exit Function
    End If
    LayoutParentExtent = CDbl(CallByName(parent, primaryProperty, VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        LayoutParentExtent = CDbl(CallByName(parent, fallbackProperty, VbGet))
    End If
    On Error GoTo 0
End Function

Private Sub AppendLayoutIssue(ByRef issueDetail As String, ByVal issueText As String)
    If Len(issueDetail) >= 500 Then Exit Sub
    If issueDetail <> "" Then issueDetail = issueDetail & ","
    issueDetail = issueDetail & issueText
End Sub

Private Sub MoveLabelByCaption(ByVal parent As Object, ByVal caption As String, ByVal leftVal As Single, _
                               ByVal topVal As Single, ByVal widthVal As Single, ByVal heightVal As Single)
    Dim ctl As MSForms.Control
    For Each ctl In parent.Controls
        If TypeName(ctl) = "Label" Then
            If StrComp(CStr(ctl.Caption), caption, vbTextCompare) = 0 Then
                ctl.Move leftVal, topVal, widthVal, heightVal
                Exit Sub
            End If
        End If
    Next ctl
End Sub

Private Function MaxDoubleForm(ByVal leftValue As Double, ByVal rightValue As Double) As Double
    If leftValue >= rightValue Then
        MaxDoubleForm = leftValue
    Else
        MaxDoubleForm = rightValue
    End If
End Function

Private Function AddList(ByVal parent As Object, ByVal name As String, ByVal leftVal As Single, ByVal topVal As Single, _
                         ByVal widthVal As Single, ByVal heightVal As Single, ByVal columns As Long, _
                         ByVal widths As String) As MSForms.ListBox
    Set AddList = parent.Controls.Add("Forms.ListBox.1", name, True)
    With AddList
        .Left = leftVal
        .Top = topVal
        .Width = widthVal
        .Height = heightVal
        .ColumnCount = columns
        .ColumnWidths = widths
        .IntegralHeight = False
    End With
End Function

Private Function AddButton(ByVal parent As Object, ByVal name As String, ByVal caption As String, ByVal leftVal As Single, _
                           ByVal topVal As Single, ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.CommandButton
    Set AddButton = parent.Controls.Add("Forms.CommandButton.1", name, True)
    With AddButton
        .Caption = caption
        .Left = leftVal
        .Top = topVal
        .Width = widthVal
        .Height = heightVal
        .TakeFocusOnClick = False
    End With
End Function

Private Function AddText(ByVal parent As Object, ByVal name As String, ByVal leftVal As Single, ByVal topVal As Single, _
                         ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.TextBox
    Set AddText = parent.Controls.Add("Forms.TextBox.1", name, True)
    With AddText
        .Left = leftVal
        .Top = topVal
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddCombo(ByVal parent As Object, ByVal name As String, ByVal leftVal As Single, ByVal topVal As Single, _
                          ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.ComboBox
    Set AddCombo = parent.Controls.Add("Forms.ComboBox.1", name, True)
    With AddCombo
        .Left = leftVal
        .Top = topVal
        .Width = widthVal
        .Height = heightVal
        .Style = fmStyleDropDownList
    End With
End Function

Private Sub ConfigureNamedIdCombo(ByVal combo As MSForms.ComboBox, ByVal displayWidth As Single)
    With combo
        .ColumnCount = 2
        .BoundColumn = 1
        .TextColumn = 2
        .ColumnWidths = "0 pt;" & CStr(displayWidth) & " pt"
    End With
End Sub

Private Sub ConfigureCompatibleTargetCombo(ByVal combo As MSForms.ComboBox, _
                                           ByVal displayWidth As Single)
    With combo
        .ColumnCount = 4
        .BoundColumn = 1
        .TextColumn = 3
        .ColumnWidths = "0 pt;0 pt;" & CStr(displayWidth) & " pt;0 pt"
    End With
End Sub

Private Sub AddLabel(ByVal parent As Object, ByVal caption As String, ByVal leftVal As Single, ByVal topVal As Single, _
                     ByVal widthVal As Single, ByVal heightVal As Single)
    Dim lbl As MSForms.Label
    Set lbl = parent.Controls.Add("Forms.Label.1", "lbl" & CleanControlName(caption) & CStr(parent.Controls.Count + 1), True)
    With lbl
        .Caption = caption
        .Left = leftVal
        .Top = topVal
        .Width = widthVal
        .Height = heightVal
        .Font.Bold = True
    End With
End Sub

Private Function CleanControlName(ByVal value As String) As String
    Dim i As Long
    Dim ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch Like "[A-Za-z0-9]" Then CleanControlName = CleanControlName & ch
    Next i
    If CleanControlName = "" Then CleanControlName = "Control"
End Function

Private Function ControlExistsByName(ByVal parent As Object, ByVal controlName As String) As Boolean
    Dim ctl As Object

    On Error Resume Next
    Set ctl = parent.Controls(controlName)
    ControlExistsByName = Not ctl Is Nothing
    On Error GoTo 0
End Function

Private Sub RefreshAllViews()
    RefreshReusableDesignLists
    RefreshReusableAssignmentState
    RefreshLoaderState
    RefreshManagerState
    RefreshOutputRegulationSettings
End Sub

Private Sub RefreshRecipeLists()
    RefreshReusableDesignLists
End Sub

Private Sub RefreshReusableDesignLists()
    Dim processes As Variant
    Dim releasedProcesses As Variant
    Dim recipes As Variant
    Dim releasedRecipes As Variant

    processes = modOperationsPrimitiveBridge.ListProcesses("")
    releasedProcesses = modOperationsPrimitiveBridge.ListProcesses("RELEASED")
    recipes = modOperationsPrimitiveBridge.ListRecipes("")
    releasedRecipes = modOperationsPrimitiveBridge.ListRecipes("RELEASED")
    FillListFromArray mLstProcesses, processes
    FillListFromArray mLstReleasedProcesses, releasedProcesses
    FillListFromArray mLstRecipes, recipes
    FillListFromArray mLstAssignRecipes, processes
    FillListFromArray mLstLoaderRecipes, releasedRecipes
End Sub

Private Sub RefreshReusableAssignmentState()
    If mLstAssignIngredients Is Nothing Then Exit Sub
    mLstAssignIngredients.Clear
    mLstAssignAllowed.Clear
    RefreshInventoryList
End Sub

Private Sub RefreshBuilderHeader()
    Dim lo As ListObject
    Set lo = ProductionTable(TABLE_BUILDER_HEADER)
    If lo Is Nothing Then Exit Sub
    mTxtRecipeName.Text = FirstRowValue(lo, "RECIPE_NAME")
    mTxtRecipeId.Text = FirstRowValue(lo, "RECIPE_ID")
    mTxtRecipeDescription.Text = FirstRowValue(lo, "DESCRIPTION")
    mTxtRecipeRowBudget.Text = NormalizeRecipeRowBudgetText(FirstRowValue(lo, "ROW_BUDGET"))
    If Trim$(mTxtRecipeId.Text) = "" And Trim$(mTxtRecipeName.Text) = "" Then
        mTxtRecipeId.Text = GenerateRecipeIdForOperatorWorkbook()
    End If
End Sub

Private Sub RefreshBuilderLines()
    Dim lo As ListObject
    Dim headers As Variant
    Dim arr As Variant
    Dim r As Long
    Dim c As Long
    Dim colIdx As Long

    headers = Array("PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT", "INGREDIENT_ID")
    mLstBuilderLines.Clear
    mBuilderLineTableRowCount = 0
    Erase mBuilderLineTableRows

    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        If Not TableArrayRowHasAnyValue(arr, r, lo, headers) Then GoTo NextBuilderTableRow
        mLstBuilderLines.AddItem CellText(arr, r, lo, CStr(headers(LBound(headers))))
        For c = LBound(headers) + 1 To UBound(headers)
            colIdx = c - LBound(headers)
            If colIdx < mLstBuilderLines.ColumnCount Then _
                mLstBuilderLines.List(mLstBuilderLines.ListCount - 1, colIdx) = CellText(arr, r, lo, CStr(headers(c)))
        Next c

        mBuilderLineTableRowCount = mBuilderLineTableRowCount + 1
        ReDim Preserve mBuilderLineTableRows(0 To mBuilderLineTableRowCount - 1)
        mBuilderLineTableRows(mBuilderLineTableRowCount - 1) = r
NextBuilderTableRow:
    Next r
End Sub

Private Function BuilderTableRowForListIndex(ByVal listIndex As Long) As Long
    If listIndex < 0 Then Exit Function
    If listIndex >= mBuilderLineTableRowCount Then Exit Function
    BuilderTableRowForListIndex = mBuilderLineTableRows(listIndex)
End Function

Private Sub RefreshRecipeUomCatalog(Optional ByVal selectedUom As String = "")
    Dim uoms As Variant
    Dim idx As Long

    If mTxtLineUom Is Nothing Then Exit Sub
    If selectedUom = "" Then selectedUom = Trim$(CStr(mTxtLineUom.Value))
    selectedUom = Trim$(selectedUom)
    mTxtLineUom.Clear
    uoms = modUomSettings.GetConfiguredUoms()
    If IsArray(uoms) Then
        For idx = LBound(uoms) To UBound(uoms)
            mTxtLineUom.AddItem CStr(uoms(idx))
        Next idx
    End If
    If selectedUom = "" Then Exit Sub

    For idx = 0 To mTxtLineUom.ListCount - 1
        If StrComp(CStr(mTxtLineUom.List(idx)), selectedUom, vbTextCompare) = 0 Then
            mTxtLineUom.ListIndex = idx
            Exit Sub
        End If
    Next idx

    ' Preserve older recipe UOMs without silently adding them to warehouse config.
    mTxtLineUom.AddItem selectedUom
    mTxtLineUom.ListIndex = mTxtLineUom.ListCount - 1
End Sub

Private Sub RefreshAssignmentState()
    Dim recipeId As String
    Dim ingredientId As String
    Dim ingredients As Variant

    BindOperatorWorkbookForRun
    recipeId = NzStr(mProduction.GetPaletteRecipeId())
    ingredientId = NzStr(mProduction.GetPaletteIngredientId())
    If recipeId <> "" Then
        ingredients = mProduction.LoadIngredientListForRecipe(recipeId)
        FillIngredientListFromArray mLstAssignIngredients, ingredients
    Else
        mLstAssignIngredients.Clear
    End If
    RefreshInventoryList
    RefreshAllowedItems
End Sub

Private Sub RefreshInventoryList(Optional ByVal forceReload As Boolean = False)
    If forceReload Or Not mInventoryCacheLoaded Then
        BindOperatorWorkbookForRun
        mInventoryRows = mProduction.LoadProductionInventoryPickerItems("")
        mInventoryCacheLoaded = True
    End If
    FillInventoryListFromArray mInventoryRows, Trim$(mTxtInventorySearch.Text)
End Sub

Private Sub RefreshAllowedItems()
    FillListFromTable mLstAssignAllowed, ProductionTable(TABLE_ASSIGN_ITEM), _
        Array("System_Key", "ITEMS", "UOM", "DESCRIPTION", "RECIPE_ID", "INGREDIENT_ID", "ITEM_CODE")
End Sub

Private Sub RefreshLoaderState()
    FillListFromTable mLstLoaderLines, ProductionTable(TABLE_LOADER_LINES), _
        Array("PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT NEEDED", "INGREDIENT_ID")
    RefreshRunProcessChoices
    If Not mLstLoaderOutput Is Nothing Then
        RefreshProductionOutputList mLstLoaderOutput
    End If
    RefreshRunPaletteState
End Sub

Private Sub RefreshManagerState()
    RefreshProductionOutputList mLstManagerOutput
    FillListFromTable mLstManagerCheck, ProductionTable(TABLE_MANAGER_CHECK), _
        Array("System_Key", "ITEM_CODE", "ITEM", "UOM", "USED", "TOTAL INV")
    RefreshRunPaletteState
End Sub

Private Sub RefreshRunPaletteState()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim choices As Variant
    Dim filterIngredientId As String
    Dim filterIngredientName As String
    Dim filterProcess As String

    If mLstRunPalette Is Nothing Then Exit Sub
    mLstRunPalette.Clear
    If Not mLstRunTree Is Nothing Then mLstRunTree.Clear
    EnsureRunItemCodeMap
    mRunItemCodeByKey.RemoveAll
    GetSelectedRunIngredientFilter filterIngredientId, filterIngredientName
    filterProcess = ActiveRunProcess()
    EnsureRunInventoryCache
    RefreshRunLocationChoices

    BindOperatorWorkbookForRun
    choices = mProduction.LoadProductionRunIngredientChoices( _
        NzStr(mProduction.GetCurrentProductionRunRecipeId()))
    If Not IsEmpty(choices) Then
        AddRunChoiceRows choices, filterIngredientId, filterIngredientName, filterProcess
        BuildRunTreeFromPaletteList
        Exit Sub
    End If

    Set ws = mProduction.GetProductionSheet()
    If ws Is Nothing Then Exit Sub

    For Each lo In ws.ListObjects
        If IsRunPaletteTable(lo) Then AddRunPaletteRows lo, filterIngredientId, filterIngredientName, filterProcess
    Next lo
    BuildRunTreeFromPaletteList
End Sub

Private Sub AddRunChoiceRows(ByVal values As Variant, Optional ByVal filterIngredientId As String = "", _
                             Optional ByVal filterIngredientName As String = "", Optional ByVal filterProcess As String = "")
    Dim r As Long
    Dim listRow As Long
    Dim rowVal As String
    Dim itemCode As String
    Dim itemVal As String
    Dim uomVal As String
    Dim locVal As String
    Dim invVal As String
    Dim ingredientId As String
    Dim ingredientName As String
    Dim processVal As String

    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub
    EnsureRunInventoryCache
    For r = LBound(values, 1) To UBound(values, 1)
        ingredientId = NzStr(values(r, 2))
        ingredientName = NzStr(values(r, 3))
        processVal = NzStr(values(r, 1))
        If Not RunProcessMatchesFilter(processVal, filterProcess) Then GoTo NextRow
        If Not RunIngredientMatchesFilter(ingredientId, ingredientName, filterIngredientId, filterIngredientName) Then GoTo NextRow

        rowVal = NzStr(values(r, 4))
        itemVal = NzStr(values(r, 5))
        uomVal = NzStr(values(r, 8))
        locVal = NzStr(values(r, 9))
        itemCode = vbNullString
        If UBound(values, 2) >= 11 Then itemCode = NzStr(values(r, 11))
        HydrateRunInventoryDisplay rowVal, itemCode, itemVal, uomVal, invVal, locVal

        mLstRunPalette.AddItem processVal
        listRow = mLstRunPalette.ListCount - 1
        mLstRunPalette.List(listRow, 1) = ingredientId
        mLstRunPalette.List(listRow, 2) = ingredientName
        mLstRunPalette.List(listRow, 3) = rowVal
        mLstRunPalette.List(listRow, 4) = itemVal
        mLstRunPalette.List(listRow, 5) = NzStr(values(r, 6))
        mLstRunPalette.List(listRow, 6) = NzStr(values(r, 7))
        mLstRunPalette.List(listRow, 7) = uomVal
        mLstRunPalette.List(listRow, 8) = invVal
        mLstRunPalette.List(listRow, 9) = locVal
        StoreRunProcess mLstRunPalette, listRow, processVal
        StoreRunBaseQty mLstRunPalette, listRow, NzStr(values(r, 10))
        StoreRunItemCode mLstRunPalette, listRow, itemCode
        ApplyRunAllocationOverride mLstRunPalette, listRow
NextRow:
    Next r
End Sub

Private Function IsRunPaletteTable(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    If StrComp(lo.Name, TABLE_MANAGER_PALETTE, vbTextCompare) = 0 Then
        If lo.Range.Row >= 500000 Then Exit Function
    End If
    IsRunPaletteTable = ProductionColumnIndex(lo, "ITEM") > 0 _
        And ProductionColumnIndex(lo, "QUANTITY") > 0 _
        And ProductionColumnIndex(lo, "PROCESS") > 0 _
        And ProductionColumnIndex(lo, "System_Key") > 0 _
        And ProductionColumnIndex(lo, "INPUT/OUTPUT") > 0
End Function

Private Sub AddRunPaletteRows(ByVal lo As ListObject, Optional ByVal filterIngredientId As String = "", _
                              Optional ByVal filterIngredientName As String = "", Optional ByVal filterProcess As String = "")
    Dim arr As Variant
    Dim r As Long
    Dim listRow As Long
    Dim rowVal As String
    Dim ingredientVal As String
    Dim ingredientId As String
    Dim processVal As String
    Dim itemVal As String
    Dim splitVal As String
    Dim qtyVal As String
    Dim uomVal As String
    Dim locVal As String
    Dim invVal As String
    Dim baseQty As String
    Dim ioVal As String
    Dim itemCode As String

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    EnsureRunInventoryCache
    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        ingredientVal = CellText(arr, r, lo, "INGREDIENT")
        ingredientId = CellText(arr, r, lo, "INGREDIENT_ID")
        processVal = CellText(arr, r, lo, "PROCESS")
        If Not RunProcessMatchesFilter(processVal, filterProcess) Then GoTo NextRow
        ioVal = UCase$(Trim$(CellText(arr, r, lo, "INPUT/OUTPUT")))
        If ioVal <> "" And ioVal <> "USED" Then GoTo NextRow
        If Not RunIngredientMatchesFilter(ingredientId, IngredientDisplayName(ingredientVal, processVal), filterIngredientId, filterIngredientName) Then GoTo NextRow

        rowVal = CellText(arr, r, lo, "System_Key")
        itemVal = CellText(arr, r, lo, "ITEM")
        splitVal = CellText(arr, r, lo, "SPLIT %")
        qtyVal = CellText(arr, r, lo, "QUANTITY")
        uomVal = CellText(arr, r, lo, "UOM")
        locVal = CellText(arr, r, lo, "LOCATION")
        baseQty = CellText(arr, r, lo, "BASE QUANTITY")
        itemCode = CellText(arr, r, lo, "ITEM_CODE")
        HydrateRunInventoryDisplay rowVal, itemCode, itemVal, uomVal, invVal, locVal

        mLstRunPalette.AddItem lo.Name
        listRow = mLstRunPalette.ListCount - 1
        mLstRunPalette.List(listRow, 1) = CStr(r)
        mLstRunPalette.List(listRow, 2) = IngredientDisplayName(ingredientVal, processVal)
        mLstRunPalette.List(listRow, 3) = rowVal
        mLstRunPalette.List(listRow, 4) = itemVal
        mLstRunPalette.List(listRow, 5) = splitVal
        mLstRunPalette.List(listRow, 6) = qtyVal
        mLstRunPalette.List(listRow, 7) = uomVal
        mLstRunPalette.List(listRow, 8) = invVal
        mLstRunPalette.List(listRow, 9) = locVal
        StoreRunProcess mLstRunPalette, listRow, processVal
        StoreRunBaseQty mLstRunPalette, listRow, baseQty
        StoreRunItemCode mLstRunPalette, listRow, itemCode
        ApplyRunAllocationOverride mLstRunPalette, listRow
NextRow:
    Next r
End Sub

Private Sub EnsureRunSplitOverrides()
    If mRunSplitOverrides Is Nothing Then Set mRunSplitOverrides = CreateObject("Scripting.Dictionary")
End Sub

Private Sub EnsureRunBaseQtyMap()
    If mRunBaseQtyByKey Is Nothing Then Set mRunBaseQtyByKey = CreateObject("Scripting.Dictionary")
End Sub

Private Sub EnsureRunProcessMap()
    If mRunProcessByKey Is Nothing Then Set mRunProcessByKey = CreateObject("Scripting.Dictionary")
End Sub

Private Sub EnsureRunItemCodeMap()
    If mRunItemCodeByKey Is Nothing Then Set mRunItemCodeByKey = CreateObject("Scripting.Dictionary")
End Sub

Private Function RunAllocationKeyFromList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    If IsRunTreeOutputRow(lst, rowIndex) Then Exit Function
    RunAllocationKeyFromList = Trim$(NzStr(lst.List(rowIndex, 0))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 1))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 3))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 4))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 7))) & "|" & _
                               Trim$(NzStr(lst.List(rowIndex, 9)))
End Function

Private Sub StoreRunItemCode(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long, ByVal itemCode As String)
    Dim key As String

    key = RunAllocationKeyFromList(lst, rowIndex)
    itemCode = Trim$(itemCode)
    If key = "" Or itemCode = "" Then Exit Sub
    EnsureRunItemCodeMap
    mRunItemCodeByKey(key) = itemCode
End Sub

Private Function RunItemCodeFromList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    Dim key As String

    If mRunItemCodeByKey Is Nothing Then Exit Function
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Function
    If mRunItemCodeByKey.Exists(key) Then RunItemCodeFromList = NzStr(mRunItemCodeByKey(key))
End Function

Private Sub StoreRunProcess(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long, ByVal processText As String)
    Dim key As String
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Sub
    EnsureRunProcessMap
    mRunProcessByKey(key) = Trim$(processText)
End Sub

Private Sub StoreRunBaseQty(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long, ByVal baseQtyText As String)
    Dim key As String
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Sub
    EnsureRunBaseQtyMap
    mRunBaseQtyByKey(key) = baseQtyText
End Sub

Private Function RunBaseQtyFromList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As Double
    Dim key As String
    Dim baseText As String

    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    If mRunBaseQtyByKey Is Nothing Then Exit Function
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Function
    If Not mRunBaseQtyByKey.Exists(key) Then Exit Function
    baseText = NzStr(mRunBaseQtyByKey(key))
    If IsNumeric(baseText) Then RunBaseQtyFromList = CDbl(baseText)
End Function

Private Function RunAllocationGroupKeyFromList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    Dim r As Long

    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    If Not mLstRunTree Is Nothing Then
        If lst Is mLstRunTree Then
            If IsRunTreeOutputRow(lst, rowIndex) Then Exit Function
            If IsRunTreeParentRow(lst, rowIndex) Then
                RunAllocationGroupKeyFromList = Trim$(NzStr(lst.List(rowIndex, 1)))
                Exit Function
            End If
            For r = rowIndex To 0 Step -1
                If IsRunTreeParentRow(lst, r) Then
                    RunAllocationGroupKeyFromList = Trim$(NzStr(lst.List(r, 1)))
                    Exit Function
                End If
            Next r
        End If
    End If
    RunAllocationGroupKeyFromList = Trim$(NzStr(lst.List(rowIndex, 1)))
    If RunAllocationGroupKeyFromList = "" Or IsNumeric(RunAllocationGroupKeyFromList) Then
        RunAllocationGroupKeyFromList = Trim$(NzStr(lst.List(rowIndex, 2)))
    End If
    If RunAllocationGroupKeyFromList <> "" Then
        RunAllocationGroupKeyFromList = ProcessKey(RunProcessFromPaletteList(lst, rowIndex)) & "|" & RunAllocationGroupKeyFromList
    End If
End Function

Private Function RunProcessFromPaletteList(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    Dim key As String

    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key <> "" Then
        If Not mRunProcessByKey Is Nothing Then
            If mRunProcessByKey.Exists(key) Then
                RunProcessFromPaletteList = Trim$(NzStr(mRunProcessByKey(key)))
                If RunProcessFromPaletteList <> "" Then Exit Function
            End If
        End If
    End If
    If IsNumeric(NzStr(lst.List(rowIndex, 1))) Then Exit Function
    RunProcessFromPaletteList = Trim$(NzStr(lst.List(rowIndex, 0)))
End Function

Private Function RunAllocationPercentForGroup(ByVal groupKey As String, Optional ByVal excludeAllocationKey As String = "") As Double
    Dim i As Long
    Dim splitText As String

    If mLstRunPalette Is Nothing Then Exit Function
    groupKey = Trim$(groupKey)
    For i = 0 To mLstRunPalette.ListCount - 1
        If StrComp(RunAllocationGroupKeyFromList(mLstRunPalette, i), groupKey, vbTextCompare) = 0 Then
            If excludeAllocationKey = "" Or RunAllocationKeyFromList(mLstRunPalette, i) <> excludeAllocationKey Then
                splitText = Trim$(NzStr(mLstRunPalette.List(i, 5)))
                If IsNumeric(splitText) Then RunAllocationPercentForGroup = RunAllocationPercentForGroup + CDbl(splitText)
            End If
        End If
    Next i
End Function

Private Function RunAllocationStatusText(ByVal groupKey As String) As String
    Dim totalPct As Double

    totalPct = RunAllocationPercentForGroup(groupKey)
    If totalPct >= 99.999 And totalPct <= 100.001 Then
        RunAllocationStatusText = "FILLED 100%"
    ElseIf totalPct > 0 Then
        RunAllocationStatusText = FormatRunNumber(totalPct) & "% filled"
    Else
        RunAllocationStatusText = "0% filled"
    End If
End Function

Private Sub StoreRunAllocationOverride(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long, _
                                       ByVal splitText As String, ByVal qtyText As String)
    Dim key As String
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Sub
    EnsureRunSplitOverrides
    mRunSplitOverrides(key) = Array(splitText, qtyText)
End Sub

Private Sub ApplyRunAllocationOverride(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long)
    Dim key As String
    Dim values As Variant

    If lst Is Nothing Then Exit Sub
    If rowIndex < 0 Then Exit Sub
    If mRunSplitOverrides Is Nothing Then Exit Sub
    key = RunAllocationKeyFromList(lst, rowIndex)
    If key = "" Then Exit Sub
    If Not mRunSplitOverrides.Exists(key) Then Exit Sub
    values = mRunSplitOverrides(key)
    lst.List(rowIndex, 5) = NzStr(values(0))
    lst.List(rowIndex, 6) = NzStr(values(1))
End Sub

Private Sub GetSelectedRunIngredientFilter(ByRef ingredientId As String, ByRef ingredientName As String)
    Dim idx As Long

    ingredientId = ""
    ingredientName = ""
    If mLstLoaderLines Is Nothing Then Exit Sub
    idx = mLstLoaderLines.ListIndex
    If idx < 0 Then Exit Sub
    ingredientName = NzStr(mLstLoaderLines.List(idx, 3))
    ingredientId = NzStr(mLstLoaderLines.List(idx, 7))
End Sub

Private Function RunIngredientMatchesFilter(ByVal rowIngredientId As String, ByVal rowIngredientName As String, _
                                            ByVal filterIngredientId As String, ByVal filterIngredientName As String) As Boolean
    rowIngredientId = Trim$(rowIngredientId)
    rowIngredientName = Trim$(rowIngredientName)
    filterIngredientId = Trim$(filterIngredientId)
    filterIngredientName = Trim$(filterIngredientName)

    If filterIngredientId = "" And filterIngredientName = "" Then
        RunIngredientMatchesFilter = True
    ElseIf filterIngredientId <> "" And rowIngredientId <> "" Then
        RunIngredientMatchesFilter = (StrComp(rowIngredientId, filterIngredientId, vbTextCompare) = 0)
    ElseIf filterIngredientName <> "" Then
        RunIngredientMatchesFilter = (StrComp(rowIngredientName, filterIngredientName, vbTextCompare) = 0)
    End If
End Function

Private Function RunProcessMatchesFilter(ByVal rowProcess As String, ByVal filterProcess As String) As Boolean
    rowProcess = Trim$(rowProcess)
    filterProcess = Trim$(filterProcess)
    If filterProcess = "" Then
        RunProcessMatchesFilter = True
    Else
        RunProcessMatchesFilter = (StrComp(rowProcess, filterProcess, vbTextCompare) = 0)
    End If
End Function

Private Sub BuildRunTreeFromPaletteList()
    Dim processes As Object
    Dim procKey As Variant
    Dim procName As String
    Dim rowIndex As Long
    Dim collapsed As Boolean

    If mLstRunTree Is Nothing Or mLstRunPalette Is Nothing Then Exit Sub
    EnsureRunTreeState
    mLstRunTree.Clear
    Set processes = OrderedRunProcesses()
    If processes Is Nothing Then Exit Sub

    For Each procKey In processes.Keys
        procName = NzStr(processes(procKey))
        collapsed = RunTreeGroupCollapsed("PROC|" & CStr(procKey))
        mLstRunTree.AddItem RUN_TREE_PARENT_MARKER
        rowIndex = mLstRunTree.ListCount - 1
        mLstRunTree.List(rowIndex, 1) = "PROC|" & CStr(procKey)
        If collapsed Then
            mLstRunTree.List(rowIndex, 2) = "[ SHOW PROCESS ]  " & ProcessCaption(procName)
        Else
            mLstRunTree.List(rowIndex, 2) = "[ HIDE PROCESS ]  " & ProcessCaption(procName)
        End If
        If Not collapsed Then
            AddRunTreeInputRowsForProcess procName
            AddRunTreeOutputRowsForProcess procName
        End If
    Next procKey
End Sub

Private Function OrderedRunProcesses() As Object
    Dim dict As Object
    Dim i As Long
    Dim procName As String
    Dim key As String
    Dim filterProcess As String

    filterProcess = ActiveRunProcess()
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 0 To mLstRunPalette.ListCount - 1
        procName = RunProcessFromPaletteList(mLstRunPalette, i)
        If Not RunProcessMatchesFilter(procName, filterProcess) Then GoTo NextPaletteProcess
        key = ProcessKey(procName)
        If Not dict.Exists(key) Then dict.Add key, procName
NextPaletteProcess:
    Next i
    If Not mLstManagerOutput Is Nothing Then
        For i = 0 To mLstManagerOutput.ListCount - 1
            procName = NzStr(mLstManagerOutput.List(i, 0))
            If Not RunProcessMatchesFilter(procName, filterProcess) Then GoTo NextOutputProcess
            key = ProcessKey(procName)
            If Not dict.Exists(key) Then dict.Add key, procName
NextOutputProcess:
        Next i
    End If
    Set OrderedRunProcesses = dict
End Function

Private Sub AddRunTreeInputRowsForProcess(ByVal procName As String)
    Dim i As Long
    Dim key As String
    Dim lastKey As String
    Dim parentText As String
    Dim childText As String
    Dim rowIndex As Long
    Dim collapsed As Boolean

    For i = 0 To mLstRunPalette.ListCount - 1
        If StrComp(RunProcessFromPaletteList(mLstRunPalette, i), procName, vbTextCompare) <> 0 Then GoTo NextPaletteRow
        key = RunTreeGroupKey(mLstRunPalette, i)
        If key <> lastKey Then
            parentText = NzStr(mLstRunPalette.List(i, 2))
            If parentText = "" Then parentText = "Ingredient"
            collapsed = RunTreeGroupCollapsed(key)
            mLstRunTree.AddItem RUN_TREE_PARENT_MARKER
            rowIndex = mLstRunTree.ListCount - 1
            mLstRunTree.List(rowIndex, 1) = key
            mLstRunTree.List(rowIndex, 2) = RunTreeParentCaption("INPUT  " & parentText, collapsed, key)
            lastKey = key
        End If

        If collapsed Then GoTo NextPaletteRow
        childText = "    " & NzStr(mLstRunPalette.List(i, 4))
        If Trim$(childText) = "" Then childText = "    System_Key " & NzStr(mLstRunPalette.List(i, 3))
        mLstRunTree.AddItem NzStr(mLstRunPalette.List(i, 0))
        rowIndex = mLstRunTree.ListCount - 1
        CopyRunPaletteListRow mLstRunPalette, i, mLstRunTree, rowIndex
        mLstRunTree.List(rowIndex, 2) = childText
NextPaletteRow:
    Next i
End Sub

Private Sub AddRunTreeOutputRowsForProcess(ByVal procName As String)
    Dim i As Long
    Dim rowIndex As Long
    Dim outputName As String
    Dim caption As String

    If mLstManagerOutput Is Nothing Then Exit Sub
    For i = 0 To mLstManagerOutput.ListCount - 1
        If StrComp(NzStr(mLstManagerOutput.List(i, 0)), procName, vbTextCompare) <> 0 Then GoTo NextOutput
        outputName = NzStr(mLstManagerOutput.List(i, 1))
        If Trim$(outputName) = "" Then GoTo NextOutput
        caption = "  OUTPUT  " & outputName
        If NzStr(mLstManagerOutput.List(i, 3)) <> "" Then caption = caption & "  Last Actual " & NzStr(mLstManagerOutput.List(i, 3))
        caption = caption & "  Batch " & NzStr(mLstManagerOutput.List(i, 4)) & _
            "  Used Goods " & NzStr(mLstManagerOutput.List(i, 5)) & _
            "  Process Total " & NzStr(mLstManagerOutput.List(i, 6))
        mLstRunTree.AddItem RUN_TREE_OUTPUT_MARKER
        rowIndex = mLstRunTree.ListCount - 1
        mLstRunTree.List(rowIndex, 1) = "OUTPUT|" & ProcessKey(procName) & "|" & NzStr(mLstManagerOutput.List(i, 8))
        mLstRunTree.List(rowIndex, 2) = caption
        mLstRunTree.List(rowIndex, 3) = NzStr(mLstManagerOutput.List(i, 8))
        mLstRunTree.List(rowIndex, 4) = outputName
        mLstRunTree.List(rowIndex, 5) = NzStr(mLstManagerOutput.List(i, 3))
        mLstRunTree.List(rowIndex, 6) = NzStr(mLstManagerOutput.List(i, 5))
        mLstRunTree.List(rowIndex, 7) = NzStr(mLstManagerOutput.List(i, 2))
        mLstRunTree.List(rowIndex, 8) = "Batch " & NzStr(mLstManagerOutput.List(i, 4))
NextOutput:
    Next i
End Sub

Private Function ProcessKey(ByVal procName As String) As String
    ProcessKey = LCase$(Trim$(procName))
    If ProcessKey = "" Then ProcessKey = "(no process)"
End Function

Private Function ProcessCaption(ByVal procName As String) As String
    ProcessCaption = Trim$(procName)
    If ProcessCaption = "" Then ProcessCaption = "Process"
End Function

Private Sub EnsureRunTreeState()
    If mRunTreeCollapsed Is Nothing Then Set mRunTreeCollapsed = CreateObject("Scripting.Dictionary")
End Sub

Private Function RunTreeGroupKey(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As String
    RunTreeGroupKey = RunAllocationGroupKeyFromList(lst, rowIndex)
End Function

Private Function RunTreeGroupCollapsed(ByVal groupKey As String) As Boolean
    EnsureRunTreeState
    If groupKey = "" Then Exit Function
    If mRunTreeCollapsed.Exists(groupKey) Then RunTreeGroupCollapsed = CBool(mRunTreeCollapsed(groupKey))
End Function

Private Function RunTreeParentCaption(ByVal parentText As String, ByVal collapsed As Boolean, ByVal groupKey As String) As String
    Dim statusText As String
    statusText = RunAllocationStatusText(groupKey)
    If collapsed Then
        RunTreeParentCaption = "[ SHOW CHOICES ]  " & parentText & "  --  " & statusText
    Else
        RunTreeParentCaption = "[ HIDE CHOICES ]  " & parentText & "  --  " & statusText
    End If
End Function

Private Function IsRunTreeParentRow(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As Boolean
    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    IsRunTreeParentRow = (NzStr(lst.List(rowIndex, 0)) = RUN_TREE_PARENT_MARKER)
End Function

Private Function IsRunTreeOutputRow(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long) As Boolean
    If lst Is Nothing Then Exit Function
    If rowIndex < 0 Then Exit Function
    IsRunTreeOutputRow = (NzStr(lst.List(rowIndex, 0)) = RUN_TREE_OUTPUT_MARKER)
End Function

Private Sub ToggleSelectedRunTreeParent()
    Dim idx As Long
    Dim key As String

    If mLstRunTree Is Nothing Then Exit Sub
    idx = mLstRunTree.ListIndex
    If Not IsRunTreeParentRow(mLstRunTree, idx) Then Exit Sub
    EnsureRunTreeState
    key = NzStr(mLstRunTree.List(idx, 1))
    If key = "" Then Exit Sub
    mRunTreeCollapsed(key) = Not RunTreeGroupCollapsed(key)
    BuildRunTreeFromPaletteList
    If RunTreeGroupCollapsed(key) Then
        ShowStatus "Ingredient choices hidden."
    Else
        ShowStatus "Ingredient choices shown."
    End If
End Sub

Private Sub SetAllRunTreeGroupsCollapsed(ByVal collapsed As Boolean)
    Dim i As Long
    Dim key As String

    If mLstRunPalette Is Nothing Then Exit Sub
    EnsureRunTreeState
    If Not collapsed Then
        mRunTreeCollapsed.RemoveAll
    Else
        For i = 0 To mLstRunPalette.ListCount - 1
            key = RunTreeGroupKey(mLstRunPalette, i)
            If key <> "" Then mRunTreeCollapsed(key) = True
        Next i
    End If
    BuildRunTreeFromPaletteList
End Sub

Private Sub CopyRunPaletteListRow(ByVal sourceList As MSForms.ListBox, ByVal sourceRow As Long, _
                                  ByVal targetList As MSForms.ListBox, ByVal targetRow As Long)
    Dim c As Long
    For c = 0 To sourceList.ColumnCount - 1
        If c < targetList.ColumnCount Then targetList.List(targetRow, c) = NzStr(sourceList.List(sourceRow, c))
    Next c
End Sub

Private Function IngredientDisplayName(ByVal ingredientVal As String, ByVal processVal As String) As String
    IngredientDisplayName = Trim$(ingredientVal)
    If IngredientDisplayName = "" Then IngredientDisplayName = Trim$(processVal)
End Function

Private Sub EnsureRunInventoryCache()
    If Not mRunInventoryCacheLoaded Then
        BindOperatorWorkbookForRun
        mRunInventoryRows = mProduction.LoadProductionRunInventoryPickerItems("")
        mRunInventoryCacheLoaded = True
    End If
End Sub

Private Sub RefreshRunLocationChoices()
    Dim dict As Object
    Dim r As Long
    Dim locVal As String
    Dim defaultLoc As String
    Dim selectedLoc As String
    Dim wasLoading As Boolean

    If mCmbRunLocation Is Nothing And mCmbTreeRunLocation Is Nothing Then Exit Sub

    selectedLoc = ActiveRunLocation()
    Set dict = CreateObject("Scripting.Dictionary")
    If Not IsEmpty(mRunInventoryRows) And IsArray(mRunInventoryRows) Then
        For r = LBound(mRunInventoryRows, 1) To UBound(mRunInventoryRows, 1)
            locVal = Trim$(NzStr(mRunInventoryRows(r, 5)))
            AddRunLocationChoice dict, locVal
        Next r
    End If
    AddRunLocationChoice dict, selectedLoc
    BindOperatorWorkbookForRun
    defaultLoc = Trim$(NzStr(mProduction.GetProductionRunDefaultLocation()))
    AddRunLocationChoice dict, defaultLoc

    wasLoading = mLoading
    mLoading = True
    PopulateRunLocationCombo mCmbRunLocation, dict, selectedLoc
    PopulateRunLocationCombo mCmbTreeRunLocation, dict, selectedLoc
    mLoading = wasLoading
End Sub

Private Sub AddRunLocationChoice(ByVal dict As Object, ByVal locationText As String)
    locationText = Trim$(locationText)
    If dict Is Nothing Or locationText = "" Then Exit Sub
    If Not dict.Exists(LCase$(locationText)) Then dict.Add LCase$(locationText), locationText
End Sub

Private Sub RefreshRunProcessChoices()
    Dim dict As Object
    Dim i As Long
    Dim procVal As String
    Dim selectedProcess As String
    Dim wasLoading As Boolean

    If mCmbRunProcess Is Nothing And mCmbTreeRunProcess Is Nothing Then Exit Sub
    If mLstLoaderLines Is Nothing Then Exit Sub

    selectedProcess = ActiveRunProcess()
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 0 To mLstLoaderLines.ListCount - 1
        procVal = Trim$(NzStr(mLstLoaderLines.List(i, 0)))
        If procVal <> "" Then
            If Not dict.Exists(LCase$(procVal)) Then dict.Add LCase$(procVal), procVal
        End If
    Next i

    wasLoading = mLoading
    mLoading = True
    PopulateRunProcessCombo mCmbRunProcess, dict, selectedProcess
    PopulateRunProcessCombo mCmbTreeRunProcess, dict, selectedProcess
    mLoading = wasLoading
End Sub

Private Sub PopulateRunProcessCombo(ByVal cmb As MSForms.ComboBox, ByVal dict As Object, ByVal selectedProcess As String)
    Dim key As Variant
    Dim i As Long

    If cmb Is Nothing Then Exit Sub
    cmb.Clear
    cmb.AddItem ""
    If Not dict Is Nothing Then
        For Each key In dict.Keys
            cmb.AddItem CStr(dict(key))
        Next key
    End If
    cmb.ListIndex = 0
    selectedProcess = Trim$(selectedProcess)
    If selectedProcess <> "" Then
        For i = 0 To cmb.ListCount - 1
            If StrComp(NzStr(cmb.List(i)), selectedProcess, vbTextCompare) = 0 Then
                cmb.ListIndex = i
                Exit For
            End If
        Next i
    End If
End Sub

Private Sub PopulateRunLocationCombo(ByVal cmb As MSForms.ComboBox, ByVal dict As Object, ByVal selectedLoc As String)
    Dim key As Variant
    Dim i As Long

    If cmb Is Nothing Then Exit Sub
    cmb.Clear
    cmb.AddItem ""
    If Not dict Is Nothing Then
        For Each key In dict.Keys
            cmb.AddItem CStr(dict(key))
        Next key
    End If
    cmb.ListIndex = 0
    selectedLoc = Trim$(selectedLoc)
    If selectedLoc <> "" Then
        For i = 0 To cmb.ListCount - 1
            If StrComp(NzStr(cmb.List(i)), selectedLoc, vbTextCompare) = 0 Then
                cmb.ListIndex = i
                Exit For
            End If
        Next i
    End If
End Sub

Private Function ActiveRunLocation() As String
    If Not mPages Is Nothing Then
        If mPages.Value = 4 Then
            ActiveRunLocation = ComboText(mCmbTreeRunLocation)
            If ActiveRunLocation <> "" Then Exit Function
        End If
    End If
    ActiveRunLocation = ComboText(mCmbRunLocation)
    If ActiveRunLocation = "" Then ActiveRunLocation = ComboText(mCmbTreeRunLocation)
End Function

Private Function ActiveRunProcess() As String
    If Not mPages Is Nothing Then
        If mPages.Value = 4 Then
            ActiveRunProcess = ComboText(mCmbTreeRunProcess)
            If ActiveRunProcess <> "" Then Exit Function
        End If
    End If
    ActiveRunProcess = ComboText(mCmbRunProcess)
    If ActiveRunProcess = "" Then ActiveRunProcess = ComboText(mCmbTreeRunProcess)
End Function

Private Function ComboText(ByVal cmb As MSForms.ComboBox) As String
    If cmb Is Nothing Then Exit Function
    If cmb.ListIndex < 0 Then Exit Function
    ComboText = Trim$(NzStr(cmb.Value))
End Function

Private Function RunLocationChoiceRequired() As Boolean
    If Not mCmbRunLocation Is Nothing Then
        If mCmbRunLocation.ListCount > 1 Then
            RunLocationChoiceRequired = True
            Exit Function
        End If
    End If
    If Not mCmbTreeRunLocation Is Nothing Then
        If mCmbTreeRunLocation.ListCount > 1 Then RunLocationChoiceRequired = True
    End If
End Function

Private Sub SyncRunLocationCombo(ByVal sourceCombo As MSForms.ComboBox, ByVal targetCombo As MSForms.ComboBox)
    Dim i As Long
    Dim locVal As String
    Dim wasLoading As Boolean

    If mLoading Then Exit Sub
    If sourceCombo Is Nothing Or targetCombo Is Nothing Then Exit Sub
    locVal = ComboText(sourceCombo)
    wasLoading = mLoading
    mLoading = True
    If targetCombo.ListCount > 0 Then
        targetCombo.ListIndex = 0
    Else
        targetCombo.Value = ""
    End If
    For i = 0 To targetCombo.ListCount - 1
        If StrComp(NzStr(targetCombo.List(i)), locVal, vbTextCompare) = 0 Then
            targetCombo.ListIndex = i
            Exit For
        End If
    Next i
    mLoading = wasLoading
End Sub

Private Sub SyncRunProcessCombo(ByVal sourceCombo As MSForms.ComboBox, ByVal targetCombo As MSForms.ComboBox)
    Dim i As Long
    Dim procVal As String
    Dim wasLoading As Boolean

    If mLoading Then Exit Sub
    If sourceCombo Is Nothing Or targetCombo Is Nothing Then Exit Sub
    procVal = ComboText(sourceCombo)
    wasLoading = mLoading
    mLoading = True
    If targetCombo.ListCount > 0 Then
        targetCombo.ListIndex = 0
    Else
        targetCombo.Value = ""
    End If
    For i = 0 To targetCombo.ListCount - 1
        If StrComp(NzStr(targetCombo.List(i)), procVal, vbTextCompare) = 0 Then
            targetCombo.ListIndex = i
            Exit For
        End If
    Next i
    mLoading = wasLoading
End Sub

Private Sub HydrateRunInventoryDisplay(ByVal rowVal As String, ByRef itemCode As String, _
                                       ByRef itemVal As String, _
                                       ByRef uomVal As String, ByRef invVal As String, _
                                       ByRef locVal As String)
    Dim r As Long
    Dim selectedRow As Long
    Dim rowKey As String
    Dim totalVal As String
    Dim rawLoc As String
    Dim preferredLoc As String
    Dim candidateCode As String

    If IsEmpty(mRunInventoryRows) Then Exit Sub
    If Not IsArray(mRunInventoryRows) Then Exit Sub

    rowKey = NormalizeRunSystemKey(rowVal)
    itemCode = Trim$(itemCode)
    If rowKey = "" And itemCode = "" Then Exit Sub
    selectedRow = -1
    preferredLoc = ActiveRunLocation()
    For r = LBound(mRunInventoryRows, 1) To UBound(mRunInventoryRows, 1)
        candidateCode = vbNullString
        If UBound(mRunInventoryRows, 2) >= 7 Then candidateCode = Trim$(NzStr(mRunInventoryRows(r, 7)))
        If (rowKey <> "" And NormalizeRunSystemKey(NzStr(mRunInventoryRows(r, 1))) = rowKey) _
           Or (itemCode <> "" And StrComp(candidateCode, itemCode, vbTextCompare) = 0) Then
            rawLoc = NzStr(mRunInventoryRows(r, 5))
            If selectedRow < 0 Then selectedRow = r
            If Trim$(rawLoc) <> "" And Trim$(NzStr(mRunInventoryRows(selectedRow, 5))) = "" Then selectedRow = r
            If preferredLoc <> "" Then
                If StrComp(Trim$(rawLoc), preferredLoc, vbTextCompare) = 0 Then
                    selectedRow = r
                    Exit For
                End If
            End If
        End If
    Next r
    If selectedRow < 0 Then Exit Sub

    If Trim$(itemVal) = "" Then itemVal = NzStr(mRunInventoryRows(selectedRow, 2))
    If Trim$(uomVal) = "" Then uomVal = NzStr(mRunInventoryRows(selectedRow, 3))
    If itemCode = "" And UBound(mRunInventoryRows, 2) >= 7 Then _
        itemCode = NzStr(mRunInventoryRows(selectedRow, 7))
    totalVal = NzStr(mRunInventoryRows(selectedRow, 4))
    rawLoc = NzStr(mRunInventoryRows(selectedRow, 5))
    invVal = RunInventoryAvailableDisplay(totalVal, uomVal)
    locVal = rawLoc
End Sub

Private Function RunInventoryAvailableDisplay(ByVal totalVal As String, ByVal uomVal As String) As String
    totalVal = Trim$(totalVal)
    uomVal = Trim$(uomVal)
    If StrComp(totalVal, "utility", vbTextCompare) = 0 Then
        RunInventoryAvailableDisplay = "utility"
        Exit Function
    End If
    If totalVal <> "" Then
        RunInventoryAvailableDisplay = totalVal
        If uomVal <> "" Then RunInventoryAvailableDisplay = RunInventoryAvailableDisplay & " " & uomVal
    End If
End Function

Private Function NormalizeRunSystemKey(ByVal value As String) As String
    NormalizeRunSystemKey = Trim$(value)
End Function

Private Sub FillListFromArray(ByVal lst As MSForms.ListBox, ByVal values As Variant)
    Dim r As Long
    Dim c As Long
    Dim lb As Long
    Dim ub As Long

    lst.Clear
    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub
    lb = LBound(values, 1)
    ub = UBound(values, 1)
    For r = lb To ub
        lst.AddItem NzStr(values(r, LBound(values, 2)))
        For c = LBound(values, 2) + 1 To UBound(values, 2)
            If c - LBound(values, 2) < lst.ColumnCount Then
                lst.List(lst.ListCount - 1, c - LBound(values, 2)) = NzStr(values(r, c))
            End If
        Next c
    Next r
End Sub

Private Sub FillIngredientListFromArray(ByVal lst As MSForms.ListBox, ByVal values As Variant)
    Dim r As Long
    Dim c As Long

    lst.Clear
    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub
    For r = LBound(values, 1) To UBound(values, 1)
        lst.AddItem NzStr(values(r, LBound(values, 2)))
        For c = LBound(values, 2) + 1 To UBound(values, 2)
            If c - LBound(values, 2) < lst.ColumnCount Then
                lst.List(lst.ListCount - 1, c - LBound(values, 2)) = NzStr(values(r, c))
            End If
        Next c
    Next r
End Sub

Private Sub FillListFromTable(ByVal lst As MSForms.ListBox, ByVal lo As ListObject, ByVal headers As Variant)
    Dim arr As Variant
    Dim r As Long
    Dim c As Long
    Dim colIdx As Long

    lst.Clear
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        If Not TableArrayRowHasAnyValue(arr, r, lo, headers) Then GoTo NextTableRow
        lst.AddItem CellText(arr, r, lo, CStr(headers(LBound(headers))))
        For c = LBound(headers) + 1 To UBound(headers)
            colIdx = c - LBound(headers)
            If colIdx < lst.ColumnCount Then lst.List(lst.ListCount - 1, colIdx) = CellText(arr, r, lo, CStr(headers(c)))
        Next c
NextTableRow:
    Next r
End Sub

Private Sub RefreshProductionOutputList(ByVal lst As MSForms.ListBox)
    Dim lo As ListObject
    Dim arr As Variant
    Dim r As Long
    Dim listRow As Long
    Dim cProc As Long
    Dim cOutput As Long
    Dim cUom As Long
    Dim cRecall As Long
    Dim cRow As Long
    Dim cItemCode As Long
    Dim procVal As String
    Dim outputVal As String
    Dim uomVal As String
    Dim recallVal As String
    Dim rowVal As String
    Dim itemCodeVal As String
    Dim identityVal As String
    Dim lastQty As Double
    Dim totalQty As Double
    Dim maxBatch As Long
    Dim loggedCount As Long

    If lst Is Nothing Then Exit Sub
    lst.Clear
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    cProc = ProductionColumnIndex(lo, "PROCESS")
    cOutput = ProductionColumnIndex(lo, "OUTPUT")
    cUom = ProductionColumnIndex(lo, "UOM")
    cRecall = ProductionColumnIndex(lo, "RECALL CODE")
    cRow = ProductionColumnIndex(lo, "System_Key")
    cItemCode = ProductionColumnIndex(lo, "ITEM_CODE")
    If cProc = 0 Or cOutput = 0 Then Exit Sub

    arr = lo.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        procVal = NzStr(arr(r, cProc))
        outputVal = NzStr(arr(r, cOutput))
        If Trim$(procVal) = "" And Trim$(outputVal) = "" Then GoTo NextRow
        If cUom > 0 Then uomVal = NzStr(arr(r, cUom)) Else uomVal = ""
        If cRecall > 0 Then recallVal = NzStr(arr(r, cRecall)) Else recallVal = ""
        If cRow > 0 Then rowVal = NzStr(arr(r, cRow)) Else rowVal = ""
        If cItemCode > 0 Then itemCodeVal = NzStr(arr(r, cItemCode)) Else itemCodeVal = ""
        If Trim$(itemCodeVal) <> "" Then
            identityVal = itemCodeVal
        Else
            identityVal = rowVal
        End If

        LoggedOutputStats identityVal, procVal, outputVal, lastQty, totalQty, maxBatch, loggedCount
        lst.AddItem procVal
        listRow = lst.ListCount - 1
        lst.List(listRow, 1) = outputVal
        lst.List(listRow, 2) = uomVal
        lst.List(listRow, 3) = IIf(loggedCount > 0, FormatRunNumber(lastQty), "")
        ' The Batch column is completed-run history: 0 before the first
        ' completion, 1 after the first completion, and so on.  The next
        ' internal batch number is calculated only when Actual Output is staged.
        lst.List(listRow, 4) = CStr(maxBatch)
        lst.List(listRow, 5) = ""
        lst.List(listRow, 6) = IIf(loggedCount > 0, FormatRunNumber(totalQty), "0")
        lst.List(listRow, 7) = recallVal
        lst.List(listRow, 8) = identityVal
NextRow:
    Next r
End Sub

Private Function TableArrayRowHasAnyValue(ByVal arr As Variant, ByVal rowIndex As Long, ByVal lo As ListObject, ByVal headers As Variant) As Boolean
    Dim c As Long

    For c = LBound(headers) To UBound(headers)
        If Trim$(CellText(arr, rowIndex, lo, CStr(headers(c)))) <> "" Then
            TableArrayRowHasAnyValue = True
            Exit Function
        End If
    Next c
End Function

Private Sub FillInventoryList(ByVal lo As ListObject, ByVal filterText As String)
    Dim arr As Variant
    Dim r As Long
    Dim rowVal As String
    Dim itemVal As String
    Dim uomVal As String
    Dim totalVal As String
    Dim locVal As String
    Dim descVal As String
    Dim itemCode As String
    Dim haystack As String

    mLstAssignInventory.Clear
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    arr = lo.DataBodyRange.Value
    filterText = LCase$(Trim$(filterText))
    For r = 1 To UBound(arr, 1)
        rowVal = CellText(arr, r, lo, "System_Key")
        itemVal = CellText(arr, r, lo, "ITEM")
        uomVal = CellText(arr, r, lo, "UOM")
        totalVal = CellText(arr, r, lo, "TOTAL INV")
        locVal = CellText(arr, r, lo, "LOCATION")
        descVal = CellText(arr, r, lo, "DESCRIPTION")
        itemCode = CellText(arr, r, lo, "ITEM_CODE")
        haystack = LCase$(rowVal & " " & itemVal & " " & descVal & " " & itemCode & " " & locVal)
        If filterText = "" Or InStr(1, haystack, filterText, vbTextCompare) > 0 Then
            mLstAssignInventory.AddItem rowVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 1) = itemVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 2) = uomVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 3) = totalVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 4) = locVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 5) = descVal
            mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 6) = itemCode
        End If
    Next r
End Sub

Private Sub FillInventoryListFromArray(ByVal values As Variant, Optional ByVal filterText As String = "")
    Dim r As Long
    Dim shown As Long
    Dim normalizedFilter As String

    mLstAssignInventory.Clear
    If IsEmpty(values) Then Exit Sub
    If Not IsArray(values) Then Exit Sub

    normalizedFilter = NormalizeInventorySearch(filterText)
    For r = LBound(values, 1) To UBound(values, 1)
        If Not InventoryRowMatchesSearch(values, r, normalizedFilter) Then GoTo NextInventoryRow
        shown = shown + 1
        mLstAssignInventory.AddItem NzStr(values(r, 1))
        If mLstAssignInventory.ColumnCount > 1 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 1) = NzStr(values(r, 2))
        If mLstAssignInventory.ColumnCount > 2 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 2) = NzStr(values(r, 3))
        If mLstAssignInventory.ColumnCount > 3 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 3) = NzStr(values(r, 4))
        If mLstAssignInventory.ColumnCount > 4 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 4) = NzStr(values(r, 5))
        If mLstAssignInventory.ColumnCount > 5 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 5) = NzStr(values(r, 6))
        If mLstAssignInventory.ColumnCount > 6 Then mLstAssignInventory.List(mLstAssignInventory.ListCount - 1, 6) = NzStr(values(r, 7))
        If shown >= ASSIGN_INVENTORY_MAX_VISIBLE Then Exit For
NextInventoryRow:
    Next r
End Sub

Private Function InventoryRowMatchesSearch(ByVal values As Variant, ByVal rowIndex As Long, ByVal normalizedFilter As String) As Boolean
    Dim haystack As String

    If normalizedFilter = "" Then
        InventoryRowMatchesSearch = True
        Exit Function
    End If

    haystack = NormalizeInventorySearch(NzStr(values(rowIndex, 1)) & " " & _
        NzStr(values(rowIndex, 2)) & " " & NzStr(values(rowIndex, 3)) & " " & _
        NzStr(values(rowIndex, 5)) & " " & NzStr(values(rowIndex, 6)) & " " & _
        NzStr(values(rowIndex, 7)))
    InventoryRowMatchesSearch = (InStr(1, haystack, normalizedFilter, vbTextCompare) > 0)
End Function

Private Function NormalizeInventorySearch(ByVal filterText As String) As String
    Dim textOut As String

    textOut = Trim$(filterText)
    If textOut = "" Then Exit Function
    textOut = Replace$(textOut, vbCr, " ")
    textOut = Replace$(textOut, vbLf, " ")
    textOut = Replace$(textOut, vbTab, " ")
    Do While InStr(textOut, "  ") > 0
        textOut = Replace$(textOut, "  ", " ")
    Loop
    NormalizeInventorySearch = LCase$(textOut)
End Function

Private Sub ResetInventoryCache()
    mInventoryRows = Empty
    mInventoryCacheLoaded = False
    mRunInventoryRows = Empty
    mRunInventoryCacheLoaded = False
End Sub

Private Function CellText(ByVal arr As Variant, ByVal rowIndex As Long, ByVal lo As ListObject, ByVal headerName As String) As String
    Dim colIndex As Long
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Function
    On Error Resume Next
    CellText = NzStr(arr(rowIndex, colIndex))
    On Error GoTo 0
End Function

Private Function ProductionTable(ByVal tableName As String) As ListObject
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then Exit Function
    If Not WorkbookHasSheet(wb, "Production") Then Exit Function
    Set ws = wb.Worksheets("Production")
    If ws Is Nothing Then Exit Function
    Set ProductionTable = mProduction.GetListObject(ws, tableName)
End Function

Private Function InventoryTable() As ListObject
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then Exit Function
    If Not WorkbookHasSheet(wb, "InventoryManagement") Then Exit Function
    Set ws = wb.Worksheets("InventoryManagement")
    If ws Is Nothing Then Exit Function
    Set InventoryTable = mProduction.GetListObject(ws, "invSys")
End Function

Private Function GenerateRecipeIdForOperatorWorkbook() As String
    Dim wb As Workbook

    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then Exit Function
    BindOperatorWorkbookForRun
    GenerateRecipeIdForOperatorWorkbook = NzStr(mProduction.GenerateRecipeId(wb))
End Function

Private Function ResolveOperatorWorkbook() As Workbook
    If IsUsableWorkbook(mOperatorWorkbook) Then
        Set ResolveOperatorWorkbook = mOperatorWorkbook
        Exit Function
    End If
End Function

Private Function RefreshProductionInventoryReadModel(ByRef reportOut As String) As Boolean
    On Error GoTo CleanFail

    Dim wb As Workbook
    Dim resultText As String
    Dim separatorAt As Long
    Dim resultCode As String

    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then
        reportOut = "Open a Production operator workbook before refreshing inventory."
        Exit Function
    End If

    BindOperatorWorkbookForRun
    resultText = NzStr(mProduction.RefreshProductionInventoryReadModelForWorkbookResult(wb))
    separatorAt = InStr(1, resultText, vbTab, vbBinaryCompare)
    If separatorAt > 0 Then
        resultCode = Left$(resultText, separatorAt - 1)
        reportOut = Mid$(resultText, separatorAt + 1)
    Else
        resultCode = resultText
        reportOut = resultText
    End If
    If Trim$(reportOut) = "" Then reportOut = IIf(resultCode = "OK", "OK", "Inventory refresh failed.")
    RefreshProductionInventoryReadModel = (StrComp(Trim$(resultCode), "OK", vbTextCompare) = 0)
    Exit Function

CleanFail:
    reportOut = "Production inventory refresh failed: " & Err.Description
End Function

Private Function IsUsableWorkbook(ByVal wb As Workbook) As Boolean
    On Error GoTo CleanFail
    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function
    IsUsableWorkbook = True
    Exit Function
CleanFail:
    IsUsableWorkbook = False
End Function

Private Function WorkbookHasSheet(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    WorkbookHasSheet = Not wb.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Sub WriteRecipeHeaderFromForm()
    Dim lo As ListObject
    Set lo = ProductionTable(TABLE_BUILDER_HEADER)
    If lo Is Nothing Then Exit Sub
    EnsureTableRow lo
    If Trim$(mTxtRecipeId.Text) = "" Then mTxtRecipeId.Text = GenerateRecipeIdForOperatorWorkbook()
    mTxtRecipeRowBudget.Text = NormalizeRecipeRowBudgetText(mTxtRecipeRowBudget.Text)
    SetFirstRowValue lo, "RECIPE_NAME", mTxtRecipeName.Text
    SetFirstRowValue lo, "RECIPE_ID", mTxtRecipeId.Text
    SetFirstRowValue lo, "DESCRIPTION", mTxtRecipeDescription.Text
    SetFirstRowValue lo, "ROW_BUDGET", CLng(Val(mTxtRecipeRowBudget.Text))
End Sub

Private Sub ClearRecipeBuilderForNew()
    Dim loHeader As ListObject
    Dim loLines As ListObject

    Set loHeader = ProductionTable(TABLE_BUILDER_HEADER)
    Set loLines = ProductionTable(TABLE_BUILDER_LINES)
    If Not loHeader Is Nothing Then
        If Not loHeader.DataBodyRange Is Nothing Then loHeader.DataBodyRange.ClearContents
    End If
    If Not loLines Is Nothing Then ClearListRows loLines
    mTxtRecipeName.Text = ""
    mTxtRecipeId.Text = GenerateRecipeIdForOperatorWorkbook()
    mTxtRecipeDescription.Text = ""
    mTxtRecipeRowBudget.Text = CStr(PRODUCTION_DEFAULT_ROW_BUDGET)
    ClearLineInputs
    RefreshBuilderLines
    ShowStatus "Ready for a new recipe."
End Sub

Private Sub ClearLineInputs()
    mTxtLineProcess.Text = ""
    If Not mCmbLineIo Is Nothing Then mCmbLineIo.ListIndex = 0
    mTxtLineIngredient.Text = ""
    mTxtLinePercent.Text = ""
    mTxtLineUom.Value = ""
    mTxtLineAmount.Text = ""
    ConfigureRecipeLineInputs
End Sub

Private Sub LoadSelectedBuilderLine()
    Dim idx As Long

    idx = mLstBuilderLines.ListIndex
    If idx < 0 Then Exit Sub
    mTxtLineProcess.Text = NzStr(mLstBuilderLines.List(idx, 0))
    SetLineIo NzStr(mLstBuilderLines.List(idx, 2))
    mTxtLineIngredient.Text = NzStr(mLstBuilderLines.List(idx, 3))
    mTxtLinePercent.Text = NzStr(mLstBuilderLines.List(idx, 4))
    RefreshRecipeUomCatalog NzStr(mLstBuilderLines.List(idx, 5))
    mTxtLineAmount.Text = NzStr(mLstBuilderLines.List(idx, 6))
End Sub

Private Sub SetLineIo(ByVal ioValue As String)
    Dim i As Long
    Dim v As String

    v = UCase$(Trim$(ioValue))
    If v = "MADE" Then v = "OUTPUT"
    For i = 0 To mCmbLineIo.ListCount - 1
        If StrComp(NzStr(mCmbLineIo.List(i)), v, vbTextCompare) = 0 Then
            mCmbLineIo.ListIndex = i
            ConfigureRecipeLineInputs
            Exit Sub
        End If
    Next i
    mCmbLineIo.ListIndex = 0
    ConfigureRecipeLineInputs
End Sub

Private Function LineIoValue() As String
    If mCmbLineIo.ListIndex >= 0 Then LineIoValue = UCase$(Trim$(NzStr(mCmbLineIo.Value)))
    If LineIoValue = "" Then LineIoValue = "USED"
End Function

Private Sub AddRecipeBuilderLine()
    Dim lo As ListObject
    Dim lr As ListRow

    If Trim$(mTxtLineIngredient.Text) = "" Then
        ShowStatus "Enter an ingredient or output name before adding a recipe line."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Then
        ShowStatus "RecipeBuilder table is missing."
        Exit Sub
    End If
    Set lr = lo.ListRows.Add(AlwaysInsert:=False)
    WriteRecipeLineToRow lo, lr.Index
    ResequenceRecipeBuilderLines lo
    RefreshBuilderLines
    ClearLineInputs
    ShowStatus "Recipe line added."
End Sub

Private Sub UpdateSelectedRecipeBuilderLine()
    If WriteSelectedRecipeBuilderLineFromForm(True) Then
        ShowStatus "Recipe line updated."
    End If
End Sub

Private Function WriteSelectedRecipeBuilderLineFromForm(Optional ByVal refreshList As Boolean = True) As Boolean
    Dim lo As ListObject
    Dim idx As Long
    Dim tableRow As Long

    idx = mLstBuilderLines.ListIndex
    If idx < 0 Then
        If refreshList Then ShowStatus "Select a recipe line to update."
        Exit Function
    End If
    tableRow = BuilderTableRowForListIndex(idx)
    If tableRow <= 0 Then Exit Function
    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    If tableRow > lo.ListRows.Count Then Exit Function
    WriteRecipeLineToRow lo, tableRow
    If refreshList Then
        RefreshBuilderLines
        If idx < mLstBuilderLines.ListCount Then mLstBuilderLines.ListIndex = idx
    End If
    WriteSelectedRecipeBuilderLineFromForm = True
End Function

Private Sub RemoveSelectedRecipeBuilderLine()
    Dim lo As ListObject
    Dim idx As Long
    Dim tableRow As Long

    idx = mLstBuilderLines.ListIndex
    If idx < 0 Then
        ShowStatus "Select a recipe line to remove."
        Exit Sub
    End If
    tableRow = BuilderTableRowForListIndex(idx)
    If tableRow <= 0 Then Exit Sub
    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    If tableRow <= lo.ListRows.Count Then lo.ListRows(tableRow).Delete
    ResequenceRecipeBuilderLines lo
    RefreshBuilderLines
    ClearLineInputs
    ShowStatus "Recipe line removed."
End Sub

Private Sub WriteRecipeLineToRow(ByVal lo As ListObject, ByVal rowIndex As Long)
    Dim ioType As String

    If lo Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub
    If Not EnsureTableRows(lo, rowIndex) Then Exit Sub
    ioType = LineIoValue()
    SetCellByHeader lo, rowIndex, "PROCESS", mTxtLineProcess.Text
    SetCellByHeader lo, rowIndex, "INPUT/OUTPUT", ioType
    SetCellByHeader lo, rowIndex, "INGREDIENT", mTxtLineIngredient.Text
    If ioType = "INSTRUCTION" Then
        SetCellByHeader lo, rowIndex, "PERCENT", vbNullString
        SetCellByHeader lo, rowIndex, "UOM", vbNullString
        SetCellByHeader lo, rowIndex, "AMOUNT", vbNullString
        SetCellByHeader lo, rowIndex, "INGREDIENT_ID", vbNullString
    Else
        SetCellByHeader lo, rowIndex, "PERCENT", mTxtLinePercent.Text
        SetCellByHeader lo, rowIndex, "UOM", CStr(mTxtLineUom.Value)
        SetCellByHeader lo, rowIndex, "AMOUNT", mTxtLineAmount.Text
        If Trim$(CellByHeader(lo, rowIndex, "INGREDIENT_ID")) = "" Then SetCellByHeader lo, rowIndex, "INGREDIENT_ID", BuildFormGuid()
    End If
    If Trim$(CellByHeader(lo, rowIndex, "RECIPE_LIST_ROW")) = "" Then SetCellByHeader lo, rowIndex, "RECIPE_LIST_ROW", rowIndex
    If Trim$(CellByHeader(lo, rowIndex, "GUID")) = "" Then SetCellByHeader lo, rowIndex, "GUID", BuildFormGuid()
End Sub

Private Sub ConfigureRecipeLineInputs()
    Dim acceptsQuantity As Boolean

    acceptsQuantity = (LineIoValue() <> "INSTRUCTION")
    If Not mTxtLinePercent Is Nothing Then mTxtLinePercent.Enabled = acceptsQuantity
    If Not mTxtLineUom Is Nothing Then mTxtLineUom.Enabled = acceptsQuantity
    If Not mBtnLineUomAdd Is Nothing Then mBtnLineUomAdd.Enabled = acceptsQuantity
    If Not mTxtLineAmount Is Nothing Then mTxtLineAmount.Enabled = acceptsQuantity
End Sub

Private Function MoveSelectedRecipeBuilderLine(ByVal moveDirection As Long, _
                                               Optional ByVal showMessages As Boolean = True) As Boolean
    Dim lo As ListObject
    Dim selectedRow As Long
    Dim targetRow As Long
    Dim selectedListIndex As Long
    Dim targetListIndex As Long
    Dim selectedValues() As Variant
    Dim targetValues() As Variant
    Dim columnIndex As Long
    Dim columnCount As Long

    If moveDirection <> -1 And moveDirection <> 1 Then Exit Function
    If mLstBuilderLines Is Nothing Then Exit Function
    If mLstBuilderLines.ListIndex < 0 Then
        If showMessages Then ShowStatus "Select a recipe line to move."
        Exit Function
    End If

    Set lo = ProductionTable(TABLE_BUILDER_LINES)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    selectedListIndex = mLstBuilderLines.ListIndex
    targetListIndex = selectedListIndex + moveDirection
    selectedRow = BuilderTableRowForListIndex(selectedListIndex)
    targetRow = BuilderTableRowForListIndex(targetListIndex)
    If selectedRow <= 0 Or targetRow <= 0 Then
        If showMessages Then ShowStatus IIf(moveDirection < 0, "The selected line is already first.", "The selected line is already last.")
        Exit Function
    End If

    columnCount = lo.ListColumns.Count
    ReDim selectedValues(1 To columnCount)
    ReDim targetValues(1 To columnCount)
    For columnIndex = 1 To columnCount
        selectedValues(columnIndex) = lo.DataBodyRange.Cells(selectedRow, columnIndex).Value2
        targetValues(columnIndex) = lo.DataBodyRange.Cells(targetRow, columnIndex).Value2
    Next columnIndex
    For columnIndex = 1 To columnCount
        lo.DataBodyRange.Cells(selectedRow, columnIndex).Value2 = targetValues(columnIndex)
        lo.DataBodyRange.Cells(targetRow, columnIndex).Value2 = selectedValues(columnIndex)
    Next columnIndex
    ResequenceRecipeBuilderLines lo
    RefreshBuilderLines
    If targetListIndex >= 0 And targetListIndex < mLstBuilderLines.ListCount Then _
        mLstBuilderLines.ListIndex = targetListIndex
    LoadSelectedBuilderLine
    If showMessages Then ShowStatus "Recipe line moved " & IIf(moveDirection < 0, "up.", "down.")
    MoveSelectedRecipeBuilderLine = True
End Function

Private Sub ResequenceRecipeBuilderLines(ByVal lo As ListObject)
    Dim rowIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For rowIndex = 1 To lo.ListRows.Count
        SetCellByHeader lo, rowIndex, "RECIPE_LIST_ROW", rowIndex
    Next rowIndex
End Sub

Private Sub SelectAssignmentRecipeFromList()
    Dim idx As Long
    Dim recipeId As String
    Dim lo As ListObject

    idx = mLstAssignRecipes.ListIndex
    If idx < 0 Then
        ShowStatus "Select a recipe first."
        Exit Sub
    End If
    recipeId = NzStr(mLstAssignRecipes.List(idx, 0))
    Set lo = ProductionTable(TABLE_ASSIGN_RECIPE)
    If lo Is Nothing Then
        ShowStatus "IP_ChooseRecipe table is missing."
        Exit Sub
    End If
    EnsureTableRow lo
    SetFirstRowValue lo, "RECIPE_ID", recipeId
    SetFirstRowValue lo, "RECIPE_NAME", NzStr(mLstAssignRecipes.List(idx, 1))
    SetFirstRowValue lo, "DESCRIPTION", NzStr(mLstAssignRecipes.List(idx, 2))
    SetFirstRowValue lo, "GUID", recipeId
    BindOperatorWorkbookForRun
    mProduction.HandlePaletteRecipeSelected recipeId
    RefreshAssignmentState
    ShowStatus "Selected assignment recipe: " & NzStr(mLstAssignRecipes.List(idx, 1))
End Sub

Private Sub SelectAssignmentIngredientFromList()
    Dim idx As Long
    Dim recipeId As String
    Dim ingredientId As String
    Dim ioVal As String
    Dim lo As ListObject

    idx = mLstAssignIngredients.ListIndex
    If idx < 0 Then
        ShowStatus "Select an ingredient first."
        Exit Sub
    End If
    ioVal = UCase$(Trim$(NzStr(mLstAssignIngredients.List(idx, 4))))
    If ioVal = "OUTPUT" Or ioVal = "MADE" Then
        ClearAssignmentIngredientSelection
        ShowStatus "Outputs do not accept ingredients. Assign acceptable inventory only to recipe inputs."
        Exit Sub
    End If
    BindOperatorWorkbookForRun
    recipeId = NzStr(mProduction.GetPaletteRecipeId())
    ingredientId = NzStr(mLstAssignIngredients.List(idx, 0))
    If recipeId = "" Or ingredientId = "" Then
        ShowStatus "Select a recipe and ingredient first."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_ASSIGN_INGREDIENT)
    If lo Is Nothing Then
        ShowStatus "IP_ChooseIngredient table is missing."
        Exit Sub
    End If
    EnsureTableRow lo
    SetFirstRowValue lo, "RECIPE_ID", recipeId
    SetFirstRowValue lo, "INGREDIENT_ID", ingredientId
    SetFirstRowValue lo, "INGREDIENT", NzStr(mLstAssignIngredients.List(idx, 1))
    SetFirstRowValue lo, "UOM", NzStr(mLstAssignIngredients.List(idx, 2))
    SetFirstRowValue lo, "PROCESS", NzStr(mLstAssignIngredients.List(idx, 3))
    SetFirstRowValue lo, "DESCRIPTION", ""
    SetFirstRowValue lo, "QUANTITY", NzStr(mLstAssignIngredients.List(idx, 5))
    BindOperatorWorkbookForRun
    mProduction.HandlePaletteIngredientSelected recipeId, ingredientId
    RefreshAllowedItems
    ShowStatus "Selected ingredient: " & NzStr(mLstAssignIngredients.List(idx, 1))
End Sub

Private Sub AddSelectedInventoryToAllowed()
    On Error GoTo FailAdd

    Dim idx As Long
    Dim lo As ListObject
    Dim lr As ListRow
    Dim recipeId As String
    Dim ingredientId As String

    idx = mLstAssignInventory.ListIndex
    If idx < 0 Then
        ShowStatus "Select an inventory row first."
        Exit Sub
    End If
    BindOperatorWorkbookForRun
    recipeId = NzStr(mProduction.GetPaletteRecipeId())
    ingredientId = NzStr(mProduction.GetPaletteIngredientId())
    If recipeId = "" Or ingredientId = "" Then
        ShowStatus "Select a recipe and ingredient before adding acceptable inventory."
        Exit Sub
    End If
    If mProduction.PaletteIngredientIsOutput(recipeId, ingredientId) Then
        ClearAssignmentIngredientSelection
        ShowStatus "Outputs do not accept ingredients. Assign acceptable inventory only to recipe inputs."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_ASSIGN_ITEM)
    If lo Is Nothing Then
        ShowStatus "IP_ChooseItem table is missing."
        Exit Sub
    End If
    Set lr = WritableAssignmentItemRow(lo)
    If lr Is Nothing Then
        ShowStatus "Acceptable inventory could not be staged because the operator table has no available row."
        Exit Sub
    End If
    SetRowValue lr, lo, "System_Key", NzStr(mLstAssignInventory.List(idx, 0))
    SetRowValue lr, lo, "ITEM_CODE", NzStr(mLstAssignInventory.List(idx, 6))
    SetRowValue lr, lo, "ITEMS", NzStr(mLstAssignInventory.List(idx, 1))
    SetRowValue lr, lo, "UOM", NzStr(mLstAssignInventory.List(idx, 2))
    SetRowValue lr, lo, "DESCRIPTION", NzStr(mLstAssignInventory.List(idx, 5))
    SetRowValue lr, lo, "RECIPE_ID", recipeId
    SetRowValue lr, lo, "INGREDIENT_ID", ingredientId
    RefreshAllowedItems
    ShowStatus "Added acceptable ingredient row " & NzStr(mLstAssignInventory.List(idx, 0)) & "."
    Exit Sub

FailAdd:
    ShowStatus "Add Acceptable failed: " & Err.Description
End Sub

Private Function WritableAssignmentItemRow(ByVal lo As ListObject) As ListRow
    Dim r As Long
    Dim priorCount As Long
    Dim insertRow As Long

    If lo Is Nothing Then Exit Function
    If Not lo.DataBodyRange Is Nothing Then
        For r = 1 To lo.ListRows.Count
            If AssignmentItemRowIsBlank(lo, r) Then
                Set WritableAssignmentItemRow = lo.ListRows(r)
                Exit Function
            End If
        Next r
    End If

    On Error Resume Next
    Set WritableAssignmentItemRow = lo.ListRows.Add(AlwaysInsert:=False)
    On Error GoTo 0
    If Not WritableAssignmentItemRow Is Nothing Then Exit Function

    On Error GoTo FailAcquire
    priorCount = lo.ListRows.Count
    insertRow = lo.Range.Row + lo.Range.Rows.Count
    lo.Parent.Rows(insertRow).Insert Shift:=xlDown
    If lo.ListRows.Count > priorCount Then
        Set WritableAssignmentItemRow = lo.ListRows(lo.ListRows.Count)
    Else
        Set WritableAssignmentItemRow = lo.ListRows.Add(AlwaysInsert:=False)
    End If
    Exit Function

FailAcquire:
    Set WritableAssignmentItemRow = Nothing
End Function

Private Function AssignmentItemRowIsBlank(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    AssignmentItemRowIsBlank = _
        Trim$(CellByHeader(lo, rowIndex, "System_Key")) = "" _
        And Trim$(CellByHeader(lo, rowIndex, "ITEM_CODE")) = "" _
        And Trim$(CellByHeader(lo, rowIndex, "ITEMS")) = "" _
        And Trim$(CellByHeader(lo, rowIndex, "RECIPE_ID")) = "" _
        And Trim$(CellByHeader(lo, rowIndex, "INGREDIENT_ID")) = ""
End Function

Private Sub ClearAssignmentIngredientSelection()
    Dim loIng As ListObject
    Dim loItems As ListObject

    Set loIng = ProductionTable(TABLE_ASSIGN_INGREDIENT)
    Set loItems = ProductionTable(TABLE_ASSIGN_ITEM)
    ClearTableContentsKeepBlank loIng
    ClearTableContentsKeepBlank loItems
    RefreshAllowedItems
End Sub

Private Sub RemoveSelectedAllowedRow()
    Dim idx As Long
    Dim lo As ListObject

    idx = mLstAssignAllowed.ListIndex
    If idx < 0 Then
        ShowStatus "Select an acceptable item row to remove."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_ASSIGN_ITEM)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If idx + 1 <= lo.ListRows.Count Then lo.ListRows(idx + 1).Delete
    RefreshAllowedItems
    ShowStatus "Removed acceptable item row."
End Sub

Private Sub LoadSelectedRecipeIntoBuilder()
    Dim idx As Long
    idx = mLstBuilderRecipes.ListIndex
    If idx < 0 Then
        ShowStatus "Select a saved recipe first."
        Exit Sub
    End If
    BindOperatorWorkbookForRun
    mProduction.LoadRecipeFromRecipes NzStr(mLstBuilderRecipes.List(idx, 0))
    RefreshBuilderHeader
    RefreshBuilderLines
    ShowStatus "Loaded recipe into builder: " & NzStr(mLstBuilderRecipes.List(idx, 1))
End Sub

Private Sub LoadSelectedRecipeIntoLoader()
    Dim idx As Long
    Dim prepared As Variant
    Dim scalePercent As Double
    Dim scaleReport As String

    idx = mLstLoaderRecipes.ListIndex
    If idx < 0 Then
        ShowStatus "Select a recipe first."
        Exit Sub
    End If
    If Not TryParseBatchScalePercent( _
            NzStr(mTxtBatchScalePercent.Text), scalePercent, scaleReport) Then
        ShowStatus scaleReport
        Exit Sub
    End If
    If Not mRunTreeCollapsed Is Nothing Then mRunTreeCollapsed.RemoveAll
    BindOperatorWorkbookForRun
    mProduction.LoadRecipeChooser NzStr(mLstLoaderRecipes.List(idx, 0))
    If Not ApplyBatchScaleToRunList(scalePercent, scaleReport) Then
        ShowStatus scaleReport
        Exit Sub
    End If
    prepared = mProduction.PrepareProductionOutputForCurrentRecipe()
    RefreshLoaderState
    RefreshManagerState
    If CBool(prepared) Then
        ShowStatus "Loaded recipe at " & FormatRunNumber(scalePercent) & _
            "% and prepared output rows: " & NzStr(mLstLoaderRecipes.List(idx, 1))
    Else
        ShowStatus "Loaded recipe, but no output rows were prepared: " & NzStr(mLstLoaderRecipes.List(idx, 1))
    End If
End Sub

Private Sub PrepareProductionOutput()
    Dim result As Variant

    BindOperatorWorkbookForRun
    result = mProduction.PrepareProductionOutputForCurrentRecipe()
    RefreshLoaderState
    RefreshManagerState
    If CBool(result) Then
        ShowStatus "Production output prepared."
    Else
        ShowStatus "Production output was not prepared. Load a recipe with OUTPUT lines first."
    End If
End Sub

Private Sub LoadSelectedProductionOutput()
    Dim idx As Long
    Dim lo As ListObject
    Dim tableRowNumber As Long

    idx = mLstManagerOutput.ListIndex
    If idx < 0 Then Exit Sub
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    tableRowNumber = SelectedProductionOutputTableRow()
    If tableRowNumber = 0 Then Exit Sub
    mTxtOutputReal.Text = CellByHeader(lo, tableRowNumber, "REAL OUTPUT")
End Sub

Private Sub ApplySelectedProductionOutput()
    Dim idx As Long
    Dim lo As ListObject
    Dim batchVal As String
    Dim tableRowNumber As Long

    idx = mLstManagerOutput.ListIndex
    If idx < 0 Then
        ShowStatus "Select a ProductionOutput row first."
        Exit Sub
    End If
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    tableRowNumber = SelectedProductionOutputTableRow()
    If tableRowNumber = 0 Then
        ShowStatus "The selected Production Output row could not be resolved."
        Exit Sub
    End If
    batchVal = CStr(NextOutputBatchNumberForListIndex(idx))
    SetCellByHeader lo, tableRowNumber, "REAL OUTPUT", mTxtOutputReal.Text
    SetCellByHeader lo, tableRowNumber, "BATCH", batchVal
    RefreshManagerState
    If idx < mLstManagerOutput.ListCount Then mLstManagerOutput.ListIndex = idx
    ShowStatus "Actual Output staged for batch " & batchVal & "."
End Sub

Private Function SelectedProductionOutputTableRow() As Long
    Dim idx As Long
    Dim lo As ListObject

    If mLstManagerOutput Is Nothing Then Exit Function
    idx = mLstManagerOutput.ListIndex
    If idx < 0 Or idx >= mLstManagerOutput.ListCount Then Exit Function
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    SelectedProductionOutputTableRow = FindProductionOutputTableRow( _
        lo, _
        NzStr(mLstManagerOutput.List(idx, 8)), _
        NzStr(mLstManagerOutput.List(idx, 0)), _
        NzStr(mLstManagerOutput.List(idx, 1)), _
        idx + 1)
End Function

Private Function FindProductionOutputTableRow(ByVal lo As ListObject, _
                                              ByVal rowVal As String, _
                                              ByVal procVal As String, _
                                              ByVal outputVal As String, _
                                              Optional ByVal fallbackRow As Long = 0) As Long
    Dim r As Long
    Dim cRow As Long
    Dim cProc As Long
    Dim cOutput As Long
    Dim wantedRow As String

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    cRow = ProductionColumnIndex(lo, "System_Key")
    cProc = ProductionColumnIndex(lo, "PROCESS")
    cOutput = ProductionColumnIndex(lo, "OUTPUT")
    wantedRow = NormalizeRunSystemKey(rowVal)

    If wantedRow <> "" And cRow > 0 Then
        For r = 1 To lo.ListRows.Count
            If NormalizeRunSystemKey(NzStr(lo.DataBodyRange.Cells(r, cRow).Value)) = wantedRow Then
                FindProductionOutputTableRow = r
                Exit Function
            End If
        Next r
    End If

    If cProc > 0 And cOutput > 0 Then
        For r = 1 To lo.ListRows.Count
            If StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cProc).Value)), Trim$(procVal), vbTextCompare) = 0 _
               And StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cOutput).Value)), Trim$(outputVal), vbTextCompare) = 0 Then
                FindProductionOutputTableRow = r
                Exit Function
            End If
        Next r
    End If

    If fallbackRow >= 1 And fallbackRow <= lo.ListRows.Count Then
        FindProductionOutputTableRow = fallbackRow
    End If
End Function

Private Function NextOutputBatchNumberForListIndex(ByVal outputIndex As Long) As Long
    Dim lastQty As Double
    Dim totalQty As Double
    Dim maxBatch As Long
    Dim loggedCount As Long

    If mLstManagerOutput Is Nothing Then Exit Function
    If outputIndex < 0 Or outputIndex >= mLstManagerOutput.ListCount Then Exit Function
    LoggedOutputStats NzStr(mLstManagerOutput.List(outputIndex, 8)), _
                      NzStr(mLstManagerOutput.List(outputIndex, 0)), _
                      NzStr(mLstManagerOutput.List(outputIndex, 1)), _
                      lastQty, totalQty, maxBatch, loggedCount
    NextOutputBatchNumberForListIndex = maxBatch + 1
End Function

Private Sub LoggedOutputStats(ByVal rowVal As String, ByVal procVal As String, ByVal outputVal As String, _
                              ByRef lastQty As Double, ByRef totalQty As Double, _
                              ByRef maxBatch As Long, ByRef loggedCount As Long)
    Dim loLog As ListObject
    Dim r As Long
    Dim cReal As Long
    Dim cBatch As Long
    Dim cTime As Long
    Dim cProc As Long
    Dim cItem As Long
    Dim cOutput As Long
    Dim cRow As Long
    Dim realVal As Double
    Dim batchVal As Long
    Dim timeVal As Variant
    Dim latestTime As Date
    Dim hasLatestTime As Boolean

    lastQty = 0
    totalQty = 0
    maxBatch = 0
    loggedCount = 0
    Set loLog = ProductionLogTable()
    If loLog Is Nothing Then Exit Sub
    If loLog.DataBodyRange Is Nothing Then Exit Sub

    cReal = ProductionColumnIndex(loLog, "REAL OUTPUT")
    If cReal = 0 Then Exit Sub
    cBatch = ProductionColumnIndex(loLog, "BATCH")
    cTime = ProductionColumnIndex(loLog, "TIMESTAMP")
    cProc = ProductionColumnIndex(loLog, "PROCESS")
    cItem = ProductionColumnIndex(loLog, "ITEM")
    cOutput = ProductionColumnIndex(loLog, "OUTPUT")
    cRow = ProductionColumnIndex(loLog, "System_Key")

    For r = 1 To loLog.ListRows.Count
        If Not ProductionLogRowMatchesOutput(loLog, r, rowVal, procVal, outputVal, cRow, cProc, cItem, cOutput) Then GoTo NextRow
        If IsNumeric(loLog.DataBodyRange.Cells(r, cReal).Value) Then
            realVal = CDbl(loLog.DataBodyRange.Cells(r, cReal).Value)
            totalQty = totalQty + realVal
            loggedCount = loggedCount + 1
            If cBatch > 0 Then
                batchVal = CLng(Val(loLog.DataBodyRange.Cells(r, cBatch).Value))
                If batchVal > maxBatch Then maxBatch = batchVal
            End If
            If cTime > 0 Then timeVal = loLog.DataBodyRange.Cells(r, cTime).Value Else timeVal = Empty
            If IsDate(timeVal) Then
                If Not hasLatestTime Or CDate(timeVal) >= latestTime Then
                    latestTime = CDate(timeVal)
                    hasLatestTime = True
                    lastQty = realVal
                End If
            ElseIf Not hasLatestTime Then
                lastQty = realVal
            End If
        End If
NextRow:
    Next r
End Sub

Private Function ProductionLogRowMatchesOutput(ByVal loLog As ListObject, ByVal rowIndex As Long, _
                                               ByVal rowVal As String, ByVal procVal As String, ByVal outputVal As String, _
                                               ByVal cRow As Long, ByVal cProc As Long, ByVal cItem As Long, ByVal cOutput As Long) As Boolean
    Dim logRowVal As String
    Dim logProc As String
    Dim logOutput As String

    If loLog Is Nothing Then Exit Function
    If loLog.DataBodyRange Is Nothing Then Exit Function
    If cRow > 0 Then
        logRowVal = NormalizeRunSystemKey(NzStr(loLog.DataBodyRange.Cells(rowIndex, cRow).Value))
        If NormalizeRunSystemKey(rowVal) <> "" And logRowVal = NormalizeRunSystemKey(rowVal) Then
            ProductionLogRowMatchesOutput = True
            Exit Function
        End If
    End If

    If cProc > 0 Then logProc = Trim$(NzStr(loLog.DataBodyRange.Cells(rowIndex, cProc).Value))
    If cItem > 0 Then logOutput = Trim$(NzStr(loLog.DataBodyRange.Cells(rowIndex, cItem).Value))
    If logOutput = "" And cOutput > 0 Then logOutput = Trim$(NzStr(loLog.DataBodyRange.Cells(rowIndex, cOutput).Value))
    ProductionLogRowMatchesOutput = (StrComp(logProc, procVal, vbTextCompare) = 0 _
        And StrComp(logOutput, outputVal, vbTextCompare) = 0)
End Function

Private Function ProductionLogTable() As ListObject
    Dim ws As Worksheet

    BindOperatorWorkbookForRun
    Set ws = mProduction.SheetExists("ProductionLog")
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set ProductionLogTable = ws.ListObjects("ProductionLog")
    If ProductionLogTable Is Nothing Then Set ProductionLogTable = ws.ListObjects("Table46")
    On Error GoTo 0
    If Not ProductionLogTable Is Nothing Then Exit Function

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If ProductionColumnIndex(lo, "PROCESS") > 0 _
           And ProductionColumnIndex(lo, "REAL OUTPUT") > 0 _
           And ProductionColumnIndex(lo, "BATCH") > 0 Then
            Set ProductionLogTable = lo
            Exit Function
        End If
    Next lo
End Function

Private Sub CheckInProductionRun()
    Dim usedPayloadJson As String
    Dim stagedTotal As Double
    Dim reusableReport As String
    Dim checkInStage As String

    On Error GoTo FailCheckIn
    checkInStage = "ReusableState"
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        If Not SyncReusableRunBatchNote(reusableReport) Then
            ShowStatus reusableReport
            Exit Sub
        End If
        If ActiveRunProcess() <> "" Then
            If modProductionReusableRun.CheckInReusableProcess(ActiveRunProcess(), _
                    ActiveRunLocation(), reusableReport) Then
                RefreshReusableRunControls False
            End If
        ElseIf modProductionReusableRun.CheckInReusableRun(ActiveRunLocation(), reusableReport) Then
            RefreshReusableRunControls False
        End If
        ShowStatus reusableReport
        Exit Sub
    End If

    checkInStage = "LoaderSelection"
    If mLstLoaderLines.ListIndex >= 0 Then
        mLstLoaderLines.ListIndex = -1
        RefreshRunPaletteState
    End If
    checkInStage = "PalettePresence"
    If mLstRunPalette Is Nothing Or mLstRunPalette.ListCount = 0 Then
        ShowStatus "Load a recipe and choose acceptable inventory before checking inventory into Production."
        Exit Sub
    End If
    checkInStage = "LocationCleanup"
    ClearMismatchedRunLocationAllocations
    checkInStage = "AllocationCompleteness"
    If Not ValidateRunAllocationsComplete() Then Exit Sub
    checkInStage = "AllocationLocations"
    If Not ValidateRunAllocationLocations() Then Exit Sub

    checkInStage = "BuildPayload"
    usedPayloadJson = BuildRunUsedPayloadJson(stagedTotal)
    If stagedTotal <= 0 Then
        ShowStatus "No inventory was checked in. Enter allocation quantities first."
        Exit Sub
    End If
    checkInStage = "WriteCheckRows"
    If Not WriteProductionCheckRowsFromRunPalette() Then
        ShowStatus "Check In failed. The Inventory Check list could not be updated. " & _
            "Refresh Production Run before completing. " & mCheckRowsDiagnostic
        Exit Sub
    End If

    checkInStage = "RefreshManager"
    RefreshManagerState
    ShowStatus "Checked in " & FormatRunNumber(stagedTotal) & " units to Production. Complete Run will consume these checked-in quantities."
    Exit Sub

FailCheckIn:
    ShowStatus "Check In failed at " & checkInStage & ": " & CStr(Err.Number) & _
        " - " & Err.Description
End Sub

Private Sub RefreshConnectionUomCatalog(Optional ByVal selectedUom As String = "")
    Dim uoms As Variant
    Dim idx As Long

    If mCmbConnectionUom Is Nothing Then Exit Sub
    If selectedUom = "" Then selectedUom = ComboText(mCmbConnectionUom)
    mCmbConnectionUom.Clear
    uoms = modUomSettings.GetConfiguredUoms()
    If IsArray(uoms) Then
        For idx = LBound(uoms) To UBound(uoms)
            mCmbConnectionUom.AddItem CStr(uoms(idx))
        Next idx
    End If
    SelectComboText mCmbConnectionUom, selectedUom
End Sub

Private Sub RefreshProcessOutputUomCatalog(Optional ByVal selectedUom As String = "")
    Dim uoms As Variant
    Dim idx As Long

    If mCmbProcessOutputUom Is Nothing Then Exit Sub
    If selectedUom = "" Then selectedUom = ComboText(mCmbProcessOutputUom)
    mCmbProcessOutputUom.Clear
    uoms = modUomSettings.GetConfiguredUoms()
    If IsArray(uoms) Then
        For idx = LBound(uoms) To UBound(uoms)
            mCmbProcessOutputUom.AddItem CStr(uoms(idx))
        Next idx
    End If
    SelectComboText mCmbProcessOutputUom, selectedUom
End Sub

Private Function OutputUomIsCatalogValue(ByVal uomName As String) As Boolean
    Dim idx As Long

    uomName = Trim$(uomName)
    If uomName = "" Then Exit Function
    For idx = 0 To mCmbProcessOutputUom.ListCount - 1
        If StrComp(NzStr(mCmbProcessOutputUom.List(idx)), uomName, vbTextCompare) = 0 Then
            OutputUomIsCatalogValue = True
            Exit Function
        End If
    Next idx
End Function

Private Sub CompleteProductionRun()
    Dim outputIndex As Long
    Dim outputRowNumber As Long
    Dim outputRowVal As String
    Dim outputProcess As String
    Dim outputName As String
    Dim processName As String
    Dim prepared As Variant
    Dim completionReport As String
    Dim completionResult As String
    Dim reportSeparator As Long
    Dim enteredRealOutput As String
    Dim lo As ListObject
    Dim reusableReport As String

    If modProductionReusableRun.ReusableRunIsLoaded() Then
        If Not SyncReusableRunBatchNote(reusableReport) Then
            ShowStatus reusableReport
            Exit Sub
        End If
        If Not StageSelectedReusableActualOutput(True) Then Exit Sub
        ShowPersistencePending "Saving the reusable Process run to the warehouse server..."
        If ActiveRunProcess() <> "" Then
            If modProductionReusableRun.CompleteReusableProcess(ActiveRunProcess(), _
                    ActiveRunLocation(), reusableReport) Then
                ResetInventoryCache
                RefreshReusableRunControls True
            End If
        ElseIf modProductionReusableRun.CompleteReusableRun(ActiveRunLocation(), reusableReport) Then
            ResetInventoryCache
            RefreshReusableRunControls True
        End If
        ShowStatus reusableReport
        Exit Sub
    End If

    If Not HasProductionCheckRows() Then
        ShowStatus "Check inventory into Production before completing the run."
        Exit Sub
    End If
    If Not ValidateRunAllocationLocations() Then Exit Sub

    enteredRealOutput = Trim$(mTxtOutputReal.Text)
    processName = ActiveRunProcess()
    If mLstManagerOutput.ListCount = 0 Then
        BindOperatorWorkbookForRun
        prepared = mProduction.PrepareProductionOutputForCurrentRecipe()
        RefreshManagerState
        If Not CBool(prepared) Then
            ShowStatus "No production output rows were prepared. Confirm the recipe has an OUTPUT line."
            Exit Sub
        End If
    End If
    If processName <> "" Then
        If Not SelectProductionOutputForProcess(processName) Then
            ShowStatus "No Production Output row matches process '" & processName & "'."
            Exit Sub
        End If
        If Trim$(mTxtOutputReal.Text) = "" And enteredRealOutput <> "" Then mTxtOutputReal.Text = enteredRealOutput
    End If

    outputIndex = mLstManagerOutput.ListIndex
    If outputIndex < 0 And mLstManagerOutput.ListCount = 1 Then
        mLstManagerOutput.ListIndex = 0
        LoadSelectedProductionOutput
        If Trim$(mTxtOutputReal.Text) = "" And enteredRealOutput <> "" Then mTxtOutputReal.Text = enteredRealOutput
        outputIndex = 0
    End If
    If outputIndex < 0 Then
        ShowStatus "Select the Production Output row for this run."
        Exit Sub
    End If
    If Trim$(mTxtOutputReal.Text) = "" Then
        ShowStatus "Enter the Actual Output quantity before completing the run."
        Exit Sub
    End If
    If Not IsNumeric(mTxtOutputReal.Text) Or CDbl(mTxtOutputReal.Text) <= 0 Then
        ShowStatus "Actual Output must be a number greater than zero."
        Exit Sub
    End If

    outputRowNumber = SelectedProductionOutputTableRow()
    If outputRowNumber = 0 Then
        ShowStatus "The selected Production Output row could not be resolved. Refresh Production Run and select the output again."
        Exit Sub
    End If
    outputRowVal = NzStr(mLstManagerOutput.List(outputIndex, 8))
    outputProcess = NzStr(mLstManagerOutput.List(outputIndex, 0))
    outputName = NzStr(mLstManagerOutput.List(outputIndex, 1))

    ApplySelectedProductionOutput
    BindOperatorWorkbookForRun
    ShowPersistencePending "Saving the completed production run to the warehouse server..."
    completionResult = CStr(mProduction.CompleteProductionRunAfterCheckInForOutputResult(outputRowNumber))
    reportSeparator = InStr(1, completionResult, vbTab, vbBinaryCompare)
    If reportSeparator > 0 Then
        completionReport = Mid$(completionResult, reportSeparator + 1)
        completionResult = Left$(completionResult, reportSeparator - 1)
    Else
        completionReport = completionResult
    End If
    If StrComp(completionResult, "OK", vbTextCompare) <> 0 Then
        If Trim$(completionReport) = "" Then completionReport = "Complete Run failed."
        ShowStatus completionReport
        MsgBox completionReport, vbExclamation, "Production Complete Run"
        Exit Sub
    End If
    ResetInventoryCache
    RefreshLoaderState
    RefreshManagerState
    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    outputRowNumber = FindProductionOutputTableRow(lo, outputRowVal, outputProcess, outputName, outputRowNumber)
    ClearProductionOutputEntry outputRowNumber
    mTxtOutputReal.Text = ""
    RefreshManagerState
    ShowStatus "Production run completed. Checked-in inventory was consumed, Actual Output was added to inventory, and the batch was logged." & IIf(Trim$(completionReport) <> "", " " & completionReport, "")
End Sub

Private Sub ClearProductionOutputEntry(ByVal outputRowNumber As Long)
    Dim lo As ListObject

    Set lo = ProductionTable(TABLE_MANAGER_OUTPUT)
    If lo Is Nothing Then Exit Sub
    If outputRowNumber < 1 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If outputRowNumber > lo.ListRows.Count Then Exit Sub
    SetCellByHeader lo, outputRowNumber, "REAL OUTPUT", ""
    SetCellByHeader lo, outputRowNumber, "BATCH", ""
End Sub

Private Function HasProductionCheckRows() As Boolean
    Dim lo As ListObject

    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    HasProductionCheckRows = (lo.ListRows.Count > 0 And Trim$(FirstNonBlankCheckRow(lo)) <> "")
End Function

Private Sub ClearProductionCheckRows()
    Dim lo As ListObject

    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If lo Is Nothing Then Exit Sub
    ClearTableContentsKeepRows lo
    If Not mLstManagerCheck Is Nothing Then mLstManagerCheck.Clear
End Sub

Private Function FirstNonBlankCheckRow(ByVal lo As ListObject) As String
    Dim r As Long
    Dim cRow As Long
    Dim cItemCode As Long
    Dim cUsed As Long
    Dim rowIdentity As String

    cRow = ProductionColumnIndex(lo, "System_Key")
    cItemCode = ProductionColumnIndex(lo, "ITEM_CODE")
    cUsed = ProductionColumnIndex(lo, "USED")
    If (cRow = 0 And cItemCode = 0) Or cUsed = 0 Then Exit Function
    For r = 1 To lo.ListRows.Count
        rowIdentity = ""
        If cItemCode > 0 Then rowIdentity = Trim$(NzStr(lo.DataBodyRange.Cells(r, cItemCode).Value))
        If rowIdentity = "" And cRow > 0 Then rowIdentity = Trim$(NzStr(lo.DataBodyRange.Cells(r, cRow).Value))
        If rowIdentity <> "" And NzDblLocal(lo.DataBodyRange.Cells(r, cUsed).Value) > 0 Then
            FirstNonBlankCheckRow = rowIdentity
            Exit Function
        End If
    Next r
End Function

Private Function SelectProductionOutputForProcess(ByVal processName As String) As Boolean
    Dim i As Long

    If mLstManagerOutput Is Nothing Then Exit Function
    processName = Trim$(processName)
    If processName = "" Then Exit Function
    For i = 0 To mLstManagerOutput.ListCount - 1
        If StrComp(Trim$(NzStr(mLstManagerOutput.List(i, 0))), processName, vbTextCompare) = 0 Then
            mLstManagerOutput.ListIndex = i
            LoadSelectedProductionOutput
            SelectProductionOutputForProcess = True
            Exit Function
        End If
    Next i
End Function

Private Function ValidateRunAllocationsComplete() As Boolean
    Dim seen As Object
    Dim items As Variant
    Dim i As Long
    Dim groupKey As String
    Dim totalPct As Double

    If mLstRunPalette Is Nothing Then Exit Function
    Set seen = CreateObject("Scripting.Dictionary")
    For i = 0 To mLstRunPalette.ListCount - 1
        groupKey = RunAllocationGroupKeyFromList(mLstRunPalette, i)
        If groupKey <> "" Then
            If Not seen.Exists(LCase$(groupKey)) Then seen.Add LCase$(groupKey), groupKey
        End If
    Next i

    items = seen.Items
    For i = LBound(items) To UBound(items)
        groupKey = CStr(items(i))
        totalPct = RunAllocationPercentForGroup(groupKey)
        If totalPct < 99.999 Or totalPct > 100.001 Then
            ShowStatus "Cannot complete run. Ingredient '" & groupKey & "' is " & FormatRunNumber(totalPct) & "% allocated; it must be 100%."
            Exit Function
        End If
    Next i
    ValidateRunAllocationsComplete = (seen.Count > 0)
End Function

Private Function ValidateRunAllocationLocations() As Boolean
    Dim runLoc As String
    Dim i As Long
    Dim qtyVal As Double
    Dim itemVal As String
    Dim invLoc As String

    If mLstRunPalette Is Nothing Then Exit Function
    runLoc = ActiveRunLocation()
    If RunLocationChoiceRequired() And runLoc = "" Then
        ShowStatus "Choose a production run location before checking inventory into Production."
        Exit Function
    End If

    For i = 0 To mLstRunPalette.ListCount - 1
        If Not IsNumeric(NzStr(mLstRunPalette.List(i, 6))) Then GoTo NextChoice
        qtyVal = CDbl(NzStr(mLstRunPalette.List(i, 6)))
        If qtyVal <= 0 Then GoTo NextChoice
        invLoc = Trim$(NzStr(mLstRunPalette.List(i, 9)))
        If Not RunChoiceLocationAllowed(runLoc, invLoc, qtyVal) Then
            itemVal = Trim$(NzStr(mLstRunPalette.List(i, 4)))
            If itemVal = "" Then itemVal = "System_Key " & Trim$(NzStr(mLstRunPalette.List(i, 3)))
            ShowStatus "Cannot use " & itemVal & " from " & invLoc & ". Production run location is " & runLoc & "; use inventory at the production location."
            Exit Function
        End If
NextChoice:
    Next i

    ValidateRunAllocationLocations = True
End Function

Private Function ClearMismatchedRunLocationAllocations(Optional ByRef clearedCount As Long = 0) As Boolean
    Dim runLoc As String
    Dim i As Long
    Dim qtyVal As Double
    Dim splitVal As Double
    Dim hasPositiveAllocation As Boolean
    Dim invLoc As String

    clearedCount = 0
    If mLstRunPalette Is Nothing Then
        ClearMismatchedRunLocationAllocations = True
        Exit Function
    End If
    runLoc = ActiveRunLocation()
    If runLoc = "" Then
        ClearMismatchedRunLocationAllocations = True
        Exit Function
    End If

    For i = 0 To mLstRunPalette.ListCount - 1
        qtyVal = 0
        splitVal = 0
        If IsNumeric(NzStr(mLstRunPalette.List(i, 6))) Then qtyVal = CDbl(NzStr(mLstRunPalette.List(i, 6)))
        If IsNumeric(NzStr(mLstRunPalette.List(i, 5))) Then splitVal = CDbl(NzStr(mLstRunPalette.List(i, 5)))
        hasPositiveAllocation = (qtyVal > 0 Or splitVal > 0)
        If hasPositiveAllocation Then
            invLoc = Trim$(NzStr(mLstRunPalette.List(i, 9)))
            If Not RunChoiceLocationAllowed(runLoc, invLoc, IIf(qtyVal > 0, qtyVal, splitVal)) Then
                ClearRunAllocationForListRow mLstRunPalette, i
                clearedCount = clearedCount + 1
            End If
        End If
    Next i

    If clearedCount > 0 Then BuildRunTreeFromPaletteList
    ClearMismatchedRunLocationAllocations = True
End Function

Private Sub ClearRunAllocationForListRow(ByVal lst As MSForms.ListBox, ByVal rowIndex As Long)
    Dim tableName As String
    Dim tableRow As Long
    Dim lo As ListObject

    If lst Is Nothing Then Exit Sub
    If rowIndex < 0 Or rowIndex >= lst.ListCount Then Exit Sub

    tableName = NzStr(lst.List(rowIndex, 0))
    tableRow = CLng(Val(NzStr(lst.List(rowIndex, 1))))
    If tableName <> "" And tableRow >= 1 Then
        Set lo = ProductionTable(tableName)
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                If tableRow <= lo.ListRows.Count Then
                    SetCellByHeader lo, tableRow, "SPLIT %", ""
                    SetCellByHeader lo, tableRow, "QUANTITY", ""
                End If
            End If
        End If
    End If

    lst.List(rowIndex, 5) = ""
    lst.List(rowIndex, 6) = ""
    StoreRunAllocationOverride lst, rowIndex, "", ""
    SyncRunAllocationToPaletteList lst, rowIndex
End Sub

Private Function RunChoiceLocationAllowed(ByVal runLocation As String, ByVal inventoryLocation As String, ByVal qtyValue As Double) As Boolean
    runLocation = Trim$(runLocation)
    inventoryLocation = Trim$(inventoryLocation)
    If qtyValue <= 0 Then
        RunChoiceLocationAllowed = True
    ElseIf runLocation = "" Or inventoryLocation = "" Then
        RunChoiceLocationAllowed = False
    Else
        RunChoiceLocationAllowed = (StrComp(runLocation, inventoryLocation, vbTextCompare) = 0)
    End If
End Function

Private Function BuildRunUsedPayloadJson(ByRef stagedTotal As Double) As String
    Dim agg As Object
    Dim i As Long
    Dim rowVal As String
    Dim systemKey As String
    Dim itemCode As String
    Dim identityKey As String
    Dim qtyVal As Double
    Dim locVal As String
    Dim key As Variant
    Dim payloadItems As Collection
    Dim payloadItem As Object

    stagedTotal = 0
    Set agg = CreateObject("Scripting.Dictionary")
    For i = 0 To mLstRunPalette.ListCount - 1
        rowVal = Trim$(NzStr(mLstRunPalette.List(i, 3)))
        itemCode = Trim$(RunItemCodeFromList(mLstRunPalette, i))
        locVal = NzStr(mLstRunPalette.List(i, 9))
        systemKey = ResolveRunSystemKey(itemCode, NzStr(mLstRunPalette.List(i, 4)), locVal)
        If systemKey = "" Then
            ShowStatus "Cannot complete run. The selected inventory row has no immutable System_Key."
            Exit Function
        End If
        identityKey = "SYS|" & systemKey
        If Not IsNumeric(NzStr(mLstRunPalette.List(i, 6))) Then GoTo NextChoice
        qtyVal = CDbl(NzStr(mLstRunPalette.List(i, 6)))
        If qtyVal <= 0 Then GoTo NextChoice
        If RunChoiceWouldExceedInventory(i, qtyVal) Then
            ShowStatus "Cannot complete run. Inventory " & IIf(rowVal <> "", "System_Key " & rowVal, itemCode) & _
                       " requires " & FormatRunNumber(qtyVal) & " but only " & _
                       NzStr(mLstRunPalette.List(i, 8)) & " is available."
            Exit Function
        End If
        If agg.Exists(identityKey) Then
            Dim existingAgg As Variant
            existingAgg = agg(identityKey)
            agg(identityKey) = Array(CDbl(existingAgg(0)) + qtyVal, locVal, systemKey, itemCode)
        Else
            agg.Add identityKey, Array(qtyVal, locVal, systemKey, itemCode)
        End If
NextChoice:
    Next i

    If agg.Count = 0 Then Exit Function
    Set payloadItems = New Collection
    For Each key In agg.Keys
        qtyVal = CDbl(agg(key)(0))
        locVal = NzStr(agg(key)(1))
        Set payloadItem = CreateObject("Scripting.Dictionary")
        payloadItem.CompareMode = vbTextCompare
        payloadItem("System_Key") = NzStr(agg(key)(2))
        payloadItem("SKU") = NzStr(agg(key)(3))
        payloadItem("Qty") = qtyVal
        payloadItem("Location") = locVal
        payloadItem("IoType") = "USED"
        payloadItem("Note") = "Production run input"
        payloadItems.Add payloadItem
        stagedTotal = stagedTotal + qtyVal
    Next key
    BuildRunUsedPayloadJson = NzStr(modProductionJson.BuildJsonArray(payloadItems))
End Function

Private Function WriteProductionCheckRowsFromRunPalette() As Boolean
    Dim lo As ListObject
    Dim agg As Object
    Dim i As Long
    Dim rowVal As String
    Dim systemKey As String
    Dim itemCode As String
    Dim identityKey As String
    Dim qtyVal As Double
    Dim entry As Variant
    Dim key As Variant
    Dim outRow As Long
    Dim writeStage As String

    On Error GoTo FailWrite
    mCheckRowsDiagnostic = ""
    writeStage = "ResolveCheckTable"
    Set lo = ProductionTable(TABLE_MANAGER_CHECK)
    If lo Is Nothing Then Exit Function
    writeStage = "CreateAggregate"
    Set agg = CreateObject("Scripting.Dictionary")

    writeStage = "AggregatePalette"
    For i = 0 To mLstRunPalette.ListCount - 1
        rowVal = Trim$(NzStr(mLstRunPalette.List(i, 3)))
        itemCode = Trim$(RunItemCodeFromList(mLstRunPalette, i))
        writeStage = "ResolveSystemKey.Row" & CStr(i + 1)
        systemKey = ResolveRunSystemKey(itemCode, NzStr(mLstRunPalette.List(i, 4)), _
                                        NzStr(mLstRunPalette.List(i, 9)))
        If systemKey = "" Then Exit Function
        identityKey = "SYS|" & systemKey
        If Not IsNumeric(NzStr(mLstRunPalette.List(i, 6))) Then GoTo NextChoice
        qtyVal = CDbl(NzStr(mLstRunPalette.List(i, 6)))
        If qtyVal <= 0 Then GoTo NextChoice
        If agg.Exists(identityKey) Then
            entry = agg(identityKey)
            entry(5) = CDbl(entry(5)) + qtyVal
            agg(identityKey) = entry
        Else
            agg.Add identityKey, Array( _
                systemKey, _
                itemCode, _
                NzStr(mLstRunPalette.List(i, 4)), _
                NzStr(mLstRunPalette.List(i, 7)), _
                NzStr(mLstRunPalette.List(i, 8)), _
                qtyVal)
        End If
NextChoice:
    Next i

    If agg.Count = 0 Then Exit Function
    writeStage = "ClearCheckTable"
    ClearTableContentsKeepRows lo
    writeStage = "EnsureCheckRows"
    If Not EnsureTableRows(lo, MaxLongLocal(agg.Count, CurrentRecipeRowBudget())) Then Exit Function
    writeStage = "WriteCheckRows"
    For Each key In agg.Keys
        outRow = outRow + 1
        entry = agg(key)
        SetCellByHeader lo, outRow, "System_Key", NzStr(entry(0))
        SetCellByHeader lo, outRow, "ITEM_CODE", NzStr(entry(1))
        SetCellByHeader lo, outRow, "ITEM", NzStr(entry(2))
        SetCellByHeader lo, outRow, "UOM", NzStr(entry(3))
        SetCellByHeader lo, outRow, "USED", CDbl(entry(5))
        SetCellByHeader lo, outRow, "TOTAL INV", NzStr(entry(4))
    Next key
    WriteProductionCheckRowsFromRunPalette = True
    Exit Function

FailWrite:
    mCheckRowsDiagnostic = "Stage=" & writeStage & "; Error=" & CStr(Err.Number) & _
        " - " & Err.Description
End Function

Private Function ResolveRunSystemKey(ByVal itemCode As String, _
                                     ByVal itemName As String, _
                                     ByVal locationValue As String) As String
    Dim lo As ListObject
    Dim entities As Variant
    Dim cSystemKey As Long
    Dim cItemCode As Long
    Dim cItem As Long
    Dim cLocation As Long
    Dim r As Long
    Dim codeMatches As Boolean
    Dim nameMatches As Boolean
    Dim locationMatches As Boolean

    Set lo = InventoryTable()
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then
            cSystemKey = ProductionColumnIndex(lo, "System_Key")
            cItemCode = ProductionColumnIndex(lo, "ITEM_CODE")
            If cItemCode = 0 Then cItemCode = ProductionColumnIndex(lo, "SKU")
            cItem = ProductionColumnIndex(lo, "ITEM")
            If cItem = 0 Then cItem = ProductionColumnIndex(lo, "ItemName")
            cLocation = ProductionColumnIndex(lo, "LOCATION")
            If cSystemKey > 0 Then
                For r = 1 To lo.ListRows.Count
                    codeMatches = False
                    If Trim$(itemCode) <> "" And cItemCode > 0 Then
                        codeMatches = (StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cItemCode).Value)), _
                                                   Trim$(itemCode), vbTextCompare) = 0)
                    End If
                    nameMatches = False
                    If Trim$(itemName) <> "" And cItem > 0 Then
                        nameMatches = (StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cItem).Value)), _
                                                   Trim$(itemName), vbTextCompare) = 0)
                    End If
                    locationMatches = True
                    If Trim$(locationValue) <> "" And cLocation > 0 Then
                        locationMatches = (StrComp(Trim$(NzStr(lo.DataBodyRange.Cells(r, cLocation).Value)), _
                                                       Trim$(locationValue), vbTextCompare) = 0)
                    End If
                    If (codeMatches Or nameMatches) And locationMatches Then
                        ResolveRunSystemKey = Trim$(NzStr(lo.DataBodyRange.Cells(r, cSystemKey).Value))
                        If ResolveRunSystemKey <> "" Then Exit Function
                    End If
                Next r
            End If
        End If
    End If

    On Error GoTo CleanFail
    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(itemCode)
    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        codeMatches = (Trim$(itemCode) <> "" And _
                       (StrComp(Trim$(NzStr(entities(r, 3))), Trim$(itemCode), vbTextCompare) = 0 Or _
                        StrComp(Trim$(NzStr(entities(r, 2))), Trim$(itemCode), vbTextCompare) = 0))
        nameMatches = (Trim$(itemName) <> "" And _
                       StrComp(Trim$(NzStr(entities(r, 4))), Trim$(itemName), vbTextCompare) = 0)
        locationMatches = (Trim$(locationValue) = "" Or _
                           StrComp(Trim$(NzStr(entities(r, 7))), Trim$(locationValue), vbTextCompare) = 0)
        If (codeMatches Or nameMatches) And locationMatches Then
            ResolveRunSystemKey = Trim$(NzStr(entities(r, 1)))
            If ResolveRunSystemKey <> "" Then Exit Function
        End If
    Next r
CleanFail:
End Function

Private Function RunChoiceWouldExceedInventory(ByVal listIndex As Long, ByVal qtyVal As Double) As Boolean
    Dim invText As String
    Dim invQty As Double

    invText = Trim$(NzStr(mLstRunPalette.List(listIndex, 8)))
    If invText = "" Then Exit Function
    If Not IsNumeric(Left$(invText, 1)) Then Exit Function
    invQty = CDbl(Val(invText))
    RunChoiceWouldExceedInventory = (qtyVal > invQty + 0.0000001)
End Function

Private Function NzDblLocal(ByVal value As Variant) As Double
    If IsError(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    If IsNumeric(value) Then NzDblLocal = CDbl(value)
End Function

Private Function CurrentRecipeRowBudget() As Long
    If mTxtRecipeRowBudget Is Nothing Then
        CurrentRecipeRowBudget = PRODUCTION_DEFAULT_ROW_BUDGET
    Else
        CurrentRecipeRowBudget = CLng(Val(NormalizeRecipeRowBudgetText(mTxtRecipeRowBudget.Text)))
    End If
End Function

Private Function NormalizeRecipeRowBudgetText(ByVal valueText As String) As String
    Dim n As Long

    n = CLng(Val(Trim$(valueText)))
    If n <= 0 Then n = PRODUCTION_DEFAULT_ROW_BUDGET
    If n > PRODUCTION_MAX_ROW_BUDGET Then n = PRODUCTION_MAX_ROW_BUDGET
    NormalizeRecipeRowBudgetText = CStr(n)
End Function

Private Function MaxLongLocal(ByVal leftValue As Long, ByVal rightValue As Long) As Long
    If leftValue >= rightValue Then
        MaxLongLocal = leftValue
    Else
        MaxLongLocal = rightValue
    End If
End Function

Private Sub LoadSelectedRunPaletteRow()
    Dim idx As Long
    Dim lst As MSForms.ListBox
    Dim splitTextBox As MSForms.TextBox
    Dim qtyTextBox As MSForms.TextBox

    Set lst = ActiveRunPaletteList()
    If lst Is Nothing Then Exit Sub
    Set splitTextBox = ActiveRunSplitTextBox()
    Set qtyTextBox = ActiveRunQtyTextBox()
    If splitTextBox Is Nothing Or qtyTextBox Is Nothing Then Exit Sub
    idx = lst.ListIndex
    If idx < 0 Then Exit Sub
    If IsRunTreeParentRow(lst, idx) Then Exit Sub
    If IsRunTreeOutputRow(lst, idx) Then Exit Sub
    mUpdatingPaletteInputs = True
    splitTextBox.Text = NzStr(lst.List(idx, 5))
    qtyTextBox.Text = NzStr(lst.List(idx, 6))
    mPaletteInputSource = ""
    mUpdatingPaletteInputs = False
    SetRunAllocationVisualState splitTextBox, qtyTextBox, RunAllocationPercentForGroup(RunAllocationGroupKeyFromList(lst, idx))
End Sub

Private Sub ApplySelectedRunPaletteSplit()
    Dim idx As Long
    Dim lst As MSForms.ListBox
    Dim tableName As String
    Dim rowIndex As Long
    Dim lo As ListObject
    Dim splitText As String
    Dim qtyText As String
    Dim baseQtyText As String
    Dim splitVal As Double
    Dim qtyVal As Double
    Dim baseQtyVal As Double
    Dim hasSplit As Boolean
    Dim hasQty As Boolean
    Dim splitTextBox As MSForms.TextBox
    Dim qtyTextBox As MSForms.TextBox
    Dim groupKey As String
    Dim allocationKey As String
    Dim otherPct As Double
    Dim newTotalPct As Double
    Dim runLoc As String
    Dim invLoc As String
    Dim reusableReport As String
    Dim requiredQty As Double
    Dim allocationApplied As Boolean

    If modProductionReusableRun.ReusableRunIsLoaded() Then
        If mLstRunPalette.ListIndex < 0 Then
            ShowStatus "Select an acceptable inventory stock row first."
            Exit Sub
        End If
        idx = mLstRunPalette.ListIndex
        splitText = Trim$(mTxtPaletteSplit.Text)
        qtyText = Trim$(mTxtPaletteQty.Text)
        requiredQty = modProductionReusableRun.ReusableRunRequirementQty( _
            NzStr(mLstRunPalette.List(idx, 0)), NzStr(mLstRunPalette.List(idx, 1)))
        If qtyText <> "" Then
            If Not TryParseNonNegativeRunNumber(qtyText, qtyVal, "Quantity") Then Exit Sub
            If requiredQty > 0 Then splitVal = qtyVal / requiredQty * 100#
        ElseIf splitText <> "" Then
            If Not TryParseNonNegativeRunNumber(splitText, splitVal, "% of Requirement") Then Exit Sub
            qtyVal = requiredQty * splitVal / 100#
        Else
            ShowStatus "Enter % of Requirement or Qty first."
            Exit Sub
        End If
        If ActiveRunLocation() = "" Then
            ShowStatus "Choose a production run location before allocating inventory."
            Exit Sub
        End If
        If StrComp(ActiveRunLocation(), NzStr(mLstRunPalette.List(idx, 9)), vbTextCompare) <> 0 Then
            ShowStatus "Allocation rejected. Inventory is at " & NzStr(mLstRunPalette.List(idx, 9)) & _
                       "; production run location is " & ActiveRunLocation() & "."
            Exit Sub
        End If
        allocationApplied = modProductionReusableRun.ApplyReusableRunStockAllocation( _
            CStr(mLstRunPalette.List(idx, 0)), CStr(mLstRunPalette.List(idx, 1)), _
            CStr(mLstRunPalette.List(idx, 3)), CDbl(qtyVal), reusableReport)
        If allocationApplied Then
            mTxtPaletteSplit.Text = FormatRunNumber(splitVal)
            mTxtPaletteQty.Text = FormatRunNumber(qtyVal)
            RefreshReusableRunControls False
        End If
        ShowStatus reusableReport
        Exit Sub
    End If

    Set lst = ActiveRunPaletteList()
    If lst Is Nothing Then Exit Sub
    Set splitTextBox = ActiveRunSplitTextBox()
    Set qtyTextBox = ActiveRunQtyTextBox()
    If splitTextBox Is Nothing Or qtyTextBox Is Nothing Then Exit Sub
    idx = lst.ListIndex
    If idx < 0 Then
        ShowStatus "Select an acceptable inventory row first."
        Exit Sub
    End If
    If IsRunTreeParentRow(lst, idx) Then
        ShowStatus "Select an acceptable inventory child row first."
        Exit Sub
    End If
    If IsRunTreeOutputRow(lst, idx) Then
        ShowStatus "Select an input inventory choice row, not an output row."
        Exit Sub
    End If

    tableName = NzStr(lst.List(idx, 0))
    rowIndex = CLng(Val(NzStr(lst.List(idx, 1))))
    runLoc = ActiveRunLocation()
    invLoc = Trim$(NzStr(lst.List(idx, 9)))

    splitText = Trim$(splitTextBox.Text)
    qtyText = Trim$(qtyTextBox.Text)
    baseQtyVal = RunBaseQtyFromList(lst, idx)

    If splitText <> "" Then
        If Not TryParseNonNegativeRunNumber(splitText, splitVal, "% of Requirement") Then Exit Sub
        hasSplit = True
    End If
    If qtyText <> "" Then
        If Not TryParseNonNegativeRunNumber(qtyText, qtyVal, "Quantity") Then Exit Sub
        hasQty = True
    End If

    If mPaletteInputSource = "QTY" And hasQty Then
        If baseQtyVal > 0 Then
            splitVal = qtyVal / baseQtyVal * 100#
            splitText = FormatRunNumber(splitVal)
            hasSplit = True
            SetPaletteTextSilently splitTextBox, splitText
        End If
    ElseIf hasSplit Then
        If baseQtyVal > 0 Then
            qtyVal = baseQtyVal * splitVal / 100#
            qtyText = FormatRunNumber(qtyVal)
            hasQty = True
            SetPaletteTextSilently qtyTextBox, qtyText
        End If
    ElseIf hasQty And baseQtyVal > 0 Then
        splitVal = qtyVal / baseQtyVal * 100#
        splitText = FormatRunNumber(splitVal)
        hasSplit = True
        SetPaletteTextSilently splitTextBox, splitText
    End If

    If Not hasSplit And Not hasQty Then
        ShowStatus "Enter % of Requirement or Qty first."
        Exit Sub
    End If
    If Not hasSplit Then
        ShowStatus "Cannot validate allocation without a % of Requirement. Select a row with a recipe requirement first."
        Exit Sub
    End If
    If (hasQty And qtyVal > 0) Or (hasSplit And splitVal > 0) Then
        If Not RunChoiceLocationAllowed(runLoc, invLoc, IIf(hasQty, qtyVal, splitVal)) Then
            ClearRunAllocationForListRow lst, idx
            BuildRunTreeFromPaletteList
            If runLoc = "" Then
                ShowStatus "Choose a production run location before allocating inventory."
            ElseIf invLoc = "" Then
                ShowStatus "This operator inventory row has no refreshed location. Click Refresh in Production Run, then select inventory at " & runLoc & "."
            Else
                ShowStatus "Allocation rejected. Inventory is at " & invLoc & "; production run location is " & runLoc & ". Use inventory at the production location. This row was cleared from the run."
            End If
            Exit Sub
        End If
    End If

    groupKey = RunAllocationGroupKeyFromList(lst, idx)
    allocationKey = RunAllocationKeyFromList(lst, idx)
    otherPct = RunAllocationPercentForGroup(groupKey, allocationKey)
    newTotalPct = otherPct + splitVal
    If newTotalPct > 100.001 Then
        SetRunAllocationVisualState splitTextBox, qtyTextBox, newTotalPct
        ShowStatus "Allocation rejected. Other choices already use " & FormatRunNumber(otherPct) & "%. This would make " & FormatRunNumber(newTotalPct) & "%; maximum is 100%."
        Exit Sub
    End If

    If tableName <> "" And rowIndex >= 1 Then
        Set lo = ProductionTable(tableName)
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                If rowIndex <= lo.ListRows.Count Then
                    If hasSplit Then SetCellByHeader lo, rowIndex, "SPLIT %", splitVal
                    If hasQty Then SetCellByHeader lo, rowIndex, "QUANTITY", qtyVal
                End If
            End If
        End If
    End If

    If hasSplit Then lst.List(idx, 5) = splitText
    If hasQty Then lst.List(idx, 6) = qtyText
    StoreRunAllocationOverride lst, idx, NzStr(lst.List(idx, 5)), NzStr(lst.List(idx, 6))
    SyncRunAllocationToPaletteList lst, idx
    BuildRunTreeFromPaletteList
    SetRunAllocationVisualState splitTextBox, qtyTextBox, newTotalPct
    ShowStatus "Acceptable inventory allocation updated. Ingredient is " & FormatRunNumber(newTotalPct) & "% filled."
End Sub

Private Function TryParseNonNegativeRunNumber(ByVal numberText As String, ByRef numberValue As Double, ByVal labelText As String) As Boolean
    If Not IsNumeric(numberText) Then
        ShowStatus labelText & " must be numeric."
        Exit Function
    End If
    numberValue = CDbl(numberText)
    If numberValue < 0 Then
        ShowStatus labelText & " cannot be negative."
        Exit Function
    End If
    TryParseNonNegativeRunNumber = True
End Function

Private Function FormatRunNumber(ByVal numberValue As Double) As String
    FormatRunNumber = Format$(numberValue, "0.###")
End Function

Private Sub SetPaletteTextSilently(ByVal txt As MSForms.TextBox, ByVal valueText As String)
    mUpdatingPaletteInputs = True
    txt.Text = valueText
    mUpdatingPaletteInputs = False
End Sub

Private Sub SetRunAllocationVisualState(ByVal splitTextBox As MSForms.TextBox, ByVal qtyTextBox As MSForms.TextBox, ByVal totalPct As Double)
    Dim fillColor As Long

    If totalPct > 100.001 Then
        fillColor = RGB(255, 210, 210)
    ElseIf totalPct >= 99.999 Then
        fillColor = RGB(210, 245, 210)
    ElseIf totalPct > 0 Then
        fillColor = RGB(255, 245, 200)
    Else
        fillColor = vbWhite
    End If
    If Not splitTextBox Is Nothing Then
        splitTextBox.BackColor = fillColor
        splitTextBox.Font.Bold = (totalPct >= 99.999 And totalPct <= 100.001)
    End If
    If Not qtyTextBox Is Nothing Then
        qtyTextBox.BackColor = fillColor
        qtyTextBox.Font.Bold = (totalPct >= 99.999 And totalPct <= 100.001)
    End If
End Sub

Private Sub SyncRunAllocationToPaletteList(ByVal sourceList As MSForms.ListBox, ByVal sourceRow As Long)
    Dim sourceKey As String
    Dim i As Long

    If mLstRunPalette Is Nothing Or sourceList Is Nothing Then Exit Sub
    sourceKey = RunAllocationKeyFromList(sourceList, sourceRow)
    If sourceKey = "" Then Exit Sub
    For i = 0 To mLstRunPalette.ListCount - 1
        If RunAllocationKeyFromList(mLstRunPalette, i) = sourceKey Then
            mLstRunPalette.List(i, 5) = NzStr(sourceList.List(sourceRow, 5))
            mLstRunPalette.List(i, 6) = NzStr(sourceList.List(sourceRow, 6))
            StoreRunAllocationOverride mLstRunPalette, i, NzStr(mLstRunPalette.List(i, 5)), NzStr(mLstRunPalette.List(i, 6))
            Exit For
        End If
    Next i
End Sub

Private Function SelectedRunPaletteBaseQty() As Double
    Dim lst As MSForms.ListBox
    Dim idx As Long
    Dim baseText As String

    Set lst = ActiveRunPaletteList()
    If lst Is Nothing Then Exit Function
    idx = lst.ListIndex
    If idx < 0 Then Exit Function
    If IsRunTreeParentRow(lst, idx) Then Exit Function
    SelectedRunPaletteBaseQty = RunBaseQtyFromList(lst, idx)
End Function

Private Function ActiveRunPaletteList() As MSForms.ListBox
    If Not mPages Is Nothing Then
        If mPages.Value = 4 Then
            Set ActiveRunPaletteList = mLstRunTree
            Exit Function
        End If
    End If
    Set ActiveRunPaletteList = mLstRunPalette
End Function

Private Function ActiveRunSplitTextBox() As MSForms.TextBox
    If Not mPages Is Nothing Then
        If mPages.Value = 4 Then
            Set ActiveRunSplitTextBox = mTxtTreePaletteSplit
            Exit Function
        End If
    End If
    Set ActiveRunSplitTextBox = mTxtPaletteSplit
End Function

Private Function ActiveRunQtyTextBox() As MSForms.TextBox
    If Not mPages Is Nothing Then
        If mPages.Value = 4 Then
            Set ActiveRunQtyTextBox = mTxtTreePaletteQty
            Exit Function
        End If
    End If
    Set ActiveRunQtyTextBox = mTxtPaletteQty
End Function

Private Sub EnsureTableRow(ByVal lo As ListObject)
    EnsureTableRows lo, 1
End Sub

Private Function EnsureTableRows(ByVal lo As ListObject, ByVal rowCount As Long) As Boolean
    On Error GoTo Fail

    If lo Is Nothing Then Exit Function
    If rowCount < 1 Then
        EnsureTableRows = True
        Exit Function
    End If
    Do While lo.ListRows.Count < rowCount
        lo.ListRows.Add AlwaysInsert:=False
    Loop
    EnsureTableRows = True
    Exit Function

Fail:
    EnsureTableRows = False
End Function

Private Sub ClearListRows(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    Do While lo.ListRows.Count > 0
        lo.ListRows(1).Delete
    Loop
    On Error GoTo 0
End Sub

Private Sub ClearTableContentsKeepRows(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.ClearContents
End Sub

Private Function CellByHeader(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal headerName As String) As String
    Dim colIndex As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.DataBodyRange.Rows.Count Then Exit Function
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Function
    CellByHeader = NzStr(lo.DataBodyRange.Cells(rowIndex, colIndex).Value)
End Function

Private Sub SetCellByHeader(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal headerName As String, ByVal value As Variant)
    Dim colIndex As Long
    If lo Is Nothing Then Exit Sub
    If rowIndex < 1 Then Exit Sub
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Sub
    If Not EnsureTableRows(lo, rowIndex) Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value = value
End Sub

Private Function BuildFormGuid() As String
    On Error Resume Next
    BuildFormGuid = NzStr(mProduction.CreateProductionGuid())
    If BuildFormGuid = "" Then
        BuildFormGuid = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd() * 1000000#))
    End If
    On Error GoTo 0
End Function

Private Function FirstRowValue(ByVal lo As ListObject, ByVal headerName As String) As String
    Dim colIndex As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Function
    FirstRowValue = NzStr(lo.DataBodyRange.Cells(1, colIndex).Value)
End Function

Private Sub SetFirstRowValue(ByVal lo As ListObject, ByVal headerName As String, ByVal value As Variant)
    Dim colIndex As Long
    If lo Is Nothing Then Exit Sub
    EnsureTableRow lo
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Sub
    If StrComp(headerName, "RECIPE_ID", vbTextCompare) = 0 Then
        lo.DataBodyRange.Cells(1, colIndex).NumberFormat = "@"
        lo.DataBodyRange.Cells(1, colIndex).Value2 = CStr(value)
    Else
        lo.DataBodyRange.Cells(1, colIndex).Value = value
    End If
End Sub

Private Sub SetRowValue(ByVal lr As ListRow, ByVal lo As ListObject, ByVal headerName As String, ByVal value As Variant)
    Dim colIndex As Long
    If lr Is Nothing Or lo Is Nothing Then Exit Sub
    colIndex = ProductionColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Sub
    If StrComp(headerName, "RECIPE_ID", vbTextCompare) = 0 Then
        lr.Range.Cells(1, colIndex).NumberFormat = "@"
        lr.Range.Cells(1, colIndex).Value2 = CStr(value)
    Else
        lr.Range.Cells(1, colIndex).Value = value
    End If
End Sub

Private Sub ClearTableContentsKeepBlank(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    EnsureTableRow lo
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.ClearContents
End Sub

Private Function ProductionColumnIndex(ByVal lo As ListObject, ByVal headerName As String) As Long
    On Error Resume Next
    ProductionColumnIndex = mProduction.ColumnIndex(lo, headerName)
    On Error GoTo 0
End Function

Private Sub BindOperatorWorkbookForRun()
    Dim wb As Workbook

    Set wb = ResolveOperatorWorkbook()
    If wb Is Nothing Then Exit Sub
    mProduction.BindProductionOperatorWorkbook wb
End Sub

Private Sub mBtnApplyBatchScale_Click()
    Dim scalePercent As Double
    Dim scaleReport As String

    If modProductionReusableRun.ReusableRunIsLoaded() Then
        If Not TryParseBatchScalePercent(mTxtBatchScalePercent.Text, scalePercent, scaleReport) Then
            ShowStatus scaleReport
            Exit Sub
        End If
        If modProductionReusableRun.ApplyReusableRunScale(scalePercent, scaleReport) Then
            RefreshReusableRunControls False
        End If
        ShowStatus scaleReport
        Exit Sub
    End If
    If mLstLoaderRecipes.ListIndex < 0 Then
        ShowStatus "Select a recipe, enter Batch scale %, then click Apply Scale."
        Exit Sub
    End If
    LoadSelectedRecipeIntoLoader
End Sub

Private Function TryParseBatchScalePercent(ByVal valueText As String, _
                                           ByRef scalePercent As Double, _
                                           ByRef report As String) As Boolean
    valueText = Trim$(valueText)
    If Not IsNumeric(valueText) Then
        report = "Batch scale must be a number from 0.001% through 1000%."
        Exit Function
    End If
    scalePercent = CDbl(valueText)
    If scalePercent < BATCH_SCALE_MIN_PERCENT Or _
       scalePercent > BATCH_SCALE_MAX_PERCENT Then
        report = "Batch scale must be from 0.001% through 1000%."
        Exit Function
    End If
    report = "OK"
    TryParseBatchScalePercent = True
End Function

Private Function ApplyBatchScaleToRunList(ByVal scalePercent As Double, _
                                          ByRef report As String) As Boolean
    On Error GoTo Failed

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim factor As Double
    Dim changedCells As Long

    If mOperatorWorkbook Is Nothing Then
        report = "Production operator workbook is not bound."
        Exit Function
    End If
    Set ws = mOperatorWorkbook.Worksheets("Production")
    factor = scalePercent / 100#
    For Each lo In ws.ListObjects
        If IsBatchScaleRunTable(lo) Then
            If LCase$(lo.Name) = "recipechooser_generated" _
               Or LCase$(lo.Name) Like "proc_*_rchooser" Then
                changedCells = changedCells + _
                    ScaleProductionTableColumn(lo, "AMOUNT NEEDED", factor)
            Else
                changedCells = changedCells + _
                    ScaleProductionTableColumn(lo, "BASE QUANTITY", factor)
                changedCells = changedCells + _
                    ScaleProductionTableColumn(lo, "QUANTITY", factor)
            End If
        End If
    Next lo
    report = "Batch scale applied: " & FormatRunNumber(scalePercent) & _
             "%; scaled cells=" & CStr(changedCells) & "."
    ApplyBatchScaleToRunList = True
    Exit Function
Failed:
    report = "Batch scaling failed: " & Err.Description
End Function

Private Function IsBatchScaleRunTable(ByVal lo As ListObject) As Boolean
    Dim tableName As String

    If lo Is Nothing Then Exit Function
    tableName = LCase$(lo.Name)
    IsBatchScaleRunTable = _
        (tableName = "recipechooser_generated") Or _
        (tableName = "inventorypalette_generated") Or _
        (tableName Like "proc_*_rchooser") Or _
        (tableName Like "proc_*_palette")
End Function

Private Function ScaleProductionTableColumn(ByVal lo As ListObject, _
                                            ByVal headerName As String, _
                                            ByVal factor As Double) As Long
    Dim columnIndex As Long
    Dim rowIndex As Long
    Dim currentValue As Variant

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    columnIndex = ProductionColumnIndex(lo, headerName)
    If columnIndex = 0 Then Exit Function
    For rowIndex = 1 To lo.ListRows.Count
        currentValue = lo.DataBodyRange.Cells(rowIndex, columnIndex).Value2
        If Not IsError(currentValue) And IsNumeric(currentValue) Then
            If Trim$(CStr(currentValue)) <> "" Then
                lo.DataBodyRange.Cells(rowIndex, columnIndex).Value2 = _
                    CDbl(currentValue) * factor
                ScaleProductionTableColumn = ScaleProductionTableColumn + 1
            End If
        End If
    Next rowIndex
End Function

Private Sub ClearProcessDraft(Optional ByVal createIdentity As Boolean = True)
    mTxtProcessName.Text = ""
    mTxtProcessId.Text = ""
    mTxtProcessVersion.Text = ""
    mTxtProcessDescription.Text = ""
    mLstProcessRequirements.Clear
    mLstProcessOutputs.Clear
    mLstProcessInstructions.Clear
    Set mProcessAlternatives = New Collection
    Set mProcessOutputRegulations = CreateObject("Scripting.Dictionary")
    mProcessOutputRegulations.CompareMode = vbTextCompare
    mSelectedProcessRequirementIndex = -1
    mSelectedProcessOutputIndex = -1
    If createIdentity Then
        mTxtProcessId.Text = NextProcessDraftBase36Id()
        mTxtProcessVersion.Text = "1"
    End If
    ClearRequirementEditor
    ClearOutputEditor
    mTxtProcessInstruction.Text = ""
End Sub

Private Sub ClearRequirementEditor()
    mTxtRequirementId.Text = NextProcessComponentBase36Id()
    mTxtRequirementName.Text = ""
    mTxtRequirementQty.Text = ""
    mTxtRequirementPercent.Text = ""
    mTxtRequirementYieldBasis.Text = ""
    mTxtRequirementUom.Text = ""
    SelectComboText mCmbRequirementQtyMode, "Enter a number"
    ApplyRequirementQtyMode
End Sub

Private Sub ClearOutputEditor()
    mTxtProcessOutputId.Text = NextProcessComponentBase36Id()
    mTxtProcessOutputName.Text = ""
    mTxtProcessOutputItemCode.Text = ""
    mTxtProcessOutputDesignId.Text = ""
    mTxtProcessOutputDesignVersion.Text = ""
    mTxtProcessOutputQty.Text = ""
    mTxtProcessOutputPercent.Text = ""
    mTxtProcessOutputYieldBasis.Text = ""
    SelectComboText mCmbProcessOutputQtyMode, "Enter a number"
    ApplyOutputQtyMode
    RefreshProcessOutputUomCatalog
    RefreshOutputDesignIdentity True
End Sub

Private Function RequirementQtyMode() As String
    RequirementQtyMode = IIf(InStr(1, ComboText(mCmbRequirementQtyMode), "Variable", vbTextCompare) > 0, "ACTUAL", "FIXED")
End Function

Private Function OutputQtyMode() As String
    OutputQtyMode = IIf(InStr(1, ComboText(mCmbProcessOutputQtyMode), "Variable", vbTextCompare) > 0, "ACTUAL", "FIXED")
End Function

Private Sub ApplyRequirementQtyMode()
    Dim isActual As Boolean
    isActual = (RequirementQtyMode() = "ACTUAL")
    mTxtRequirementQty.Locked = isActual
    mTxtRequirementPercent.Locked = isActual
    mTxtRequirementYieldBasis.Locked = isActual
    If isActual Then
        mTxtRequirementQty.Text = ""
        mTxtRequirementPercent.Text = ""
        mTxtRequirementYieldBasis.Text = ""
    End If
End Sub

Private Sub ApplyOutputQtyMode()
    Dim isActual As Boolean
    isActual = (OutputQtyMode() = "ACTUAL")
    mTxtProcessOutputQty.Locked = isActual
    mTxtProcessOutputPercent.Locked = isActual
    mTxtProcessOutputYieldBasis.Locked = isActual
    If isActual Then
        mTxtProcessOutputQty.Text = ""
        mTxtProcessOutputPercent.Text = ""
        mTxtProcessOutputYieldBasis.Text = ""
    End If
End Sub

Private Sub RefreshOutputDesignIdentity(Optional ByVal forceGenerated As Boolean = False)
    Dim designId As String

    If Trim$(mTxtProcessId.Text) = "" Or Trim$(mTxtProcessOutputId.Text) = "" Then Exit Sub
    designId = "D-" & UCase$(Trim$(mTxtProcessId.Text)) & "-" & _
        UCase$(Trim$(mTxtProcessOutputId.Text))
    If forceGenerated Or Trim$(mTxtProcessOutputDesignId.Text) = "" Then _
        mTxtProcessOutputDesignId.Text = designId
    If forceGenerated Or Trim$(mTxtProcessOutputDesignVersion.Text) = "" Then _
        mTxtProcessOutputDesignVersion.Text = Trim$(mTxtProcessVersion.Text)
    If forceGenerated Or Trim$(mTxtProcessOutputItemCode.Text) = "" Then _
        mTxtProcessOutputItemCode.Text = designId
End Sub

Private Sub LoadSelectedProcessDefinition(ByVal reuseAsNewVersion As Boolean)
    Dim idx As Long

    idx = mLstProcesses.ListIndex
    If idx < 0 Then
        ShowStatus "Select a saved Process first."
        Exit Sub
    End If
    LoadProcessDefinitionIntoDesigner NzStr(mLstProcesses.List(idx, 0)), _
        NzStr(mLstProcesses.List(idx, 1)), reuseAsNewVersion
End Sub

Private Function LoadProcessDefinitionIntoDesigner(ByVal processId As String, _
                                                   ByVal processVersion As String, _
                                                   ByVal reuseAsNewVersion As Boolean) As Boolean
    Dim jsonText As String
    Dim parseReport As String
    Dim records As Collection
    Dim record As Object
    Dim rowIndex As Long

    jsonText = modOperationsPrimitiveBridge.GetProcessVersion(processId, processVersion)
    Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords(jsonText, parseReport)
    If records Is Nothing Then
        ShowStatus "Process load failed: " & parseReport
        Exit Function
    End If
    ClearProcessDraft False
    For Each record In records
        Select Case UCase$(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"))
            Case "PROCESS"
                mTxtProcessId.Text = modProductionReusableDesigns.ReusableRecordText(record, "ProcessId")
                mTxtProcessVersion.Text = modProductionReusableDesigns.ReusableRecordText(record, "ProcessVersion")
                mTxtProcessName.Text = modProductionReusableDesigns.ReusableRecordText(record, "ProcessName")
                mTxtProcessDescription.Text = modProductionReusableDesigns.ReusableRecordText(record, "Description")
            Case "REQUIREMENT"
                rowIndex = AddProcessRequirementRecord(record)
            Case "ALTERNATIVE"
                mProcessAlternatives.Add CloneReusableRecord(record)
            Case "OUTPUT"
                rowIndex = AddProcessOutputRecord(record)
            Case "INSTRUCTION"
                mLstProcessInstructions.AddItem modProductionReusableDesigns.ReusableRecordText(record, "InstructionOrdinal")
                mLstProcessInstructions.List(mLstProcessInstructions.ListCount - 1, 1) = _
                    modProductionReusableDesigns.ReusableRecordText(record, "Instruction")
        End Select
    Next record
    If reuseAsNewVersion Then
        SetEditableProcessDraftVersion _
            modProductionReusableDesigns.NextReusableDefinitionVersion(processId, True)
        ShowStatus "Process " & processId & " is ready to edit as new DRAFT version " & _
            mTxtProcessVersion.Text & ". The saved version remains unchanged."
    Else
        ShowStatus "Viewed immutable Process " & processId & " version " & processVersion & _
            ". Choose Edit as New Version before saving changes, or Send Process to Sheet."
    End If
    NormalizeProcessComponentIdentities
    LoadProcessDefinitionIntoDesigner = True
End Function

Private Sub SetEditableProcessDraftVersion(ByVal processVersion As String)
    Dim rowIndex As Long
    Dim outputId As String

    processVersion = Trim$(processVersion)
    If processVersion = "" Then Exit Sub
    mTxtProcessVersion.Text = processVersion
    For rowIndex = 0 To mLstProcessOutputs.ListCount - 1
        outputId = UCase$(Trim$(NzStr(mLstProcessOutputs.List(rowIndex, 0))))
        If outputId <> "" Then
            mLstProcessOutputs.List(rowIndex, 3) = "D-" & _
                UCase$(Trim$(mTxtProcessId.Text)) & "-" & outputId
            mLstProcessOutputs.List(rowIndex, 4) = processVersion
        End If
    Next rowIndex
    ClearOutputEditor
End Sub

Private Function AddProcessRequirementRecord(ByVal record As Object) As Long
    mLstProcessRequirements.AddItem modProductionReusableDesigns.ReusableRecordText(record, "RequirementId")
    AddProcessRequirementRecord = mLstProcessRequirements.ListCount - 1
    With mLstProcessRequirements
        .List(AddProcessRequirementRecord, 1) = modProductionReusableDesigns.ReusableRecordText(record, "RequirementName")
        .List(AddProcessRequirementRecord, 2) = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Qty"))
        .List(AddProcessRequirementRecord, 3) = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Percent"))
        .List(AddProcessRequirementRecord, 4) = modProductionReusableDesigns.ReusableRecordText(record, "YieldBasis")
        .List(AddProcessRequirementRecord, 5) = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
        .List(AddProcessRequirementRecord, 6) = UCase$(NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "RequirementQtyMode")))
        If .List(AddProcessRequirementRecord, 6) = "" Then .List(AddProcessRequirementRecord, 6) = "FIXED"
    End With
End Function

Private Function AddProcessOutputRecord(ByVal record As Object) As Long
    Dim qtyText As String
    Dim percentText As String
    Dim yieldBasisText As String

    qtyText = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Qty"))
    percentText = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Percent"))
    yieldBasisText = modProductionReusableDesigns.ReusableRecordText(record, "YieldBasis")
    If StrComp(NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "OutputQtyMode")), "ACTUAL", vbTextCompare) <> 0 Then
        percentText = NormalizedOutputYieldPercent(qtyText, percentText)
        yieldBasisText = NormalizedOutputYieldBasis(qtyText, percentText, yieldBasisText)
    Else
        qtyText = "": percentText = "": yieldBasisText = ""
    End If
    mLstProcessOutputs.AddItem modProductionReusableDesigns.ReusableRecordText(record, "OutputId")
    AddProcessOutputRecord = mLstProcessOutputs.ListCount - 1
    With mLstProcessOutputs
        .List(AddProcessOutputRecord, 1) = modProductionReusableDesigns.ReusableRecordText(record, "OutputName")
        .List(AddProcessOutputRecord, 2) = modProductionReusableDesigns.ReusableRecordText(record, "ITEM_CODE")
        .List(AddProcessOutputRecord, 3) = modProductionReusableDesigns.ReusableRecordText(record, "ComponentDesignId")
        .List(AddProcessOutputRecord, 4) = modProductionReusableDesigns.ReusableRecordText(record, "ComponentDesignVersion")
        .List(AddProcessOutputRecord, 5) = qtyText
        .List(AddProcessOutputRecord, 6) = percentText
        .List(AddProcessOutputRecord, 7) = yieldBasisText
        .List(AddProcessOutputRecord, 8) = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
        .List(AddProcessOutputRecord, 9) = UCase$(NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "OutputQtyMode")))
        If .List(AddProcessOutputRecord, 9) = "" Then .List(AddProcessOutputRecord, 9) = "FIXED"
        SetProcessOutputRegulation modProductionReusableDesigns.ReusableRecordText(record, "OutputId"), _
            ReusableRecordBoolean(record, "OutputRegulationEnabled"), _
            NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "OutputFloorQty")), _
            NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "OutputCeilingQty"))
    End With
End Function

Private Function LoadProcessPayloadIntoDesigner(ByVal processId As String, _
                                                ByVal processVersion As String, _
                                                ByVal processName As String, _
                                                ByVal description As String, _
                                                ByVal payloadJson As String, _
                                                ByRef report As String) As Boolean
    Dim records As Collection
    Dim record As Object
    Dim parseReport As String

    Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords(payloadJson, parseReport)
    If records Is Nothing Then
        report = parseReport
        Exit Function
    End If
    ClearProcessDraft False
    mTxtProcessId.Text = processId
    mTxtProcessVersion.Text = processVersion
    mTxtProcessName.Text = processName
    mTxtProcessDescription.Text = description
    For Each record In records
        Select Case UCase$(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"))
            Case "REQUIREMENT"
                AddProcessRequirementRecord record
            Case "ALTERNATIVE"
                mProcessAlternatives.Add CloneReusableRecord(record)
            Case "OUTPUT"
                AddProcessOutputRecord record
            Case "INSTRUCTION"
                mLstProcessInstructions.AddItem _
                    modProductionReusableDesigns.ReusableRecordText(record, "InstructionOrdinal")
                mLstProcessInstructions.List(mLstProcessInstructions.ListCount - 1, 1) = _
                    modProductionReusableDesigns.ReusableRecordText(record, "Instruction")
        End Select
    Next record
    NormalizeProcessComponentIdentities
    ClearRequirementEditor
    ClearOutputEditor
    report = "Process worksheet draft loaded into Process Designer."
    LoadProcessPayloadIntoDesigner = True
End Function

Private Sub NormalizeProcessComponentIdentities()
    Dim usedIds As Object
    Dim retainedRequirementRows As Object
    Dim retainedOutputRows As Object
    Dim requirementIdChanges As Object
    Dim alternative As Variant
    Dim rowIndex As Long
    Dim componentId As String
    Dim originalId As String
    Dim usedArray As Variant

    Set usedIds = CreateObject("Scripting.Dictionary")
    usedIds.CompareMode = vbTextCompare
    Set retainedRequirementRows = CreateObject("Scripting.Dictionary")
    Set retainedOutputRows = CreateObject("Scripting.Dictionary")
    Set requirementIdChanges = CreateObject("Scripting.Dictionary")
    requirementIdChanges.CompareMode = vbTextCompare
    For rowIndex = 0 To mLstProcessRequirements.ListCount - 1
        componentId = UCase$(Trim$(NzStr(mLstProcessRequirements.List(rowIndex, 0))))
        If mProduction.IsBase36Identifier(componentId) _
           And Not usedIds.Exists(componentId) Then
            usedIds.Add componentId, True
            retainedRequirementRows.Add CStr(rowIndex), True
        End If
    Next rowIndex
    For rowIndex = 0 To mLstProcessOutputs.ListCount - 1
        componentId = UCase$(Trim$(NzStr(mLstProcessOutputs.List(rowIndex, 0))))
        If mProduction.IsBase36Identifier(componentId) _
           And Not usedIds.Exists(componentId) Then
            usedIds.Add componentId, True
            retainedOutputRows.Add CStr(rowIndex), True
        End If
    Next rowIndex
    For rowIndex = 0 To mLstProcessRequirements.ListCount - 1
        If Not retainedRequirementRows.Exists(CStr(rowIndex)) Then
            originalId = UCase$(Trim$(NzStr(mLstProcessRequirements.List(rowIndex, 0))))
            If usedIds.Count = 0 Then
                usedArray = Array("")
            Else
                usedArray = usedIds.Keys
            End If
            componentId = mProduction.NextBase36Identifier(usedArray)
            mLstProcessRequirements.List(rowIndex, 0) = componentId
            If originalId <> "" Then requirementIdChanges(originalId) = componentId
            usedIds.Add componentId, True
        End If
    Next rowIndex
    For Each alternative In mProcessAlternatives
        originalId = UCase$(Trim$(modProductionReusableDesigns.ReusableRecordText( _
            alternative, "RequirementId")))
        If requirementIdChanges.Exists(originalId) Then _
            alternative("RequirementId") = requirementIdChanges(originalId)
    Next alternative
    For rowIndex = 0 To mLstProcessOutputs.ListCount - 1
        If Not retainedOutputRows.Exists(CStr(rowIndex)) Then
            If usedIds.Count = 0 Then
                usedArray = Array("")
            Else
                usedArray = usedIds.Keys
            End If
            componentId = mProduction.NextBase36Identifier(usedArray)
            mLstProcessOutputs.List(rowIndex, 0) = componentId
            mLstProcessOutputs.List(rowIndex, 3) = "D-" & _
                UCase$(Trim$(mTxtProcessId.Text)) & "-" & componentId
            mLstProcessOutputs.List(rowIndex, 4) = Trim$(mTxtProcessVersion.Text)
            usedIds.Add componentId, True
        End If
    Next rowIndex
End Sub

Private Function CloneReusableRecord(ByVal source As Object) As Object
    Dim target As Object
    Dim key As Variant

    Set target = CreateObject("Scripting.Dictionary")
    target.CompareMode = vbTextCompare
    For Each key In source.Keys
        target(CStr(key)) = source(key)
    Next key
    Set CloneReusableRecord = target
End Function

Private Sub WriteRequirementEditorToList(ByVal updateExisting As Boolean)
    Dim idx As Long
    Dim wasLoading As Boolean

    idx = mLstProcessRequirements.ListIndex
    If updateExisting And Trim$(mTxtRequirementId.Text) <> "" Then
        idx = FindProcessComponentEditorRow( _
            mLstProcessRequirements, Trim$(mTxtRequirementId.Text))
    End If
    If updateExisting And idx < 0 Then
        If mSelectedProcessRequirementIndex >= 0 _
           And mSelectedProcessRequirementIndex < mLstProcessRequirements.ListCount Then
            idx = mSelectedProcessRequirementIndex
        End If
    End If
    If Not updateExisting Or idx < 0 Then
        If Trim$(mTxtRequirementId.Text) = "" _
           Or ProcessComponentIdentityExists(mTxtRequirementId.Text) Then
            mTxtRequirementId.Text = NextProcessComponentBase36Id()
        End If
    Else
        mTxtRequirementId.Text = NzStr(mLstProcessRequirements.List(idx, 0))
    End If

    If Trim$(mTxtRequirementId.Text) = "" Or Trim$(mTxtRequirementName.Text) = "" Or _
       Trim$(mTxtRequirementUom.Text) = "" Then
        ShowStatus "A requirement needs ID, name, and UOM."
        Exit Sub
    End If
    If RequirementQtyMode() <> "ACTUAL" And Not PositiveTextValue(mTxtRequirementQty.Text) And Not PositiveTextValue(mTxtRequirementPercent.Text) Then
        ShowStatus "A requirement needs a positive quantity or percentage."
        Exit Sub
    End If
    If RequirementQtyMode() <> "ACTUAL" And PositiveTextValue(mTxtRequirementPercent.Text) _
       And Not PositiveTextValue(mTxtRequirementYieldBasis.Text) Then
        ShowStatus "A percentage requirement needs a positive batch basis quantity."
        Exit Sub
    End If
    If Not ValidateWholeQuantityForUom(mTxtRequirementQty.Text, _
            mTxtRequirementUom.Text, "Requirement quantity") Then Exit Sub
    If Not ValidateWholeQuantityForUom(mTxtRequirementYieldBasis.Text, _
            mTxtRequirementUom.Text, "Requirement batch basis") Then Exit Sub
    If Not updateExisting Or idx < 0 Then
        mLstProcessRequirements.AddItem ""
        idx = mLstProcessRequirements.ListCount - 1
    End If
    wasLoading = mLoading
    mLoading = True
    With mLstProcessRequirements
        .List(idx, 0) = Trim$(mTxtRequirementId.Text)
        .List(idx, 1) = Trim$(mTxtRequirementName.Text)
        .List(idx, 2) = Trim$(mTxtRequirementQty.Text)
        .List(idx, 3) = Trim$(mTxtRequirementPercent.Text)
        .List(idx, 4) = Trim$(mTxtRequirementYieldBasis.Text)
        .List(idx, 5) = Trim$(mTxtRequirementUom.Text)
        .List(idx, 6) = RequirementQtyMode()
        .ListIndex = idx
    End With
    mLoading = wasLoading
    mSelectedProcessRequirementIndex = idx
    If Not updateExisting Then ClearRequirementEditor
    ShowStatus "Requirement staged in the Process draft."
End Sub

Private Sub WriteOutputEditorToList(ByVal updateExisting As Boolean)
    Dim idx As Long
    Dim priorQty As String
    Dim priorPercent As String
    Dim priorYieldBasis As String
    Dim wasLoading As Boolean

    idx = mLstProcessOutputs.ListIndex
    If updateExisting And Trim$(mTxtProcessOutputId.Text) <> "" Then
        idx = FindProcessComponentEditorRow( _
            mLstProcessOutputs, Trim$(mTxtProcessOutputId.Text))
    End If
    If updateExisting And idx < 0 Then
        If mSelectedProcessOutputIndex >= 0 _
           And mSelectedProcessOutputIndex < mLstProcessOutputs.ListCount Then
            idx = mSelectedProcessOutputIndex
        End If
    End If
    If Not updateExisting Or idx < 0 Then
        If Trim$(mTxtProcessOutputId.Text) = "" _
           Or ProcessComponentIdentityExists(mTxtProcessOutputId.Text) Then
            mTxtProcessOutputId.Text = NextProcessComponentBase36Id()
        End If
    Else
        mTxtProcessOutputId.Text = NzStr(mLstProcessOutputs.List(idx, 0))
        priorQty = Trim$(NzStr(mLstProcessOutputs.List(idx, 5)))
        priorPercent = Trim$(NzStr(mLstProcessOutputs.List(idx, 6)))
        priorYieldBasis = Trim$(NzStr(mLstProcessOutputs.List(idx, 7)))
        If PositiveTextValue(mTxtProcessOutputQty.Text) _
           And StrComp(Trim$(mTxtProcessOutputQty.Text), priorQty, vbBinaryCompare) <> 0 _
           And Abs(Val(priorPercent) - 100#) < 0.00000001 _
           And Abs(Val(priorYieldBasis) - Val(priorQty)) < 0.00000001 _
           And Abs(Val(Trim$(mTxtProcessOutputPercent.Text)) - 100#) < 0.00000001 _
           And Abs(Val(Trim$(mTxtProcessOutputYieldBasis.Text)) - Val(priorYieldBasis)) < 0.00000001 Then
            mTxtProcessOutputYieldBasis.Text = Trim$(mTxtProcessOutputQty.Text)
        End If
    End If
    RefreshOutputDesignIdentity False
    If OutputQtyMode() <> "ACTUAL" Then NormalizeOutputYieldEditorDefaults

    If Trim$(mTxtProcessOutputId.Text) = "" Or Trim$(mTxtProcessOutputName.Text) = "" Or _
       Trim$(mTxtProcessOutputItemCode.Text) = "" Or ComboText(mCmbProcessOutputUom) = "" Then
        ShowStatus "An output needs its generated identity, name, and UOM."
        Exit Sub
    End If
    If Not OutputUomIsCatalogValue(ComboText(mCmbProcessOutputUom)) Then
        ShowStatus "Select an output UOM from the Recipe UOM Catalog."
        Exit Sub
    End If
    If OutputQtyMode() <> "ACTUAL" And Not PositiveTextValue(mTxtProcessOutputQty.Text) And Not PositiveTextValue(mTxtProcessOutputPercent.Text) Then
        ShowStatus "An output needs a positive quantity or percentage."
        Exit Sub
    End If
    If OutputQtyMode() <> "ACTUAL" And PositiveTextValue(mTxtProcessOutputPercent.Text) _
       And Not PositiveTextValue(mTxtProcessOutputYieldBasis.Text) Then
        ShowStatus "A percentage output needs a positive yield basis."
        Exit Sub
    End If
    If Not ValidateWholeQuantityForUom(mTxtProcessOutputQty.Text, _
            ComboText(mCmbProcessOutputUom), "Output quantity") Then Exit Sub
    If Not ValidateWholeQuantityForUom(mTxtProcessOutputYieldBasis.Text, _
            ComboText(mCmbProcessOutputUom), "Output yield basis") Then Exit Sub
    If Not updateExisting Or idx < 0 Then
        mLstProcessOutputs.AddItem ""
        idx = mLstProcessOutputs.ListCount - 1
    End If
    wasLoading = mLoading
    mLoading = True
    With mLstProcessOutputs
        .List(idx, 0) = Trim$(mTxtProcessOutputId.Text)
        .List(idx, 1) = Trim$(mTxtProcessOutputName.Text)
        .List(idx, 2) = Trim$(mTxtProcessOutputItemCode.Text)
        .List(idx, 3) = Trim$(mTxtProcessOutputDesignId.Text)
        .List(idx, 4) = Trim$(mTxtProcessOutputDesignVersion.Text)
        .List(idx, 5) = Trim$(mTxtProcessOutputQty.Text)
        .List(idx, 6) = Trim$(mTxtProcessOutputPercent.Text)
        .List(idx, 7) = Trim$(mTxtProcessOutputYieldBasis.Text)
        .List(idx, 8) = ComboText(mCmbProcessOutputUom)
        .List(idx, 9) = OutputQtyMode()
        .ListIndex = idx
    End With
    mLoading = wasLoading
    mSelectedProcessOutputIndex = idx
    If Not updateExisting Then ClearOutputEditor
    ShowStatus "Output staged in the Process draft."
End Sub

Private Function FindProcessComponentEditorRow(ByVal listControl As MSForms.ListBox, _
                                               ByVal componentId As String) As Long
    Dim rowIndex As Long

    FindProcessComponentEditorRow = -1
    If listControl Is Nothing Then Exit Function
    componentId = UCase$(Trim$(componentId))
    If componentId = "" Then Exit Function
    For rowIndex = 0 To listControl.ListCount - 1
        If StrComp(UCase$(Trim$(NzStr(listControl.List(rowIndex, 0)))), _
                   componentId, vbTextCompare) = 0 Then
            FindProcessComponentEditorRow = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub NormalizeOutputYieldEditorDefaults()
    Dim qtyText As String
    Dim percentText As String
    Dim yieldBasisText As String

    qtyText = Trim$(mTxtProcessOutputQty.Text)
    percentText = Trim$(mTxtProcessOutputPercent.Text)
    yieldBasisText = Trim$(mTxtProcessOutputYieldBasis.Text)
    mTxtProcessOutputPercent.Text = NormalizedOutputYieldPercent(qtyText, percentText)
    mTxtProcessOutputYieldBasis.Text = _
        NormalizedOutputYieldBasis(qtyText, mTxtProcessOutputPercent.Text, yieldBasisText)
End Sub

Private Function NormalizedOutputYieldPercent(ByVal qtyText As String, _
                                              ByVal percentText As String) As String
    percentText = Trim$(percentText)
    If PositiveTextValue(percentText) Then
        NormalizedOutputYieldPercent = percentText
    ElseIf PositiveTextValue(qtyText) Then
        NormalizedOutputYieldPercent = "100"
    End If
End Function

Private Function NormalizedOutputYieldBasis(ByVal qtyText As String, _
                                            ByVal percentText As String, _
                                            ByVal yieldBasisText As String) As String
    yieldBasisText = Trim$(yieldBasisText)
    If PositiveTextValue(yieldBasisText) Then
        NormalizedOutputYieldBasis = yieldBasisText
    ElseIf PositiveTextValue(qtyText) And PositiveTextValue(percentText) Then
        NormalizedOutputYieldBasis = Trim$(qtyText)
    End If
End Function

Private Function NextProcessComponentBase36Id() As String
    Dim usedIds() As String
    Dim rowIndex As Long
    Dim valueIndex As Long
    Dim usedCount As Long

    usedCount = mLstProcessRequirements.ListCount + mLstProcessOutputs.ListCount
    If usedCount = 0 Then
        NextProcessComponentBase36Id = mProduction.NextBase36Identifier(Array(""))
        Exit Function
    End If
    ReDim usedIds(0 To usedCount - 1)
    For rowIndex = 0 To mLstProcessRequirements.ListCount - 1
        usedIds(valueIndex) = NzStr(mLstProcessRequirements.List(rowIndex, 0))
        valueIndex = valueIndex + 1
    Next rowIndex
    For rowIndex = 0 To mLstProcessOutputs.ListCount - 1
        usedIds(valueIndex) = NzStr(mLstProcessOutputs.List(rowIndex, 0))
        valueIndex = valueIndex + 1
    Next rowIndex
    NextProcessComponentBase36Id = mProduction.NextBase36Identifier(usedIds)
End Function

Private Function ProcessComponentIdentityExists(ByVal identityText As String) As Boolean
    ProcessComponentIdentityExists = _
        ListIdentityExists(mLstProcessRequirements, 0, identityText) Or _
        ListIdentityExists(mLstProcessOutputs, 0, identityText)
End Function

Private Function NextListBase36Id(ByVal listControl As MSForms.ListBox, _
                                  ByVal identityColumn As Long) As String
    Dim usedIds() As String
    Dim rowIndex As Long

    If listControl Is Nothing Or listControl.ListCount = 0 Then
        NextListBase36Id = mProduction.NextBase36Identifier(Array(""))
        Exit Function
    End If
    ReDim usedIds(0 To listControl.ListCount - 1)
    For rowIndex = 0 To listControl.ListCount - 1
        usedIds(rowIndex) = NzStr(listControl.List(rowIndex, identityColumn))
    Next rowIndex
    NextListBase36Id = mProduction.NextBase36Identifier(usedIds)
End Function

Private Function NextProcessDraftBase36Id() As String
    Dim usedIds As Collection
    Dim usedArray() As String
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim valueIndex As Long

    Set usedIds = New Collection
    If Not mLstProcesses Is Nothing Then
        For rowIndex = 0 To mLstProcesses.ListCount - 1
            usedIds.Add NzStr(mLstProcesses.List(rowIndex, 0))
        Next rowIndex
    End If
    If Not mOperatorWorkbook Is Nothing Then
        For Each ws In mOperatorWorkbook.Worksheets
            For Each lo In ws.ListObjects
                If LCase$(Left$(lo.Name, 15)) = "invsys_process_" Then _
                    usedIds.Add Mid$(lo.Name, 16, 3)
            Next lo
        Next ws
    End If
    If usedIds.Count = 0 Then
        NextProcessDraftBase36Id = mProduction.NextBase36Identifier(Array(""))
        Exit Function
    End If
    ReDim usedArray(0 To usedIds.Count - 1)
    For valueIndex = 1 To usedIds.Count
        usedArray(valueIndex - 1) = usedIds(valueIndex)
    Next valueIndex
    NextProcessDraftBase36Id = mProduction.NextBase36Identifier(usedArray)
End Function

Private Function ListIdentityExists(ByVal listControl As MSForms.ListBox, _
                                    ByVal identityColumn As Long, _
                                    ByVal identityValue As String) As Boolean
    Dim rowIndex As Long

    If listControl Is Nothing Then Exit Function
    For rowIndex = 0 To listControl.ListCount - 1
        If StrComp(Trim$(NzStr(listControl.List(rowIndex, identityColumn))), _
                   Trim$(identityValue), vbTextCompare) = 0 Then
            ListIdentityExists = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function PositiveTextValue(ByVal textValue As String) As Boolean
    If Trim$(textValue) = "" Or Not IsNumeric(textValue) Then Exit Function
    PositiveTextValue = (CDbl(textValue) > 0)
End Function

Private Function ValidateWholeQuantityForUom(ByVal quantityText As String, _
                                             ByVal uomName As String, _
                                             ByVal fieldName As String) As Boolean
    Dim quantityValue As Double
    Dim report As String

    quantityText = Trim$(quantityText)
    If quantityText = "" Or Not IsNumeric(quantityText) Then
        ValidateWholeQuantityForUom = True
        Exit Function
    End If
    quantityValue = CDbl(quantityText)
    If Not modUomSettings.ValidateQuantityForUom(quantityValue, uomName, report) Then
        ShowStatus fieldName & ": " & report
        Exit Function
    End If
    ValidateWholeQuantityForUom = True
End Function

Private Sub RemoveSelectedListRow(ByVal listControl As MSForms.ListBox)
    If listControl Is Nothing Then Exit Sub
    If listControl.ListIndex >= 0 Then listControl.RemoveItem listControl.ListIndex
End Sub

Private Sub MoveSelectedListRow(ByVal listControl As MSForms.ListBox, ByVal direction As Long)
    Dim sourceIndex As Long
    Dim targetIndex As Long
    Dim columnIndex As Long
    Dim tempValue As Variant

    If listControl Is Nothing Then Exit Sub
    sourceIndex = listControl.ListIndex
    targetIndex = sourceIndex + direction
    If sourceIndex < 0 Or targetIndex < 0 Or targetIndex >= listControl.ListCount Then Exit Sub
    For columnIndex = 0 To listControl.ColumnCount - 1
        tempValue = listControl.List(sourceIndex, columnIndex)
        listControl.List(sourceIndex, columnIndex) = listControl.List(targetIndex, columnIndex)
        listControl.List(targetIndex, columnIndex) = tempValue
    Next columnIndex
    listControl.ListIndex = targetIndex
    RenumberInstructionOrdinals
End Sub

Private Sub RenumberInstructionOrdinals()
    Dim i As Long
    If mLstProcessInstructions Is Nothing Then Exit Sub
    For i = 0 To mLstProcessInstructions.ListCount - 1
        mLstProcessInstructions.List(i, 0) = CStr(i + 1)
    Next i
End Sub

Private Function ValidateProcessDraft(ByRef report As String) As Boolean
    Dim i As Long
    Dim usedComponentIds As Object
    Dim componentId As String

    If Trim$(mTxtProcessId.Text) = "" Or Trim$(mTxtProcessVersion.Text) = "" Or _
       Trim$(mTxtProcessName.Text) = "" Then
        report = "Process ID, version, and name are required."
        Exit Function
    End If
    If mLstProcessOutputs.ListCount = 0 Then
        report = "Every Process must declare at least one output."
        Exit Function
    End If
    Set usedComponentIds = CreateObject("Scripting.Dictionary")
    usedComponentIds.CompareMode = vbTextCompare
    For i = 0 To mLstProcessRequirements.ListCount - 1
        If Trim$(NzStr(mLstProcessRequirements.List(i, 0))) = "" Or _
           Trim$(NzStr(mLstProcessRequirements.List(i, 1))) = "" Or _
           Trim$(NzStr(mLstProcessRequirements.List(i, 5))) = "" Then
            report = "Each Process requirement requires identity, name, and UOM."
            Exit Function
        End If
        componentId = UCase$(Trim$(NzStr(mLstProcessRequirements.List(i, 0))))
        If Not mProduction.IsBase36Identifier(componentId) Then
            report = "Invalid Process requirement ID " & componentId & "."
            Exit Function
        End If
        If usedComponentIds.Exists(componentId) Then
            report = "Duplicate Process requirement/output ID " & componentId & "."
            Exit Function
        End If
        usedComponentIds.Add componentId, True
        If StrComp(NzStr(mLstProcessRequirements.List(i, 6)), "ACTUAL", vbTextCompare) <> 0 _
           And Not PositiveTextValue(NzStr(mLstProcessRequirements.List(i, 2))) _
           And Not PositiveTextValue(NzStr(mLstProcessRequirements.List(i, 3))) Then
            report = "Each Process requirement requires a positive quantity or percentage."
            Exit Function
        End If
        If StrComp(NzStr(mLstProcessRequirements.List(i, 6)), "ACTUAL", vbTextCompare) <> 0 _
           And PositiveTextValue(NzStr(mLstProcessRequirements.List(i, 3))) _
           And Not PositiveTextValue(NzStr(mLstProcessRequirements.List(i, 4))) Then
            report = "A percentage Process requirement requires a positive batch basis quantity."
            Exit Function
        End If
    Next i
    For i = 0 To mLstProcessOutputs.ListCount - 1
        If Trim$(NzStr(mLstProcessOutputs.List(i, 0))) = "" Or _
           Trim$(NzStr(mLstProcessOutputs.List(i, 1))) = "" Or _
           Trim$(NzStr(mLstProcessOutputs.List(i, 2))) = "" Or _
           Trim$(NzStr(mLstProcessOutputs.List(i, 8))) = "" Then
            report = "Each Process output requires identity, name, item code, and UOM."
            Exit Function
        End If
        componentId = UCase$(Trim$(NzStr(mLstProcessOutputs.List(i, 0))))
        If Not mProduction.IsBase36Identifier(componentId) Then
            report = "Invalid Process output ID " & componentId & "."
            Exit Function
        End If
        If usedComponentIds.Exists(componentId) Then
            report = "Duplicate Process requirement/output ID " & componentId & "."
            Exit Function
        End If
        usedComponentIds.Add componentId, True
        If StrComp(NzStr(mLstProcessOutputs.List(i, 9)), "ACTUAL", vbTextCompare) <> 0 _
           And Not PositiveTextValue(NzStr(mLstProcessOutputs.List(i, 5))) _
           And Not PositiveTextValue(NzStr(mLstProcessOutputs.List(i, 6))) Then
            report = "Each Process output requires a positive quantity or percentage."
            Exit Function
        End If
        If StrComp(NzStr(mLstProcessOutputs.List(i, 9)), "ACTUAL", vbTextCompare) <> 0 _
           And PositiveTextValue(NzStr(mLstProcessOutputs.List(i, 6))) _
           And Not PositiveTextValue(NzStr(mLstProcessOutputs.List(i, 7))) Then
            report = "A percentage Process output requires a positive yield basis."
            Exit Function
        End If
    Next i
    report = "Process draft is valid with " & CStr(mLstProcessRequirements.ListCount) & _
             " requirement(s), " & CStr(mLstProcessOutputs.ListCount) & _
             " output(s), and " & CStr(mLstProcessInstructions.ListCount) & " instruction(s)."
    ValidateProcessDraft = True
End Function

Private Function BuildProcessPayload() As String
    Dim records As New Collection
    Dim record As Object
    Dim alternative As Variant
    Dim i As Long

    Set record = NewReusableRecord("PROCESS")
    record("ProcessName") = Trim$(mTxtProcessName.Text)
    record("Description") = Trim$(mTxtProcessDescription.Text)
    records.Add record
    For i = 0 To mLstProcessRequirements.ListCount - 1
        Set record = NewReusableRecord("REQUIREMENT")
        record("RequirementId") = NzStr(mLstProcessRequirements.List(i, 0))
        record("RequirementName") = NzStr(mLstProcessRequirements.List(i, 1))
        AddNumericReusableField record, "Qty", NzStr(mLstProcessRequirements.List(i, 2))
        AddNumericReusableField record, "Percent", NzStr(mLstProcessRequirements.List(i, 3))
        AddNumericReusableField record, "YieldBasis", NzStr(mLstProcessRequirements.List(i, 4))
        record("UOM") = NzStr(mLstProcessRequirements.List(i, 5))
        record("RequirementQtyMode") = NzStr(mLstProcessRequirements.List(i, 6))
        records.Add record
    Next i
    If Not mProcessAlternatives Is Nothing Then
        For Each alternative In mProcessAlternatives
            records.Add CloneReusableRecord(alternative)
        Next alternative
    End If
    For i = 0 To mLstProcessOutputs.ListCount - 1
        Set record = NewReusableRecord("OUTPUT")
        record("OutputId") = NzStr(mLstProcessOutputs.List(i, 0))
        record("OutputName") = NzStr(mLstProcessOutputs.List(i, 1))
        record("ITEM_CODE") = NzStr(mLstProcessOutputs.List(i, 2))
        record("ComponentDesignId") = NzStr(mLstProcessOutputs.List(i, 3))
        record("ComponentDesignVersion") = NzStr(mLstProcessOutputs.List(i, 4))
        AddNumericReusableField record, "Qty", NzStr(mLstProcessOutputs.List(i, 5))
        AddNumericReusableField record, "Percent", NzStr(mLstProcessOutputs.List(i, 6))
        AddNumericReusableField record, "YieldBasis", NzStr(mLstProcessOutputs.List(i, 7))
        record("UOM") = NzStr(mLstProcessOutputs.List(i, 8))
        record("OutputQtyMode") = NzStr(mLstProcessOutputs.List(i, 9))
        record("OutputRegulationEnabled") = (StrComp(ProcessOutputRegulationValue(NzStr(mLstProcessOutputs.List(i, 0)), 0), "True", vbTextCompare) = 0)
        AddNumericReusableField record, "OutputFloorQty", ProcessOutputRegulationValue(NzStr(mLstProcessOutputs.List(i, 0)), 1)
        AddNumericReusableField record, "OutputCeilingQty", ProcessOutputRegulationValue(NzStr(mLstProcessOutputs.List(i, 0)), 2)
        records.Add record
    Next i
    For i = 0 To mLstProcessInstructions.ListCount - 1
        Set record = NewReusableRecord("INSTRUCTION")
        record("InstructionOrdinal") = i + 1
        record("Instruction") = NzStr(mLstProcessInstructions.List(i, 1))
        records.Add record
    Next i
    BuildProcessPayload = modProductionJson.BuildJsonArray(records)
End Function

Private Function NewReusableRecord(ByVal recordType As String) As Object
    Dim record As Object
    Set record = CreateObject("Scripting.Dictionary")
    record.CompareMode = vbTextCompare
    record("RecordType") = recordType
    Set NewReusableRecord = record
End Function

Private Sub AddNumericReusableField(ByVal record As Object, ByVal fieldName As String, _
                                    ByVal textValue As String)
    If Trim$(textValue) <> "" And IsNumeric(textValue) Then record(fieldName) = CDbl(textValue)
End Sub

Private Function SubmitProcessAction(ByVal eventType As String, _
                                     Optional ByVal payloadJson As String = "") As Boolean
    Dim report As String
    Dim quietStarted As Boolean

    On Error GoTo Failed
    If Not mOperatorWorkbook Is Nothing Then
        modOperationsPrimitiveBridge.BeginQuietUiForWorkbook mOperatorWorkbook.Name
        quietStarted = True
    End If
    ShowPersistencePending "Applying Process lifecycle action to warehouse storage..."
    SubmitProcessAction = modProductionReusableDesigns.SubmitReusableDesignEvent( _
        eventType, Trim$(mTxtProcessId.Text), Trim$(mTxtProcessVersion.Text), _
        payloadJson, "Production Process Designer", report)
    If quietStarted Then modUiQuiet.EndQuietUi
    RefreshReusableDesignLists
    ShowStatus report
    Exit Function
Failed:
    If quietStarted Then modUiQuiet.EndQuietUi
    ShowStatus "Process action failed: " & Err.Description
End Function

Private Sub ClearRecipeDraft(Optional ByVal createIdentity As Boolean = True)
    mTxtReusableRecipeName.Text = ""
    mTxtReusableRecipeId.Text = ""
    mTxtReusableRecipeVersion.Text = ""
    mTxtReusableRecipeDescription.Text = ""
    mLstRecipeNodes.Clear
    mLstRecipeConnections.Clear
    Set mRecipeOutputRegulations = New Collection
    RefreshRecipeConnectionDisplay
    mLstRecipeValidation.Clear
    ClearConnectionEditor
    If createIdentity Then EnsureRecipeDraftIdentity
End Sub

Private Sub EnsureRecipeDraftIdentity()
    If Trim$(mTxtReusableRecipeId.Text) = "" Then
        mTxtReusableRecipeId.Text = NextListBase36Id(mLstRecipes, 0)
    End If
    If Trim$(mTxtReusableRecipeVersion.Text) = "" Then
        mTxtReusableRecipeVersion.Text = _
            modProductionReusableDesigns.NextReusableDefinitionVersion( _
                Trim$(mTxtReusableRecipeId.Text), False)
    End If
End Sub

Private Function RecipeVersionIsValid(ByVal versionText As String) As Boolean
    Dim i As Long
    Dim ch As String

    On Error GoTo InvalidVersion
    versionText = Trim$(versionText)
    If versionText = "" Then Exit Function
    For i = 1 To Len(versionText)
        ch = Mid$(versionText, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i
    RecipeVersionIsValid = (CDbl(versionText) > 0#)
    Exit Function

InvalidVersion:
    RecipeVersionIsValid = False
End Function

Private Sub ClearConnectionEditor()
    mCmbConnectionFromNode.Clear
    mCmbConnectionOutput.Clear
    mCmbConnectionToNode.Clear
    mCmbConnectionRequirement.Clear
    mTxtConnectionQty.Text = ""
    mTxtConnectionPercent.Text = ""
    RefreshConnectionUomCatalog
End Sub

Private Sub LoadSelectedRecipeDefinition()
    Dim idx As Long
    idx = mLstRecipes.ListIndex
    If idx < 0 Then
        ShowStatus "Select a saved Recipe first."
        Exit Sub
    End If
    LoadRecipeDefinitionIntoDesigner NzStr(mLstRecipes.List(idx, 0)), _
        NzStr(mLstRecipes.List(idx, 1))
End Sub

Private Function LoadRecipeDefinitionIntoDesigner(ByVal recipeId As String, _
                                                  ByVal recipeVersion As String) As Boolean
    Dim jsonText As String
    Dim parseReport As String
    Dim records As Collection
    Dim record As Object
    Dim rowIndex As Long

    jsonText = modOperationsPrimitiveBridge.GetRecipeGraph(recipeId, recipeVersion)
    Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords(jsonText, parseReport)
    If records Is Nothing Then
        ShowStatus "Recipe load failed: " & parseReport
        Exit Function
    End If
    ClearRecipeDraft False
    For Each record In records
        Select Case UCase$(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"))
            Case "RECIPE"
                mTxtReusableRecipeId.Text = modProductionReusableDesigns.ReusableRecordText(record, "RecipeId")
                mTxtReusableRecipeVersion.Text = modProductionReusableDesigns.ReusableRecordText(record, "RecipeVersion")
                mTxtReusableRecipeName.Text = modProductionReusableDesigns.ReusableRecordText(record, "RecipeName")
                mTxtReusableRecipeDescription.Text = modProductionReusableDesigns.ReusableRecordText(record, "Description")
            Case "PROCESS_NODE"
                mLstRecipeNodes.AddItem modProductionReusableDesigns.ReusableRecordText(record, "ProcessNodeId")
                rowIndex = mLstRecipeNodes.ListCount - 1
                mLstRecipeNodes.List(rowIndex, 1) = modProductionReusableDesigns.ReusableRecordText(record, "ProcessId")
                mLstRecipeNodes.List(rowIndex, 2) = modProductionReusableDesigns.ReusableRecordText(record, "ProcessVersion")
                mLstRecipeNodes.List(rowIndex, 3) = ProcessNameForIdentity( _
                    mLstRecipeNodes.List(rowIndex, 1), mLstRecipeNodes.List(rowIndex, 2))
                mLstRecipeNodes.List(rowIndex, 4) = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "ExecutionOrdinal"))
            Case "CONNECTION"
                mLstRecipeConnections.AddItem modProductionReusableDesigns.ReusableRecordText(record, "FromProcessNodeId")
                rowIndex = mLstRecipeConnections.ListCount - 1
                mLstRecipeConnections.List(rowIndex, 1) = modProductionReusableDesigns.ReusableRecordText(record, "FromOutputId")
                mLstRecipeConnections.List(rowIndex, 2) = modProductionReusableDesigns.ReusableRecordText(record, "ToProcessNodeId")
                mLstRecipeConnections.List(rowIndex, 3) = modProductionReusableDesigns.ReusableRecordText(record, "ToRequirementId")
                mLstRecipeConnections.List(rowIndex, 4) = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Qty"))
                mLstRecipeConnections.List(rowIndex, 5) = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Percent"))
                mLstRecipeConnections.List(rowIndex, 6) = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
            Case "OUTPUT_REGULATION"
                If mRecipeOutputRegulations Is Nothing Then Set mRecipeOutputRegulations = New Collection
                mRecipeOutputRegulations.Add CloneReusableRecord(record)
        End Select
    Next record
    RefreshConnectionNodeCombos
    RefreshRecipeConnectionDisplay
    RefreshOutputRegulationSettings
    ShowStatus "Loaded Recipe " & recipeId & " version " & recipeVersion & "."
    LoadRecipeDefinitionIntoDesigner = True
End Function

Private Function ProcessNameForIdentity(ByVal processId As String, ByVal processVersion As String) As String
    Dim rows As Variant
    Dim r As Long

    rows = modOperationsPrimitiveBridge.ListProcesses("")
    If Not IsArray(rows) Then Exit Function
    On Error GoTo NoRows
    For r = LBound(rows, 1) To UBound(rows, 1)
        If StrComp(NzStr(rows(r, 1)), processId, vbTextCompare) = 0 _
           And StrComp(NzStr(rows(r, 2)), processVersion, vbTextCompare) = 0 Then
            ProcessNameForIdentity = NzStr(rows(r, 3))
            Exit Function
        End If
    Next r
NoRows:
End Function

Private Sub AddSelectedReleasedProcessToRecipe()
    Dim idx As Long
    Dim rowIndex As Long
    Dim nodeId As String

    idx = mLstReleasedProcesses.ListIndex
    If idx < 0 Then
        ShowStatus "Select a released Process first."
        Exit Sub
    End If
    nodeId = NextRecipeNodeId()
    mLstRecipeNodes.AddItem nodeId
    rowIndex = mLstRecipeNodes.ListCount - 1
    mLstRecipeNodes.List(rowIndex, 1) = NzStr(mLstReleasedProcesses.List(idx, 0))
    mLstRecipeNodes.List(rowIndex, 2) = NzStr(mLstReleasedProcesses.List(idx, 1))
    mLstRecipeNodes.List(rowIndex, 3) = NzStr(mLstReleasedProcesses.List(idx, 2))
    mLstRecipeNodes.List(rowIndex, 4) = CStr(rowIndex + 1)
    mLstRecipeNodes.ListIndex = rowIndex
    RefreshConnectionNodeCombos
    ShowStatus "Added released Process as Recipe node " & nodeId & "."
End Sub

Private Function NextRecipeNodeId() As String
    Dim candidate As Long
    candidate = mLstRecipeNodes.ListCount + 1
    Do While RecipeNodeIndex("N" & CStr(candidate)) >= 0
        candidate = candidate + 1
    Loop
    NextRecipeNodeId = "N" & CStr(candidate)
End Function

Private Function RecipeNodeIndex(ByVal nodeId As String) As Long
    Dim i As Long
    RecipeNodeIndex = -1
    For i = 0 To mLstRecipeNodes.ListCount - 1
        If StrComp(NzStr(mLstRecipeNodes.List(i, 0)), nodeId, vbTextCompare) = 0 Then
            RecipeNodeIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub RemoveSelectedRecipeNode()
    Dim nodeId As String
    Dim i As Long

    If mLstRecipeNodes.ListIndex < 0 Then Exit Sub
    nodeId = NzStr(mLstRecipeNodes.List(mLstRecipeNodes.ListIndex, 0))
    For i = mLstRecipeConnections.ListCount - 1 To 0 Step -1
        If StrComp(NzStr(mLstRecipeConnections.List(i, 0)), nodeId, vbTextCompare) = 0 _
           Or StrComp(NzStr(mLstRecipeConnections.List(i, 2)), nodeId, vbTextCompare) = 0 Then
            mLstRecipeConnections.RemoveItem i
        End If
    Next i
    mLstRecipeNodes.RemoveItem mLstRecipeNodes.ListIndex
    RenumberRecipeExecutionOrder
    RefreshConnectionNodeCombos
    RefreshRecipeConnectionDisplay
End Sub

Private Sub RenumberRecipeExecutionOrder()
    Dim i As Long
    For i = 0 To mLstRecipeNodes.ListCount - 1
        mLstRecipeNodes.List(i, 4) = CStr(i + 1)
    Next i
End Sub

Private Sub RefreshConnectionNodeCombos()
    Dim i As Long
    Dim selectedFrom As String
    Dim selectedTo As String
    Dim selectedRequirement As String
    Dim selectedOutput As String

    selectedFrom = ComboText(mCmbConnectionFromNode)
    selectedTo = ConnectionTargetNodeId()
    selectedRequirement = ConnectionRequirementId()
    selectedOutput = ConnectionOutputId()
    mLoading = True
    mCmbConnectionFromNode.Clear
    mCmbConnectionToNode.Clear
    mCmbConnectionRequirement.Clear
    For i = 0 To mLstRecipeNodes.ListCount - 1
        mCmbConnectionFromNode.AddItem NzStr(mLstRecipeNodes.List(i, 0))
        mCmbConnectionFromNode.List(mCmbConnectionFromNode.ListCount - 1, 1) = _
            NzStr(mLstRecipeNodes.List(i, 3))
    Next i
    SelectComboText mCmbConnectionFromNode, selectedFrom
    If mCmbConnectionFromNode.ListIndex < 0 And mCmbConnectionFromNode.ListCount > 0 Then mCmbConnectionFromNode.ListIndex = 0
    FillConnectionRecordChoices mCmbConnectionOutput, ComboText(mCmbConnectionFromNode), _
        "OUTPUT", "OutputId", "OutputName"
    SelectComboText mCmbConnectionOutput, selectedOutput
    mLoading = False
    RefreshCompatibleDownstreamChoices selectedTo, selectedRequirement
End Sub

Private Sub SelectComboText(ByVal combo As MSForms.ComboBox, ByVal textValue As String)
    Dim i As Long
    For i = 0 To combo.ListCount - 1
        If StrComp(NzStr(combo.List(i)), textValue, vbTextCompare) = 0 Then
            combo.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub RefreshConnectionOutputChoices()
    Dim selectedOutput As String

    selectedOutput = ConnectionOutputId()
    mLoading = True
    FillConnectionRecordChoices mCmbConnectionOutput, ComboText(mCmbConnectionFromNode), _
        "OUTPUT", "OutputId", "OutputName"
    SelectComboText mCmbConnectionOutput, selectedOutput
    mLoading = False
    RefreshCompatibleDownstreamChoices
End Sub

Private Sub RefreshConnectionRequirementChoices()
    BindSelectedCompatibleRequirement
End Sub

Private Sub RefreshCompatibleDownstreamChoices(Optional ByVal selectedNodeId As String = "", _
                                               Optional ByVal selectedRequirementId As String = "")
    Dim outputItemCode As String
    Dim outputUom As String
    Dim outputQty As String
    Dim outputPercent As String
    Dim sourceNodeId As String
    Dim nodeIndex As Long
    Dim records As Collection
    Dim record As Object
    Dim report As String
    Dim requirementId As String
    Dim requirementName As String
    Dim requirementUom As String
    Dim compatibleCount As Long
    Dim matchedRequirementId As String
    Dim matchedRequirementName As String
    Dim rowIndex As Long
    Dim ignoreConnectionIndex As Long

    If mCmbConnectionToNode Is Nothing Then Exit Sub
    sourceNodeId = ComboText(mCmbConnectionFromNode)
    If selectedNodeId = "" Then selectedNodeId = ConnectionTargetNodeId()
    If selectedRequirementId = "" Then selectedRequirementId = ConnectionRequirementId()
    mLoading = True
    mCmbConnectionToNode.Clear
    mCmbConnectionRequirement.Clear
    mLoading = False
    If Not TrySelectedOutputRoutingInfo(outputItemCode, outputUom, outputQty, outputPercent) Then Exit Sub

    ignoreConnectionIndex = mLstRecipeConnections.ListIndex
    For nodeIndex = 0 To mLstRecipeNodes.ListCount - 1
        If StrComp(NzStr(mLstRecipeNodes.List(nodeIndex, 0)), sourceNodeId, vbTextCompare) <> 0 Then
            Set records = ProcessRecordsForRecipeNode(nodeIndex, report)
            compatibleCount = 0
            matchedRequirementId = ""
            matchedRequirementName = ""
            If Not records Is Nothing Then
                For Each record In records
                    If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"), _
                            "REQUIREMENT", vbTextCompare) = 0 Then
                        requirementId = modProductionReusableDesigns.ReusableRecordText(record, "RequirementId")
                        requirementName = modProductionReusableDesigns.ReusableRecordText(record, "RequirementName")
                        requirementUom = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
                        If StrComp(requirementUom, outputUom, vbTextCompare) = 0 _
                           And ProcessRecordsHaveAlternativeItem(records, requirementId, outputItemCode) _
                           And Not RecipeRequirementAlreadyConnected( _
                               NzStr(mLstRecipeNodes.List(nodeIndex, 0)), requirementId, _
                               ignoreConnectionIndex) Then
                            compatibleCount = compatibleCount + 1
                            matchedRequirementId = requirementId
                            matchedRequirementName = requirementName
                        End If
                    End If
                Next record
            End If
            If compatibleCount = 1 Then
                mCmbConnectionToNode.AddItem NzStr(mLstRecipeNodes.List(nodeIndex, 0))
                rowIndex = mCmbConnectionToNode.ListCount - 1
                mCmbConnectionToNode.List(rowIndex, 1) = matchedRequirementId
                mCmbConnectionToNode.List(rowIndex, 2) = NzStr(mLstRecipeNodes.List(nodeIndex, 3))
                mCmbConnectionToNode.List(rowIndex, 3) = matchedRequirementName
            End If
        End If
    Next nodeIndex
    SelectCompatibleTarget selectedNodeId, selectedRequirementId
    If mCmbConnectionToNode.ListIndex < 0 And mCmbConnectionToNode.ListCount = 1 Then
        mCmbConnectionToNode.ListIndex = 0
    End If
    BindSelectedCompatibleRequirement
End Sub

Private Function TrySelectedOutputRoutingInfo(ByRef itemCode As String, _
                                              ByRef outputUom As String, _
                                              ByRef outputQty As String, _
                                              ByRef outputPercent As String) As Boolean
    Dim nodeIndex As Long
    Dim records As Collection
    Dim record As Object
    Dim report As String
    Dim outputId As String

    nodeIndex = RecipeNodeIndex(ComboText(mCmbConnectionFromNode))
    outputId = ConnectionOutputId()
    If nodeIndex < 0 Or outputId = "" Then Exit Function
    Set records = ProcessRecordsForRecipeNode(nodeIndex, report)
    If records Is Nothing Then Exit Function
    For Each record In records
        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"), _
                   "OUTPUT", vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(record, "OutputId"), _
                       outputId, vbTextCompare) = 0 Then
            itemCode = modProductionReusableDesigns.ReusableRecordText(record, "ITEM_CODE")
            outputUom = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
            outputQty = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Qty"))
            outputPercent = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Percent"))
            TrySelectedOutputRoutingInfo = (itemCode <> "" And outputUom <> "")
            Exit Function
        End If
    Next record
End Function

Private Function ProcessRecordsHaveAlternativeItem(ByVal records As Collection, _
                                                   ByVal requirementId As String, _
                                                   ByVal itemCode As String) As Boolean
    Dim record As Object

    For Each record In records
        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"), _
                   "ALTERNATIVE", vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RequirementId"), _
                       requirementId, vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(record, "ITEM_CODE"), _
                       itemCode, vbTextCompare) = 0 Then
            ProcessRecordsHaveAlternativeItem = True
            Exit Function
        End If
    Next record
End Function

Private Function RecipeRequirementAlreadyConnected(ByVal nodeId As String, _
                                                   ByVal requirementId As String, _
                                                   ByVal ignoreIndex As Long) As Boolean
    Dim i As Long

    For i = 0 To mLstRecipeConnections.ListCount - 1
        If i <> ignoreIndex _
           And StrComp(NzStr(mLstRecipeConnections.List(i, 2)), nodeId, vbTextCompare) = 0 _
           And StrComp(NzStr(mLstRecipeConnections.List(i, 3)), requirementId, vbTextCompare) = 0 Then
            RecipeRequirementAlreadyConnected = True
            Exit Function
        End If
    Next i
End Function

Private Sub SelectCompatibleTarget(ByVal nodeId As String, ByVal requirementId As String)
    Dim i As Long

    For i = 0 To mCmbConnectionToNode.ListCount - 1
        If StrComp(NzStr(mCmbConnectionToNode.List(i, 0)), nodeId, vbTextCompare) = 0 _
           And StrComp(NzStr(mCmbConnectionToNode.List(i, 1)), requirementId, vbTextCompare) = 0 Then
            mCmbConnectionToNode.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

Private Sub BindSelectedCompatibleRequirement()
    Dim nodeIndex As Long
    Dim records As Collection
    Dim record As Object
    Dim report As String
    Dim requirementId As String
    Dim requirementUom As String
    Dim requirementQty As String
    Dim requirementPercent As String

    mCmbConnectionRequirement.Clear
    requirementId = ConnectionRequirementId()
    If requirementId = "" Then Exit Sub
    mCmbConnectionRequirement.AddItem requirementId
    mCmbConnectionRequirement.ListIndex = 0
    nodeIndex = RecipeNodeIndex(ConnectionTargetNodeId())
    If nodeIndex < 0 Then Exit Sub
    Set records = ProcessRecordsForRecipeNode(nodeIndex, report)
    If records Is Nothing Then Exit Sub
    For Each record In records
        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"), _
                   "REQUIREMENT", vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RequirementId"), _
                       requirementId, vbTextCompare) = 0 Then
            requirementQty = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Qty"))
            requirementPercent = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Percent"))
            requirementUom = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
            mTxtConnectionQty.Text = requirementQty
            mTxtConnectionPercent.Text = requirementPercent
            RefreshConnectionUomCatalog requirementUom
            Exit Sub
        End If
    Next record
End Sub

Private Sub FillConnectionRecordChoices(ByVal combo As MSForms.ComboBox, ByVal nodeId As String, _
                                        ByVal recordType As String, ByVal idField As String, _
                                        Optional ByVal displayField As String = "")
    Dim nodeIndex As Long
    Dim records As Collection
    Dim record As Object
    Dim report As String
    Dim jsonText As String
    Dim rowIndex As Long

    If combo Is Nothing Then Exit Sub
    combo.Clear
    nodeIndex = RecipeNodeIndex(nodeId)
    If nodeIndex < 0 Then Exit Sub
    jsonText = modOperationsPrimitiveBridge.GetProcessVersion( _
        NzStr(mLstRecipeNodes.List(nodeIndex, 1)), NzStr(mLstRecipeNodes.List(nodeIndex, 2)))
    Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords(jsonText, report)
    If records Is Nothing Then Exit Sub
    For Each record In records
        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"), _
                recordType, vbTextCompare) = 0 Then
            combo.AddItem modProductionReusableDesigns.ReusableRecordText(record, idField)
            rowIndex = combo.ListCount - 1
            If displayField <> "" And combo.ColumnCount > 1 Then
                combo.List(rowIndex, 1) = _
                    modProductionReusableDesigns.ReusableRecordText(record, displayField)
            End If
        End If
    Next record
    If combo.ListCount > 0 Then combo.ListIndex = 0
End Sub

Private Function RecipeNodeDisplayName(ByVal nodeId As String) As String
    Dim nodeIndex As Long

    nodeIndex = RecipeNodeIndex(nodeId)
    If nodeIndex < 0 Then
        RecipeNodeDisplayName = "(missing Process)"
    Else
        RecipeNodeDisplayName = Trim$(NzStr(mLstRecipeNodes.List(nodeIndex, 3)))
        If RecipeNodeDisplayName = "" Then RecipeNodeDisplayName = "(unnamed Process)"
    End If
End Function

Private Function RecipeNodeRecordDisplayName(ByVal nodeId As String, _
                                             ByVal recordType As String, _
                                             ByVal idField As String, _
                                             ByVal idValue As String, _
                                             ByVal displayField As String) As String
    Dim nodeIndex As Long
    Dim records As Collection
    Dim record As Object
    Dim report As String
    Dim jsonText As String

    nodeIndex = RecipeNodeIndex(nodeId)
    If nodeIndex < 0 Then GoTo MissingName
    jsonText = modOperationsPrimitiveBridge.GetProcessVersion( _
        NzStr(mLstRecipeNodes.List(nodeIndex, 1)), NzStr(mLstRecipeNodes.List(nodeIndex, 2)))
    Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords(jsonText, report)
    If records Is Nothing Then GoTo MissingName
    For Each record In records
        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"), _
                   recordType, vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(record, idField), _
                       idValue, vbTextCompare) = 0 Then
            RecipeNodeRecordDisplayName = _
                modProductionReusableDesigns.ReusableRecordText(record, displayField)
            If Trim$(RecipeNodeRecordDisplayName) <> "" Then Exit Function
        End If
    Next record

MissingName:
    If StrComp(recordType, "OUTPUT", vbTextCompare) = 0 Then
        RecipeNodeRecordDisplayName = "(missing output)"
    Else
        RecipeNodeRecordDisplayName = "(missing input requirement)"
    End If
End Function

Private Sub RefreshRecipeConnectionDisplay(Optional ByVal selectedIndex As Long = -1)
    Dim stageMap As Object
    Dim stageNumber As Long
    Dim nodeIndex As Long
    Dim nodeId As String
    Dim records As Collection
    Dim record As Object
    Dim report As String
    Dim outputId As String
    Dim outputName As String
    Dim outputQty As String
    Dim outputPercent As String
    Dim outputUom As String
    Dim i As Long
    Dim rowIndex As Long
    Dim hasOutgoing As Boolean
    Dim wasLoading As Boolean

    If mLstRecipeConnectionDisplay Is Nothing Then Exit Sub
    wasLoading = mLoading
    mLoading = True
    mLstRecipeConnectionDisplay.Clear
    Set stageMap = BuildRecipeStageMap()
    For stageNumber = 1 To mLstRecipeNodes.ListCount
        For nodeIndex = 0 To mLstRecipeNodes.ListCount - 1
            nodeId = NzStr(mLstRecipeNodes.List(nodeIndex, 0))
            If stageMap.Exists(nodeId) And CLng(stageMap(nodeId)) = stageNumber Then
                Set records = ProcessRecordsForRecipeNode(nodeIndex, report)
                If Not records Is Nothing Then
                    For Each record In records
                        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"), _
                                   "OUTPUT", vbTextCompare) = 0 Then
                            outputId = modProductionReusableDesigns.ReusableRecordText(record, "OutputId")
                            outputName = modProductionReusableDesigns.ReusableRecordText(record, "OutputName")
                            outputQty = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Qty"))
                            outputPercent = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Percent"))
                            outputPercent = NormalizedOutputYieldPercent(outputQty, outputPercent)
                            outputUom = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
                            hasOutgoing = False
                            For i = 0 To mLstRecipeConnections.ListCount - 1
                                If StrComp(NzStr(mLstRecipeConnections.List(i, 0)), nodeId, vbTextCompare) = 0 _
                                   And StrComp(NzStr(mLstRecipeConnections.List(i, 1)), outputId, vbTextCompare) = 0 Then
                                    AddRecipeOutputFlowRow stageNumber, nodeId, outputId, outputName, _
                                        RecipeNodeDisplayName(NzStr(mLstRecipeConnections.List(i, 2))), _
                                        outputQty, outputPercent, outputUom, i
                                    hasOutgoing = True
                                End If
                            Next i
                            If Not hasOutgoing Then
                                AddRecipeOutputFlowRow stageNumber, nodeId, outputId, outputName, _
                                    "Finished inventory", outputQty, outputPercent, outputUom, -1
                            End If
                        End If
                    Next record
                End If
            End If
        Next nodeIndex
    Next stageNumber
    If selectedIndex >= 0 Then
        For rowIndex = 0 To mLstRecipeConnectionDisplay.ListCount - 1
            If CLng(Val(NzStr(mLstRecipeConnectionDisplay.List(rowIndex, 7)))) = selectedIndex Then
                mLstRecipeConnectionDisplay.ListIndex = rowIndex
                Exit For
            End If
        Next rowIndex
    End If
    mLoading = wasLoading
End Sub

Private Sub AddRecipeOutputFlowRow(ByVal stageNumber As Long, ByVal sourceNodeId As String, _
                                   ByVal outputId As String, ByVal outputName As String, _
                                   ByVal feedsProcess As String, ByVal qty As String, _
                                   ByVal percent As String, ByVal uom As String, _
                                   ByVal connectionIndex As Long)
    Dim rowIndex As Long

    mLstRecipeConnectionDisplay.AddItem "Stage " & CStr(stageNumber)
    rowIndex = mLstRecipeConnectionDisplay.ListCount - 1
    mLstRecipeConnectionDisplay.List(rowIndex, 1) = RecipeNodeDisplayName(sourceNodeId)
    mLstRecipeConnectionDisplay.List(rowIndex, 2) = outputName
    mLstRecipeConnectionDisplay.List(rowIndex, 3) = feedsProcess
    mLstRecipeConnectionDisplay.List(rowIndex, 4) = qty
    mLstRecipeConnectionDisplay.List(rowIndex, 5) = percent
    mLstRecipeConnectionDisplay.List(rowIndex, 6) = uom
    mLstRecipeConnectionDisplay.List(rowIndex, 7) = CStr(connectionIndex)
    mLstRecipeConnectionDisplay.List(rowIndex, 8) = sourceNodeId
    mLstRecipeConnectionDisplay.List(rowIndex, 9) = outputId
End Sub

Private Function BuildRecipeStageMap() As Object
    Dim stages As Object
    Dim i As Long
    Dim pass As Long
    Dim sourceNodeId As String
    Dim targetNodeId As String
    Dim sourceStage As Long
    Dim targetStage As Long

    Set stages = CreateObject("Scripting.Dictionary")
    stages.CompareMode = vbTextCompare
    For i = 0 To mLstRecipeNodes.ListCount - 1
        stages(NzStr(mLstRecipeNodes.List(i, 0))) = 1
    Next i
    For pass = 1 To mLstRecipeNodes.ListCount
        For i = 0 To mLstRecipeConnections.ListCount - 1
            sourceNodeId = NzStr(mLstRecipeConnections.List(i, 0))
            targetNodeId = NzStr(mLstRecipeConnections.List(i, 2))
            If stages.Exists(sourceNodeId) And stages.Exists(targetNodeId) Then
                sourceStage = CLng(stages(sourceNodeId))
                targetStage = sourceStage + 1
                If CLng(stages(targetNodeId)) < targetStage Then stages(targetNodeId) = targetStage
            End If
        Next i
    Next pass
    Set BuildRecipeStageMap = stages
End Function

Private Function ConnectionOutputId() As String
    If mCmbConnectionOutput Is Nothing Then Exit Function
    If mCmbConnectionOutput.ListIndex < 0 Then Exit Function
    ConnectionOutputId = Trim$(NzStr( _
        mCmbConnectionOutput.List(mCmbConnectionOutput.ListIndex, 0)))
End Function

Private Function ConnectionOutputDisplayName() As String
    If mCmbConnectionOutput Is Nothing Then Exit Function
    If mCmbConnectionOutput.ListIndex < 0 Then Exit Function
    ConnectionOutputDisplayName = Trim$(NzStr( _
        mCmbConnectionOutput.List(mCmbConnectionOutput.ListIndex, 1)))
End Function

Private Function ConnectionTargetNodeId() As String
    If mCmbConnectionToNode Is Nothing Then Exit Function
    If mCmbConnectionToNode.ListIndex < 0 Then Exit Function
    ConnectionTargetNodeId = Trim$(NzStr( _
        mCmbConnectionToNode.List(mCmbConnectionToNode.ListIndex, 0)))
End Function

Private Function ConnectionRequirementId() As String
    If mCmbConnectionToNode Is Nothing Then Exit Function
    If mCmbConnectionToNode.ListIndex < 0 Then Exit Function
    ConnectionRequirementId = Trim$(NzStr( _
        mCmbConnectionToNode.List(mCmbConnectionToNode.ListIndex, 1)))
End Function

Private Sub WriteConnectionEditorToList(ByVal updateExisting As Boolean)
    Dim idx As Long
    Dim ignoreIndex As Long

    If ComboText(mCmbConnectionFromNode) = "" Or ConnectionOutputId() = "" Or _
       ConnectionTargetNodeId() = "" Or ConnectionRequirementId() = "" Then
        ShowStatus "Select an output and one compatible Feeds Process target."
        Exit Sub
    End If
    If StrComp(ComboText(mCmbConnectionFromNode), ConnectionTargetNodeId(), vbTextCompare) = 0 Then
        ShowStatus "A Process output cannot connect back to the same Recipe node."
        Exit Sub
    End If
    If Not PositiveTextValue(mTxtConnectionQty.Text) And Not PositiveTextValue(mTxtConnectionPercent.Text) Then
        ShowStatus "A connection needs a positive quantity or percentage."
        Exit Sub
    End If
    If ComboText(mCmbConnectionUom) = "" Then
        ShowStatus "Select a connection UOM from the Recipe UOM Catalog."
        Exit Sub
    End If
    If Not ValidateWholeQuantityForUom(mTxtConnectionQty.Text, _
            ComboText(mCmbConnectionUom), "Connection quantity") Then Exit Sub
    idx = mLstRecipeConnections.ListIndex
    ignoreIndex = -1
    If updateExisting Then ignoreIndex = idx
    If RecipeRequirementAlreadyConnected(ConnectionTargetNodeId(), _
            ConnectionRequirementId(), ignoreIndex) Then
        ShowStatus "That Feeds Process input is already supplied by another output."
        Exit Sub
    End If
    If Not updateExisting Or idx < 0 Then
        mLstRecipeConnections.AddItem ""
        idx = mLstRecipeConnections.ListCount - 1
    End If
    With mLstRecipeConnections
        .List(idx, 0) = ComboText(mCmbConnectionFromNode)
        .List(idx, 1) = ConnectionOutputId()
        .List(idx, 2) = ConnectionTargetNodeId()
        .List(idx, 3) = ConnectionRequirementId()
        .List(idx, 4) = Trim$(mTxtConnectionQty.Text)
        .List(idx, 5) = Trim$(mTxtConnectionPercent.Text)
        .List(idx, 6) = ComboText(mCmbConnectionUom)
        .ListIndex = idx
    End With
    RefreshRecipeConnectionDisplay idx
    ShowStatus "Recipe connection staged."
End Sub

Private Function ValidateRecipeDraft(ByRef report As String, _
                                     Optional ByVal requireResolved As Boolean = False) As Boolean
    Dim i As Long
    Dim fromIndex As Long
    Dim toIndex As Long

    mLstRecipeValidation.Clear
    If Trim$(mTxtReusableRecipeId.Text) = "" Or Trim$(mTxtReusableRecipeVersion.Text) = "" Or _
       Trim$(mTxtReusableRecipeName.Text) = "" Then
        report = "Recipe ID, version, and name are required."
        AddRecipeValidationIssue "IDENTITY", report
        Exit Function
    End If
    If Not RecipeVersionIsValid(mTxtReusableRecipeVersion.Text) Then
        report = "Recipe version must be a positive whole number."
        AddRecipeValidationIssue "VERSION", report
        Exit Function
    End If
    If mLstRecipeNodes.ListCount = 0 Then
        report = "Every Recipe must select at least one released Process version."
        AddRecipeValidationIssue "PROCESS_REQUIRED", report
        Exit Function
    End If
    For i = 0 To mLstRecipeConnections.ListCount - 1
        fromIndex = RecipeNodeIndex(NzStr(mLstRecipeConnections.List(i, 0)))
        toIndex = RecipeNodeIndex(NzStr(mLstRecipeConnections.List(i, 2)))
        If fromIndex < 0 Or toIndex < 0 Then
            report = "A connection references a missing Process node."
            AddRecipeValidationIssue "MISSING_NODE", report
            Exit Function
        End If
        If CLng(Val(NzStr(mLstRecipeNodes.List(fromIndex, 4)))) >= _
           CLng(Val(NzStr(mLstRecipeNodes.List(toIndex, 4)))) Then
            report = "Execution order must place each source Process before its downstream Process."
            AddRecipeValidationIssue "EXECUTION_ORDER", report
            Exit Function
        End If
    Next i
    If Not RecipeGraphIsAcyclic() Then
        report = "Recipe Process connections contain a circular dependency."
        AddRecipeValidationIssue "CIRCULAR_DEPENDENCY", report
        Exit Function
    End If
    If requireResolved And Not RecipeRequirementsResolved(report) Then
        AddRecipeValidationIssue "UNRESOLVED_INPUT", report
        Exit Function
    End If
    report = "Recipe graph is valid: nodes=" & CStr(mLstRecipeNodes.ListCount) & _
             "; connections=" & CStr(mLstRecipeConnections.ListCount) & "."
    AddRecipeValidationIssue "OK", report
    ValidateRecipeDraft = True
End Function

Private Sub AddRecipeValidationIssue(ByVal code As String, ByVal detail As String)
    mLstRecipeValidation.AddItem code
    mLstRecipeValidation.List(mLstRecipeValidation.ListCount - 1, 1) = detail
End Sub

Private Function RecipeGraphIsAcyclic() As Boolean
    Dim indegree As Object
    Dim processed As Object
    Dim i As Long
    Dim changed As Boolean
    Dim nodeId As String
    Dim sourceId As String
    Dim targetId As String
    Dim processedCount As Long
    Dim edgeIndex As Long

    Set indegree = CreateObject("Scripting.Dictionary")
    indegree.CompareMode = vbTextCompare
    Set processed = CreateObject("Scripting.Dictionary")
    processed.CompareMode = vbTextCompare
    For i = 0 To mLstRecipeNodes.ListCount - 1
        indegree(NzStr(mLstRecipeNodes.List(i, 0))) = 0
    Next i
    For i = 0 To mLstRecipeConnections.ListCount - 1
        targetId = NzStr(mLstRecipeConnections.List(i, 2))
        If indegree.Exists(targetId) Then indegree(targetId) = CLng(indegree(targetId)) + 1
    Next i
    Do
        changed = False
        For i = 0 To mLstRecipeNodes.ListCount - 1
            nodeId = NzStr(mLstRecipeNodes.List(i, 0))
            If Not processed.Exists(nodeId) And CLng(indegree(nodeId)) = 0 Then
                processed(nodeId) = True
                processedCount = processedCount + 1
                changed = True
                For edgeIndex = 0 To mLstRecipeConnections.ListCount - 1
                    sourceId = NzStr(mLstRecipeConnections.List(edgeIndex, 0))
                    targetId = NzStr(mLstRecipeConnections.List(edgeIndex, 2))
                    If StrComp(sourceId, nodeId, vbTextCompare) = 0 And indegree.Exists(targetId) Then
                        indegree(targetId) = CLng(indegree(targetId)) - 1
                    End If
                Next edgeIndex
            End If
        Next i
    Loop While changed
    RecipeGraphIsAcyclic = (processedCount = mLstRecipeNodes.ListCount)
End Function

Private Function RecipeRequirementsResolved(ByRef report As String) As Boolean
    Dim nodeIndex As Long
    Dim records As Collection
    Dim record As Object
    Dim reqId As String
    Dim parseReport As String

    For nodeIndex = 0 To mLstRecipeNodes.ListCount - 1
        Set records = ProcessRecordsForRecipeNode(nodeIndex, parseReport)
        If records Is Nothing Then
            report = parseReport
            Exit Function
        End If
        For Each record In records
            If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"), _
                    "REQUIREMENT", vbTextCompare) = 0 Then
                reqId = modProductionReusableDesigns.ReusableRecordText(record, "RequirementId")
                If Not RecipeHasIncomingRequirement(NzStr(mLstRecipeNodes.List(nodeIndex, 0)), reqId) _
                   And Not ProcessRecordsHaveAlternative(records, reqId) Then
                    report = "Unresolved requirement " & reqId & " on Process node " & _
                             NzStr(mLstRecipeNodes.List(nodeIndex, 0)) & "."
                    Exit Function
                End If
            End If
        Next record
    Next nodeIndex
    RecipeRequirementsResolved = True
End Function

Private Function ProcessRecordsForRecipeNode(ByVal nodeIndex As Long, _
                                             ByRef report As String) As Collection
    Dim jsonText As String
    jsonText = modOperationsPrimitiveBridge.GetProcessVersion( _
        NzStr(mLstRecipeNodes.List(nodeIndex, 1)), NzStr(mLstRecipeNodes.List(nodeIndex, 2)))
    Set ProcessRecordsForRecipeNode = _
        modProductionReusableDesigns.ParseReusableDefinitionRecords(jsonText, report)
End Function

Private Function RecipeHasIncomingRequirement(ByVal nodeId As String, ByVal requirementId As String) As Boolean
    Dim i As Long
    For i = 0 To mLstRecipeConnections.ListCount - 1
        If StrComp(NzStr(mLstRecipeConnections.List(i, 2)), nodeId, vbTextCompare) = 0 _
           And StrComp(NzStr(mLstRecipeConnections.List(i, 3)), requirementId, vbTextCompare) = 0 Then
            RecipeHasIncomingRequirement = True
            Exit Function
        End If
    Next i
End Function

Private Function ProcessRecordsHaveAlternative(ByVal records As Collection, _
                                               ByVal requirementId As String) As Boolean
    Dim record As Object
    For Each record In records
        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"), _
                "ALTERNATIVE", vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(record, "RequirementId"), _
                requirementId, vbTextCompare) = 0 Then
            ProcessRecordsHaveAlternative = True
            Exit Function
        End If
    Next record
End Function

Private Function BuildRecipePayload() As String
    Dim records As New Collection
    Dim record As Object
    Dim i As Long

    Set record = NewReusableRecord("RECIPE")
    record("RecipeName") = Trim$(mTxtReusableRecipeName.Text)
    record("Description") = Trim$(mTxtReusableRecipeDescription.Text)
    records.Add record
    For i = 0 To mLstRecipeNodes.ListCount - 1
        Set record = NewReusableRecord("PROCESS_NODE")
        record("ProcessNodeId") = NzStr(mLstRecipeNodes.List(i, 0))
        record("ProcessId") = NzStr(mLstRecipeNodes.List(i, 1))
        record("ProcessVersion") = NzStr(mLstRecipeNodes.List(i, 2))
        record("ExecutionOrdinal") = i + 1
        records.Add record
    Next i
    For i = 0 To mLstRecipeConnections.ListCount - 1
        Set record = NewReusableRecord("CONNECTION")
        record("FromProcessNodeId") = NzStr(mLstRecipeConnections.List(i, 0))
        record("FromOutputId") = NzStr(mLstRecipeConnections.List(i, 1))
        record("ToProcessNodeId") = NzStr(mLstRecipeConnections.List(i, 2))
        record("ToRequirementId") = NzStr(mLstRecipeConnections.List(i, 3))
        AddNumericReusableField record, "Qty", NzStr(mLstRecipeConnections.List(i, 4))
        AddNumericReusableField record, "Percent", NzStr(mLstRecipeConnections.List(i, 5))
        record("UOM") = NzStr(mLstRecipeConnections.List(i, 6))
        records.Add record
    Next i
    If Not mRecipeOutputRegulations Is Nothing Then
        For Each record In mRecipeOutputRegulations
            records.Add CloneReusableRecord(record)
        Next record
    End If
    BuildRecipePayload = modProductionJson.BuildJsonArray(records)
End Function

Private Function SubmitRecipeAction(ByVal eventType As String, _
                                    Optional ByVal payloadJson As String = "") As Boolean
    Dim report As String
    Dim quietStarted As Boolean

    On Error GoTo Failed
    If Not mOperatorWorkbook Is Nothing Then
        modOperationsPrimitiveBridge.BeginQuietUiForWorkbook mOperatorWorkbook.Name
        quietStarted = True
    End If
    ShowPersistencePending "Applying Recipe lifecycle action to warehouse storage..."
    SubmitRecipeAction = modProductionReusableDesigns.SubmitReusableDesignEvent( _
        eventType, Trim$(mTxtReusableRecipeId.Text), Trim$(mTxtReusableRecipeVersion.Text), _
        payloadJson, "Production Recipe Designer", report)
    If quietStarted Then modUiQuiet.EndQuietUi
    RefreshReusableDesignLists
    ShowStatus report
    Exit Function
Failed:
    If quietStarted Then modUiQuiet.EndQuietUi
    ShowStatus "Recipe action failed: " & Err.Description
End Function

Private Sub SelectReusableAssignmentProcess()
    Dim idx As Long
    Dim records As Collection
    Dim record As Object
    Dim jsonText As String
    Dim parseReport As String
    Dim rowIndex As Long

    idx = mLstAssignRecipes.ListIndex
    If idx < 0 Then
        ShowStatus "Select a Process version first."
        Exit Sub
    End If
    jsonText = modOperationsPrimitiveBridge.GetProcessVersion( _
        NzStr(mLstAssignRecipes.List(idx, 0)), NzStr(mLstAssignRecipes.List(idx, 1)))
    Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords(jsonText, parseReport)
    If records Is Nothing Then
        ShowStatus "Process requirements could not be loaded: " & parseReport
        Exit Sub
    End If
    mLstAssignIngredients.Clear
    Set mProcessAlternatives = New Collection
    For Each record In records
        Select Case UCase$(modProductionReusableDesigns.ReusableRecordText(record, "RecordType"))
            Case "REQUIREMENT"
                mLstAssignIngredients.AddItem modProductionReusableDesigns.ReusableRecordText(record, "RequirementId")
                rowIndex = mLstAssignIngredients.ListCount - 1
                mLstAssignIngredients.List(rowIndex, 1) = modProductionReusableDesigns.ReusableRecordText(record, "RequirementName")
                mLstAssignIngredients.List(rowIndex, 2) = modProductionReusableDesigns.ReusableRecordText(record, "UOM")
                mLstAssignIngredients.List(rowIndex, 3) = NzStr(mLstAssignRecipes.List(idx, 2))
                mLstAssignIngredients.List(rowIndex, 4) = "REQUIREMENT"
                mLstAssignIngredients.List(rowIndex, 5) = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Qty"))
                mLstAssignIngredients.List(rowIndex, 6) = NzStr(modProductionReusableDesigns.ReusableRecordValue(record, "Percent"))
            Case "ALTERNATIVE"
                mProcessAlternatives.Add CloneReusableRecord(record)
        End Select
    Next record
    RefreshReusableAllowedItems
    ShowStatus "Selected Process " & NzStr(mLstAssignRecipes.List(idx, 0)) & _
               " version " & NzStr(mLstAssignRecipes.List(idx, 1)) & "."
End Sub

Private Sub SelectReusableAssignmentRequirement()
    If mLstAssignIngredients.ListIndex < 0 Then
        ShowStatus "Select an ingredient requirement first."
    Else
        RefreshReusableAllowedItems
        ShowStatus "Selected requirement " & NzStr(mLstAssignIngredients.List(mLstAssignIngredients.ListIndex, 0)) & "."
    End If
End Sub

Private Sub RefreshReusableAllowedItems()
    Dim alternative As Variant
    Dim requirementId As String
    Dim itemCode As String
    Dim itemName As String
    Dim itemUom As String
    Dim rowIndex As Long

    mLstAssignAllowed.Clear
    If mLstAssignIngredients.ListIndex >= 0 Then _
        requirementId = NzStr(mLstAssignIngredients.List(mLstAssignIngredients.ListIndex, 0))
    If mProcessAlternatives Is Nothing Then Exit Sub
    For Each alternative In mProcessAlternatives
        If requirementId = "" Or StrComp(modProductionReusableDesigns.ReusableRecordText( _
                alternative, "RequirementId"), requirementId, vbTextCompare) = 0 Then
            mLstAssignAllowed.AddItem modProductionReusableDesigns.ReusableRecordText(alternative, "RequirementId")
            rowIndex = mLstAssignAllowed.ListCount - 1
            itemCode = modProductionReusableDesigns.ReusableRecordText(alternative, "ITEM_CODE")
            itemName = ManagedItemDisplayForAssignmentCode(itemCode, itemUom)
            If itemName = "" Then itemName = itemCode
            mLstAssignAllowed.List(rowIndex, 1) = itemName
            mLstAssignAllowed.List(rowIndex, 2) = itemUom
            mLstAssignAllowed.List(rowIndex, 3) = itemCode
            mLstAssignAllowed.List(rowIndex, 6) = itemCode
        End If
    Next alternative
End Sub

Private Function ManagedItemDisplayForAssignmentCode(ByVal itemCode As String, _
                                                      ByRef itemUom As String) As String
    Dim i As Long
    Dim r As Long

    itemCode = Trim$(itemCode)
    itemUom = ""
    If itemCode = "" Then Exit Function
    If Not mLstAssignInventory Is Nothing Then
        For i = 0 To mLstAssignInventory.ListCount - 1
            If StrComp(NzStr(mLstAssignInventory.List(i, 6)), itemCode, vbTextCompare) = 0 Then
                ManagedItemDisplayForAssignmentCode = NzStr(mLstAssignInventory.List(i, 1))
                itemUom = NzStr(mLstAssignInventory.List(i, 2))
                Exit Function
            End If
        Next i
    End If
    If IsArray(mInventoryRows) Then
        For r = LBound(mInventoryRows, 1) To UBound(mInventoryRows, 1)
            If StrComp(NzStr(mInventoryRows(r, 7)), itemCode, vbTextCompare) = 0 Then
                ManagedItemDisplayForAssignmentCode = NzStr(mInventoryRows(r, 2))
                itemUom = NzStr(mInventoryRows(r, 3))
                Exit Function
            End If
        Next r
    End If
End Function

Private Sub AddReusableInventoryAlternative()
    Dim inventoryIndex As Long
    Dim requirementId As String
    Dim itemCode As String
    Dim alternative As Object
    Dim existing As Variant

    If mLstAssignIngredients.ListIndex < 0 Then
        ShowStatus "Select an ingredient requirement first."
        Exit Sub
    End If
    inventoryIndex = mLstAssignInventory.ListIndex
    If inventoryIndex < 0 Then
        ShowStatus "Select a managed item first."
        Exit Sub
    End If
    requirementId = NzStr(mLstAssignIngredients.List(mLstAssignIngredients.ListIndex, 0))
    itemCode = NzStr(mLstAssignInventory.List(inventoryIndex, 6))
    If itemCode = "" Then
        ShowStatus "The selected inventory row has no managed item code."
        Exit Sub
    End If
    For Each existing In mProcessAlternatives
        If StrComp(modProductionReusableDesigns.ReusableRecordText(existing, "RequirementId"), _
                requirementId, vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(existing, "ITEM_CODE"), _
                itemCode, vbTextCompare) = 0 Then
            ShowStatus "That acceptable item is already assigned."
            Exit Sub
        End If
    Next existing
    Set alternative = NewReusableRecord("ALTERNATIVE")
    alternative("RequirementId") = requirementId
    alternative("ITEM_CODE") = itemCode
    mProcessAlternatives.Add alternative
    RefreshReusableAllowedItems
    ShowStatus "Added acceptable managed item " & itemCode & "."
End Sub

Private Sub RemoveReusableInventoryAlternative()
    Dim visibleIndex As Long
    Dim requirementId As String
    Dim itemCode As String
    Dim i As Long

    visibleIndex = mLstAssignAllowed.ListIndex
    If visibleIndex < 0 Then Exit Sub
    requirementId = NzStr(mLstAssignAllowed.List(visibleIndex, 0))
    itemCode = NzStr(mLstAssignAllowed.List(visibleIndex, 6))
    For i = mProcessAlternatives.Count To 1 Step -1
        If StrComp(modProductionReusableDesigns.ReusableRecordText(mProcessAlternatives(i), _
                "RequirementId"), requirementId, vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(mProcessAlternatives(i), _
                "ITEM_CODE"), itemCode, vbTextCompare) = 0 Then
            mProcessAlternatives.Remove i
            Exit For
        End If
    Next i
    RefreshReusableAllowedItems
End Sub

Private Function SaveReusableAssignments() As Boolean
    Dim processIndex As Long
    Dim sourceId As String
    Dim sourceVersion As String
    Dim preserved As New Collection
    Dim alternative As Variant
    Dim report As String

    processIndex = mLstAssignRecipes.ListIndex
    If processIndex < 0 Then
        ShowStatus "Select a Process version first."
        Exit Function
    End If
    For Each alternative In mProcessAlternatives
        preserved.Add CloneReusableRecord(alternative)
    Next alternative
    sourceId = NzStr(mLstAssignRecipes.List(processIndex, 0))
    sourceVersion = NzStr(mLstAssignRecipes.List(processIndex, 1))
    If Not LoadProcessDefinitionIntoDesigner(sourceId, sourceVersion, True) Then Exit Function
    Set mProcessAlternatives = preserved
    If Not ValidateProcessDraft(report) Then
        ShowStatus report
        Exit Function
    End If
    SaveReusableAssignments = SubmitProcessAction("PROCESS_SAVE", BuildProcessPayload())
    If SaveReusableAssignments Then
        ShowStatus "Acceptable alternatives saved as Process " & sourceId & _
                   " version " & mTxtProcessVersion.Text & "."
    End If
End Function

Public Function TestBatchScaleContract() As String
    Dim parsed As Double
    Dim report As String
    Dim minimumOk As Boolean
    Dim normalOk As Boolean
    Dim maximumOk As Boolean
    Dim lowRejected As Boolean
    Dim highRejected As Boolean

    minimumOk = TryParseBatchScalePercent("0.001", parsed, report) And _
                (Abs((10# * parsed / 100#) - 0.0001) < 0.000000001)
    normalOk = TryParseBatchScalePercent("100", parsed, report) And _
               (Abs((10# * parsed / 100#) - 10#) < 0.000000001)
    maximumOk = TryParseBatchScalePercent("1000", parsed, report) And _
                (Abs((10# * parsed / 100#) - 100#) < 0.000000001)
    lowRejected = Not TryParseBatchScalePercent("0.0009", parsed, report)
    highRejected = Not TryParseBatchScalePercent("1000.001", parsed, report)

    If minimumOk And normalOk And maximumOk And lowRejected And highRejected Then
        TestBatchScaleContract = _
            "OK|Min=.001%|Default=100%|Max=1000%|BoundsRejected=True|ListOnly=True"
    Else
        TestBatchScaleContract = "FAIL|Batch scale contract did not hold."
    End If
End Function

Public Function TestReusableProductionSurfaceContract() As String
    Dim pageCaptions As String
    Dim pageIndex As Long
    Dim hasProcessDesigner As Boolean
    Dim hasRecipeDesigner As Boolean
    Dim hasIngredientsAssignment As Boolean
    Dim hasRunList As Boolean
    Dim hasRunTree As Boolean
    Dim hasProductionSettings As Boolean
    Dim hasLegacyBuilder As Boolean

    If Not mBuilt Then BuildLayout
    For pageIndex = 0 To mPages.Pages.Count - 1
        If pageCaptions <> "" Then pageCaptions = pageCaptions & ","
        pageCaptions = pageCaptions & CStr(mPages.Pages(pageIndex).Caption)
        Select Case Trim$(CStr(mPages.Pages(pageIndex).Caption))
            Case "Process Designer": hasProcessDesigner = True
            Case "Recipe Designer": hasRecipeDesigner = True
            Case "Ingredients Assignment": hasIngredientsAssignment = True
            Case "Production Run - List": hasRunList = True
            Case "Production Run - Tree": hasRunTree = True
            Case "Production Settings": hasProductionSettings = True
            Case "Recipe Builder": hasLegacyBuilder = True
        End Select
    Next pageIndex

    If mPages.Pages.Count = 6 _
       And hasProcessDesigner _
       And hasRecipeDesigner _
       And hasIngredientsAssignment _
       And hasRunList _
       And hasRunTree _
       And hasProductionSettings _
       And Not hasLegacyBuilder Then
        TestReusableProductionSurfaceContract = _
            "OK|Pages=6|ProcessDesigner=True|RecipeDesigner=True|" & _
            "IngredientsAssignment=True|RunList=True|RunTreeExperimental=True|ProductionSettings=True|" & _
            "LegacyRecipeBuilder=False"
    Else
        TestReusableProductionSurfaceContract = _
            "FAIL|Pages=" & CStr(mPages.Pages.Count) & _
            "|Captions=" & pageCaptions & _
            "|ProcessDesigner=" & CStr(hasProcessDesigner) & _
            "|RecipeDesigner=" & CStr(hasRecipeDesigner) & _
            "|IngredientsAssignment=" & CStr(hasIngredientsAssignment) & _
            "|RunList=" & CStr(hasRunList) & _
            "|RunTreeExperimental=" & CStr(hasRunTree) & _
            "|ProductionSettings=" & CStr(hasProductionSettings) & _
            "|LegacyRecipeBuilder=" & CStr(hasLegacyBuilder)
    End If
End Function

Public Function TestProcessWorksheetRoundTripContract() As String
    Dim processIdGenerated As Boolean
    Dim recipeIdentityInitialized As Boolean
    Dim recipeIdGenerated As Boolean
    Dim recipeVersionGenerated As Boolean
    Dim recipeIdLocked As Boolean
    Dim recipeVersionEditable As Boolean
    Dim requirementIdGenerated As Boolean
    Dim outputIdGenerated As Boolean
    Dim identityControlsLocked As Boolean
    Dim worksheetControlVisible As Boolean
    Dim worksheetHandlerReached As Boolean
    Dim formulaEvidence As String
    Dim formulaCorrect As Boolean
    Dim mixedUomAccepted As Boolean
    Dim mixedUomRowsPreserved As Boolean
    Dim tableRemoved As Boolean
    Dim repeatRoundTrip As Boolean
    Dim actionReport As String
    Dim boundWorkbookName As String
    Dim tableName As String
    Dim tableCountBeforeMixed As Long
    Dim tableCountBeforeRepeat As Long

    If Not mBuilt Then BuildLayout
    If Not mOperatorWorkbook Is Nothing Then boundWorkbookName = mOperatorWorkbook.Name

    recipeIdentityInitialized = _
        mProduction.IsBase36Identifier(mTxtReusableRecipeId.Text) And _
        (Trim$(mTxtReusableRecipeVersion.Text) = "1")

    mBtnProcessNew_Click
    processIdGenerated = mProduction.IsBase36Identifier(mTxtProcessId.Text)
    identityControlsLocked = mTxtProcessId.Locked And mTxtProcessVersion.Locked

    mTxtRequirementName.Text = "Sugar"
    mTxtRequirementQty.Text = "100"
    mTxtRequirementUom.Text = "LB"
    mBtnProcessRequirementAdd_Click
    If mLstProcessRequirements.ListCount > 0 Then
        requirementIdGenerated = mProduction.IsBase36Identifier( _
            NzStr(mLstProcessRequirements.List(mLstProcessRequirements.ListCount - 1, 0)))
    End If

    mTxtProcessOutputName.Text = "Finished Product"
    mTxtProcessOutputQty.Text = "100"
    RefreshProcessOutputUomCatalog "LB"
    mBtnProcessOutputAdd_Click
    If mLstProcessOutputs.ListCount > 0 Then
        outputIdGenerated = mProduction.IsBase36Identifier( _
            NzStr(mLstProcessOutputs.List(mLstProcessOutputs.ListCount - 1, 0)))
    End If

    mBtnRecipeNew_Click
    recipeIdGenerated = mProduction.IsBase36Identifier(mTxtReusableRecipeId.Text)
    recipeVersionGenerated = (Trim$(mTxtReusableRecipeVersion.Text) = "1")
    recipeIdLocked = mTxtReusableRecipeId.Locked
    recipeVersionEditable = Not mTxtReusableRecipeVersion.Locked
    recipeIdentityInitialized = recipeIdGenerated And recipeVersionGenerated
    identityControlsLocked = identityControlsLocked And _
        recipeIdLocked And _
        mTxtProcessOutputDesignId.Locked And mTxtProcessOutputDesignVersion.Locked

    worksheetControlVisible = Not mBtnProcessWorksheetCreate Is Nothing And _
        Not mBtnProcessWorksheetRetrieve Is Nothing
    If worksheetControlVisible Then
        ClearProcessDraft True
        mTxtProcessName.Text = "Formula Worksheet Process"
        mBtnProcessWorksheetCreate_Click
        If modProductionProcessWorksheet.FindOutstandingProcessWorksheetTable( _
                mOperatorWorkbook, tableName, actionReport) Then
            Call modProductionProcessWorksheet.PopulateFormulationExampleForTest( _
                mOperatorWorkbook, tableName, False, actionReport)
            formulaEvidence = modProductionProcessWorksheet.ReadFormulaEvidenceForTest( _
                mOperatorWorkbook, tableName)
            formulaCorrect = formulaEvidence Like _
                "OK|Basis=611.2|Sugar=16.4|Flour=32.7|BakingPowder=1.8|Water=49.1|Total=100.0"
            Call modProductionProcessWorksheet.PopulateFormulationExampleForTest( _
                mOperatorWorkbook, tableName, True, actionReport)
            tableCountBeforeMixed = _
                modProductionProcessWorksheet.CountProcessWorksheetTables(mOperatorWorkbook)
            Call modProductionProcessWorksheet.SelectProcessWorksheetTableForTest( _
                mOperatorWorkbook, tableName)
            mBtnProcessWorksheetRetrieve_Click
            mixedUomAccepted = (modProductionProcessWorksheet.CountProcessWorksheetTables( _
                mOperatorWorkbook) = tableCountBeforeMixed - 1) And _
                Not ProcessWorksheetTableExistsForTest(tableName) And _
                (mLstProcessRequirements.ListCount = 3) And _
                (mLstProcessOutputs.ListCount = 1) And _
                (InStr(1, TestStatusText(), "Retrieved 1 selected Process table", _
                    vbTextCompare) > 0)
            If mixedUomAccepted Then
                mixedUomRowsPreserved = _
                    UCase$(NzStr(mLstProcessRequirements.List(0, 5))) = "LB" And _
                    UCase$(NzStr(mLstProcessRequirements.List(1, 5))) = "EA" And _
                    UCase$(NzStr(mLstProcessRequirements.List(2, 5))) = "EA" And _
                    Abs(CDbl(mLstProcessRequirements.List(0, 4)) - 4.5) < 0.05 And _
                    Abs(CDbl(mLstProcessRequirements.List(0, 3)) - 100#) < 0.05 And _
                    Abs(CDbl(mLstProcessRequirements.List(1, 4)) - 2#) < 0.05 And _
                    Abs(CDbl(mLstProcessRequirements.List(1, 3)) - 50#) < 0.05 And _
                    Abs(CDbl(mLstProcessRequirements.List(2, 3)) - 50#) < 0.05 And _
                    UCase$(NzStr(mLstProcessOutputs.List(0, 8))) = "EA"
            End If
            tableRemoved = mixedUomAccepted
            If tableRemoved Then
                tableCountBeforeRepeat = _
                    modProductionProcessWorksheet.CountProcessWorksheetTables(mOperatorWorkbook)
                mBtnProcessWorksheetCreate_Click
                If modProductionProcessWorksheet.FindSelectedProcessWorksheetTable( _
                        mOperatorWorkbook, tableName, actionReport) Then
                    Call modProductionProcessWorksheet.SelectProcessWorksheetTableForTest( _
                        mOperatorWorkbook, tableName)
                    mBtnProcessWorksheetRetrieve_Click
                    repeatRoundTrip = (modProductionProcessWorksheet.CountProcessWorksheetTables( _
                        mOperatorWorkbook) = tableCountBeforeRepeat) And _
                        Not ProcessWorksheetTableExistsForTest(tableName)
                End If
            End If
        End If
        worksheetHandlerReached = mixedUomAccepted And mixedUomRowsPreserved And formulaCorrect And _
            tableRemoved And repeatRoundTrip
    End If

    If boundWorkbookName <> "" And recipeIdentityInitialized _
       And processIdGenerated And recipeIdGenerated _
       And recipeVersionGenerated And recipeIdLocked And recipeVersionEditable _
       And requirementIdGenerated And outputIdGenerated And identityControlsLocked _
       And worksheetControlVisible And worksheetHandlerReached Then
        TestProcessWorksheetRoundTripContract = _
            "OK|BoundWorkbook=" & boundWorkbookName & _
            "|RecipeIdentityInitialized=True" & _
            "|ProcessIdGenerated=True|RecipeIdGenerated=True" & _
            "|RecipeVersionGenerated=True|RecipeIdLocked=True" & _
            "|RecipeVersionEditable=True" & _
            "|RequirementIdGenerated=True|OutputIdGenerated=True" & _
            "|IdentityControlsLocked=True|WorksheetHandler=True" & _
            "|MixedUomAccepted=True|MixedUomRowsPreserved=True|Formula=" & formulaEvidence & _
            "|TableRemoved=True|RepeatRoundTrip=True"
    Else
        TestProcessWorksheetRoundTripContract = _
            "FAIL|BoundWorkbook=" & boundWorkbookName & _
            "|RecipeIdentityInitialized=" & CStr(recipeIdentityInitialized) & _
            "|ProcessIdGenerated=" & CStr(processIdGenerated) & _
            "|RecipeIdGenerated=" & CStr(recipeIdGenerated) & _
            "|RecipeVersionGenerated=" & CStr(recipeVersionGenerated) & _
            "|RecipeIdLocked=" & CStr(recipeIdLocked) & _
            "|RecipeVersionEditable=" & CStr(recipeVersionEditable) & _
            "|RequirementIdGenerated=" & CStr(requirementIdGenerated) & _
            "|OutputIdGenerated=" & CStr(outputIdGenerated) & _
            "|IdentityControlsLocked=" & CStr(identityControlsLocked) & _
            "|WorksheetControl=" & CStr(worksheetControlVisible) & _
            "|WorksheetHandler=" & CStr(worksheetHandlerReached) & _
            "|MixedUomAccepted=" & CStr(mixedUomAccepted) & _
            "|MixedUomRowsPreserved=" & CStr(mixedUomRowsPreserved) & _
            "|Formula=" & formulaEvidence & _
            "|TableRemoved=" & CStr(tableRemoved) & _
            "|RepeatRoundTrip=" & CStr(repeatRoundTrip) & _
            "|Status=" & Replace$(Replace$(TestStatusText(), vbCr, " "), vbLf, " ")
    End If
End Function

Public Function TestProcessWorksheetWorkbenchContract() As String
    Dim lo As ListObject
    Dim itemSearchCell As Range
    Dim firstCreated As Boolean
    Dim separateActions As Boolean
    Dim multipleTables As Boolean
    Dim selectedOnly As Boolean
    Dim recordTypeDropdown As Boolean
    Dim calculatedPercent As Boolean
    Dim generatedDesign As Boolean
    Dim itemCodeRemoved As Boolean
    Dim assignments As Boolean
    Dim itemSearch As Boolean
    Dim recordTypeColumn As Long
    Dim percentColumn As Long
    Dim basisColumn As Long
    Dim designColumn As Long
    Dim outputRow As Long
    Dim tableCountAfterFirst As Long
    Dim tableCountAfterSecond As Long
    Dim tableCountAfterThird As Long
    Dim tableCountAfterRetrieve As Long
    Dim firstTable As String
    Dim secondTable As String
    Dim thirdTable As String
    Dim actionReport As String

    If Not mBuilt Then BuildLayout
    ClearProcessDraft True
    mTxtProcessName.Text = "Workbench Process One"
    mBtnProcessWorksheetCreate_Click
    Call modProductionProcessWorksheet.FindSelectedProcessWorksheetTable( _
        mOperatorWorkbook, firstTable, actionReport)
    tableCountAfterFirst = modProductionProcessWorksheet.CountProcessWorksheetTables(mOperatorWorkbook)
    firstCreated = (tableCountAfterFirst = 1)
    Call modProductionProcessWorksheet.PopulateFormulationExampleForTest( _
        mOperatorWorkbook, firstTable, False, actionReport)
    Set lo = ProcessWorksheetTableForTest(firstTable)
    If Not lo Is Nothing Then
        recordTypeColumn = ProcessWorksheetColumnForTest(lo, "Record Type")
        percentColumn = ProcessWorksheetColumnForTest(lo, "Percent")
        basisColumn = ProcessWorksheetColumnForTest(lo, "Basis Qty")
        designColumn = ProcessWorksheetColumnForTest(lo, "Design ID")
        If recordTypeColumn > 0 Then
            On Error Resume Next
            recordTypeDropdown = (lo.DataBodyRange.Cells(1, recordTypeColumn).Validation.Type = xlValidateList)
            On Error GoTo 0
        End If
        If percentColumn > 0 And basisColumn > 0 Then
            lo.DataBodyRange.Cells(1, recordTypeColumn).Value2 = "INPUT"
            lo.DataBodyRange.Cells(1, ProcessWorksheetColumnForTest(lo, "Name")).Value2 = "Pasted Ingredient"
            lo.DataBodyRange.Cells(1, ProcessWorksheetColumnForTest(lo, "Qty")).Value2 = 50
            lo.DataBodyRange.Cells(1, ProcessWorksheetColumnForTest(lo, "UOM")).Value2 = "LB"
            Application.Calculate
            calculatedPercent = lo.DataBodyRange.Cells(1, percentColumn).HasFormula And _
                lo.DataBodyRange.Cells(1, basisColumn).HasFormula
        End If
        itemCodeRemoved = (ProcessWorksheetColumnForTest(lo, "Item Code") = 0)
        assignments = ProcessWorksheetColumnForTest(lo, "Requirement ID") > 0 And _
            ProcessWorksheetColumnForTest(lo, "Acceptable Managed Item 1") > 0 And _
            ProcessWorksheetColumnForTest(lo, "Accepted SKU 1") > 0
        If assignments Then
            Set itemSearchCell = lo.DataBodyRange.Cells(1, _
                ProcessWorksheetColumnForTest(lo, "Acceptable Managed Item 1"))
            itemSearch = modProductionProcessWorksheet.IsProcessWorksheetItemSearchTarget( _
                itemSearchCell)
        End If
        If recordTypeColumn > 0 Then
            For outputRow = 1 To lo.ListRows.Count
                If UCase$(Trim$(CStr(lo.DataBodyRange.Cells(outputRow, recordTypeColumn).Value2))) = "OUTPUT" Then Exit For
            Next outputRow
            If outputRow <= lo.ListRows.Count And designColumn > 0 Then
                generatedDesign = Trim$(CStr(lo.DataBodyRange.Cells(outputRow, designColumn).Value2)) <> ""
            End If
        End If
    End If

    mTxtProcessName.Text = "Workbench Process Two"
    mBtnProcessWorksheetCreate_Click
    Call modProductionProcessWorksheet.FindSelectedProcessWorksheetTable( _
        mOperatorWorkbook, secondTable, actionReport)
    Call modProductionProcessWorksheet.PopulateFormulationExampleForTest( _
        mOperatorWorkbook, secondTable, False, actionReport)
    tableCountAfterSecond = modProductionProcessWorksheet.CountProcessWorksheetTables(mOperatorWorkbook)

    mTxtProcessName.Text = "Workbench Process Three"
    mBtnProcessWorksheetCreate_Click
    Call modProductionProcessWorksheet.FindSelectedProcessWorksheetTable( _
        mOperatorWorkbook, thirdTable, actionReport)
    Call modProductionProcessWorksheet.PopulateFormulationExampleForTest( _
        mOperatorWorkbook, thirdTable, False, actionReport)
    tableCountAfterThird = modProductionProcessWorksheet.CountProcessWorksheetTables(mOperatorWorkbook)

    separateActions = ProductionPageControlExists("btnProcessWorksheetCreate") And _
        ProductionPageControlExists("btnProcessWorksheetRetrieve")
    multipleTables = (tableCountAfterSecond = 2 And tableCountAfterThird = 3)
    If modProductionProcessWorksheet.SelectProcessWorksheetTableForTest( _
            mOperatorWorkbook, firstTable) Then
        mBtnProcessWorksheetRetrieve_Click
        tableCountAfterRetrieve = _
            modProductionProcessWorksheet.CountProcessWorksheetTables(mOperatorWorkbook)
        selectedOnly = (tableCountAfterRetrieve = 2) And _
            Not ProcessWorksheetTableExistsForTest(firstTable) And _
            ProcessWorksheetTableExistsForTest(secondTable) And _
            ProcessWorksheetTableExistsForTest(thirdTable)
    End If

    If firstCreated And separateActions And multipleTables And selectedOnly And _
       recordTypeDropdown And calculatedPercent And generatedDesign And _
       itemCodeRemoved And assignments And itemSearch Then
        TestProcessWorksheetWorkbenchContract = _
            "OK|SeparateActions=True|MultipleTables=True|SelectedOnly=True" & _
            "|RecordTypeDropdown=True|CalculatedPercent=True|GeneratedDesign=True" & _
            "|ItemCodeRemoved=True|Assignments=True|ItemSearch=True"
    Else
        TestProcessWorksheetWorkbenchContract = _
            "FAIL|SeparateActions=" & CStr(separateActions) & _
            "|MultipleTables=" & CStr(multipleTables) & _
            "|SelectedOnly=" & CStr(selectedOnly) & _
            "|RecordTypeDropdown=" & CStr(recordTypeDropdown) & _
            "|CalculatedPercent=" & CStr(calculatedPercent) & _
            "|GeneratedDesign=" & CStr(generatedDesign) & _
            "|ItemCodeRemoved=" & CStr(itemCodeRemoved) & _
            "|Assignments=" & CStr(assignments) & _
            "|ItemSearch=" & CStr(itemSearch) & _
            "|Tables=" & CStr(tableCountAfterRetrieve)
    End If
End Function

Public Function TestProcessEditExportContract() As String
    Dim token As String
    Dim processId As String
    Dim rowIndex As Long
    Dim outputRow As Long
    Dim recordTypeColumn As Long
    Dim outputNameColumn As Long
    Dim outputQtyColumn As Long
    Dim tableName As String
    Dim exportedProcessId As String
    Dim exportedProcessVersion As String
    Dim exportedProcessName As String
    Dim exportedDescription As String
    Dim exportedPayload As String
    Dim actionReport As String
    Dim lo As ListObject
    Dim viewedImmutable As Boolean
    Dim releasedProcessEditable As Boolean
    Dim outputDesignVersionRebased As Boolean
    Dim outputYieldRebased As Boolean
    Dim existingProcessExported As Boolean
    Dim exportRoundTrip As Boolean
    Dim testStage As String
    Dim stagedRequirement0 As String
    Dim stagedRequirement1 As String
    Dim stagedRequirement2 As String
    Dim stagedOutputQty As String
    Dim stagedOutputYield As String
    Dim stagedStatus As String
    Dim editorValues As String

    On Error GoTo Failed
    testStage = "BuildLayout"
    If Not mBuilt Then BuildLayout
    mPages.Value = 0
    mReusableActionTestInProgress = True
    token = UCase$(Right$(CleanControlName(BuildFormGuid()), 10))

    testStage = "CreateSourceDraft"
    ClearProcessDraft True
    processId = Trim$(mTxtProcessId.Text)
    mTxtProcessVersion.Text = "1"
    mTxtProcessName.Text = "Bottled Product Edit " & token

    mTxtRequirementId.Text = "001"
    mTxtRequirementName.Text = "Concentrate"
    mTxtRequirementQty.Text = "4.5"
    mTxtRequirementUom.Text = "LB"
    mBtnProcessRequirementAdd_Click
    mTxtRequirementId.Text = "002"
    mTxtRequirementName.Text = "Bottle"
    mTxtRequirementQty.Text = "1"
    mTxtRequirementUom.Text = "EA"
    mBtnProcessRequirementAdd_Click
    mTxtRequirementId.Text = "003"
    mTxtRequirementName.Text = "Cap"
    mTxtRequirementQty.Text = "1"
    mTxtRequirementUom.Text = "EA"
    mBtnProcessRequirementAdd_Click

    mTxtProcessOutputId.Text = "004"
    mTxtProcessOutputName.Text = "64oz Bottled Product"
    mTxtProcessOutputItemCode.Text = "SKU-BOTTLED-" & token
    mTxtProcessOutputQty.Text = "1"
    RefreshProcessOutputUomCatalog "EA"
    mBtnProcessOutputAdd_Click
    testStage = "SaveSourceDraft"
    mBtnProcessSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then GoTo ContractFailed
    testStage = "ReleaseSource"
    mBtnProcessRelease_Click
    If InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) = 0 Then GoTo ContractFailed

    rowIndex = FindIdentityListRow(mLstProcesses, processId, "1")
    If rowIndex < 0 Then GoTo ContractFailed
    mLstProcesses.ListIndex = rowIndex
    testStage = "ViewSource"
    mBtnProcessLoad_Click
    viewedImmutable = (Trim$(mTxtProcessVersion.Text) = "1") And _
        (InStr(1, TestStatusText(), "Viewed immutable Process", vbTextCompare) > 0)

    testStage = "EditSuccessor"
    mBtnProcessReuse_Click
    outputDesignVersionRebased = _
        (mLstProcessOutputs.ListCount = 1) And _
        (StrComp(NzStr(mLstProcessOutputs.List(0, 4)), "2", vbTextCompare) = 0)

    testStage = "UpdateSuccessor"
    mLstProcessRequirements.ListIndex = 0
    mLstProcessRequirements_Click
    mTxtRequirementQty.SetFocus
    mTxtRequirementQty.Text = "600"
    editorValues = "R0=" & mTxtRequirementQty.Text
    mBtnProcessRequirementUpdate_Click
    mLstProcessRequirements.ListIndex = 1
    mLstProcessRequirements_Click
    mTxtRequirementQty.SetFocus
    mTxtRequirementQty.Text = "133"
    editorValues = editorValues & ",R1=" & mTxtRequirementQty.Text
    mBtnProcessRequirementUpdate_Click
    mLstProcessRequirements.ListIndex = 2
    mLstProcessRequirements_Click
    mTxtRequirementQty.SetFocus
    mTxtRequirementQty.Text = "133"
    editorValues = editorValues & ",R2=" & mTxtRequirementQty.Text
    mBtnProcessRequirementUpdate_Click
    mLstProcessOutputs.ListIndex = 0
    mLstProcessOutputs_Click
    mTxtProcessOutputQty.SetFocus
    mTxtProcessOutputQty.Text = "133"
    editorValues = editorValues & ",O=" & mTxtProcessOutputQty.Text
    mBtnProcessOutputUpdate_Click
    stagedRequirement0 = NzStr(mLstProcessRequirements.List(0, 2))
    stagedRequirement1 = NzStr(mLstProcessRequirements.List(1, 2))
    stagedRequirement2 = NzStr(mLstProcessRequirements.List(2, 2))
    stagedOutputQty = NzStr(mLstProcessOutputs.List(0, 5))
    stagedOutputYield = NzStr(mLstProcessOutputs.List(0, 7))
    stagedStatus = TestStatusText()
    outputYieldRebased = _
        (Abs(Val(stagedOutputYield) - 133#) < 0.00000001)
    releasedProcessEditable = viewedImmutable And _
        (Trim$(mTxtProcessVersion.Text) = "2") And _
        (Abs(Val(NzStr(mLstProcessRequirements.List(0, 2))) - 600#) < 0.00000001) And _
        (Abs(Val(NzStr(mLstProcessRequirements.List(1, 2))) - 133#) < 0.00000001) And _
        (Abs(Val(NzStr(mLstProcessRequirements.List(2, 2))) - 133#) < 0.00000001) And _
        (Abs(Val(NzStr(mLstProcessOutputs.List(0, 5))) - 133#) < 0.00000001)

    testStage = "SendToSheet"
    mBtnProcessWorksheetCreate_Click
    testStage = "FindExportedTable"
    If Not modProductionProcessWorksheet.FindSelectedProcessWorksheetTable( _
            mOperatorWorkbook, tableName, actionReport) Then GoTo ContractFailed
    Set lo = ProcessWorksheetTableForTest(tableName)
    If lo Is Nothing Then GoTo ContractFailed
    recordTypeColumn = ProcessWorksheetColumnForTest(lo, "Record Type")
    outputNameColumn = ProcessWorksheetColumnForTest(lo, "Name")
    outputQtyColumn = ProcessWorksheetColumnForTest(lo, "Qty")
    For outputRow = 1 To lo.ListRows.Count
        If UCase$(Trim$(CStr(lo.DataBodyRange.Cells(outputRow, recordTypeColumn).Value2))) = _
                "OUTPUT" Then Exit For
    Next outputRow
    If outputRow > lo.ListRows.Count Then GoTo ContractFailed
    testStage = "ReadExportedTable"
    existingProcessExported = modProductionProcessWorksheet.ReadProcessDraftFromWorksheet( _
        mOperatorWorkbook, tableName, exportedProcessId, exportedProcessVersion, _
        exportedProcessName, exportedDescription, exportedPayload, actionReport) And _
        (StrComp(exportedProcessId, processId, vbTextCompare) = 0) And _
        (exportedProcessVersion = "2") And _
        (Abs(Val(CStr(lo.DataBodyRange.Cells(outputRow, outputQtyColumn).Value2)) - 133#) < 0.00000001)

    lo.DataBodyRange.Cells(outputRow, outputNameColumn).Value2 = "64oz Bottled Product Edited"
    lo.DataBodyRange.Cells(outputRow, outputNameColumn).Select
    testStage = "RetrieveEditedTable"
    mBtnProcessWorksheetRetrieve_Click
    exportRoundTrip = _
        Not ProcessWorksheetTableExistsForTest(tableName) And _
        (Trim$(mTxtProcessVersion.Text) = "2") And _
        (mLstProcessOutputs.ListCount = 1) And _
        (StrComp(NzStr(mLstProcessOutputs.List(0, 1)), _
            "64oz Bottled Product Edited", vbTextCompare) = 0) And _
        (InStr(1, TestStatusText(), "Retrieved 1 selected Process table", vbTextCompare) > 0)

    If releasedProcessEditable And existingProcessExported And exportRoundTrip And _
       outputDesignVersionRebased And outputYieldRebased Then
        TestProcessEditExportContract = _
            "OK|ReleasedProcessEditable=True|ExistingProcessExported=True" & _
            "|ExportRoundTrip=True|OutputDesignVersionRebased=True" & _
            "|OutputYieldRebased=True"
    Else
ContractFailed:
        TestProcessEditExportContract = _
            "FAIL|ReleasedProcessEditable=" & CStr(releasedProcessEditable) & _
            "|ExistingProcessExported=" & CStr(existingProcessExported) & _
            "|ExportRoundTrip=" & CStr(exportRoundTrip) & _
            "|OutputDesignVersionRebased=" & CStr(outputDesignVersionRebased) & _
            "|OutputYieldRebased=" & CStr(outputYieldRebased) & _
            "|Version=" & Trim$(mTxtProcessVersion.Text) & _
            "|Req0=" & NzStr(mLstProcessRequirements.List(0, 2)) & _
            "|Req1=" & NzStr(mLstProcessRequirements.List(1, 2)) & _
            "|Req2=" & NzStr(mLstProcessRequirements.List(2, 2)) & _
            "|OutputQty=" & NzStr(mLstProcessOutputs.List(0, 5)) & _
            "|OutputYield=" & NzStr(mLstProcessOutputs.List(0, 7)) & _
            "|StagedReqs=" & stagedRequirement0 & "," & stagedRequirement1 & "," & _
                stagedRequirement2 & _
            "|StagedOutput=" & stagedOutputQty & "," & stagedOutputYield & _
            "|Editors=" & editorValues & _
            "|StagedStatus=" & stagedStatus & _
            "|ExportVersion=" & exportedProcessVersion & _
            "|Status=" & TestStatusText()
    End If
    mReusableActionTestInProgress = False
    Exit Function

Failed:
    mReusableActionTestInProgress = False
    TestProcessEditExportContract = _
        "FAIL|Stage=" & testStage & "|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function TestProcessWorksheetBulkImportContract() As String
    Dim lo As ListObject
    Dim firstTable As String
    Dim secondTable As String
    Dim firstProcessId As String
    Dim secondProcessId As String
    Dim observedInputId As String
    Dim observedOutputId As String
    Dim observedRequirementId As String
    Dim observedDesignId As String
    Dim observedIdFormat As String
    Dim actionReport As String
    Dim outputRow As Long
    Dim recordTypeColumn As Long
    Dim itemSearchCell As Range
    Dim textSafeIds As Boolean
    Dim requirementIds As Boolean
    Dim uomCatalog As Boolean
    Dim numberedAlternatives As Boolean
    Dim addedAlternative As Boolean
    Dim pickerOpened As Boolean
    Dim pickerInventoryRows As Boolean
    Dim multiAreaSelection As Boolean
    Dim multiTableDrafts As Boolean

    If Not mBuilt Then BuildLayout
    ClearProcessDraft True
    mTxtProcessName.Text = "Slice 4aa Bulk Process One"
    mBtnProcessWorksheetCreate_Click
    Call modProductionProcessWorksheet.FindSelectedProcessWorksheetTable( _
        mOperatorWorkbook, firstTable, actionReport)
    Set lo = ProcessWorksheetTableForTest(firstTable)
    If lo Is Nothing Then GoTo ReportResult
    firstProcessId = CStr(lo.Parent.Cells(lo.HeaderRowRange.Row - 4, 5).Value2)
    Call modProductionProcessWorksheet.PopulateFormulationExampleForTest( _
        mOperatorWorkbook, firstTable, False, actionReport)
    Application.Calculate
    recordTypeColumn = ProcessWorksheetColumnForTest(lo, "Record Type")
    For outputRow = 1 To lo.ListRows.Count
        If UCase$(Trim$(CStr(lo.DataBodyRange.Cells(outputRow, recordTypeColumn).Value2))) = "OUTPUT" Then Exit For
    Next outputRow
    observedInputId = CStr(lo.DataBodyRange.Cells(1, _
        ProcessWorksheetColumnForTest(lo, "ID")).Value2)
    observedIdFormat = CStr(lo.DataBodyRange.Cells(1, _
        ProcessWorksheetColumnForTest(lo, "ID")).NumberFormat)
    observedRequirementId = CStr(lo.DataBodyRange.Cells(1, _
        ProcessWorksheetColumnForTest(lo, "Requirement ID")).Value2)
    observedDesignId = CStr(lo.DataBodyRange.Cells(outputRow, _
        ProcessWorksheetColumnForTest(lo, "Design ID")).Value2)
    observedOutputId = CStr(lo.DataBodyRange.Cells(outputRow, _
        ProcessWorksheetColumnForTest(lo, "ID")).Value2)
    textSafeIds = (firstProcessId Like "[0-9A-Z][0-9A-Z][0-9A-Z]") And _
        (observedInputId = "001") And (observedIdFormat = "@") And _
        mProduction.IsBase36Identifier(observedOutputId) And _
        (observedOutputId <> observedInputId) And _
        (observedDesignId = "D-" & firstProcessId & "-" & observedOutputId)
    requirementIds = lo.DataBodyRange.Cells(1, _
        ProcessWorksheetColumnForTest(lo, "Requirement ID")).HasFormula And _
        (CStr(lo.DataBodyRange.Cells(1, _
            ProcessWorksheetColumnForTest(lo, "Requirement ID")).Value2) = "001")
    On Error Resume Next
    uomCatalog = (lo.DataBodyRange.Cells(1, _
        ProcessWorksheetColumnForTest(lo, "UOM")).Validation.Type = xlValidateList) And _
        (InStr(1, lo.DataBodyRange.Cells(1, _
            ProcessWorksheetColumnForTest(lo, "UOM")).Validation.Formula1, "LB", vbTextCompare) > 0)
    On Error GoTo 0
    numberedAlternatives = ProcessWorksheetColumnForTest(lo, "Acceptable Managed Item 1") > 0 And _
        ProcessWorksheetColumnForTest(lo, "Acceptable Managed Item 4") > 0 And _
        ProcessWorksheetColumnForTest(lo, "Accepted SKU 4") > 0
    lo.DataBodyRange.Cells(1, recordTypeColumn).Select
    mBtnProcessWorksheetAddAlternative_Click
    addedAlternative = (ProcessWorksheetColumnForTest(lo, "Acceptable Managed Item 5") > 0)

    Set itemSearchCell = lo.DataBodyRange.Cells(1, _
        ProcessWorksheetColumnForTest(lo, "Acceptable Managed Item 1"))
    lo.DataBodyRange.Cells(1, ProcessWorksheetColumnForTest(lo, "Name")).Select
    itemSearchCell.Select
    DoEvents
    pickerOpened = mProduction.ProductionProcessItemSearchVisibleForTest()
    pickerInventoryRows = _
        (mProduction.ProductionProcessItemSearchResultCountForTest() > 0)
    mProduction.CloseProductionProcessItemSearchForTest

    ClearProcessDraft True
    mTxtProcessName.Text = "Slice 4aa Bulk Process Two"
    mBtnProcessWorksheetCreate_Click
    Call modProductionProcessWorksheet.FindSelectedProcessWorksheetTable( _
        mOperatorWorkbook, secondTable, actionReport)
    Set lo = ProcessWorksheetTableForTest(secondTable)
    If Not lo Is Nothing Then
        secondProcessId = CStr(lo.Parent.Cells(lo.HeaderRowRange.Row - 4, 5).Value2)
        Call modProductionProcessWorksheet.PopulateFormulationExampleForTest( _
            mOperatorWorkbook, secondTable, False, actionReport)
    End If
    multiAreaSelection = modProductionProcessWorksheet.SelectProcessWorksheetTablesForTest( _
        mOperatorWorkbook, firstTable, secondTable)
    If multiAreaSelection Then mBtnProcessWorksheetRetrieve_Click
    RefreshReusableDesignLists
    multiTableDrafts = (modProductionProcessWorksheet.CountProcessWorksheetTables( _
        mOperatorWorkbook) = 2) And _
        Not ProcessWorksheetTableExistsForTest(firstTable) And _
        Not ProcessWorksheetTableExistsForTest(secondTable) And _
        (FindIdentityListRow(mLstProcesses, firstProcessId, "1") >= 0) And _
        (FindIdentityListRow(mLstProcesses, secondProcessId, "1") >= 0)

ReportResult:
    If textSafeIds And requirementIds And uomCatalog And numberedAlternatives _
       And addedAlternative And pickerOpened And pickerInventoryRows And _
       multiAreaSelection And multiTableDrafts Then
        TestProcessWorksheetBulkImportContract = _
            "OK|TextSafeIds=True|RequirementIds=True|UomCatalog=True" & _
            "|NumberedAlternatives=True|AddedAlternative=True|PickerOpened=True" & _
            "|PickerInventoryRows=True" & _
            "|MultiAreaSelection=True|MultiTableDrafts=True"
    Else
        TestProcessWorksheetBulkImportContract = _
            "FAIL|TextSafeIds=" & CStr(textSafeIds) & _
            "|RequirementIds=" & CStr(requirementIds) & _
            "|UomCatalog=" & CStr(uomCatalog) & _
            "|NumberedAlternatives=" & CStr(numberedAlternatives) & _
            "|AddedAlternative=" & CStr(addedAlternative) & _
            "|PickerOpened=" & CStr(pickerOpened) & _
            "|PickerInventoryRows=" & CStr(pickerInventoryRows) & _
            "|MultiAreaSelection=" & CStr(multiAreaSelection) & _
            "|MultiTableDrafts=" & CStr(multiTableDrafts) & _
            "|ProcessId=" & firstProcessId & _
            "|InputId=" & observedInputId & _
            "|InputIdFormat=" & observedIdFormat & _
            "|RequirementIdValue=" & observedRequirementId & _
            "|DesignIdValue=" & observedDesignId & _
            "|Status=" & Replace$(Replace$(TestStatusText(), vbCr, " "), vbLf, " ")
    End If
End Function

Public Function TestProcessWorksheetOutputPickerContract() As String
    Dim lo As ListObject
    Dim tableName As String
    Dim actionReport As String
    Dim outputRow As Long
    Dim recordTypeColumn As Long
    Dim idColumn As Long
    Dim nameColumn As Long
    Dim outputItemColumn As Long
    Dim outputSkuColumn As Long
    Dim outputItemCell As Range
    Dim outputNameCell As Range
    Dim selectedOutputName As String
    Dim selectedOutputItem As String
    Dim selectedOutputSku As String
    Dim outputPickerOpened As Boolean
    Dim outputPickerCommitted As Boolean
    Dim outputSkuHidden As Boolean
    Dim outputSkuRoundTrip As Boolean
    Dim outputNameRetained As Boolean
    Dim outputNamePickerSuppressed As Boolean
    Dim uniqueRowIds As Boolean
    Dim firstAssignedIdRetained As Boolean
    Dim noPhysicalKey As Boolean
    Dim firstAssignedId As String
    Dim priorEvents As Boolean
    Dim rowIndex As Long

    If Not mBuilt Then BuildLayout
    ClearProcessDraft True
    mTxtProcessName.Text = "Slice 4ac Output Picker Process"
    mBtnProcessWorksheetCreate_Click
    Call modProductionProcessWorksheet.FindSelectedProcessWorksheetTable( _
        mOperatorWorkbook, tableName, actionReport)
    Set lo = ProcessWorksheetTableForTest(tableName)
    If lo Is Nothing Then GoTo ReportResult
    recordTypeColumn = ProcessWorksheetColumnForTest(lo, "Record Type")
    idColumn = ProcessWorksheetColumnForTest(lo, "ID")
    nameColumn = ProcessWorksheetColumnForTest(lo, "Name")
    outputItemColumn = ProcessWorksheetColumnForTest(lo, "Acceptable Managed Item 1")
    outputSkuColumn = ProcessWorksheetColumnForTest(lo, "Output SKU")
    If recordTypeColumn = 0 Or idColumn = 0 Or nameColumn = 0 Or outputItemColumn = 0 _
       Or outputSkuColumn = 0 Then GoTo ReportResult

    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    For rowIndex = 1 To lo.ListRows.Count
        lo.DataBodyRange.Cells(rowIndex, recordTypeColumn).ClearContents
        lo.DataBodyRange.Cells(rowIndex, idColumn).ClearContents
    Next rowIndex
    Application.EnableEvents = priorEvents
    lo.DataBodyRange.Cells(7, recordTypeColumn).Value2 = "OUTPUT"
    mProduction.HandleProductionChange lo.DataBodyRange.Cells(7, recordTypeColumn)
    firstAssignedId = Trim$(CStr(lo.DataBodyRange.Cells(7, idColumn).Value2))
    uniqueRowIds = ProcessWorksheetRowIdsUniqueForTest(lo, recordTypeColumn, idColumn)
    lo.DataBodyRange.Cells(1, recordTypeColumn).Value2 = "INPUT"
    mProduction.HandleProductionChange lo.DataBodyRange.Cells(1, recordTypeColumn)
    uniqueRowIds = uniqueRowIds And _
        ProcessWorksheetRowIdsUniqueForTest(lo, recordTypeColumn, idColumn)
    lo.DataBodyRange.Cells(3, recordTypeColumn).Value2 = "INSTRUCTION"
    mProduction.HandleProductionChange lo.DataBodyRange.Cells(3, recordTypeColumn)
    uniqueRowIds = uniqueRowIds And _
        ProcessWorksheetRowIdsUniqueForTest(lo, recordTypeColumn, idColumn)
    firstAssignedIdRetained = firstAssignedId <> "" And _
        (Trim$(CStr(lo.DataBodyRange.Cells(7, idColumn).Value2)) = firstAssignedId)

    Call modProductionProcessWorksheet.PopulateFormulationExampleForTest( _
        mOperatorWorkbook, tableName, False, actionReport)
    For outputRow = 1 To lo.ListRows.Count
        If UCase$(Trim$(CStr(lo.DataBodyRange.Cells(outputRow, _
                recordTypeColumn).Value2))) = "OUTPUT" Then Exit For
    Next outputRow
    If outputRow > lo.ListRows.Count Then GoTo ReportResult

    Set outputNameCell = lo.DataBodyRange.Cells(outputRow, nameColumn)
    mProduction.CloseProductionProcessItemSearchForTest
    mProduction.ShowProductionProcessItemSearch outputNameCell
    DoEvents
    outputNamePickerSuppressed = _
        Not mProduction.ProductionProcessItemSearchVisibleForTest()
    lo.DataBodyRange.Cells(outputRow, outputItemColumn).ClearContents
    lo.DataBodyRange.Cells(outputRow, _
        ProcessWorksheetColumnForTest(lo, "Accepted SKU 1")).ClearContents
    lo.DataBodyRange.Cells(outputRow, outputSkuColumn).ClearContents
    selectedOutputName = Trim$(CStr(lo.DataBodyRange.Cells(outputRow, nameColumn).Value2))
    Set outputItemCell = lo.DataBodyRange.Cells(outputRow, outputItemColumn)
    outputItemCell.Select
    DoEvents
    outputPickerOpened = mProduction.ProductionProcessItemSearchVisibleForTest() And _
        (mProduction.ProductionProcessItemSearchResultCountForTest() > 0)
    If outputPickerOpened Then outputPickerCommitted = _
        mProduction.CommitFirstProductionProcessItemSearchResultForTest()
    selectedOutputItem = Trim$(CStr(lo.DataBodyRange.Cells(outputRow, outputItemColumn).Value2))
    selectedOutputSku = Trim$(CStr(lo.DataBodyRange.Cells(outputRow, outputSkuColumn).Value2))
    outputPickerCommitted = outputPickerCommitted And _
        selectedOutputItem <> "" And selectedOutputSku <> ""
    outputNameRetained = (Trim$(CStr(lo.DataBodyRange.Cells(outputRow, _
        nameColumn).Value2)) = selectedOutputName) And selectedOutputName <> ""
    outputSkuHidden = lo.ListColumns("Output SKU").Range.EntireColumn.Hidden
    noPhysicalKey = (ProcessWorksheetColumnForTest(lo, "System_Key") = 0)

    If modProductionProcessWorksheet.SelectProcessWorksheetTableForTest( _
            mOperatorWorkbook, tableName) Then
        mBtnProcessWorksheetRetrieve_Click
        If mLstProcessOutputs.ListCount > 0 Then
            outputSkuRoundTrip = _
                (StrComp(NzStr(mLstProcessOutputs.List(0, 1)), _
                    selectedOutputName, vbTextCompare) = 0) And _
                (StrComp(NzStr(mLstProcessOutputs.List(0, 2)), _
                    selectedOutputSku, vbTextCompare) = 0) And _
                Not ProcessWorksheetTableExistsForTest(tableName)
        End If
    End If

ReportResult:
    mProduction.CloseProductionProcessItemSearchForTest
    If outputPickerOpened And outputPickerCommitted And outputSkuHidden And _
       outputSkuRoundTrip And outputNameRetained And outputNamePickerSuppressed And _
       uniqueRowIds And firstAssignedIdRetained And noPhysicalKey Then
        TestProcessWorksheetOutputPickerContract = _
            "OK|OutputPickerOpened=True|OutputPickerCommitted=True" & _
            "|OutputSkuHidden=True|OutputSkuRoundTrip=True" & _
            "|OutputNameRetained=True|OutputNamePickerSuppressed=True" & _
            "|UniqueRowIds=True|FirstAssignedIdRetained=True|NoPhysicalKey=True"
    Else
        TestProcessWorksheetOutputPickerContract = _
            "FAIL|OutputPickerOpened=" & CStr(outputPickerOpened) & _
            "|OutputPickerCommitted=" & CStr(outputPickerCommitted) & _
            "|OutputSkuHidden=" & CStr(outputSkuHidden) & _
            "|OutputSkuRoundTrip=" & CStr(outputSkuRoundTrip) & _
            "|OutputNameRetained=" & CStr(outputNameRetained) & _
            "|OutputNamePickerSuppressed=" & CStr(outputNamePickerSuppressed) & _
            "|UniqueRowIds=" & CStr(uniqueRowIds) & _
            "|FirstAssignedIdRetained=" & CStr(firstAssignedIdRetained) & _
            "|NoPhysicalKey=" & CStr(noPhysicalKey) & _
            "|OutputName=" & selectedOutputName & _
            "|OutputItem=" & selectedOutputItem & _
            "|OutputSku=" & selectedOutputSku & _
            "|Status=" & Replace$(Replace$(TestStatusText(), vbCr, " "), vbLf, " ")
    End If
End Function

Private Function ProcessWorksheetRowIdsUniqueForTest(ByVal lo As ListObject, _
                                                     ByVal recordTypeColumn As Long, _
                                                     ByVal idColumn As Long) As Boolean
    Dim used As Object
    Dim rowIndex As Long
    Dim recordType As String
    Dim rowId As String

    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare
    For rowIndex = 1 To lo.ListRows.Count
        recordType = UCase$(Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, _
            recordTypeColumn).Value2)))
        Select Case recordType
            Case "INPUT", "REQUIREMENT", "OUTPUT", "INSTRUCTION"
                rowId = UCase$(Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, _
                    idColumn).Value2)))
                If Not mProduction.IsBase36Identifier(rowId) Then Exit Function
                If used.Exists(rowId) Then Exit Function
                used.Add rowId, True
        End Select
    Next rowIndex
    ProcessWorksheetRowIdsUniqueForTest = (used.Count > 0)
End Function

Private Function ProcessWorksheetTableForTest(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    If mOperatorWorkbook Is Nothing Then Exit Function
    For Each ws In mOperatorWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set ProcessWorksheetTableForTest = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function ProcessWorksheetTableExistsForTest(ByVal tableName As String) As Boolean
    Dim lo As ListObject

    Set lo = ProcessWorksheetTableForTest(tableName)
    ProcessWorksheetTableExistsForTest = Not lo Is Nothing
End Function

Private Function ProcessWorksheetColumnForTest(ByVal lo As ListObject, _
                                               ByVal headerText As String) As Long
    On Error Resume Next
    ProcessWorksheetColumnForTest = lo.ListColumns(headerText).Index
    On Error GoTo 0
End Function

Private Function FindRecipeOutputFlowRowForTest(ByVal producedBy As String, _
                                                ByVal outputName As String, _
                                                ByVal feedsProcess As String) As Long
    Dim i As Long

    FindRecipeOutputFlowRowForTest = -1
    For i = 0 To mLstRecipeConnectionDisplay.ListCount - 1
        If StrComp(NzStr(mLstRecipeConnectionDisplay.List(i, 1)), producedBy, vbTextCompare) = 0 _
           And StrComp(NzStr(mLstRecipeConnectionDisplay.List(i, 2)), outputName, vbTextCompare) = 0 _
           And StrComp(NzStr(mLstRecipeConnectionDisplay.List(i, 3)), feedsProcess, vbTextCompare) = 0 Then
            FindRecipeOutputFlowRowForTest = i
            Exit Function
        End If
    Next i
End Function

Private Function CompatibleTargetContainsForTest(ByVal nodeId As String) As Boolean
    Dim i As Long

    For i = 0 To mCmbConnectionToNode.ListCount - 1
        If StrComp(NzStr(mCmbConnectionToNode.List(i, 0)), nodeId, vbTextCompare) = 0 Then
            CompatibleTargetContainsForTest = True
            Exit Function
        End If
    Next i
End Function

Public Function TestReusableProductionFormActionContract() As String
    Dim requiredControls As Variant
    Dim controlName As Variant
    Dim missingControls As String
    Dim boundWorkbookName As String

    If Not mBuilt Then BuildLayout
    If Not mOperatorWorkbook Is Nothing Then boundWorkbookName = mOperatorWorkbook.Name
    requiredControls = Array( _
        "btnProcessSave", "btnProcessRelease", "btnProcessObsolete", "btnProcessReuse", _
        "btnProcessWorksheetCreate", "btnProcessWorksheetRetrieve", _
        "btnProcessWorksheetAddAlternative", "cmbProcessOutputUom", _
        "btnRecipeAddProcess", "btnRecipeConnect", "btnRecipeMoveUp", _
        "btnRecipeSave", "btnRecipeRelease", "btnRecipeObsolete", _
        "btnAssignSave", "chkRunTargetOutputScale", "cmbRunTargetOutput", _
        "txtRunTargetOutputQty")

    For Each controlName In requiredControls
        If Not ProductionPageControlExists(CStr(controlName)) Then
            If missingControls <> "" Then missingControls = missingControls & ","
            missingControls = missingControls & CStr(controlName)
        End If
    Next controlName

    If boundWorkbookName = "" Then
        TestReusableProductionFormActionContract = "FAIL|BoundWorkbook=Missing"
    ElseIf missingControls <> "" Then
        TestReusableProductionFormActionContract = _
            "FAIL|BoundWorkbook=" & boundWorkbookName & _
            "|MissingControls=" & missingControls & _
            "|HandlersExercised=False"
    Else
        TestReusableProductionFormActionContract = _
            ExerciseReusableProductionFormActions(boundWorkbookName)
    End If
End Function

Private Function ExerciseReusableProductionFormActions(ByVal boundWorkbookName As String) As String
    On Error GoTo Failed

    Dim token As String
    Dim sourceId As String
    Dim sinkId As String
    Dim recipeId As String
    Dim rowIndex As Long
    Dim processSaved As Boolean
    Dim processReleased As Boolean
    Dim processObsoleted As Boolean
    Dim processReused As Boolean
    Dim alternativesSaved As Boolean
    Dim recipeConnected As Boolean
    Dim recipeOrdered As Boolean
    Dim recipeSaved As Boolean
    Dim recipeReleased As Boolean
    Dim recipeObsoleted As Boolean
    Dim recipeIdGenerated As Boolean
    Dim recipeVersionGenerated As Boolean
    Dim recipeIdLocked As Boolean
    Dim recipeVersionEditable As Boolean
    Dim editedRecipeVersionRetained As Boolean
    Dim recipeOutputNameVisible As Boolean
    Dim recipeOutputIdPreserved As Boolean
    Dim recipeUomCatalog As Boolean
    Dim recipeConnectionUpdated As Boolean
    Dim recipeConnectionId As String
    Dim recipeConnectionQty As String
    Dim recipeConnectionUom As String
    Dim recipeNodeNamesVisible As Boolean
    Dim recipeRequirementNameVisible As Boolean
    Dim recipeConnectionNamesVisible As Boolean
    Dim recipeConnectionHeaders As Boolean
    Dim recipeConnectionsFullWidth As Boolean
    Dim recipeFinishedOutputGuidance As Boolean
    Dim recipeConnectionSelected As Boolean
    Dim recipeDisconnected As Boolean
    Dim recipeSelfReferenceRejected As Boolean
    Dim recipeOutputFirstRouting As Boolean
    Dim recipeCompatibleTargetsOnly As Boolean
    Dim recipeRequirementInternallyBound As Boolean
    Dim recipeNoIngredientDropdown As Boolean
    Dim recipeForkConvergenceVisible As Boolean
    Dim recipeTerminalOutputVisible As Boolean
    Dim recipeStagesDerived As Boolean
    Dim outputYieldDefaults As Boolean
    Dim outputFlowUsesProcessYield As Boolean
    Dim processAssignmentHeaders As Boolean
    Dim assignmentKeyReadable As Boolean
    Dim acceptableItemsNamed As Boolean
    Dim processOutputEditorCompact As Boolean
    Dim processOutputUomCatalog As Boolean

    mReusableActionTestInProgress = True
    token = UCase$(Right$(CleanControlName(BuildFormGuid()), 10))
    sourceId = "PROC-SRC-" & token
    sinkId = "PROC-SINK-" & token
    RefreshProcessOutputUomCatalog "LB"
    processOutputEditorCompact = _
        Not mTxtProcessOutputItemCode.Visible And _
        (mTxtProcessOutputItemCode.Width <= 1) And _
        (Abs(mTxtProcessOutputDesignId.Left - 463) < 0.1) And _
        (Abs(mTxtProcessOutputDesignVersion.Left - 515) < 0.1) And _
        (Abs(mTxtProcessOutputQty.Left - 550) < 0.1) And _
        (Abs(mTxtProcessOutputPercent.Left - 590) < 0.1) And _
        (Abs(mTxtProcessOutputYieldBasis.Left - 630) < 0.1) And _
        (Abs(mCmbProcessOutputUom.Left - 674) < 0.1) And _
        (Abs(mCmbProcessOutputUom.Top - mTxtProcessOutputQty.Top) < 0.1)
    processOutputUomCatalog = _
        (mCmbProcessOutputUom.Style = fmStyleDropDownList) And _
        OutputUomIsCatalogValue("LB") And _
        (StrComp(ComboText(mCmbProcessOutputUom), "LB", vbTextCompare) = 0)
    processAssignmentHeaders = _
        ControlExistsByName(mPages.Pages(0), "hdrSavedProcesses1") And _
        ControlExistsByName(mPages.Pages(0), "hdrProcessRequirements1") And _
        ControlExistsByName(mPages.Pages(0), "hdrProcessOutputs1") And _
        ControlExistsByName(mPages.Pages(0), "hdrProcessInstructions1") And _
        ControlExistsByName(mPages.Pages(2), "hdrAssignmentProcesses2") And _
        ControlExistsByName(mPages.Pages(2), "hdrAssignmentRequirements2") And _
        ControlExistsByName(mPages.Pages(2), "hdrAssignmentInventory1") And _
        ControlExistsByName(mPages.Pages(2), "hdrAssignmentAllowed2")
    assignmentKeyReadable = AssignmentSystemKeyReadable()

    ClearProcessDraft False
    mTxtProcessId.Text = sourceId
    mTxtProcessVersion.Text = "1"
    mTxtProcessName.Text = "Reusable Source " & token
    mTxtProcessOutputId.Text = "001"
    mTxtProcessOutputName.Text = "Source Output"
    mTxtProcessOutputItemCode.Text = "SKU-SOURCE-" & token
    mTxtProcessOutputQty.Text = "5"
    RefreshProcessOutputUomCatalog "LB"
    mBtnProcessOutputAdd_Click
    If mLstProcessOutputs.ListCount > 0 Then
        mLstProcessOutputs.ListIndex = 0
        mLstProcessOutputs_Click
        processOutputUomCatalog = processOutputUomCatalog And _
            (StrComp(ComboText(mCmbProcessOutputUom), "LB", vbTextCompare) = 0)
        mBtnProcessOutputUpdate_Click
        outputYieldDefaults = _
            (Abs(Val(NzStr(mLstProcessOutputs.List(0, 6))) - 100#) < 0.00000001) And _
            (Abs(Val(NzStr(mLstProcessOutputs.List(0, 7))) - 5#) < 0.00000001)
        processOutputUomCatalog = processOutputUomCatalog And _
            (StrComp(NzStr(mLstProcessOutputs.List(0, 8)), "LB", vbTextCompare) = 0)
    End If
    mTxtProcessInstruction.Text = "Produce the reusable source output."
    mBtnProcessInstructionAdd_Click
    mBtnProcessSave_Click
    processSaved = (InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) > 0)
    If Not processSaved Then GoTo ActionFailed
    mBtnProcessRelease_Click
    processReleased = (InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) > 0)
    If Not processReleased Then GoTo ActionFailed

    rowIndex = FindIdentityListRow(mLstProcesses, sourceId, "1")
    If rowIndex < 0 Then GoTo ActionFailed
    mLstProcesses.ListIndex = rowIndex
    mBtnProcessReuse_Click
    processReused = (Trim$(mTxtProcessVersion.Text) = "2")
    If Not processReused Then GoTo ActionFailed
    mBtnProcessSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then GoTo ActionFailed
    mBtnProcessObsolete_Click
    processObsoleted = (InStr(1, TestStatusText(), " is OBSOLETE", vbTextCompare) > 0)
    If Not processObsoleted Then GoTo ActionFailed

    ClearProcessDraft False
    mTxtProcessId.Text = sinkId
    mTxtProcessVersion.Text = "1"
    mTxtProcessName.Text = "Reusable Sink " & token
    mTxtRequirementId.Text = "001"
    mTxtRequirementName.Text = "Source ingredient"
    mTxtRequirementQty.Text = "5"
    mTxtRequirementUom.Text = "LB"
    mBtnProcessRequirementAdd_Click
    mTxtProcessOutputId.Text = "002"
    mTxtProcessOutputName.Text = "Finished Output"
    mTxtProcessOutputItemCode.Text = "SKU-FINISHED-" & token
    mTxtProcessOutputQty.Text = "5"
    RefreshProcessOutputUomCatalog "LB"
    mBtnProcessOutputAdd_Click
    mBtnProcessSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then GoTo ActionFailed

    rowIndex = FindIdentityListRow(mLstAssignRecipes, sinkId, "1")
    If rowIndex < 0 Then GoTo ActionFailed
    mLstAssignRecipes.ListIndex = rowIndex
    mLstAssignRecipes_Click
    If mLstAssignIngredients.ListCount = 0 Then GoTo ActionFailed
    mLstAssignIngredients.ListIndex = 0
    mLstAssignIngredients_Click
    mLstAssignInventory.Clear
    mLstAssignInventory.AddItem "SYS-ACTION-" & token
    mLstAssignInventory.List(0, 1) = "Acceptable source ingredient"
    mLstAssignInventory.List(0, 2) = "LB"
    mLstAssignInventory.List(0, 6) = "SKU-SOURCE-" & token
    mLstAssignInventory.ListIndex = 0
    mBtnAssignAdd_Click
    If mLstAssignAllowed.ListCount = 0 Then GoTo ActionFailed
    acceptableItemsNamed = _
        (StrComp(NzStr(mLstAssignAllowed.List(0, 1)), _
            "Acceptable source ingredient", vbTextCompare) = 0) And _
        (StrComp(NzStr(mLstAssignAllowed.List(0, 2)), "LB", vbTextCompare) = 0) And _
        (StrComp(NzStr(mLstAssignAllowed.List(0, 3)), _
            "SKU-SOURCE-" & token, vbTextCompare) = 0)
    mBtnAssignSave_Click
    alternativesSaved = (InStr(1, TestStatusText(), "Acceptable alternatives saved", vbTextCompare) > 0)
    If Not alternativesSaved Then GoTo ActionFailed
    mBtnProcessRelease_Click
    If InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) = 0 Then GoTo ActionFailed

    mBtnRecipeNew_Click
    recipeId = Trim$(mTxtReusableRecipeId.Text)
    recipeIdGenerated = mProduction.IsBase36Identifier(recipeId)
    recipeVersionGenerated = (Trim$(mTxtReusableRecipeVersion.Text) = "1")
    recipeIdLocked = mTxtReusableRecipeId.Locked
    recipeVersionEditable = Not mTxtReusableRecipeVersion.Locked
    If Not recipeIdGenerated Or Not recipeVersionGenerated Or _
       Not recipeIdLocked Or Not recipeVersionEditable Then GoTo ActionFailed
    mTxtReusableRecipeVersion.Text = "9"
    mTxtReusableRecipeName.Text = "Reusable Recipe " & token
    RefreshReusableDesignLists
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, sourceId, "1")
    If rowIndex < 0 Then GoTo ActionFailed
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, sinkId, "2")
    If rowIndex < 0 Then GoTo ActionFailed
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    If mLstRecipeNodes.ListCount <> 2 Then GoTo ActionFailed

    recipeConnectionHeaders = _
        ControlExistsByName(mPages.Pages(1), "hdrRecipeNodes1") And _
        ControlExistsByName(mPages.Pages(1), "hdrRecipeNodes4") And _
        ControlExistsByName(mPages.Pages(1), "hdrRecipeConnections1") And _
        ControlExistsByName(mPages.Pages(1), "hdrRecipeConnections7")
    recipeConnectionsFullWidth = mLstRecipeConnectionDisplay.Visible And _
        (mLstRecipeConnectionDisplay.Left <= 12) And _
        (mLstRecipeConnectionDisplay.Width >= 1000) And _
        (mLstRecipeConnectionDisplay.Top > mLstRecipeNodes.Top + mLstRecipeNodes.Height)
    recipeFinishedOutputGuidance = _
        (StrComp(mLblRecipeFinishedOutput.Caption, _
            "Final output: leave unconnected; Production creates it as finished inventory.", _
            vbTextCompare) = 0)

    SelectComboText mCmbConnectionFromNode, NzStr(mLstRecipeNodes.List(0, 0))
    mCmbConnectionFromNode_Change
    SelectComboText mCmbConnectionOutput, "001"
    mCmbConnectionOutput_Change
    recipeOutputFirstRouting = (mCmbConnectionToNode.ListCount = 1)
    recipeCompatibleTargetsOnly = recipeOutputFirstRouting And _
        (StrComp(ConnectionTargetNodeId(), NzStr(mLstRecipeNodes.List(1, 0)), _
            vbTextCompare) = 0)
    recipeRequirementInternallyBound = _
        (StrComp(ConnectionRequirementId(), "001", vbTextCompare) = 0) And _
        (StrComp(ComboText(mCmbConnectionRequirement), "001", vbTextCompare) = 0)
    recipeNoIngredientDropdown = Not mCmbConnectionRequirement.Visible
    recipeNodeNamesVisible = _
        (mCmbConnectionFromNode.ColumnCount = 2) And _
        (mCmbConnectionToNode.ColumnCount = 4) And _
        (StrComp(NzStr(mCmbConnectionFromNode.List(mCmbConnectionFromNode.ListIndex, 1)), _
            "Reusable Source " & token, vbTextCompare) = 0) And _
        (StrComp(NzStr(mCmbConnectionToNode.List(mCmbConnectionToNode.ListIndex, 2)), _
            "Reusable Sink " & token, vbTextCompare) = 0)
    recipeOutputNameVisible = _
        (StrComp(ConnectionOutputDisplayName(), "Source Output", vbTextCompare) = 0)
    recipeRequirementNameVisible = _
        (StrComp(NzStr(mCmbConnectionToNode.List( _
            mCmbConnectionToNode.ListIndex, 3)), "Source ingredient", vbTextCompare) = 0)
    recipeUomCatalog = (mCmbConnectionUom.Style = fmStyleDropDownList) And _
        (ComboText(mCmbConnectionUom) = "LB") And _
        (mCmbConnectionUom.ListCount > 0)
    mBtnRecipeConnect_Click
    recipeConnected = (mLstRecipeConnections.ListCount = 1) And _
        (mLstRecipeConnectionDisplay.ListCount = 2)
    rowIndex = FindRecipeOutputFlowRowForTest( _
        "Reusable Source " & token, "Source Output", "Reusable Sink " & token)
    recipeConnectionNamesVisible = recipeConnected And rowIndex >= 0 And _
        (StrComp(NzStr(mLstRecipeConnectionDisplay.List(rowIndex, 0)), _
            "Stage 1", vbTextCompare) = 0)
    rowIndex = FindRecipeOutputFlowRowForTest( _
        "Reusable Sink " & token, "Finished Output", "Finished inventory")
    recipeTerminalOutputVisible = (rowIndex >= 0)
    If recipeTerminalOutputVisible Then
        recipeStagesDerived = (StrComp(NzStr( _
            mLstRecipeConnectionDisplay.List(rowIndex, 0)), "Stage 2", vbTextCompare) = 0)
    End If
    If recipeConnected Then
        recipeOutputIdPreserved = _
            (StrComp(NzStr(mLstRecipeConnections.List(0, 1)), "001", vbTextCompare) = 0)
        rowIndex = FindRecipeOutputFlowRowForTest( _
            "Reusable Source " & token, "Source Output", "Reusable Sink " & token)
        mLstRecipeConnectionDisplay.ListIndex = rowIndex
        mLstRecipeConnectionDisplay_Click
        recipeConnectionSelected = _
            (StrComp(ComboText(mCmbConnectionFromNode), _
                NzStr(mLstRecipeNodes.List(0, 0)), vbTextCompare) = 0) And _
            (StrComp(ConnectionOutputId(), "001", vbTextCompare) = 0) And _
            (StrComp(ConnectionTargetNodeId(), _
                NzStr(mLstRecipeNodes.List(1, 0)), vbTextCompare) = 0) And _
            (StrComp(ConnectionRequirementId(), "001", vbTextCompare) = 0)
        mBtnRecipeUpdateConnection_Click
        If mLstRecipeConnections.ListCount > 0 Then
            recipeConnectionId = NzStr(mLstRecipeConnections.List(0, 1))
            recipeConnectionQty = NzStr(mLstRecipeConnections.List(0, 4))
            recipeConnectionUom = NzStr(mLstRecipeConnections.List(0, 6))
        End If
        recipeConnectionUpdated = (mLstRecipeConnections.ListCount = 1) And _
            (Abs(Val(recipeConnectionQty) - 5#) < 0.00000001) And _
            (StrComp(recipeConnectionId, "001", vbTextCompare) = 0) And _
            (StrComp(recipeConnectionUom, "LB", vbTextCompare) = 0)
    End If
    If Not recipeConnected Then GoTo ActionFailed
    If Not recipeOutputNameVisible Or Not recipeOutputIdPreserved Or _
       Not recipeUomCatalog Or Not recipeConnectionUpdated Or _
       Not recipeNodeNamesVisible Or Not recipeRequirementNameVisible Or _
       Not recipeConnectionNamesVisible Or Not recipeConnectionHeaders Or _
       Not recipeConnectionsFullWidth Or Not recipeFinishedOutputGuidance Or _
       Not recipeConnectionSelected Or Not recipeOutputFirstRouting Or _
       Not recipeCompatibleTargetsOnly Or Not recipeRequirementInternallyBound Or _
       Not recipeNoIngredientDropdown Or Not recipeTerminalOutputVisible Or _
       Not recipeStagesDerived Then GoTo ActionFailed

    SelectComboText mCmbConnectionFromNode, NzStr(mLstRecipeNodes.List(1, 0))
    mCmbConnectionFromNode_Change
    SelectComboText mCmbConnectionOutput, "002"
    mCmbConnectionOutput_Change
    recipeSelfReferenceRejected = _
        Not CompatibleTargetContainsForTest(NzStr(mLstRecipeNodes.List(1, 0)))
    If Not recipeSelfReferenceRejected Then GoTo ActionFailed

    rowIndex = FindRecipeOutputFlowRowForTest( _
        "Reusable Source " & token, "Source Output", "Reusable Sink " & token)
    mLstRecipeConnectionDisplay.ListIndex = rowIndex
    mLstRecipeConnectionDisplay_Click
    mBtnRecipeDisconnect_Click
    recipeDisconnected = (mLstRecipeConnections.ListCount = 0) And _
        (FindRecipeOutputFlowRowForTest("Reusable Source " & token, _
            "Source Output", "Reusable Sink " & token) < 0)
    If Not recipeDisconnected Then GoTo ActionFailed

    SelectComboText mCmbConnectionFromNode, NzStr(mLstRecipeNodes.List(0, 0))
    mCmbConnectionFromNode_Change
    SelectComboText mCmbConnectionOutput, "001"
    mCmbConnectionOutput_Change
    mBtnRecipeConnect_Click
    recipeConnected = recipeConnected And (mLstRecipeConnections.ListCount = 1) And _
        (mLstRecipeConnectionDisplay.ListCount = 2)
    If Not recipeConnected Then GoTo ActionFailed
    mLstRecipeNodes.ListIndex = 0
    mBtnRecipeMoveDown_Click
    mBtnRecipeAutoOrder_Click
    recipeOrdered = (StrComp(NzStr(mLstRecipeNodes.List(0, 1)), sourceId, vbTextCompare) = 0)
    If Not recipeOrdered Then GoTo ActionFailed
    mBtnRecipeSave_Click
    recipeSaved = (InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) > 0)
    If Not recipeSaved Then GoTo ActionFailed
    mBtnRecipeRelease_Click
    recipeReleased = (InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) > 0)
    If Not recipeReleased Then GoTo ActionFailed
    editedRecipeVersionRetained = (Trim$(mTxtReusableRecipeVersion.Text) = "9")
    If Not editedRecipeVersionRetained Then GoTo ActionFailed
    mBtnRecipeObsolete_Click
    recipeObsoleted = (InStr(1, TestStatusText(), " is OBSOLETE", vbTextCompare) > 0)
    If Not recipeObsoleted Then GoTo ActionFailed

    mReusableTestSourceId = sourceId
    mReusableTestSinkId = sinkId
    mReusableTestRecipeId = recipeId
    recipeForkConvergenceVisible = ExerciseRecipeForkProjectionForTest( _
        token, sourceId, outputFlowUsesProcessYield)
    If Not recipeForkConvergenceVisible Then GoTo ActionFailed

    ExerciseReusableProductionFormActions = _
        "OK|BoundWorkbook=" & boundWorkbookName & _
        "|ProcessSaved=" & CStr(processSaved) & _
        "|ProcessReleased=" & CStr(processReleased) & _
        "|ProcessObsoleted=" & CStr(processObsoleted) & _
        "|ProcessReused=" & CStr(processReused) & _
        "|RecipeConnected=" & CStr(recipeConnected) & _
        "|RecipeOrdered=" & CStr(recipeOrdered) & _
        "|RecipeSaved=" & CStr(recipeSaved) & _
        "|RecipeReleased=" & CStr(recipeReleased) & _
        "|RecipeObsoleted=" & CStr(recipeObsoleted) & _
        "|RecipeIdGenerated=" & CStr(recipeIdGenerated) & _
        "|RecipeVersionGenerated=" & CStr(recipeVersionGenerated) & _
        "|RecipeIdLocked=" & CStr(recipeIdLocked) & _
        "|RecipeVersionEditable=" & CStr(recipeVersionEditable) & _
        "|EditedRecipeVersionRetained=" & CStr(editedRecipeVersionRetained)
    ExerciseReusableProductionFormActions = ExerciseReusableProductionFormActions & _
        "|RecipeOutputNameVisible=" & CStr(recipeOutputNameVisible) & _
        "|RecipeOutputIdPreserved=" & CStr(recipeOutputIdPreserved) & _
        "|RecipeUomCatalog=" & CStr(recipeUomCatalog) & _
        "|RecipeConnectionUpdated=" & CStr(recipeConnectionUpdated) & _
        "|RecipeNodeNamesVisible=" & CStr(recipeNodeNamesVisible) & _
        "|RecipeRequirementNameVisible=" & CStr(recipeRequirementNameVisible) & _
        "|RecipeConnectionNamesVisible=" & CStr(recipeConnectionNamesVisible) & _
        "|RecipeConnectionHeaders=" & CStr(recipeConnectionHeaders) & _
        "|RecipeConnectionsFullWidth=" & CStr(recipeConnectionsFullWidth) & _
        "|RecipeFinishedOutputGuidance=" & CStr(recipeFinishedOutputGuidance) & _
        "|RecipeConnectionSelected=" & CStr(recipeConnectionSelected) & _
        "|RecipeDisconnected=" & CStr(recipeDisconnected) & _
        "|RecipeSelfReferenceRejected=" & CStr(recipeSelfReferenceRejected) & _
        "|RecipeOutputFirstRouting=" & CStr(recipeOutputFirstRouting) & _
        "|RecipeCompatibleTargetsOnly=" & CStr(recipeCompatibleTargetsOnly) & _
        "|RecipeRequirementInternallyBound=" & CStr(recipeRequirementInternallyBound) & _
        "|RecipeNoIngredientDropdown=" & CStr(recipeNoIngredientDropdown)
    ExerciseReusableProductionFormActions = ExerciseReusableProductionFormActions & _
        "|RecipeForkConvergenceVisible=" & CStr(recipeForkConvergenceVisible) & _
        "|RecipeTerminalOutputVisible=" & CStr(recipeTerminalOutputVisible) & _
        "|RecipeStagesDerived=" & CStr(recipeStagesDerived) & _
        "|AlternativesSaved=" & CStr(alternativesSaved) & _
        "|OutputYieldDefaults=" & CStr(outputYieldDefaults) & _
        "|OutputFlowUsesProcessYield=" & CStr(outputFlowUsesProcessYield) & _
        "|ProcessAssignmentHeaders=" & CStr(processAssignmentHeaders) & _
        "|AssignmentSystemKeyReadable=" & CStr(assignmentKeyReadable) & _
        "|AcceptableItemsNamed=" & CStr(acceptableItemsNamed) & _
        "|ProcessOutputEditorCompact=" & CStr(processOutputEditorCompact) & _
        "|ProcessOutputUomCatalog=" & CStr(processOutputUomCatalog)
CleanExit:
    mReusableActionTestInProgress = False
    Exit Function

ActionFailed:
    ExerciseReusableProductionFormActions = "FAIL|BoundWorkbook=" & boundWorkbookName & _
        "|RecipeOutputNameVisible=" & CStr(recipeOutputNameVisible) & _
        "|RecipeOutputIdPreserved=" & CStr(recipeOutputIdPreserved) & _
        "|RecipeUomCatalog=" & CStr(recipeUomCatalog) & _
        "|RecipeConnectionUpdated=" & CStr(recipeConnectionUpdated) & _
        "|RecipeNodeNamesVisible=" & CStr(recipeNodeNamesVisible) & _
        "|RecipeRequirementNameVisible=" & CStr(recipeRequirementNameVisible) & _
        "|RecipeConnectionNamesVisible=" & CStr(recipeConnectionNamesVisible) & _
        "|RecipeConnectionHeaders=" & CStr(recipeConnectionHeaders) & _
        "|RecipeConnectionsFullWidth=" & CStr(recipeConnectionsFullWidth) & _
        "|RecipeFinishedOutputGuidance=" & CStr(recipeFinishedOutputGuidance)
    ExerciseReusableProductionFormActions = ExerciseReusableProductionFormActions & _
        "|RecipeConnectionSelected=" & CStr(recipeConnectionSelected) & _
        "|RecipeDisconnected=" & CStr(recipeDisconnected) & _
        "|RecipeSelfReferenceRejected=" & CStr(recipeSelfReferenceRejected) & _
        "|RecipeOutputFirstRouting=" & CStr(recipeOutputFirstRouting) & _
        "|RecipeCompatibleTargetsOnly=" & CStr(recipeCompatibleTargetsOnly) & _
        "|RecipeRequirementInternallyBound=" & CStr(recipeRequirementInternallyBound) & _
        "|RecipeNoIngredientDropdown=" & CStr(recipeNoIngredientDropdown) & _
        "|RecipeForkConvergenceVisible=" & CStr(recipeForkConvergenceVisible) & _
        "|RecipeTerminalOutputVisible=" & CStr(recipeTerminalOutputVisible) & _
        "|RecipeStagesDerived=" & CStr(recipeStagesDerived) & _
        "|OutputYieldDefaults=" & CStr(outputYieldDefaults) & _
        "|OutputFlowUsesProcessYield=" & CStr(outputFlowUsesProcessYield) & _
        "|ProcessAssignmentHeaders=" & CStr(processAssignmentHeaders) & _
        "|AssignmentSystemKeyReadable=" & CStr(assignmentKeyReadable) & _
        "|AcceptableItemsNamed=" & CStr(acceptableItemsNamed) & _
        "|ProcessOutputEditorCompact=" & CStr(processOutputEditorCompact) & _
        "|ProcessOutputUomCatalog=" & CStr(processOutputUomCatalog) & _
        "|ConnectionRows=" & CStr(mLstRecipeConnections.ListCount) & _
        "|ConnectionId=" & recipeConnectionId & _
        "|ConnectionQty=" & recipeConnectionQty & _
        "|ConnectionUom=" & recipeConnectionUom & _
        "|Status=" & Replace$(Replace$(TestStatusText(), vbCr, " "), vbLf, " ")
    GoTo CleanExit
Failed:
    ExerciseReusableProductionFormActions = "FAIL|" & CStr(Err.Number) & "|" & Err.Description & _
        "|Status=" & Replace$(Replace$(TestStatusText(), vbCr, " "), vbLf, " ")
    Resume CleanExit
End Function

Private Function ExerciseRecipeForkProjectionForTest(ByVal token As String, _
                                                     ByVal sourceId As String, _
                                                     ByRef outputYieldVisible As Boolean) As Boolean
    On Error GoTo Failed

    Dim branchId As String
    Dim convergeId As String
    Dim rowIndex As Long
    Dim alternative As Object
    Dim sourceNodeId As String
    Dim branchNodeId As String
    Dim convergeNodeId As String
    Dim stageMap As Object
    Dim sourceRow As Long
    Dim branchRow As Long
    Dim terminalRow As Long

    branchId = "PROC-BRANCH-" & token
    convergeId = "PROC-CONVERGE-" & token

    ClearProcessDraft False
    mTxtProcessId.Text = branchId
    mTxtProcessVersion.Text = "1"
    mTxtProcessName.Text = "Parallel Source " & token
    mTxtProcessOutputId.Text = "001"
    mTxtProcessOutputName.Text = "Parallel Output"
    mTxtProcessOutputItemCode.Text = "SKU-BRANCH-" & token
    mTxtProcessOutputQty.Text = "75.025"
    RefreshProcessOutputUomCatalog "LB"
    mBtnProcessOutputAdd_Click
    mBtnProcessSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then Exit Function
    mBtnProcessRelease_Click
    If InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) = 0 Then Exit Function

    ClearProcessDraft False
    mTxtProcessId.Text = convergeId
    mTxtProcessVersion.Text = "1"
    mTxtProcessName.Text = "Converged Product " & token
    mTxtRequirementId.Text = "001"
    mTxtRequirementName.Text = "Source input"
    mTxtRequirementQty.Text = "5"
    mTxtRequirementUom.Text = "LB"
    mBtnProcessRequirementAdd_Click
    mTxtRequirementId.Text = "002"
    mTxtRequirementName.Text = "Parallel input"
    mTxtRequirementQty.Text = "5"
    mTxtRequirementUom.Text = "LB"
    mBtnProcessRequirementAdd_Click
    Set alternative = NewReusableRecord("ALTERNATIVE")
    alternative("RequirementId") = "001"
    alternative("ITEM_CODE") = "SKU-SOURCE-" & token
    mProcessAlternatives.Add alternative
    Set alternative = NewReusableRecord("ALTERNATIVE")
    alternative("RequirementId") = "002"
    alternative("ITEM_CODE") = "SKU-BRANCH-" & token
    mProcessAlternatives.Add alternative
    mTxtProcessOutputId.Text = "003"
    mTxtProcessOutputName.Text = "Converged Output"
    mTxtProcessOutputItemCode.Text = "SKU-CONVERGED-" & token
    mTxtProcessOutputQty.Text = "10"
    RefreshProcessOutputUomCatalog "LB"
    mBtnProcessOutputAdd_Click
    mBtnProcessSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then Exit Function
    mBtnProcessRelease_Click
    If InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) = 0 Then Exit Function

    mBtnRecipeNew_Click
    mTxtReusableRecipeName.Text = "Fork Projection " & token
    RefreshReusableDesignLists
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, sourceId, "1")
    If rowIndex < 0 Then Exit Function
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, branchId, "1")
    If rowIndex < 0 Then Exit Function
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, convergeId, "1")
    If rowIndex < 0 Then Exit Function
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    If mLstRecipeNodes.ListCount <> 3 Then Exit Function

    sourceNodeId = NzStr(mLstRecipeNodes.List(0, 0))
    branchNodeId = NzStr(mLstRecipeNodes.List(1, 0))
    convergeNodeId = NzStr(mLstRecipeNodes.List(2, 0))

    SelectComboText mCmbConnectionFromNode, sourceNodeId
    mCmbConnectionFromNode_Change
    SelectComboText mCmbConnectionOutput, "001"
    mCmbConnectionOutput_Change
    If StrComp(ConnectionTargetNodeId(), convergeNodeId, vbTextCompare) <> 0 _
       Or StrComp(ConnectionRequirementId(), "001", vbTextCompare) <> 0 Then Exit Function
    mBtnRecipeConnect_Click

    SelectComboText mCmbConnectionFromNode, branchNodeId
    mCmbConnectionFromNode_Change
    SelectComboText mCmbConnectionOutput, "001"
    mCmbConnectionOutput_Change
    If StrComp(ConnectionTargetNodeId(), convergeNodeId, vbTextCompare) <> 0 _
       Or StrComp(ConnectionRequirementId(), "002", vbTextCompare) <> 0 Then Exit Function
    mBtnRecipeConnect_Click
    If mLstRecipeConnections.ListCount <> 2 Then Exit Function

    mBtnRecipeAutoOrder_Click
    Set stageMap = BuildRecipeStageMap()
    If CLng(stageMap(sourceNodeId)) <> 1 Or CLng(stageMap(branchNodeId)) <> 1 _
       Or CLng(stageMap(convergeNodeId)) <> 2 Then Exit Function

    sourceRow = FindRecipeOutputFlowRowForTest( _
        "Reusable Source " & token, "Source Output", "Converged Product " & token)
    branchRow = FindRecipeOutputFlowRowForTest( _
        "Parallel Source " & token, "Parallel Output", "Converged Product " & token)
    terminalRow = FindRecipeOutputFlowRowForTest( _
        "Converged Product " & token, "Converged Output", "Finished inventory")
    If sourceRow < 0 Or branchRow < 0 Or terminalRow < 0 Then Exit Function
    If StrComp(NzStr(mLstRecipeConnectionDisplay.List(sourceRow, 0)), _
               "Stage 1", vbTextCompare) <> 0 Then Exit Function
    If StrComp(NzStr(mLstRecipeConnectionDisplay.List(branchRow, 0)), _
               "Stage 1", vbTextCompare) <> 0 Then Exit Function
    outputYieldVisible = _
        (Abs(Val(NzStr(mLstRecipeConnectionDisplay.List(branchRow, 4))) - 75.025) < 0.00000001) And _
        (Abs(Val(NzStr(mLstRecipeConnectionDisplay.List(branchRow, 5))) - 100#) < 0.00000001)
    If Not outputYieldVisible Then Exit Function
    If StrComp(NzStr(mLstRecipeConnectionDisplay.List(terminalRow, 0)), _
               "Stage 2", vbTextCompare) <> 0 Then Exit Function

    ExerciseRecipeForkProjectionForTest = True
    Exit Function
Failed:
    ExerciseRecipeForkProjectionForTest = False
End Function

Public Function TestReusableProductionRunActionContract() As String
    On Error GoTo Failed

    Dim setupReport As String
    Dim rowIndex As Long
    Dim alternative As Object
    Dim rawSystemKey As String
    Dim rawStartQty As Double
    Dim runLocation As String
    Dim scaleMin As Boolean
    Dim scaleDefault As Boolean
    Dim scaleMax As Boolean
    Dim exactInputKeys As Boolean
    Dim distinctOutputKeys As Boolean
    Dim intermediateConsumed As Boolean
    Dim coProductRemaining As Boolean
    Dim checkedIn As Boolean
    Dim completed As Boolean
    Dim refreshed As Boolean
    Dim nextBatch As Boolean
    Dim sourceNodeId As String
    Dim sinkNodeId As String
    Dim processName As String
    Dim firstBatchKeys As Object
    Dim secondBatchKeys As Object
    Dim key As Variant
    Dim i As Long
    Dim fixtureStage As String
    Dim staleSystemKey As String
    Dim staleStartQty As Double
    Dim insufficiencyRejected As Boolean
    Dim staleRejected As Boolean
    Dim actualOutputAccepted As Boolean
    Dim lastActualDisplayed As Boolean
    Dim actualInventoryQty As Boolean
    Dim systemKeyHeadersReadable As Boolean
    Dim batchHistoryRows As Boolean
    Dim processTotalReady As Boolean
    Dim utilityDisplay As Boolean
    Dim multiProcessRunPlan As Boolean
    Dim targetOutputScaleStub As Boolean
    Dim locationStockBuckets As Boolean
    Dim locationStockExactExpansion As Boolean
    Dim rawStockStartQty As Double
    Dim selectedProcessOnly As Boolean
    Dim runInstructionsVisible As Boolean
    Dim wholeRecipeStatus As Boolean
    Dim eightPaletteRows As Boolean
    Dim sourceProcessName As String
    Dim sinkProcessName As String
    Dim upstreamWaitThenReady As Boolean
    Dim routedInputVisible As Boolean
    Dim routedInputDetails As Boolean
    Dim groupedUsedGoods As Boolean

    mReusableActionTestInProgress = True
    systemKeyHeadersReadable = ProductionRunSystemKeyHeadersReadable()
    utilityDisplay = (StrComp(modProductionReusableRun.ReusableRunInventoryDisplayForTest( _
        "FALSE", "UTILITY", "raw", 19400#), "Utility", vbTextCompare) = 0)
    If mReusableTestSourceId = "" Then
        setupReport = ExerciseReusableProductionFormActions(IIf(mOperatorWorkbook Is Nothing, _
            "", mOperatorWorkbook.Name))
        If Left$(setupReport, 3) <> "OK|" Then
            TestReusableProductionRunActionContract = "FAIL|Fixture=" & setupReport
            GoTo CleanExit
        End If
    End If
    mReusableActionTestInProgress = True

    RefreshReusableDesignLists
    rowIndex = FindIdentityListRow(mLstProcesses, mReusableTestSourceId, "1")
    If rowIndex < 0 Then GoTo FixtureFailed
    mLstProcesses.ListIndex = rowIndex
    mBtnProcessReuse_Click
    mTxtRequirementId.Text = "002"
    mTxtRequirementName.Text = "Raw inventory"
    mTxtRequirementQty.Text = "5"
    mTxtRequirementUom.Text = "LB"
    mBtnProcessRequirementAdd_Click
    Set alternative = NewReusableRecord("ALTERNATIVE")
    alternative("RequirementId") = "002"
    alternative("ITEM_CODE") = "SKU-RUN-RAW"
    mProcessAlternatives.Add alternative
    Set alternative = NewReusableRecord("ALTERNATIVE")
    alternative("RequirementId") = "002"
    alternative("ITEM_CODE") = "SKU-RUN-STALE"
    mProcessAlternatives.Add alternative
    mTxtProcessOutputId.Text = "003"
    mTxtProcessOutputName.Text = "Co Product"
    mTxtProcessOutputItemCode.Text = "SKU-RUN-CO"
    mTxtProcessOutputQty.Text = ""
    mTxtProcessOutputPercent.Text = "20"
    mTxtProcessOutputYieldBasis.Text = "10"
    RefreshProcessOutputUomCatalog "LB"
    mBtnProcessOutputAdd_Click
    mTxtProcessOutputId.Text = "004"
    mTxtProcessOutputName.Text = "Run Supply"
    mTxtProcessOutputItemCode.Text = "SKU-RUN-SUPPLY"
    mTxtProcessOutputQty.Text = "1"
    mTxtProcessOutputPercent.Text = ""
    mTxtProcessOutputYieldBasis.Text = ""
    RefreshProcessOutputUomCatalog "LB"
    mBtnProcessOutputAdd_Click
    mBtnProcessSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then GoTo FixtureFailed
    mBtnProcessRelease_Click
    If InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) = 0 Then GoTo FixtureFailed

    RefreshReusableDesignLists
    rowIndex = FindIdentityListRow(mLstProcesses, mReusableTestSinkId, "2")
    If rowIndex < 0 Then GoTo FixtureFailed
    mLstProcesses.ListIndex = rowIndex
    mBtnProcessReuse_Click
    mTxtRequirementId.Text = "003"
    mTxtRequirementName.Text = "Run supply"
    mTxtRequirementQty.Text = "1"
    mTxtRequirementUom.Text = "LB"
    mBtnProcessRequirementAdd_Click
    Set alternative = NewReusableRecord("ALTERNATIVE")
    alternative("RequirementId") = "003"
    alternative("ITEM_CODE") = "SKU-RUN-SUPPLY"
    mProcessAlternatives.Add alternative
    mTxtProcessInstruction.Text = "Finish the reusable output."
    mBtnProcessInstructionAdd_Click
    mBtnProcessSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then GoTo FixtureFailed
    mBtnProcessRelease_Click
    If InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) = 0 Then GoTo FixtureFailed

    ClearRecipeDraft False
    mTxtReusableRecipeId.Text = mReusableTestRecipeId
    mTxtReusableRecipeVersion.Text = "2"
    mTxtReusableRecipeName.Text = "Reusable Multi-output Run Recipe"
    RefreshReusableDesignLists
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, mReusableTestSourceId, "3")
    If rowIndex < 0 Then GoTo FixtureFailed
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, mReusableTestSinkId, "3")
    If rowIndex < 0 Then GoTo FixtureFailed
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    SelectComboText mCmbConnectionFromNode, NzStr(mLstRecipeNodes.List(0, 0))
    mCmbConnectionFromNode_Change
    SelectComboText mCmbConnectionOutput, "001"
    mCmbConnectionOutput_Change
    mBtnRecipeConnect_Click
    mBtnRecipeSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then GoTo FixtureFailed
    mBtnRecipeRelease_Click
    If InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) = 0 Then GoTo FixtureFailed

    fixtureStage = "CreateRawInventory"
    If mCmbRunLocation.ListCount > 0 Then
        runLocation = NzStr(mCmbRunLocation.List(0))
    Else
        runLocation = "PRODUCTION"
        mCmbRunLocation.AddItem runLocation
        mCmbTreeRunLocation.AddItem runLocation
    End If
    If Not CreateReusableRunRawInventory(runLocation, rawSystemKey, rawStartQty, setupReport) Then GoTo FixtureFailed
    rawStockStartQty = ReusableRunFixtureAvailableQty("SKU-RUN-RAW", runLocation)
    If Not ResolveReusableRunFixtureEntity("SKU-RUN-STALE", runLocation, _
            staleSystemKey, staleStartQty, setupReport) Then GoTo FixtureFailed

    fixtureStage = "LoadRecipe"
    RefreshReusableDesignLists
    rowIndex = FindIdentityListRow(mLstLoaderRecipes, mReusableTestRecipeId, "2")
    If rowIndex < 0 Then GoTo FixtureFailed
    mLstLoaderRecipes.ListIndex = rowIndex
    mBtnLoaderLoad_Click
    If Not modProductionReusableRun.ReusableRunIsLoaded() Then GoTo FixtureFailed
    For i = 0 To mLstLoaderLines.ListCount - 1
        If InStr(1, NzStr(mLstLoaderLines.List(i, 0)), "Source", vbTextCompare) > 0 Then _
            sourceNodeId = NzStr(mLstLoaderLines.List(i, 1))
        If InStr(1, NzStr(mLstLoaderLines.List(i, 0)), "Sink", vbTextCompare) > 0 Then _
            sinkNodeId = NzStr(mLstLoaderLines.List(i, 1))
    Next i
    If sourceNodeId = "" Or sinkNodeId = "" Then GoTo FixtureFailed
    sourceProcessName = ReusableRunProcessNameForNode(sourceNodeId)
    sinkProcessName = ReusableRunProcessNameForNode(sinkNodeId)
    If sourceProcessName = "" Or sinkProcessName = "" Then GoTo FixtureFailed
    multiProcessRunPlan = _
        ControlExistsByName(mPages.Pages(3), "hdrLoaderLines1") And _
        ControlExistsByName(mPages.Pages(3), "hdrLoaderLines7")
    For i = 0 To mLstRunPalette.ListCount - 1
        processName = ReusableRunProcessNameForNode(NzStr(mLstRunPalette.List(i, 0)))
        If processName <> "" And InStr(1, NzStr(mLstRunPalette.List(i, 2)), _
                processName & " / ", vbTextCompare) = 1 Then
            multiProcessRunPlan = multiProcessRunPlan And True
            Exit For
        End If
    Next i
    If i >= mLstRunPalette.ListCount Then multiProcessRunPlan = False
    wholeRecipeStatus = LoaderHasReusableStatus("! INSUFFICIENT")
    upstreamWaitThenReady = LoaderHasReusableStatus("WAITING UPSTREAM")
    eightPaletteRows = (mLstRunPalette.Height >= 90)
    SelectComboText mCmbRunProcess, sourceProcessName
    mCmbRunProcess_Change
    runInstructionsVisible = _
        (mLstRunInstructions.ListCount > 0) And _
        (InStr(1, NzStr(mLstRunInstructions.List(0, 1)), _
            "Produce the reusable source output", vbTextCompare) > 0)
    locationStockBuckets = _
        (ReusablePaletteRowsForItemName("Production Raw Material") = 1)
    targetOutputScaleStub = _
        (StrComp(mChkRunTargetOutputScale.Caption, _
            "Scale from target output Qty (coming later)", vbTextCompare) = 0) And _
        Not mChkRunTargetOutputScale.Enabled And _
        Not mCmbRunTargetOutput.Enabled And _
        Not mTxtRunTargetOutputQty.Enabled And _
        (mCmbRunTargetOutput.ListCount > 0)

    fixtureStage = "ScaleBounds"
    mTxtBatchScalePercent.Text = "0.001"
    mBtnApplyBatchScale_Click
    scaleMin = (Abs(modProductionReusableRun.ReusableRunScalePercent() - 0.001) < 0.00000001)
    mTxtBatchScalePercent.Text = "1000"
    mBtnApplyBatchScale_Click
    scaleMax = (Abs(modProductionReusableRun.ReusableRunScalePercent() - 1000#) < 0.00000001)
    mTxtBatchScalePercent.Text = "100"
    mBtnApplyBatchScale_Click
    scaleDefault = (Abs(modProductionReusableRun.ReusableRunScalePercent() - 100#) < 0.00000001)
    If Not scaleMin Or Not scaleDefault Or Not scaleMax Then GoTo FixtureFailed

    fixtureStage = "SufficiencyAndStale"
    SelectComboText mCmbRunLocation, runLocation
    mCmbRunLocation_Change
    mLstRunPalette.ListIndex = FindReusablePaletteSystemKey(staleSystemKey)
    If mLstRunPalette.ListIndex < 0 Then GoTo FixtureFailed
    mTxtPaletteQty.Text = CStr(staleStartQty + 1#)
    mBtnRunApplyPalette_Click
    insufficiencyRejected = (InStr(1, TestStatusText(), "exceeds the scaled requirement", vbTextCompare) > 0 _
        Or InStr(1, TestStatusText(), "exceeds exact entity availability", vbTextCompare) > 0)
    If Not insufficiencyRejected Then GoTo FixtureFailed
    mLstRunPalette.ListIndex = FindReusablePaletteSystemKey(staleSystemKey)
    mTxtPaletteQty.Text = "5"
    mBtnRunApplyPalette_Click
    If Not ConsumeReusableRunFixtureEntity(staleSystemKey, "SKU-RUN-STALE", 3#, runLocation, setupReport) Then GoTo FixtureFailed
    mBtnManagerCheckIn_Click
    staleRejected = (InStr(1, TestStatusText(), "Stale allocation rejected", vbTextCompare) > 0)
    If Not staleRejected Then GoTo FixtureFailed
    mTxtBatchScalePercent.Text = "100"
    mBtnApplyBatchScale_Click

    fixtureStage = "Batch1Allocate"
    SelectComboText mCmbRunProcess, sourceProcessName
    mCmbRunProcess_Change
    SelectComboText mCmbRunLocation, runLocation
    mCmbRunLocation_Change
    If mLstRunPalette.ListCount = 0 Then GoTo FixtureFailed
    mLstRunPalette.ListIndex = FindReusablePaletteItemName("Production Raw Material")
    If mLstRunPalette.ListIndex < 0 Then GoTo FixtureFailed
    mTxtPaletteSplit.Text = "100"
    mTxtPaletteQty.Text = "5"
    mBtnRunApplyPalette_Click
    locationStockExactExpansion = _
        (modProductionReusableRun.ReusableRunExactAllocationCountForRequirement( _
            sourceNodeId, "002") >= 2)
    exactInputKeys = locationStockExactExpansion
    fixtureStage = "Batch1CheckIn"
    mBtnManagerCheckIn_Click
    checkedIn = modProductionReusableRun.ReusableRunIsCheckedIn()
    If Not checkedIn Then GoTo FixtureFailed
    fixtureStage = "Batch1Complete"
    actualOutputAccepted = StageReusableTestActualOutputs(4.25)
    If Not actualOutputAccepted Then GoTo FixtureFailed
    mBtnManagerApplyOutput_Click
    selectedProcessOnly = _
        modProductionReusableRun.ReusableRunIsProcessComplete(sourceProcessName) And _
        Not modProductionReusableRun.ReusableRunIsProcessComplete(sinkProcessName) And _
        Not modProductionReusableRun.ReusableRunIsCompleted()
    If Not selectedProcessOnly Then GoTo FixtureFailed
    mBtnManagerRefresh_Click
    SelectComboText mCmbRunProcess, sinkProcessName
    mCmbRunProcess_Change
    upstreamWaitThenReady = upstreamWaitThenReady And Not LoaderHasReusableStatus("WAITING UPSTREAM")
    wholeRecipeStatus = wholeRecipeStatus And LoaderHasReusableStatus("NEEDS ALLOCATION")
    fixtureStage = "Batch1SinkAllocate"
    SelectComboText mCmbRunProcess, sinkProcessName
    mCmbRunProcess_Change
    mLstRunPalette.ListIndex = FindReusablePaletteItemName("Run Supply")
    If mLstRunPalette.ListIndex < 0 Then GoTo FixtureFailed
    mTxtPaletteSplit.Text = "100"
    mTxtPaletteQty.Text = "1"
    mBtnRunApplyPalette_Click
    exactInputKeys = exactInputKeys And _
        (modProductionReusableRun.ReusableRunExactAllocationCountForRequirement( _
            sinkNodeId, "003") >= 1)
    fixtureStage = "Batch1SinkCheckIn"
    mBtnManagerCheckIn_Click
    If Not modProductionReusableRun.ReusableRunIsCheckedIn() Then GoTo FixtureFailed
    routedInputVisible = ManagerCheckHasRoutedInputForTest( _
        modProductionReusableRun.ReusableRunOutputSystemKey(sourceNodeId, "001"), _
        sourceProcessName, "Source Output", "Source ingredient")
    routedInputDetails = routedInputVisible And _
        (mLstManagerCheck.ColumnCount = 9) And _
        (mLstRunPalette.ListCount = 1)
    groupedUsedGoods = (StrComp(modProductionReusableRun.ReusableRunUsedGoodsDisplayForTest( _
        "lb", 5#, "EA", 12#), "12 EA; 5 LB", vbTextCompare) = 0)
    fixtureStage = "Batch1SinkComplete"
    actualOutputAccepted = actualOutputAccepted And _
        StageReusableTestActualOutputs(4.25)
    If Not actualOutputAccepted Then GoTo FixtureFailed
    mBtnManagerApplyOutput_Click
    completed = modProductionReusableRun.ReusableRunIsCompleted()
    If Not completed Then GoTo FixtureFailed
    Set firstBatchKeys = CaptureReusableOutputKeys()
    If firstBatchKeys.Count <> 4 Then GoTo FixtureFailed
    intermediateConsumed = (Abs(modProductionReusableRun.ReusableRunExactEntityQty( _
        modProductionReusableRun.ReusableRunOutputSystemKey(sourceNodeId, "001"))) < 0.0000001)
    coProductRemaining = (Abs(modProductionReusableRun.ReusableRunExactEntityQty( _
        modProductionReusableRun.ReusableRunOutputSystemKey(sourceNodeId, "003")) - 2#) < 0.0000001)
    lastActualDisplayed = (Abs(ReusableOutputListValueForBatch( _
        "Finished Output", 1, 3) - 4.25) < 0.0000001)
    actualInventoryQty = (Abs(modProductionReusableRun.ReusableRunExactEntityQty( _
        modProductionReusableRun.ReusableRunOutputSystemKey(sinkNodeId, "002")) - 4.25) < 0.0000001)
    exactInputKeys = exactInputKeys And _
        (Abs(ReusableRunFixtureAvailableQty("SKU-RUN-RAW", runLocation) - _
             (rawStockStartQty - 5#)) < 0.0000001)
    fixtureStage = "RefreshAndNext"
    mBtnManagerRefresh_Click
    refreshed = (InStr(1, TestStatusText(), "refreshed", vbTextCompare) > 0)
    mBtnManagerNext_Click
    nextBatch = (modProductionReusableRun.ReusableRunBatchNumber() = 2 And _
                  Not modProductionReusableRun.ReusableRunIsCheckedIn())
    batchHistoryRows = (ReusableOutputRowsForBatch(1) = 4 And _
                        ReusableActiveOutputRowCount() = 4)
    If Not nextBatch Then GoTo FixtureFailed

    fixtureStage = "Batch2Allocate"
    SelectComboText mCmbRunProcess, sourceProcessName
    mCmbRunProcess_Change
    SelectComboText mCmbRunLocation, runLocation
    mCmbRunLocation_Change
    mLstRunPalette.ListIndex = FindReusablePaletteItemName("Production Raw Material")
    If mLstRunPalette.ListIndex < 0 Then GoTo FixtureFailed
    mTxtPaletteSplit.Text = "100"
    mTxtPaletteQty.Text = "5"
    mBtnRunApplyPalette_Click
    fixtureStage = "Batch2CheckIn"
    mBtnManagerCheckIn_Click
    If Not modProductionReusableRun.ReusableRunIsCheckedIn() Then GoTo FixtureFailed
    fixtureStage = "Batch2Complete"
    actualOutputAccepted = actualOutputAccepted And StageReusableTestActualOutputs(4.5)
    If Not actualOutputAccepted Then GoTo FixtureFailed
    mBtnManagerApplyOutput_Click
    If Not modProductionReusableRun.ReusableRunIsProcessComplete(sourceProcessName) Or _
       modProductionReusableRun.ReusableRunIsCompleted() Then GoTo FixtureFailed
    fixtureStage = "Batch2SinkAllocate"
    SelectComboText mCmbRunProcess, sinkProcessName
    mCmbRunProcess_Change
    mLstRunPalette.ListIndex = FindReusablePaletteItemName("Run Supply")
    If mLstRunPalette.ListIndex < 0 Then GoTo FixtureFailed
    mTxtPaletteSplit.Text = "100"
    mTxtPaletteQty.Text = "1"
    mBtnRunApplyPalette_Click
    fixtureStage = "Batch2SinkCheckIn"
    mBtnManagerCheckIn_Click
    If Not modProductionReusableRun.ReusableRunIsCheckedIn() Then GoTo FixtureFailed
    fixtureStage = "Batch2SinkComplete"
    actualOutputAccepted = actualOutputAccepted And StageReusableTestActualOutputs(4.5)
    If Not actualOutputAccepted Then GoTo FixtureFailed
    mBtnManagerApplyOutput_Click
    If Not modProductionReusableRun.ReusableRunIsCompleted() Then GoTo FixtureFailed
    batchHistoryRows = batchHistoryRows And _
        (ReusableOutputRowsForBatch(1) = 4 And ReusableOutputRowsForBatch(2) = 4)
    processTotalReady = (Abs(ReusableOutputListValueForBatch( _
        "Finished Output", 2, 6) - 8.75) < 0.0000001)
    Set secondBatchKeys = CaptureReusableOutputKeys()
    If secondBatchKeys.Count <> 4 Then GoTo FixtureFailed
    distinctOutputKeys = True
    For Each key In secondBatchKeys.Keys
        If firstBatchKeys.Exists(CStr(key)) Then distinctOutputKeys = False
    Next key
    distinctOutputKeys = distinctOutputKeys And (firstBatchKeys.Count + secondBatchKeys.Count = 8)
    intermediateConsumed = intermediateConsumed And _
        (Abs(modProductionReusableRun.ReusableRunExactEntityQty( _
            modProductionReusableRun.ReusableRunOutputSystemKey(sourceNodeId, "001"))) < 0.0000001)
    coProductRemaining = coProductRemaining And _
        (Abs(modProductionReusableRun.ReusableRunExactEntityQty( _
            modProductionReusableRun.ReusableRunOutputSystemKey(sourceNodeId, "003")) - 2#) < 0.0000001)
    lastActualDisplayed = lastActualDisplayed And _
        (Abs(ReusableOutputListValueForBatch( _
            "Finished Output", 2, 3) - 4.5) < 0.0000001)
    actualInventoryQty = actualInventoryQty And _
        (Abs(modProductionReusableRun.ReusableRunExactEntityQty( _
            modProductionReusableRun.ReusableRunOutputSystemKey(sinkNodeId, "002")) - 4.5) < 0.0000001)
    exactInputKeys = exactInputKeys And _
        (Abs(ReusableRunFixtureAvailableQty("SKU-RUN-RAW", runLocation) - _
             (rawStockStartQty - 10#)) < 0.0000001)

    TestReusableProductionRunActionContract = _
        "OK|ReusableRecipe=" & mReusableTestRecipeId & _
        "|Version=2|Batches=2|ScaleMin=" & CStr(scaleMin) & _
        "|ScaleDefault=" & CStr(scaleDefault) & "|ScaleMax=" & CStr(scaleMax) & _
        "|ExactInputKeys=" & CStr(exactInputKeys) & _
        "|InsufficiencyRejected=" & CStr(insufficiencyRejected) & _
        "|StaleRejected=" & CStr(staleRejected) & _
        "|DistinctOutputKeys=" & CStr(distinctOutputKeys) & _
        "|IntermediateConsumed=" & CStr(intermediateConsumed) & _
        "|CoProductRemaining=" & CStr(coProductRemaining) & _
        "|PercentageYieldBasis=" & CStr(coProductRemaining) & _
        "|ActualOutputAccepted=" & CStr(actualOutputAccepted) & _
        "|LastActualDisplayed=" & CStr(lastActualDisplayed) & _
        "|ActualInventoryQty=" & CStr(actualInventoryQty) & _
        "|SystemKeyHeadersReadable=" & CStr(systemKeyHeadersReadable) & _
        "|BatchHistoryRows=" & CStr(batchHistoryRows) & _
        "|ProcessTotal=" & CStr(processTotalReady) & _
        "|UtilityDisplay=" & CStr(utilityDisplay)
    TestReusableProductionRunActionContract = _
        TestReusableProductionRunActionContract & _
        "|MultiProcessRunPlan=" & CStr(multiProcessRunPlan) & _
        "|TargetOutputScaleStub=" & CStr(targetOutputScaleStub) & _
        "|LocationStockBuckets=" & CStr(locationStockBuckets) & _
        "|LocationStockExactExpansion=" & CStr(locationStockExactExpansion) & _
        "|SelectedProcessOnly=" & CStr(selectedProcessOnly) & _
        "|RunInstructionsVisible=" & CStr(runInstructionsVisible) & _
        "|WholeRecipeStatus=" & CStr(wholeRecipeStatus) & _
        "|EightPaletteRows=" & CStr(eightPaletteRows) & _
        "|UpstreamWaitThenReady=" & CStr(upstreamWaitThenReady) & _
        "|RoutedInputVisible=" & CStr(routedInputVisible) & _
        "|RoutedInputDetails=" & CStr(routedInputDetails) & _
        "|GroupedUsedGoods=" & CStr(groupedUsedGoods) & _
        "|CheckedIn=" & CStr(checkedIn) & "|Completed=" & CStr(completed) & _
        "|Refreshed=" & CStr(refreshed) & "|NextBatch=" & CStr(nextBatch)
CleanExit:
    mReusableActionTestInProgress = False
    Exit Function
FixtureFailed:
    TestReusableProductionRunActionContract = "FAIL|FixtureStage=" & fixtureStage & _
        "|Setup=" & Replace$(Replace$(setupReport, vbCr, " "), vbLf, " ") & _
        "|Status=" & Replace$(Replace$(TestStatusText(), vbCr, " "), vbLf, " ")
    GoTo CleanExit
Failed:
    TestReusableProductionRunActionContract = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
    Resume CleanExit
End Function

Public Function TestChaiForkConvergenceRunActionContract() As String
    On Error GoTo Failed

    Dim token As String
    Dim teaId As String
    Dim spiceId As String
    Dim cookId As String
    Dim bottlingId As String
    Dim recipeId As String
    Dim rowIndex As Long
    Dim runLocation As String
    Dim rawSystemKey As String
    Dim rawQty As Double
    Dim setupReport As String
    Dim teaNodeId As String
    Dim spiceNodeId As String
    Dim cookNodeId As String
    Dim bottlingNodeId As String
    Dim initialWaiting As Boolean
    Dim batchNoteText As String
    Dim batchNoteFrozen As Boolean
    Dim batchNoteRetained As Boolean
    Dim runIdentityAfterTea As String
    Dim runIdentityAfterRefresh As String
    Dim runIdentityAfterNavigation As String
    Dim brewedTeaKey As String
    Dim spiceBlendKey As String
    Dim concentrateKey As String
    Dim bottledKey As String
    Dim routedCookInputs As Boolean
    Dim routedBottleInput As Boolean
    Dim upstreamConsumed As Boolean
    Dim finalBottlingCompleted As Boolean
    Dim allFourComplete As Boolean
    Dim runNotRestarted As Boolean
    Dim finalKeyCreated As Boolean
    Dim i As Long
    Dim fixtureStage As String

    If Not mBuilt Then BuildLayout
    mReusableActionTestInProgress = True
    token = UCase$(Right$(CleanControlName(BuildFormGuid()), 10))
    teaId = "CHAI-TEA-" & token
    spiceId = "CHAI-SPICE-" & token
    cookId = "CHAI-COOK-" & token
    bottlingId = "CHAI-BOTTLE-" & token

    fixtureStage = "CreateTea"
    If Not CreateChaiProcessForTest(teaId, "Chai Tea Brewing " & token, _
            Array(Array("Black Tea", "SKU-RUN-RAW", "LB", "1")), _
            "Brewed Black Tea", "SKU-CHAI-TEA-" & token, 1#, "LB") Then GoTo FixtureFailed
    fixtureStage = "CreateSpice"
    If Not CreateChaiProcessForTest(spiceId, "Chai Spice Blending " & token, _
            Array(Array("Chai Spices", "SKU-RUN-RAW", "LB", "1")), _
            "Classic Chai Spice Blend", "SKU-CHAI-SPICE-" & token, 1#, "LB") Then GoTo FixtureFailed
    fixtureStage = "CreateConvergence"
    If Not CreateChaiProcessForTest(cookId, "Chai Convergence " & token, _
            Array(Array("Brewed Black Tea", "SKU-CHAI-TEA-" & token, "LB", "1"), _
                  Array("Classic Chai Spice Blend", "SKU-CHAI-SPICE-" & token, "LB", "1")), _
            "Chai Concentrate", "SKU-CHAI-CONCENTRATE-" & token, 2#, "LB") Then GoTo FixtureFailed
    fixtureStage = "CreateBottling"
    If Not CreateChaiProcessForTest(bottlingId, "Chai Final Bottling " & token, _
            Array(Array("Chai Concentrate", "SKU-CHAI-CONCENTRATE-" & token, "LB", "2")), _
            "64oz Classic Chai Bottle", "SKU-CHAI-64OZ-" & token, 1#, "EA") Then GoTo FixtureFailed

    fixtureStage = "CreateRecipe"
    ClearRecipeDraft False
    mBtnRecipeNew_Click
    recipeId = Trim$(mTxtReusableRecipeId.Text)
    mTxtReusableRecipeName.Text = "Classic Chai Fork Convergence " & token
    RefreshReusableDesignLists
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, teaId, "1")
    If rowIndex < 0 Then GoTo FixtureFailed
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, spiceId, "1")
    If rowIndex < 0 Then GoTo FixtureFailed
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, cookId, "1")
    If rowIndex < 0 Then GoTo FixtureFailed
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    rowIndex = FindIdentityListRow(mLstReleasedProcesses, bottlingId, "1")
    If rowIndex < 0 Then GoTo FixtureFailed
    mLstReleasedProcesses.ListIndex = rowIndex
    mBtnRecipeAddProcess_Click
    If mLstRecipeNodes.ListCount <> 4 Then GoTo FixtureFailed
    teaNodeId = NzStr(mLstRecipeNodes.List(0, 0))
    spiceNodeId = NzStr(mLstRecipeNodes.List(1, 0))
    cookNodeId = NzStr(mLstRecipeNodes.List(2, 0))
    bottlingNodeId = NzStr(mLstRecipeNodes.List(3, 0))
    fixtureStage = "ConnectTea"
    If Not ConnectChaiTestNodes(teaNodeId, "001") Then GoTo FixtureFailed
    fixtureStage = "ConnectSpice"
    If Not ConnectChaiTestNodes(spiceNodeId, "001") Then GoTo FixtureFailed
    fixtureStage = "ConnectBottling"
    If Not ConnectChaiTestNodes(cookNodeId, "001") Then GoTo FixtureFailed
    fixtureStage = "VerifyConnections=" & CStr(mLstRecipeConnections.ListCount)
    If mLstRecipeConnections.ListCount <> 3 Then GoTo FixtureFailed
    mBtnRecipeAutoOrder_Click
    mBtnRecipeSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then GoTo FixtureFailed
    mBtnRecipeRelease_Click
    If InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) = 0 Then GoTo FixtureFailed

    If mCmbRunLocation.ListCount > 0 Then
        runLocation = NzStr(mCmbRunLocation.List(0))
    Else
        runLocation = "PRODUCTION"
        mCmbRunLocation.AddItem runLocation
        mCmbTreeRunLocation.AddItem runLocation
    End If
    fixtureStage = "ResolveRaw"
    If Not ResolveReusableRunFixtureEntity("SKU-RUN-RAW", runLocation, _
            rawSystemKey, rawQty, setupReport) Then GoTo FixtureFailed
    If rawQty < 2# Then GoTo FixtureFailed

    fixtureStage = "LoadRecipe"
    RefreshReusableDesignLists
    rowIndex = FindIdentityListRow(mLstLoaderRecipes, recipeId, "1")
    If rowIndex < 0 Then GoTo FixtureFailed
    mLstLoaderRecipes.ListIndex = rowIndex
    mBtnLoaderLoad_Click
    If Not modProductionReusableRun.ReusableRunIsLoaded() Then GoTo FixtureFailed
    For i = 0 To mLstLoaderLines.ListCount - 1
        If StrComp(NzStr(mLstLoaderLines.List(i, 0)), _
                "Chai Tea Brewing " & token, vbTextCompare) = 0 Then _
            teaNodeId = NzStr(mLstLoaderLines.List(i, 1))
        If StrComp(NzStr(mLstLoaderLines.List(i, 0)), _
                "Chai Spice Blending " & token, vbTextCompare) = 0 Then _
            spiceNodeId = NzStr(mLstLoaderLines.List(i, 1))
        If StrComp(NzStr(mLstLoaderLines.List(i, 0)), _
                "Chai Convergence " & token, vbTextCompare) = 0 Then _
            cookNodeId = NzStr(mLstLoaderLines.List(i, 1))
        If StrComp(NzStr(mLstLoaderLines.List(i, 0)), _
                "Chai Final Bottling " & token, vbTextCompare) = 0 Then _
            bottlingNodeId = NzStr(mLstLoaderLines.List(i, 1))
    Next i
    If teaNodeId = "" Or spiceNodeId = "" Or cookNodeId = "" Or _
       bottlingNodeId = "" Then GoTo FixtureFailed
    initialWaiting = LoaderHasReusableStatus("WAITING UPSTREAM")
    If Not initialWaiting Then GoTo FixtureFailed

    batchNoteText = "Slice 4bc Chai batch " & token
    mTxtRunBatchNote.Text = batchNoteText

    fixtureStage = "CompleteTea"
    If Not CompleteChaiTestProcess("Chai Tea Brewing " & token, 1#) Then GoTo FixtureFailed
    brewedTeaKey = modProductionReusableRun.ReusableRunOutputSystemKey(teaNodeId, "001")
    If brewedTeaKey = "" Then brewedTeaKey = ReusableOutputSystemKeyForTest( _
        "Chai Tea Brewing " & token, "Brewed Black Tea")
    batchNoteFrozen = modProductionReusableRun.ReusableRunBatchNoteFrozen() And _
        StrComp(modProductionReusableRun.ReusableRunBatchNote(), batchNoteText, vbBinaryCompare) = 0
    runIdentityAfterTea = modProductionReusableRun.ReusableRunIdentityForTest()
    mBtnManagerRefresh_Click
    runIdentityAfterRefresh = modProductionReusableRun.ReusableRunIdentityForTest()
    SelectComboText mCmbRunProcess, "Chai Spice Blending " & token
    mCmbRunProcess_Change
    runIdentityAfterNavigation = modProductionReusableRun.ReusableRunIdentityForTest()
    batchNoteRetained = batchNoteFrozen And _
        StrComp(mTxtRunBatchNote.Text, batchNoteText, vbBinaryCompare) = 0 And _
        Not mTxtRunBatchNote.Enabled
    fixtureStage = "CompleteSpice"
    If Not CompleteChaiTestProcess("Chai Spice Blending " & token, 1#) Then GoTo FixtureFailed
    spiceBlendKey = modProductionReusableRun.ReusableRunOutputSystemKey(spiceNodeId, "001")
    If spiceBlendKey = "" Then spiceBlendKey = ReusableOutputSystemKeyForTest( _
        "Chai Spice Blending " & token, "Classic Chai Spice Blend")

    fixtureStage = "CompleteConvergence"
    SelectComboText mCmbRunProcess, "Chai Convergence " & token
    mCmbRunProcess_Change
    mBtnManagerCheckIn_Click
    routedCookInputs = modProductionReusableRun.ReusableRunIsCheckedIn() And _
        ManagerCheckHasRoutedInputForTest(brewedTeaKey, "Chai Tea Brewing " & token, _
            "Brewed Black Tea", "Brewed Black Tea") And _
        ManagerCheckHasRoutedInputForTest(spiceBlendKey, "Chai Spice Blending " & token, _
            "Classic Chai Spice Blend", "Classic Chai Spice Blend")
    If Not routedCookInputs Then _
        setupReport = "BrewedKey=" & brewedTeaKey & "|SpiceKey=" & spiceBlendKey & _
            "|CheckRows=" & ManagerCheckRowsSummaryForTest()
    If Not routedCookInputs Then GoTo FixtureFailed
    If Not StageReusableTestActualOutputs(2#) Then GoTo FixtureFailed
    mBtnManagerApplyOutput_Click
    If Not modProductionReusableRun.ReusableRunIsProcessComplete("Chai Convergence " & token) Then GoTo FixtureFailed
    concentrateKey = modProductionReusableRun.ReusableRunOutputSystemKey(cookNodeId, "001")
    If concentrateKey = "" Then concentrateKey = ReusableOutputSystemKeyForTest( _
        "Chai Convergence " & token, "Chai Concentrate")
    upstreamConsumed = _
        (Abs(modProductionReusableRun.ReusableRunExactEntityQty(brewedTeaKey)) < 0.0000001) And _
        (Abs(modProductionReusableRun.ReusableRunExactEntityQty(spiceBlendKey)) < 0.0000001)
    If Not upstreamConsumed Then GoTo FixtureFailed

    fixtureStage = "CompleteBottling"
    mBtnManagerRefresh_Click
    SelectComboText mCmbRunProcess, "Chai Final Bottling " & token
    mCmbRunProcess_Change
    mBtnManagerCheckIn_Click
    routedBottleInput = modProductionReusableRun.ReusableRunIsCheckedIn() And _
        ManagerCheckHasRoutedInputForTest(concentrateKey, "Chai Convergence " & token, _
            "Chai Concentrate", "Chai Concentrate")
    If Not routedBottleInput Then GoTo FixtureFailed
    If Not StageReusableTestActualOutputs(1#) Then GoTo FixtureFailed
    mBtnManagerApplyOutput_Click
    bottledKey = modProductionReusableRun.ReusableRunOutputSystemKey(bottlingNodeId, "001")
    If bottledKey = "" Then bottledKey = ReusableOutputSystemKeyForTest( _
        "Chai Final Bottling " & token, "64oz Classic Chai Bottle")
    finalKeyCreated = (bottledKey <> "") And _
        (Abs(modProductionReusableRun.ReusableRunExactEntityQty(bottledKey) - 1#) < 0.0000001)
    finalBottlingCompleted = modProductionReusableRun.ReusableRunIsProcessComplete( _
        "Chai Final Bottling " & token) And modProductionReusableRun.ReusableRunIsCompleted()
    allFourComplete = finalBottlingCompleted And _
        modProductionReusableRun.ReusableRunIsProcessComplete("Chai Tea Brewing " & token) And _
        modProductionReusableRun.ReusableRunIsProcessComplete("Chai Spice Blending " & token) And _
        modProductionReusableRun.ReusableRunIsProcessComplete("Chai Convergence " & token)
    runNotRestarted = (runIdentityAfterTea <> "") And _
        (StrComp(runIdentityAfterTea, runIdentityAfterRefresh, vbTextCompare) = 0) And _
        (StrComp(runIdentityAfterTea, runIdentityAfterNavigation, vbTextCompare) = 0) And _
        (InStr(1, runIdentityAfterTea, "|Batch=1", vbTextCompare) > 0)

    TestChaiForkConvergenceRunActionContract = "OK|Recipe=" & recipeId & _
        "|ChaiInitialWaitingUpstream=" & CStr(initialWaiting) & _
        "|ChaiBatchNoteFrozen=" & CStr(batchNoteFrozen) & _
        "|ChaiBatchNoteRetained=" & CStr(batchNoteRetained) & _
        "|ChaiRoutedConvergenceInputs=" & CStr(routedCookInputs) & _
        "|ChaiUpstreamExactKeysConsumed=" & CStr(upstreamConsumed) & _
        "|ChaiFinalBottlingRoutedInput=" & CStr(routedBottleInput) & _
        "|ChaiFinalOutputNewKey=" & CStr(finalKeyCreated) & _
        "|ChaiFourProcessesCompleted=" & CStr(allFourComplete) & _
        "|ChaiFinalBottlingCompleted=" & CStr(finalBottlingCompleted) & _
        "|ChaiRunNotRestarted=" & CStr(runNotRestarted)
CleanExit:
    mReusableActionTestInProgress = False
    Exit Function
FixtureFailed:
    TestChaiForkConvergenceRunActionContract = "FAIL|FixtureStage=" & fixtureStage & _
        "|Fixture=" & setupReport & _
        "|Status=" & Replace$(Replace$(TestStatusText(), vbCr, " "), vbLf, " ")
    GoTo CleanExit
Failed:
    TestChaiForkConvergenceRunActionContract = "FAIL|" & CStr(Err.Number) & _
        "|" & Err.Description
    Resume CleanExit
End Function

Private Function CreateChaiProcessForTest(ByVal processId As String, _
                                          ByVal processName As String, _
                                          ByVal requirements As Variant, _
                                          ByVal outputName As String, _
                                          ByVal outputItemCode As String, _
                                          ByVal outputQty As Double, _
                                          ByVal outputUom As String) As Boolean
    Dim requirement As Variant
    Dim alternative As Object
    Dim requirementId As String

    ClearProcessDraft False
    mTxtProcessId.Text = processId
    mTxtProcessVersion.Text = "1"
    mTxtProcessName.Text = processName
    For Each requirement In requirements
        requirementId = Format$(mLstProcessRequirements.ListCount + 1, "000")
        mTxtRequirementId.Text = requirementId
        mTxtRequirementName.Text = CStr(requirement(0))
        mTxtRequirementQty.Text = CStr(requirement(3))
        mTxtRequirementUom.Text = CStr(requirement(2))
        mBtnProcessRequirementAdd_Click
        Set alternative = NewReusableRecord("ALTERNATIVE")
        alternative("RequirementId") = requirementId
        alternative("ITEM_CODE") = CStr(requirement(1))
        mProcessAlternatives.Add alternative
    Next requirement
    mTxtProcessOutputId.Text = "001"
    mTxtProcessOutputName.Text = outputName
    mTxtProcessOutputItemCode.Text = outputItemCode
    mTxtProcessOutputQty.Text = FormatRunNumber(outputQty)
    RefreshProcessOutputUomCatalog outputUom
    mBtnProcessOutputAdd_Click
    mBtnProcessSave_Click
    If InStr(1, TestStatusText(), " is DRAFT", vbTextCompare) = 0 Then Exit Function
    mBtnProcessRelease_Click
    CreateChaiProcessForTest = (InStr(1, TestStatusText(), " is RELEASED", vbTextCompare) > 0)
End Function

Private Function ConnectChaiTestNodes(ByVal sourceNodeId As String, _
                                      ByVal outputId As String) As Boolean
    SelectComboText mCmbConnectionFromNode, sourceNodeId
    mCmbConnectionFromNode_Change
    SelectComboText mCmbConnectionOutput, outputId
    mCmbConnectionOutput_Change
    If Trim$(ConnectionTargetNodeId()) = "" Or Trim$(ConnectionRequirementId()) = "" Then Exit Function
    mBtnRecipeConnect_Click
    ConnectChaiTestNodes = _
        (InStr(1, TestStatusText(), "connection staged", vbTextCompare) > 0) Or _
        (InStr(1, TestStatusText(), "connected", vbTextCompare) > 0)
End Function

Private Function CompleteChaiTestProcess(ByVal processName As String, _
                                         ByVal externalQty As Double) As Boolean
    SelectComboText mCmbRunProcess, processName
    mCmbRunProcess_Change
    If externalQty > 0# Then
        mLstRunPalette.ListIndex = FindReusablePaletteItemName("Production Raw Material")
        If mLstRunPalette.ListIndex < 0 Then Exit Function
        mTxtPaletteSplit.Text = "100"
        mTxtPaletteQty.Text = FormatRunNumber(externalQty)
        mBtnRunApplyPalette_Click
    End If
    mBtnManagerCheckIn_Click
    If Not modProductionReusableRun.ReusableRunIsCheckedIn() Then Exit Function
    If Not StageReusableTestActualOutputs(1#) Then Exit Function
    mBtnManagerApplyOutput_Click
    CompleteChaiTestProcess = modProductionReusableRun.ReusableRunIsProcessComplete(processName)
End Function

Public Function TestEaWholeUnitActionContract() As String
    On Error GoTo Failed

    Dim beforeCount As Long
    Dim requirementRejected As Boolean

    If Not mBuilt Then BuildLayout
    beforeCount = mLstProcessRequirements.ListCount
    mTxtRequirementId.Text = "EA-FRACTION"
    mTxtRequirementName.Text = "Fractional count test"
    mTxtRequirementQty.Text = "1.5"
    mTxtRequirementPercent.Text = ""
    mTxtRequirementYieldBasis.Text = ""
    mTxtRequirementUom.Text = "ea"
    mBtnProcessRequirementAdd_Click
    requirementRejected = (mLstProcessRequirements.ListCount = beforeCount And _
        InStr(1, TestStatusText(), "requires a whole quantity", vbTextCompare) > 0)

    TestEaWholeUnitActionContract = IIf(requirementRejected, "OK", "FAIL") & _
        "|RequirementHandlerRejected=" & CStr(requirementRejected)
    Exit Function

Failed:
    TestEaWholeUnitActionContract = "FAIL|Error=" & CStr(Err.Number) & _
        " " & Err.Description
End Function

Public Function TestProductionOutputRegulationActionContract() As String
    Dim originalScope As String
    Dim defaultSaved As Boolean
    Dim wholeEaRejected As Boolean
    Dim settingsHelpUsable As Boolean
    Dim testStage As String

    On Error GoTo Failed

    If Not mBuilt Then BuildLayout
    settingsHelpUsable = ControlExistsByName(mPages.Pages(5), "txtOutputRegulationHelp") _
        And InStr(1, mTxtOutputRegulationHelp.Text, "Production Run only reads the released Recipe version", vbTextCompare) > 0
    originalScope = ComboText(mCmbOutputRegulationScope)
    testStage = "ClearProcessDraft"
    ClearProcessDraft False
    mTxtProcessId.Text = "REG-TEST"
    mTxtProcessVersion.Text = "1"
    mTxtProcessName.Text = "Output Regulation Test"
    mTxtProcessOutputId.Text = "001"
    mTxtProcessOutputName.Text = "Regulated Output"
    mTxtProcessOutputItemCode.Text = "SKU-REG-TEST"
    mTxtProcessOutputQty.Text = "610"
    RefreshProcessOutputUomCatalog "LB"
    testStage = "AddOutput"
    mBtnProcessOutputAdd_Click
    mCmbOutputRegulationScope.ListIndex = 0
    testStage = "RefreshDefaultSettings"
    RefreshOutputRegulationSettings
    If mLstOutputRegulations.ListCount > 0 Then
        mLstOutputRegulations.ListIndex = 0
        mLstOutputRegulations_Click
        mChkOutputRegulated.Value = True
        mTxtOutputRegulationFloor.Text = "600"
        mTxtOutputRegulationCeiling.Text = "610"
        testStage = "ApplyLbRegulation"
        mBtnOutputRegulationApply_Click
        defaultSaved = (StrComp(ProcessOutputRegulationValue("001", 0), "True", vbTextCompare) = 0 _
            And Abs(CDbl(ProcessOutputRegulationValue("001", 1)) - 600#) < 0.0000001 _
            And Abs(CDbl(ProcessOutputRegulationValue("001", 2)) - 610#) < 0.0000001)
        mLstProcessOutputs.List(0, 8) = "EA"
        testStage = "RefreshEaSettings"
        RefreshOutputRegulationSettings
        mLstOutputRegulations.ListIndex = 0
        mLstOutputRegulations_Click
        mTxtOutputRegulationFloor.Text = "1.5"
        mTxtOutputRegulationCeiling.Text = "2"
        testStage = "ApplyFractionalEa"
        mBtnOutputRegulationApply_Click
        wholeEaRejected = (InStr(1, TestStatusText(), "whole", vbTextCompare) > 0)
    End If
    If originalScope <> "" Then
        mCmbOutputRegulationScope.Value = originalScope
        RefreshOutputRegulationSettings
    End If
    TestProductionOutputRegulationActionContract = IIf(defaultSaved And wholeEaRejected And settingsHelpUsable, "OK", "FAIL") & _
        "|SettingsPage=True|VersionedDefault=" & CStr(defaultSaved) & _
        "|FractionalEaRejected=" & CStr(wholeEaRejected) & _
        "|WorkflowInstructions=" & CStr(settingsHelpUsable) & _
        "|CompletionHandler=ValidateReusableActualOutput"
    Exit Function
Failed:
    TestProductionOutputRegulationActionContract = "FAIL|Stage=" & testStage & _
        "|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function TestProductionExternalStockUomConversionActionContract() As String
    Dim factor As Double
    Dim catalogVersion As String
    Dim report As String
    Dim buttonsPresent As Boolean
    Dim lbToOz As Boolean
    Dim eaRejected As Boolean
    Dim actionWorkbook As Workbook
    Dim priorWorkbook As Workbook
    Dim catalogTablePresent As Boolean
    Dim priorAlerts As Boolean
    Dim testStage As String

    On Error GoTo Failed
    priorAlerts = Application.DisplayAlerts
    testStage = "BuildLayout"
    If Not mBuilt Then BuildLayout
    testStage = "Controls"
    buttonsPresent = Not mBtnUomCatalogSend Is Nothing And Not mBtnUomCatalogRetrieve Is Nothing
    testStage = "Conversions"
    lbToOz = modUomSettings.GetUomConversion("LB", "OZ", factor, catalogVersion, report) _
        And Abs(factor - 16#) < 0.0000001 And catalogVersion <> ""
    eaRejected = Not modUomSettings.GetUomConversion("EA", "OZ", factor, catalogVersion, report)
    testStage = "CreateWorkbook"
    Set priorWorkbook = mOperatorWorkbook
    Set actionWorkbook = Application.Workbooks.Add
    Set mOperatorWorkbook = actionWorkbook
    testStage = "ExportHandler"
    mBtnUomCatalogSend_Click
    testStage = "ExportTable"
    On Error Resume Next
    catalogTablePresent = Not actionWorkbook.Worksheets("invSys UOM Catalog"). _
        ListObjects("tblInvSysUomCatalog") Is Nothing
    On Error GoTo Failed
    testStage = "CloseWorkbook"
    Application.DisplayAlerts = False
    actionWorkbook.Close False
    Application.DisplayAlerts = priorAlerts
    Set mOperatorWorkbook = priorWorkbook
    TestProductionExternalStockUomConversionActionContract = _
        IIf(buttonsPresent And lbToOz And eaRejected And catalogTablePresent, "OK", "FAIL") & _
        "|SettingsCatalogButtons=" & CStr(buttonsPresent) & _
        "|LbToOz=" & CStr(lbToOz) & _
        "|EaRejected=" & CStr(eaRejected) & _
        "|SheetAction=" & CStr(catalogTablePresent)
    Exit Function
Failed:
    On Error Resume Next
    Application.DisplayAlerts = priorAlerts
    If Not actionWorkbook Is Nothing Then actionWorkbook.Close False
    Set mOperatorWorkbook = priorWorkbook
    On Error GoTo 0
    TestProductionExternalStockUomConversionActionContract = "FAIL|Stage=" & testStage & _
        "|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function TestProductionVariableQuantityModeActionContract() As String
    Dim payloadJson As String
    Dim requirementsActual As Boolean
    Dim outputsActual As Boolean
    Dim testStage As String

    On Error GoTo Failed
    If Not mBuilt Then BuildLayout
    testStage = "ClearDraft"
    ClearProcessDraft False
    mTxtProcessId.Text = "VAR"
    mTxtProcessVersion.Text = "1"
    mTxtProcessName.Text = "Variable Quantity Test"
    testStage = "AddActualRequirement"
    mTxtRequirementId.Text = "001"
    mTxtRequirementName.Text = "Measured Kitchen Input"
    mTxtRequirementUom.Text = "LB"
    SelectComboText mCmbRequirementQtyMode, "Variable -- determined at Check In"
    mBtnProcessRequirementAdd_Click
    requirementsActual = (mLstProcessRequirements.ListCount = 1 And _
        StrComp(NzStr(mLstProcessRequirements.List(0, 6)), "ACTUAL", vbTextCompare) = 0)
    testStage = "AddActualOutput"
    mTxtProcessOutputId.Text = "002"
    mTxtProcessOutputName.Text = "Measured Kitchen Output"
    mTxtProcessOutputItemCode.Text = "SKU-VAR-TEST"
    RefreshProcessOutputUomCatalog "LB"
    SelectComboText mCmbProcessOutputQtyMode, "Variable -- determined by Actual Output"
    mBtnProcessOutputAdd_Click
    outputsActual = (mLstProcessOutputs.ListCount = 1 And _
        StrComp(NzStr(mLstProcessOutputs.List(0, 9)), "ACTUAL", vbTextCompare) = 0)
    testStage = "BuildPayload"
    payloadJson = BuildProcessPayload()
    TestProductionVariableQuantityModeActionContract = _
        IIf(requirementsActual And outputsActual _
            And InStr(1, payloadJson, """RequirementQtyMode"":""ACTUAL""", vbTextCompare) > 0 _
            And InStr(1, payloadJson, """OutputQtyMode"":""ACTUAL""", vbTextCompare) > 0, "OK", "FAIL") & _
        "|RequirementActual=" & CStr(requirementsActual) & _
        "|OutputActual=" & CStr(outputsActual) & _
        "|WorksheetQtyMode=FIXED|ACTUAL"
    Exit Function
Failed:
    TestProductionVariableQuantityModeActionContract = "FAIL|Stage=" & testStage & _
        "|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function TestReusableProductionRestartActionContract( _
    ByVal recipeId As String, ByVal recipeVersion As String, _
    ByVal expectedWorkbookFullName As String) As String
    On Error GoTo Failed

    Dim rowIndex As Long
    Dim boundWorkbookFullName As String
    Dim sameWorkbook As Boolean
    Dim worksheetRediscovered As Boolean
    Dim worksheetRetrieved As Boolean
    Dim multipleTablesRediscovered As Boolean
    Dim selectedOnly As Boolean
    Dim allRetrieved As Boolean
    Dim tableName As String
    Dim worksheetReport As String

    If Not mBuilt Then BuildLayout
    If Not mOperatorWorkbook Is Nothing Then
        boundWorkbookFullName = mOperatorWorkbook.FullName
    End If
    sameWorkbook = (boundWorkbookFullName <> "" And _
        StrComp(boundWorkbookFullName, expectedWorkbookFullName, vbTextCompare) = 0)

    RefreshReusableDesignLists
    rowIndex = FindIdentityListRow(mLstLoaderRecipes, recipeId, recipeVersion)
    If rowIndex < 0 Then
        TestReusableProductionRestartActionContract = _
            "FAIL|RecipeFound=False|RecipeId=" & recipeId & _
            "|Version=" & recipeVersion & "|SameWorkbook=" & CStr(sameWorkbook)
        Exit Function
    End If

    mLstLoaderRecipes.ListIndex = rowIndex
    mBtnLoaderLoad_Click
    If Not modProductionReusableRun.ReusableRunIsLoaded() Then
        TestReusableProductionRestartActionContract = _
            "FAIL|RecipeFound=True|Loaded=False|RecipeId=" & recipeId & _
            "|Version=" & recipeVersion & "|SameWorkbook=" & CStr(sameWorkbook) & _
            "|Status=" & Replace$(Replace$(TestStatusText(), vbCr, " "), vbLf, " ")
        Exit Function
    End If

    multipleTablesRediscovered = _
        (modProductionProcessWorksheet.CountProcessWorksheetTables(mOperatorWorkbook) = 2)
    worksheetRediscovered = modProductionProcessWorksheet.FindOutstandingProcessWorksheetTable( _
        mOperatorWorkbook, tableName, worksheetReport)
    If worksheetRediscovered Then
        Call modProductionProcessWorksheet.SelectProcessWorksheetTableForTest( _
            mOperatorWorkbook, tableName)
        mBtnProcessWorksheetRetrieve_Click
        selectedOnly = _
            (modProductionProcessWorksheet.CountProcessWorksheetTables(mOperatorWorkbook) = 1)
        worksheetRetrieved = selectedOnly
    End If
    If selectedOnly And modProductionProcessWorksheet.FindOutstandingProcessWorksheetTable( _
            mOperatorWorkbook, tableName, worksheetReport) Then
        Call modProductionProcessWorksheet.SelectProcessWorksheetTableForTest( _
            mOperatorWorkbook, tableName)
        mBtnProcessWorksheetRetrieve_Click
        allRetrieved = _
            (modProductionProcessWorksheet.CountProcessWorksheetTables(mOperatorWorkbook) = 0)
    End If
    If Not multipleTablesRediscovered Or Not worksheetRediscovered Or _
       Not worksheetRetrieved Or Not selectedOnly Or Not allRetrieved Then
        TestReusableProductionRestartActionContract = _
            "FAIL|RecipeFound=True|Loaded=True|RecipeId=" & recipeId & _
            "|Version=" & recipeVersion & "|SameWorkbook=" & CStr(sameWorkbook) & _
            "|WorksheetRediscovered=" & CStr(worksheetRediscovered) & _
            "|WorksheetRetrieved=" & CStr(worksheetRetrieved) & _
            "|MultipleTablesRediscovered=" & CStr(multipleTablesRediscovered) & _
            "|SelectedOnly=" & CStr(selectedOnly) & _
            "|AllRetrieved=" & CStr(allRetrieved) & _
            "|Status=" & Replace$(Replace$(TestStatusText(), vbCr, " "), vbLf, " ")
        Exit Function
    End If

    TestReusableProductionRestartActionContract = _
        "OK|RecipeFound=True|Loaded=True|RecipeId=" & recipeId & _
        "|Version=" & recipeVersion & "|SameWorkbook=" & CStr(sameWorkbook) & _
        "|WorksheetRediscovered=" & CStr(worksheetRediscovered) & _
        "|WorksheetRetrieved=" & CStr(worksheetRetrieved) & _
        "|MultipleTablesRediscovered=" & CStr(multipleTablesRediscovered) & _
        "|SelectedOnly=" & CStr(selectedOnly) & _
        "|AllRetrieved=" & CStr(allRetrieved) & _
        "|BoundWorkbook=" & boundWorkbookFullName
    Exit Function
Failed:
    TestReusableProductionRestartActionContract = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function TestProductionBatchNoteHandlerContract() As String
    Dim report As String
    Dim accepted As Boolean
    Dim frozen As Boolean

    If Not mBuilt Then BuildLayout
    mTxtRunBatchNote.Text = "Slice 4bc batch note"
    accepted = modProductionReusableRun.SetReusableRunBatchNote(mTxtRunBatchNote.Text, report)
    frozen = modProductionReusableRun.ReusableRunBatchNoteFrozen()
    TestProductionBatchNoteHandlerContract = "OK|Control=True|Accepted=" & CStr(accepted) & _
        "|BatchNote=" & modProductionReusableRun.ReusableRunBatchNote() & _
        "|Frozen=" & CStr(frozen) & "|EventNoteContract=" & _
        CStr(InStr(1, RunBatchNoteTestEventContract(), "BatchNote=", vbTextCompare) > 0)
End Function

Private Function RunBatchNoteTestEventContract() As String
    RunBatchNoteTestEventContract = "PRODUCTION_REUSABLE_RUN|BatchNote=" & _
        modProductionReusableRun.ReusableRunBatchNote()
End Function

Private Function CreateReusableRunRawInventory(ByRef runLocation As String, _
                                               ByRef systemKeyOut As String, _
                                               ByRef startingQtyOut As Double, _
                                               ByRef report As String) As Boolean
    CreateReusableRunRawInventory = ResolveReusableRunFixtureEntity( _
        "SKU-RUN-RAW", runLocation, systemKeyOut, startingQtyOut, report)
End Function

Private Function ResolveReusableRunFixtureEntity(ByVal itemCode As String, _
                                                 ByRef runLocation As String, _
                                                 ByRef systemKeyOut As String, _
                                                 ByRef startingQtyOut As Double, _
                                                 ByRef report As String) As Boolean
    Dim entities As Variant
    Dim r As Long
    Dim entityLocation As String
    Dim inventoryWb As Workbook
    Dim rebuildReport As String

    Set inventoryWb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge("")
    If Not inventoryWb Is Nothing Then
        If Not modInventoryDomainBridge.EnsureInventorySchemaBridge(inventoryWb, rebuildReport) Then
            report = rebuildReport
            Exit Function
        End If
        If Not modInventoryDomainBridge.RebuildInventoryProjectionsBridge(inventoryWb, rebuildReport) Then
            report = rebuildReport
            Exit Function
        End If
    End If

    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(itemCode)
    If IsArray(entities) Then
        For r = LBound(entities, 1) To UBound(entities, 1)
            If StrComp(Trim$(NzStr(entities(r, 2))), itemCode, vbTextCompare) = 0 _
               Or StrComp(Trim$(NzStr(entities(r, 3))), itemCode, vbTextCompare) = 0 Then
                systemKeyOut = Trim$(NzStr(entities(r, 1)))
                If IsNumeric(entities(r, 6)) Then startingQtyOut = CDbl(entities(r, 6))
                entityLocation = Trim$(NzStr(entities(r, 7)))
                If entityLocation <> "" And StrComp(entityLocation, runLocation, vbTextCompare) <> 0 Then
                    SelectComboText mCmbRunLocation, entityLocation
                    runLocation = entityLocation
                End If
                report = "Raw exact-key inventory fixture resolved."
                ResolveReusableRunFixtureEntity = (systemKeyOut <> "" And startingQtyOut > 0#)
                Exit Function
            End If
        Next r
    End If
    report = "Exact-key inventory fixture " & itemCode & " is not available."
End Function

Private Function ConsumeReusableRunFixtureEntity(ByVal systemKey As String, _
                                                 ByVal itemCode As String, _
                                                 ByVal qty As Double, _
                                                 ByVal runLocation As String, _
                                                 ByRef report As String) As Boolean
    Dim items As New Collection
    Dim item As Object
    Dim eventId As String
    Dim queueError As String
    Dim processorReport As String
    Dim appliedCount As Long

    Set item = modProductionJson.CreateProductionDeltaPayloadItem( _
        systemKey, itemCode, qty, runLocation, "Reusable run stale-allocation fixture", "USED")
    items.Add item
    If Not modRoleEventWriter.QueuePayloadEventCurrent("PROD_CONSUME", "", _
            modProductionJson.BuildJsonArray(items), "Reusable run stale-allocation fixture", _
            eventId, queueError) Then
        report = queueError
        Exit Function
    End If
    appliedCount = modProcessor.RunBatch(modConfig.GetWarehouseId(), 0, processorReport)
    If appliedCount < 1 Then
        report = processorReport
        Exit Function
    End If
    report = "Stale allocation fixture consumed exact entity."
    ConsumeReusableRunFixtureEntity = True
End Function

Private Function FindReusablePaletteSystemKey(ByVal systemKey As String) As Long
    Dim i As Long
    FindReusablePaletteSystemKey = -1
    For i = 0 To mLstRunPalette.ListCount - 1
        If StrComp(NzStr(mLstRunPalette.List(i, 3)), systemKey, vbTextCompare) = 0 Then
            FindReusablePaletteSystemKey = i
            Exit Function
        End If
    Next i
End Function

Private Function FindReusablePaletteItemName(ByVal itemName As String) As Long
    Dim i As Long

    FindReusablePaletteItemName = -1
    For i = 0 To mLstRunPalette.ListCount - 1
        If StrComp(NzStr(mLstRunPalette.List(i, 4)), itemName, vbTextCompare) = 0 Then
            FindReusablePaletteItemName = i
            Exit Function
        End If
    Next i
End Function

Private Function ReusablePaletteRowsForItemName(ByVal itemName As String) As Long
    Dim i As Long

    For i = 0 To mLstRunPalette.ListCount - 1
        If StrComp(NzStr(mLstRunPalette.List(i, 4)), itemName, vbTextCompare) = 0 Then _
            ReusablePaletteRowsForItemName = ReusablePaletteRowsForItemName + 1
    Next i
End Function

Private Function LoaderHasReusableStatus(ByVal expectedStatus As String) As Boolean
    Dim i As Long

    For i = 0 To mLstLoaderLines.ListCount - 1
        If StrComp(Trim$(NzStr(mLstLoaderLines.List(i, 7))), _
                   expectedStatus, vbTextCompare) = 0 Then
            LoaderHasReusableStatus = True
            Exit Function
        End If
    Next i
End Function

Private Function ManagerCheckHasRoutedInputForTest(ByVal systemKey As String, _
                                                    ByVal sourceProcess As String, _
                                                    ByVal sourceOutput As String, _
                                                    ByVal requirementName As String) As Boolean
    Dim i As Long

    For i = 0 To mLstManagerCheck.ListCount - 1
        If StrComp(NzStr(mLstManagerCheck.List(i, 0)), "ROUTED", vbTextCompare) = 0 _
           And StrComp(NzStr(mLstManagerCheck.List(i, 3)), systemKey, vbTextCompare) = 0 _
           And InStr(1, NzStr(mLstManagerCheck.List(i, 1)), requirementName, vbTextCompare) > 0 _
           And InStr(1, NzStr(mLstManagerCheck.List(i, 2)), sourceProcess, vbTextCompare) > 0 _
           And InStr(1, NzStr(mLstManagerCheck.List(i, 2)), sourceOutput, vbTextCompare) > 0 _
           And IsNumeric(mLstManagerCheck.List(i, 7)) _
           And IsNumeric(mLstManagerCheck.List(i, 8)) Then
            ManagerCheckHasRoutedInputForTest = True
            Exit Function
        End If
    Next i
End Function

Private Function ManagerCheckRowsSummaryForTest() As String
    Dim i As Long

    For i = 0 To mLstManagerCheck.ListCount - 1
        If ManagerCheckRowsSummaryForTest <> "" Then _
            ManagerCheckRowsSummaryForTest = ManagerCheckRowsSummaryForTest & ";"
        ManagerCheckRowsSummaryForTest = ManagerCheckRowsSummaryForTest & _
            NzStr(mLstManagerCheck.List(i, 0)) & "/" & _
            NzStr(mLstManagerCheck.List(i, 1)) & "/" & _
            NzStr(mLstManagerCheck.List(i, 2)) & "/" & _
            NzStr(mLstManagerCheck.List(i, 3)) & "/" & _
            NzStr(mLstManagerCheck.List(i, 7)) & "/" & _
            NzStr(mLstManagerCheck.List(i, 8))
    Next i
End Function

Private Function ReusableRunFixtureAvailableQty(ByVal itemCode As String, _
                                                ByVal locationValue As String) As Double
    Dim entities As Variant
    Dim r As Long

    entities = modInventoryDomainBridge.ListAvailableInventoryEntitiesBridge(itemCode)
    If Not IsArray(entities) Then Exit Function
    For r = LBound(entities, 1) To UBound(entities, 1)
        If (StrComp(Trim$(NzStr(entities(r, 2))), itemCode, vbTextCompare) = 0 _
            Or StrComp(Trim$(NzStr(entities(r, 3))), itemCode, vbTextCompare) = 0) _
           And StrComp(Trim$(NzStr(entities(r, 7))), locationValue, vbTextCompare) = 0 Then
            If IsNumeric(entities(r, 6)) Then _
                ReusableRunFixtureAvailableQty = ReusableRunFixtureAvailableQty + CDbl(entities(r, 6))
        End If
    Next r
End Function

Private Function CaptureReusableOutputKeys() As Object
    Dim result As Object
    Dim i As Long
    Dim systemKey As String
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    For i = 0 To mLstManagerOutput.ListCount - 1
        If CLng(Val(NzStr(mLstManagerOutput.List(i, 4)))) = _
                modProductionReusableRun.ReusableRunBatchNumber() Then
            systemKey = Trim$(NzStr(mLstManagerOutput.List(i, 8)))
            If systemKey <> "" Then result(systemKey) = True
        End If
    Next i
    Set CaptureReusableOutputKeys = result
End Function

Private Function ReusableOutputSystemKeyForTest(ByVal processName As String, _
                                                ByVal outputName As String) As String
    Dim i As Long

    For i = 0 To mLstManagerOutput.ListCount - 1
        If StrComp(NzStr(mLstManagerOutput.List(i, 0)), processName, vbTextCompare) = 0 _
           And StrComp(NzStr(mLstManagerOutput.List(i, 1)), outputName, vbTextCompare) = 0 Then
            ReusableOutputSystemKeyForTest = Trim$(NzStr(mLstManagerOutput.List(i, 8)))
            Exit Function
        End If
    Next i
End Function

Private Function StageReusableTestActualOutputs(ByVal finishedActualQty As Double) As Boolean
    Dim i As Long
    Dim stagedQty As Double
    Dim outputIndex As Long

    For i = 0 To mLstManagerOutput.ListCount - 1
        outputIndex = modProductionReusableRun.ReusableRunOutputDefinitionIndex(i + 1)
        If outputIndex = 0 Then GoTo NextOutput
        mLstManagerOutput.ListIndex = i
        mLstManagerOutput_Click
        stagedQty = CDbl(modProductionReusableRun.ReusableRunPlannedOutput(outputIndex))
        If StrComp(NzStr(mLstManagerOutput.List(i, 1)), _
                   "Finished Output", vbTextCompare) = 0 Then stagedQty = finishedActualQty
        mTxtOutputReal.Text = FormatRunNumber(stagedQty)
        mTxtOutputReal_Change
        If Abs(CDbl(modProductionReusableRun.ReusableRunActualOutput(outputIndex)) - stagedQty) > _
           0.0000001 Then Exit Function
NextOutput:
    Next i
    StageReusableTestActualOutputs = True
End Function

Private Function ReusableOutputRowsForBatch(ByVal batchNumber As Long) As Long
    Dim i As Long

    For i = 0 To mLstManagerOutput.ListCount - 1
        If CLng(Val(NzStr(mLstManagerOutput.List(i, 4)))) = batchNumber _
           And Trim$(NzStr(mLstManagerOutput.List(i, 8))) <> "" Then _
            ReusableOutputRowsForBatch = ReusableOutputRowsForBatch + 1
    Next i
End Function

Private Function ReusableActiveOutputRowCount() As Long
    Dim i As Long

    For i = 0 To mLstManagerOutput.ListCount - 1
        If modProductionReusableRun.ReusableRunOutputDefinitionIndex(i + 1) > 0 Then _
            ReusableActiveOutputRowCount = ReusableActiveOutputRowCount + 1
    Next i
End Function

Private Function ReusableOutputListValueForBatch(ByVal outputName As String, _
                                                 ByVal batchNumber As Long, _
                                                 ByVal columnIndex As Long) As Double
    Dim i As Long

    For i = 0 To mLstManagerOutput.ListCount - 1
        If StrComp(NzStr(mLstManagerOutput.List(i, 1)), outputName, vbTextCompare) = 0 _
           And CLng(Val(NzStr(mLstManagerOutput.List(i, 4)))) = batchNumber Then
            If IsNumeric(mLstManagerOutput.List(i, columnIndex)) Then _
                ReusableOutputListValueForBatch = CDbl(mLstManagerOutput.List(i, columnIndex))
            Exit Function
        End If
    Next i
End Function

Private Function ProductionRunSystemKeyHeadersReadable() As Boolean
    On Error GoTo NotReadable

    Dim checkHeader As MSForms.Label
    Dim outputHeader As MSForms.Label

    Set checkHeader = mPages.Pages(3).Controls("hdrManagerCheck4")
    Set outputHeader = mPages.Pages(3).Controls("hdrManagerOutput9")
    ProductionRunSystemKeyHeadersReadable = _
        AssignmentSystemKeyReadable() And _
        (checkHeader.Caption = "System_Key" And checkHeader.Width >= 120) And _
        (outputHeader.Caption = "System_Key" And outputHeader.Width >= 120)
    Exit Function

NotReadable:
    ProductionRunSystemKeyHeadersReadable = False
End Function

Private Function AssignmentSystemKeyReadable() As Boolean
    On Error GoTo NotReadable

    Dim assignmentHeader As MSForms.Label

    Set assignmentHeader = mPages.Pages(2).Controls("hdrAssignmentInventory1")
    AssignmentSystemKeyReadable = _
        (assignmentHeader.Caption = "System_Key" And assignmentHeader.Width >= 120)
    Exit Function

NotReadable:
    AssignmentSystemKeyReadable = False
End Function

Private Function FindIdentityListRow(ByVal listControl As MSForms.ListBox, _
                                     ByVal definitionId As String, _
                                     ByVal definitionVersion As String) As Long
    Dim i As Long
    FindIdentityListRow = -1
    If listControl Is Nothing Then Exit Function
    For i = 0 To listControl.ListCount - 1
        If StrComp(NzStr(listControl.List(i, 0)), definitionId, vbTextCompare) = 0 _
           And StrComp(NzStr(listControl.List(i, 1)), definitionVersion, vbTextCompare) = 0 Then
            FindIdentityListRow = i
            Exit Function
        End If
    Next i
End Function

Private Function ProductionPageControlExists(ByVal controlName As String) As Boolean
    Dim pageIndex As Long

    If mPages Is Nothing Then Exit Function
    For pageIndex = 0 To mPages.Pages.Count - 1
        If ControlExistsByName(mPages.Pages(pageIndex), controlName) Then
            ProductionPageControlExists = True
            Exit Function
        End If
    Next pageIndex
End Function

Private Sub ShowStatus(ByVal messageText As String)
    If mTxtStatus Is Nothing Then Exit Sub
    mTxtStatus.Text = messageText
End Sub

Private Sub ShowPersistencePending(ByVal messageText As String)
    ShowStatus messageText
    Me.Repaint
    DoEvents
End Sub

Private Function NzStr(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    If IsNull(value) Then Exit Function
    If IsEmpty(value) Then Exit Function
    NzStr = CStr(value)
End Function

Private Sub mLstProcesses_Click()
    If mLoading Then Exit Sub
    If mLstProcesses.ListIndex < 0 Then Exit Sub
    ShowStatus "Selected Process " & NzStr(mLstProcesses.List(mLstProcesses.ListIndex, 0)) & _
               " version " & NzStr(mLstProcesses.List(mLstProcesses.ListIndex, 1)) & "."
End Sub

Private Sub mBtnProcessRefresh_Click()
    RefreshReusableDesignLists
    ShowStatus "Process Designer refreshed."
End Sub

Private Sub mBtnProcessNew_Click()
    ClearProcessDraft True
    ShowStatus "New Process draft started."
End Sub

Private Sub mBtnProcessLoad_Click()
    LoadSelectedProcessDefinition False
End Sub

Private Sub mBtnProcessReuse_Click()
    LoadSelectedProcessDefinition True
End Sub

Private Sub mBtnProcessValidate_Click()
    Dim report As String
    Call ValidateProcessDraft(report)
    ShowStatus report
End Sub

Private Sub mBtnProcessSave_Click()
    Dim report As String
    If Not ValidateProcessDraft(report) Then
        ShowStatus report
        Exit Sub
    End If
    SubmitProcessAction "PROCESS_SAVE", BuildProcessPayload()
End Sub

Private Sub mBtnProcessRelease_Click()
    If Not mReusableActionTestInProgress Then
        If MsgBox("Release this immutable Process version?", vbQuestion Or vbYesNo Or vbDefaultButton2, _
                  "Process Designer") <> vbYes Then Exit Sub
    End If
    SubmitProcessAction "PROCESS_RELEASE"
End Sub

Private Sub mBtnProcessObsolete_Click()
    If Not mReusableActionTestInProgress Then
        If MsgBox("Obsolete this Process version?", vbExclamation Or vbYesNo Or vbDefaultButton2, _
                  "Process Designer") <> vbYes Then Exit Sub
    End If
    SubmitProcessAction "PROCESS_OBSOLETE"
End Sub

Private Sub mBtnProcessClear_Click()
    ClearProcessDraft True
    ShowStatus "Process Designer cleared."
End Sub

Private Sub mBtnProcessWorksheetCreate_Click()
    Dim report As String
    Dim existingRow As Long
    Dim tableName As String

    On Error GoTo Failed
    If mOperatorWorkbook Is Nothing Then
        ShowStatus "The captured Production operator workbook is unavailable."
        Exit Sub
    End If

    If Trim$(mTxtProcessId.Text) = "" Then _
        mTxtProcessId.Text = NextProcessDraftBase36Id()
    If Trim$(mTxtProcessVersion.Text) = "" Then mTxtProcessVersion.Text = "1"
    existingRow = FindIdentityListRow(mLstProcesses, _
        Trim$(mTxtProcessId.Text), Trim$(mTxtProcessVersion.Text))
    If existingRow >= 0 Then
        SetEditableProcessDraftVersion _
            modProductionReusableDesigns.NextReusableDefinitionVersion( _
                Trim$(mTxtProcessId.Text), True)
    End If
    If modProductionProcessWorksheet.SendProcessDraftToWorksheet( _
        mOperatorWorkbook, Trim$(mTxtProcessId.Text), Trim$(mTxtProcessVersion.Text), _
        Trim$(mTxtProcessName.Text), Trim$(mTxtProcessDescription.Text), _
        BuildProcessPayload(), tableName, report) Then
        ClearProcessDraft True
    End If
    ShowStatus report
    Exit Sub
Failed:
    ShowStatus "Process worksheet creation failed: " & Err.Description
End Sub

Private Sub mBtnProcessWorksheetAddAlternative_Click()
    Dim report As String

    If mOperatorWorkbook Is Nothing Then
        ShowStatus "The captured Production operator workbook is unavailable."
        Exit Sub
    End If
    Call modProductionProcessWorksheet.AddAcceptableItemPairToSelectedTable( _
        mOperatorWorkbook, report)
    ShowStatus report
End Sub

Private Sub mBtnProcessWorksheetRetrieve_Click()
    Dim report As String
    Dim deleteReport As String
    Dim processId As String
    Dim processVersion As String
    Dim processName As String
    Dim description As String
    Dim payloadJson As String
    Dim oldProcessId As String
    Dim oldProcessVersion As String
    Dim oldProcessName As String
    Dim oldDescription As String
    Dim oldPayload As String
    Dim validationReport As String
    Dim tableName As String
    Dim tableNames As Collection
    Dim imports As New Collection
    Dim importRecord As Object
    Dim tableNameValue As Variant
    Dim succeeded As Long
    Dim failed As Long
    Dim summary As String
    Dim failureDetails As String

    On Error GoTo Failed
    If mOperatorWorkbook Is Nothing Then
        ShowStatus "The captured Production operator workbook is unavailable."
        Exit Sub
    End If
    Set tableNames = modProductionProcessWorksheet.FindSelectedProcessWorksheetTables( _
        mOperatorWorkbook, report)
    If tableNames Is Nothing Or tableNames.Count = 0 Then
        ShowStatus report
        Exit Sub
    End If

    oldProcessId = Trim$(mTxtProcessId.Text)
    oldProcessVersion = Trim$(mTxtProcessVersion.Text)
    oldProcessName = Trim$(mTxtProcessName.Text)
    oldDescription = Trim$(mTxtProcessDescription.Text)
    oldPayload = BuildProcessPayload()
    ' Validate every selected table before the first Designs Domain write.
    For Each tableNameValue In tableNames
        tableName = CStr(tableNameValue)
        If Not modProductionProcessWorksheet.ReadProcessDraftFromWorksheet( _
            mOperatorWorkbook, tableName, processId, processVersion, _
            processName, description, payloadJson, report) Then GoTo ValidationFailed
        If Not LoadProcessPayloadIntoDesigner(processId, processVersion, processName, _
                                               description, payloadJson, report) Then
            report = "Process worksheet retrieval failed: " & report
            GoTo ValidationFailed
        End If
        If Not ValidateProcessDraft(validationReport) Then
            report = "Process worksheet retrieval failed: " & validationReport
            GoTo ValidationFailed
        End If
        Set importRecord = CreateObject("Scripting.Dictionary")
        importRecord.CompareMode = vbTextCompare
        importRecord("TableName") = tableName
        importRecord("ProcessId") = processId
        importRecord("ProcessVersion") = processVersion
        importRecord("ProcessName") = processName
        importRecord("Description") = description
        importRecord("Payload") = payloadJson
        imports.Add importRecord
    Next tableNameValue

    For Each importRecord In imports
        If Not LoadProcessPayloadIntoDesigner(CStr(importRecord("ProcessId")), _
                CStr(importRecord("ProcessVersion")), CStr(importRecord("ProcessName")), _
                CStr(importRecord("Description")), CStr(importRecord("Payload")), report) Then
            failed = failed + 1
            failureDetails = report
            GoTo NextImport
        End If
        If SubmitProcessAction("PROCESS_SAVE", CStr(importRecord("Payload"))) Then
            If modProductionProcessWorksheet.DeleteProcessWorksheetTable( _
                    mOperatorWorkbook, CStr(importRecord("TableName")), deleteReport) Then
                succeeded = succeeded + 1
            Else
                failed = failed + 1
                failureDetails = deleteReport
            End If
        Else
            failed = failed + 1
            failureDetails = TestStatusText()
        End If
NextImport:
    Next importRecord
    summary = "Retrieved " & CStr(succeeded) & " selected Process table(s) as DRAFT"
    If failed > 0 Then summary = summary & "; " & CStr(failed) & " table(s) remain"
    summary = summary & "."
    If failureDetails <> "" Then summary = summary & " " & failureDetails
    ShowStatus summary
    Exit Sub

ValidationFailed:
    Call LoadProcessPayloadIntoDesigner(oldProcessId, oldProcessVersion, _
        oldProcessName, oldDescription, oldPayload, validationReport)
    ShowStatus CStr(tableName) & ": " & report & _
        " No selected table was saved or removed."
    Exit Sub
Failed:
    ShowStatus "Process worksheet retrieval failed: " & Err.Description
End Sub

Private Sub mLstProcessRequirements_Click()
    Dim idx As Long
    If mLoading Then Exit Sub
    idx = mLstProcessRequirements.ListIndex
    If idx < 0 Then Exit Sub
    mSelectedProcessRequirementIndex = idx
    mTxtRequirementId.Text = NzStr(mLstProcessRequirements.List(idx, 0))
    mTxtRequirementName.Text = NzStr(mLstProcessRequirements.List(idx, 1))
    mTxtRequirementQty.Text = NzStr(mLstProcessRequirements.List(idx, 2))
    mTxtRequirementPercent.Text = NzStr(mLstProcessRequirements.List(idx, 3))
    mTxtRequirementYieldBasis.Text = NzStr(mLstProcessRequirements.List(idx, 4))
    mTxtRequirementUom.Text = NzStr(mLstProcessRequirements.List(idx, 5))
    If StrComp(NzStr(mLstProcessRequirements.List(idx, 6)), "ACTUAL", vbTextCompare) = 0 Then
        SelectComboText mCmbRequirementQtyMode, "Variable -- determined at Check In"
    Else
        SelectComboText mCmbRequirementQtyMode, "Enter a number"
    End If
    ApplyRequirementQtyMode
End Sub

Private Sub mBtnProcessRequirementAdd_Click()
    WriteRequirementEditorToList False
End Sub

Private Sub mBtnProcessRequirementUpdate_Click()
    WriteRequirementEditorToList True
End Sub

Private Sub mBtnProcessRequirementRemove_Click()
    RemoveSelectedListRow mLstProcessRequirements
    mSelectedProcessRequirementIndex = -1
    ClearRequirementEditor
End Sub

Private Sub mBtnProcessRequirementUp_Click()
    MoveSelectedListRow mLstProcessRequirements, -1
End Sub

Private Sub mBtnProcessRequirementDown_Click()
    MoveSelectedListRow mLstProcessRequirements, 1
End Sub

Private Sub mLstProcessOutputs_Click()
    Dim idx As Long
    If mLoading Then Exit Sub
    idx = mLstProcessOutputs.ListIndex
    If idx < 0 Then Exit Sub
    mSelectedProcessOutputIndex = idx
    mTxtProcessOutputId.Text = NzStr(mLstProcessOutputs.List(idx, 0))
    mTxtProcessOutputName.Text = NzStr(mLstProcessOutputs.List(idx, 1))
    mTxtProcessOutputItemCode.Text = NzStr(mLstProcessOutputs.List(idx, 2))
    mTxtProcessOutputDesignId.Text = NzStr(mLstProcessOutputs.List(idx, 3))
    mTxtProcessOutputDesignVersion.Text = NzStr(mLstProcessOutputs.List(idx, 4))
    mTxtProcessOutputQty.Text = NzStr(mLstProcessOutputs.List(idx, 5))
    mTxtProcessOutputPercent.Text = NzStr(mLstProcessOutputs.List(idx, 6))
    mTxtProcessOutputYieldBasis.Text = NzStr(mLstProcessOutputs.List(idx, 7))
    RefreshProcessOutputUomCatalog NzStr(mLstProcessOutputs.List(idx, 8))
    If StrComp(NzStr(mLstProcessOutputs.List(idx, 9)), "ACTUAL", vbTextCompare) = 0 Then
        SelectComboText mCmbProcessOutputQtyMode, "Variable -- determined by Actual Output"
    Else
        SelectComboText mCmbProcessOutputQtyMode, "Enter a number"
    End If
    ApplyOutputQtyMode
End Sub

Private Sub mCmbRequirementQtyMode_Change()
    If Not mLoading Then ApplyRequirementQtyMode
End Sub

Private Sub mCmbProcessOutputQtyMode_Change()
    If Not mLoading Then ApplyOutputQtyMode
End Sub

Private Sub mCmbOutputRegulationScope_Change()
    If Not mLoading Then RefreshOutputRegulationSettings
End Sub

Private Sub mCmbOutputRegulationNode_Change()
    If OutputRegulationRecipeScope() And Not mLoading Then RefreshOutputRegulationSettings
End Sub

Private Sub mLstOutputRegulations_Click()
    Dim idx As Long
    If mLoading Then Exit Sub
    idx = mLstOutputRegulations.ListIndex
    If idx < 0 Then Exit Sub
    mChkOutputRegulated.Value = (StrComp(NzStr(mLstOutputRegulations.List(idx, 3)), "True", vbTextCompare) = 0)
    mTxtOutputRegulationFloor.Text = NzStr(mLstOutputRegulations.List(idx, 4))
    mTxtOutputRegulationCeiling.Text = NzStr(mLstOutputRegulations.List(idx, 5))
End Sub

Private Sub mBtnOutputRegulationApply_Click()
    Dim idx As Long, sourceRow As Long, record As Object
    idx = mLstOutputRegulations.ListIndex
    If idx < 0 Then
        ShowStatus "Select an output to regulate."
        Exit Sub
    End If
    If mChkOutputRegulated.Value Then
        If Not PositiveTextValue(mTxtOutputRegulationFloor.Text) _
           Or Not PositiveTextValue(mTxtOutputRegulationCeiling.Text) _
           Or CDbl(mTxtOutputRegulationFloor.Text) > CDbl(mTxtOutputRegulationCeiling.Text) Then
            ShowStatus "Enabled output regulation requires positive floor and ceiling with floor not above ceiling."
            Exit Sub
        End If
        If Not ValidateWholeQuantityForUom(mTxtOutputRegulationFloor.Text, _
                NzStr(mLstOutputRegulations.List(idx, 2)), "Output regulation floor") Then Exit Sub
        If Not ValidateWholeQuantityForUom(mTxtOutputRegulationCeiling.Text, _
                NzStr(mLstOutputRegulations.List(idx, 2)), "Output regulation ceiling") Then Exit Sub
    End If
    If OutputRegulationRecipeScope() Then
        Set record = NewReusableRecord("OUTPUT_REGULATION")
        record("ProcessNodeId") = NzStr(mLstOutputRegulations.List(idx, 6))
        record("ProcessId") = NzStr(mLstOutputRegulations.List(idx, 7))
        record("ProcessVersion") = NzStr(mLstOutputRegulations.List(idx, 8))
        record("OutputId") = NzStr(mLstOutputRegulations.List(idx, 0))
        record("OutputRegulationEnabled") = CBool(mChkOutputRegulated.Value)
        AddNumericReusableField record, "OutputFloorQty", mTxtOutputRegulationFloor.Text
        AddNumericReusableField record, "OutputCeilingQty", mTxtOutputRegulationCeiling.Text
        RemoveRecipeOutputRegulation record("ProcessNodeId"), record("OutputId")
        If mRecipeOutputRegulations Is Nothing Then Set mRecipeOutputRegulations = New Collection
        mRecipeOutputRegulations.Add record
    Else
        sourceRow = CLng(Val(NzStr(mLstOutputRegulations.List(idx, 6))))
        SetProcessOutputRegulation NzStr(mLstProcessOutputs.List(sourceRow, 0)), _
            CBool(mChkOutputRegulated.Value), mTxtOutputRegulationFloor.Text, mTxtOutputRegulationCeiling.Text
    End If
    RefreshOutputRegulationSettings
    ShowStatus "Output regulation staged in the versioned " & IIf(OutputRegulationRecipeScope(), "Recipe override.", "Process default.")
End Sub

Private Sub mBtnOutputRegulationClear_Click()
    Dim idx As Long, sourceRow As Long
    idx = mLstOutputRegulations.ListIndex
    If idx < 0 Then Exit Sub
    If OutputRegulationRecipeScope() Then
        RemoveRecipeOutputRegulation NzStr(mLstOutputRegulations.List(idx, 6)), NzStr(mLstOutputRegulations.List(idx, 0))
    Else
        sourceRow = CLng(Val(NzStr(mLstOutputRegulations.List(idx, 6))))
        SetProcessOutputRegulation NzStr(mLstProcessOutputs.List(sourceRow, 0)), False, "", ""
    End If
    RefreshOutputRegulationSettings
    ShowStatus "Output regulation cleared."
End Sub

Private Sub mBtnUomCatalogSend_Click()
    Dim report As String
    If Not modProductionUomCatalog.SendUomCatalogToWorksheet(mOperatorWorkbook, report) Then
        ShowStatus "UOM Catalog export failed: " & report
        Exit Sub
    End If
    ShowStatus report
End Sub

Private Sub mBtnUomCatalogRetrieve_Click()
    Dim report As String
    If Not modProductionUomCatalog.RetrieveUomCatalogFromWorksheet(mOperatorWorkbook, report) Then
        ShowStatus "UOM Catalog retrieval failed: " & report
        Exit Sub
    End If
    RefreshProcessOutputUomCatalog
    RefreshRecipeUomCatalog
    RefreshConnectionUomCatalog
    ShowStatus report
End Sub

Private Sub RemoveRecipeOutputRegulation(ByVal nodeId As String, ByVal outputId As String)
    Dim i As Long, record As Object
    If mRecipeOutputRegulations Is Nothing Then Exit Sub
    For i = mRecipeOutputRegulations.Count To 1 Step -1
        Set record = mRecipeOutputRegulations(i)
        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(record, "OutputId"), outputId, vbTextCompare) = 0 Then
            mRecipeOutputRegulations.Remove i
        End If
    Next i
End Sub

Private Sub mBtnProcessOutputAdd_Click()
    WriteOutputEditorToList False
End Sub

Private Sub mBtnProcessOutputUpdate_Click()
    WriteOutputEditorToList True
End Sub

Private Sub mBtnProcessOutputRemove_Click()
    If mLstProcessOutputs.ListIndex >= 0 And Not mProcessOutputRegulations Is Nothing Then
        If mProcessOutputRegulations.Exists(NzStr(mLstProcessOutputs.List(mLstProcessOutputs.ListIndex, 0))) Then _
            mProcessOutputRegulations.Remove NzStr(mLstProcessOutputs.List(mLstProcessOutputs.ListIndex, 0))
    End If
    RemoveSelectedListRow mLstProcessOutputs
    mSelectedProcessOutputIndex = -1
    ClearOutputEditor
End Sub

Private Sub mBtnProcessOutputUp_Click()
    MoveSelectedListRow mLstProcessOutputs, -1
End Sub

Private Sub mBtnProcessOutputDown_Click()
    MoveSelectedListRow mLstProcessOutputs, 1
End Sub

Private Sub mLstProcessInstructions_Click()
    If mLstProcessInstructions.ListIndex >= 0 Then _
        mTxtProcessInstruction.Text = NzStr(mLstProcessInstructions.List(mLstProcessInstructions.ListIndex, 1))
End Sub

Private Sub mBtnProcessInstructionAdd_Click()
    If Trim$(mTxtProcessInstruction.Text) = "" Then Exit Sub
    mLstProcessInstructions.AddItem CStr(mLstProcessInstructions.ListCount + 1)
    mLstProcessInstructions.List(mLstProcessInstructions.ListCount - 1, 1) = Trim$(mTxtProcessInstruction.Text)
End Sub

Private Sub mBtnProcessInstructionUpdate_Click()
    If mLstProcessInstructions.ListIndex >= 0 Then _
        mLstProcessInstructions.List(mLstProcessInstructions.ListIndex, 1) = Trim$(mTxtProcessInstruction.Text)
End Sub

Private Sub mBtnProcessInstructionRemove_Click()
    RemoveSelectedListRow mLstProcessInstructions
    RenumberInstructionOrdinals
End Sub

Private Sub mBtnProcessInstructionUp_Click()
    MoveSelectedListRow mLstProcessInstructions, -1
End Sub

Private Sub mBtnProcessInstructionDown_Click()
    MoveSelectedListRow mLstProcessInstructions, 1
End Sub

Private Sub mBtnRecipeRefresh_Click()
    RefreshReusableDesignLists
    ShowStatus "Recipe Designer refreshed."
End Sub

Private Sub mBtnRecipeNew_Click()
    ClearRecipeDraft True
    ShowStatus "New Recipe draft started."
End Sub

Private Sub mBtnRecipeLoad_Click()
    LoadSelectedRecipeDefinition
End Sub

Private Sub mBtnRecipeAddProcess_Click()
    AddSelectedReleasedProcessToRecipe
End Sub

Private Sub mBtnRecipeRemoveProcess_Click()
    RemoveSelectedRecipeNode
End Sub

Private Sub mBtnRecipeConnect_Click()
    WriteConnectionEditorToList False
End Sub

Private Sub mBtnRecipeUpdateConnection_Click()
    WriteConnectionEditorToList True
End Sub

Private Sub mBtnRecipeDisconnect_Click()
    Dim idx As Long
    Dim displayIndex As Long

    displayIndex = mLstRecipeConnectionDisplay.ListIndex
    If displayIndex >= 0 Then
        idx = CLng(Val(NzStr(mLstRecipeConnectionDisplay.List(displayIndex, 7))))
    Else
        idx = -1
    End If
    If idx < 0 Then idx = mLstRecipeConnections.ListIndex
    If idx < 0 Or idx >= mLstRecipeConnections.ListCount Then Exit Sub
    mLstRecipeConnections.RemoveItem idx
    RefreshRecipeConnectionDisplay
    ClearConnectionEditor
    RefreshConnectionNodeCombos
    ShowStatus "Recipe connection removed."
End Sub

Private Sub mBtnRecipeMoveUp_Click()
    MoveSelectedListRow mLstRecipeNodes, -1
    RenumberRecipeExecutionOrder
    RefreshConnectionNodeCombos
End Sub

Private Sub mBtnRecipeMoveDown_Click()
    MoveSelectedListRow mLstRecipeNodes, 1
    RenumberRecipeExecutionOrder
    RefreshConnectionNodeCombos
End Sub

Private Sub mBtnRecipeAutoOrder_Click()
    AutoOrderRecipeNodes
End Sub

Private Sub mBtnRecipeValidate_Click()
    Dim report As String
    EnsureRecipeDraftIdentity
    Call ValidateRecipeDraft(report, True)
    ShowStatus report
End Sub

Private Sub mBtnRecipeSave_Click()
    Dim report As String
    EnsureRecipeDraftIdentity
    If Not ValidateRecipeDraft(report, False) Then
        ShowStatus report
        Exit Sub
    End If
    SubmitRecipeAction "RECIPE_SAVE", BuildRecipePayload()
End Sub

Private Sub mBtnRecipeRelease_Click()
    Dim report As String
    EnsureRecipeDraftIdentity
    If Not ValidateRecipeDraft(report, True) Then
        ShowStatus report
        Exit Sub
    End If
    If Not mReusableActionTestInProgress Then
        If MsgBox("Release this immutable Recipe version?", vbQuestion Or vbYesNo Or vbDefaultButton2, _
                  "Recipe Designer") <> vbYes Then Exit Sub
    End If
    SubmitRecipeAction "RECIPE_RELEASE"
End Sub

Private Sub mBtnRecipeObsolete_Click()
    If Not mReusableActionTestInProgress Then
        If MsgBox("Obsolete this Recipe version?", vbExclamation Or vbYesNo Or vbDefaultButton2, _
                  "Recipe Designer") <> vbYes Then Exit Sub
    End If
    SubmitRecipeAction "RECIPE_OBSOLETE"
End Sub

Private Sub mBtnRecipeClear_Click()
    ClearRecipeDraft True
    ShowStatus "Recipe Designer cleared."
End Sub

Private Sub mCmbConnectionFromNode_Change()
    If mLoading Then Exit Sub
    RefreshConnectionOutputChoices
End Sub

Private Sub mCmbConnectionOutput_Change()
    If mLoading Then Exit Sub
    RefreshCompatibleDownstreamChoices
End Sub

Private Sub mCmbConnectionToNode_Change()
    If mLoading Then Exit Sub
    BindSelectedCompatibleRequirement
End Sub

Private Sub mLstRecipeConnections_Click()
    Dim idx As Long
    idx = mLstRecipeConnections.ListIndex
    LoadConnectionEditorFromIndex idx
End Sub

Private Sub mLstRecipeConnectionDisplay_Click()
    Dim idx As Long
    Dim displayIndex As Long
    Dim sourceNodeId As String
    Dim outputId As String

    If mLoading Then Exit Sub
    displayIndex = mLstRecipeConnectionDisplay.ListIndex
    If displayIndex < 0 Then Exit Sub
    idx = CLng(Val(NzStr(mLstRecipeConnectionDisplay.List(displayIndex, 7))))
    If idx >= 0 And idx < mLstRecipeConnections.ListCount Then
        mLstRecipeConnections.ListIndex = idx
        LoadConnectionEditorFromIndex idx
    Else
        mLstRecipeConnections.ListIndex = -1
        sourceNodeId = NzStr(mLstRecipeConnectionDisplay.List(displayIndex, 8))
        outputId = NzStr(mLstRecipeConnectionDisplay.List(displayIndex, 9))
        SelectComboText mCmbConnectionFromNode, sourceNodeId
        RefreshConnectionOutputChoices
        SelectComboText mCmbConnectionOutput, outputId
        RefreshCompatibleDownstreamChoices
        ShowStatus "This output is currently Finished inventory; choose Feeds Process to route it."
    End If
End Sub

Private Sub LoadConnectionEditorFromIndex(ByVal idx As Long)
    If idx < 0 Or idx >= mLstRecipeConnections.ListCount Then Exit Sub
    SelectComboText mCmbConnectionFromNode, NzStr(mLstRecipeConnections.List(idx, 0))
    RefreshConnectionOutputChoices
    SelectComboText mCmbConnectionOutput, NzStr(mLstRecipeConnections.List(idx, 1))
    RefreshCompatibleDownstreamChoices NzStr(mLstRecipeConnections.List(idx, 2)), _
        NzStr(mLstRecipeConnections.List(idx, 3))
    mTxtConnectionQty.Text = NzStr(mLstRecipeConnections.List(idx, 4))
    mTxtConnectionPercent.Text = NzStr(mLstRecipeConnections.List(idx, 5))
    RefreshConnectionUomCatalog NzStr(mLstRecipeConnections.List(idx, 6))
End Sub

Private Sub AutoOrderRecipeNodes()
    Dim pass As Long
    Dim i As Long
    Dim fromIndex As Long
    Dim toIndex As Long
    Dim changed As Boolean

    For pass = 1 To mLstRecipeNodes.ListCount * mLstRecipeNodes.ListCount
        changed = False
        For i = 0 To mLstRecipeConnections.ListCount - 1
            fromIndex = RecipeNodeIndex(NzStr(mLstRecipeConnections.List(i, 0)))
            toIndex = RecipeNodeIndex(NzStr(mLstRecipeConnections.List(i, 2)))
            If fromIndex >= toIndex And fromIndex >= 0 And toIndex >= 0 Then
                mLstRecipeNodes.ListIndex = fromIndex
                MoveSelectedListRow mLstRecipeNodes, -1
                changed = True
            End If
        Next i
        If Not changed Then Exit For
    Next pass
    RenumberRecipeExecutionOrder
    RefreshConnectionNodeCombos
    RefreshRecipeConnectionDisplay
    If RecipeGraphIsAcyclic() Then
        ShowStatus "Recipe execution order updated."
    Else
        ShowStatus "Auto Order cannot resolve a circular dependency."
    End If
End Sub

Private Sub mLstBuilderRecipes_Click()
    If mLoading Then Exit Sub
    If mLstBuilderRecipes.ListIndex < 0 Then Exit Sub
    mTxtRecipeId.Text = NzStr(mLstBuilderRecipes.List(mLstBuilderRecipes.ListIndex, 0))
    mTxtRecipeName.Text = NzStr(mLstBuilderRecipes.List(mLstBuilderRecipes.ListIndex, 1))
    mTxtRecipeDescription.Text = NzStr(mLstBuilderRecipes.List(mLstBuilderRecipes.ListIndex, 2))
End Sub

Private Sub mLstBuilderLines_Click()
    If mLoading Then Exit Sub
    LoadSelectedBuilderLine
End Sub

Private Sub mCmbLineIo_Change()
    ConfigureRecipeLineInputs
End Sub

Private Sub mBtnBuilderRefresh_Click()
    RefreshRecipeLists
    RefreshBuilderHeader
    RefreshBuilderLines
    ShowStatus "Recipe Builder refreshed."
End Sub

Private Sub mBtnBuilderNew_Click()
    ClearRecipeBuilderForNew
End Sub

Private Sub mBtnBuilderLoad_Click()
    LoadSelectedRecipeIntoBuilder
End Sub

Private Sub mBtnBuilderSave_Click()
    WriteRecipeHeaderFromForm
    WriteSelectedRecipeBuilderLineFromForm False
    BindOperatorWorkbookForRun
    ShowPersistencePending "Saving the recipe to warehouse storage..."
    mProduction.BtnSaveRecipe
    RefreshRecipeLists
    RefreshBuilderLines
    RefreshAssignmentState
    RefreshLoaderState
    RefreshManagerState
    ShowStatus "Save Recipe completed."
End Sub

Private Sub mBtnBuilderProcess_Click()
    WriteRecipeHeaderFromForm
    BindOperatorWorkbookForRun
    mProduction.BtnBuildRecipeProcessTables
    RefreshBuilderLines
    ShowStatus "Process table action completed."
End Sub

Private Sub mBtnBuilderFormulas_Click()
    WriteRecipeHeaderFromForm
    BindOperatorWorkbookForRun
    ShowPersistencePending "Saving recipe formulas to warehouse storage..."
    mProduction.BtnSaveFormulas
    ShowStatus "Save Formulas completed."
End Sub

Private Sub mBtnBuilderClear_Click()
    BindOperatorWorkbookForRun
    mProduction.BtnClearRecipeBuilder
    RefreshBuilderHeader
    RefreshBuilderLines
    ShowStatus "Recipe Builder cleared."
End Sub

Private Sub mBtnBuilderRelease_Click()
    Dim recipeId As String
    Dim recipeName As String
    Dim releaseReport As String

    recipeId = Trim$(mTxtRecipeId.Text)
    recipeName = Trim$(mTxtRecipeName.Text)
    If recipeId = "" Then
        ShowStatus "Select or load a saved recipe before releasing it."
        Exit Sub
    End If
    If MsgBox("Release " & IIf(recipeName = "", recipeId, recipeName) & _
              " for production runs?" & vbCrLf & vbCrLf & _
              "This publishes the latest immutable recipe version.", _
              vbQuestion Or vbYesNo Or vbDefaultButton2, _
              "Production Designs") <> vbYes Then Exit Sub

    BindOperatorWorkbookForRun
    ShowPersistencePending "Publishing the released recipe design..."
    releaseReport = NzStr(mProduction.ReleaseRecipeForProduction(recipeId))
    RefreshRecipeLists
    ShowStatus releaseReport
End Sub

Private Sub mBtnLineAdd_Click()
    AddRecipeBuilderLine
End Sub

Private Sub mBtnLineUpdate_Click()
    UpdateSelectedRecipeBuilderLine
End Sub

Private Sub mBtnLineRemove_Click()
    RemoveSelectedRecipeBuilderLine
End Sub

Private Sub mBtnLineMoveUp_Click()
    MoveSelectedRecipeBuilderLine -1
End Sub

Private Sub mBtnLineMoveDown_Click()
    MoveSelectedRecipeBuilderLine 1
End Sub

Private Sub mBtnLineUomAdd_Click()
    Dim uomName As String
    Dim report As String

    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("PROD_POST", report) Then
        ShowStatus report
        Exit Sub
    End If
    uomName = InputBox("Enter the UOM to add to this warehouse catalog:", _
                       "Recipe Builder - Add UOM", CStr(mTxtLineUom.Value))
    uomName = modUomSettings.NormalizeConfiguredUomName(uomName)
    If uomName = "" Then Exit Sub

    If modUomSettings.AddConfiguredUom(uomName, report) Then
        RefreshRecipeUomCatalog uomName
        ShowStatus report
    Else
        ShowStatus report
    End If
End Sub

Private Sub mLstAssignRecipes_Click()
    If mLoading Then Exit Sub
    SelectReusableAssignmentProcess
End Sub

Private Sub mLstAssignIngredients_Click()
    If mLoading Then Exit Sub
    SelectReusableAssignmentRequirement
End Sub

Private Sub mTxtInventorySearch_Change()
    If mLoading Then Exit Sub
    RefreshInventoryList
End Sub

Private Sub mCmbRunLocation_Change()
    Dim clearedCount As Long

    If mLoading Then Exit Sub
    SyncRunLocationCombo mCmbRunLocation, mCmbTreeRunLocation
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        RefreshReusableRunControls False
        Exit Sub
    End If
    ClearMismatchedRunLocationAllocations clearedCount
    If clearedCount > 0 Then ShowStatus "Cleared " & CStr(clearedCount) & " allocation(s) that were not at the selected production run location."
End Sub

Private Sub mCmbTreeRunLocation_Change()
    Dim clearedCount As Long

    If mLoading Then Exit Sub
    SyncRunLocationCombo mCmbTreeRunLocation, mCmbRunLocation
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        RefreshReusableRunControls False
        Exit Sub
    End If
    ClearMismatchedRunLocationAllocations clearedCount
    If clearedCount > 0 Then ShowStatus "Cleared " & CStr(clearedCount) & " allocation(s) that were not at the selected production run location."
End Sub

Private Sub mTxtPaletteSplit_Change()
    PaletteSplitTextChanged mTxtPaletteSplit, mTxtPaletteQty
End Sub

Private Sub mTxtTreePaletteSplit_Change()
    PaletteSplitTextChanged mTxtTreePaletteSplit, mTxtTreePaletteQty
End Sub

Private Sub mTxtPaletteQty_Change()
    PaletteQtyTextChanged mTxtPaletteQty, mTxtPaletteSplit
End Sub

Private Sub mTxtTreePaletteQty_Change()
    PaletteQtyTextChanged mTxtTreePaletteQty, mTxtTreePaletteSplit
End Sub

Private Sub mTxtOutputReal_Change()
    If mLoading Then Exit Sub
    If modProductionReusableRun.ReusableRunIsLoaded() Then _
        StageSelectedReusableActualOutput False
End Sub

Private Sub PaletteSplitTextChanged(ByVal splitTextBox As MSForms.TextBox, ByVal qtyTextBox As MSForms.TextBox)
    Dim baseQty As Double
    Dim splitVal As Double
    Dim lst As MSForms.ListBox
    Dim idx As Long
    Dim totalPct As Double

    If mLoading Or mUpdatingPaletteInputs Then Exit Sub
    mPaletteInputSource = "SPLIT"
    If splitTextBox Is Nothing Or qtyTextBox Is Nothing Then Exit Sub
    If Trim$(splitTextBox.Text) = "" Then Exit Sub
    If Not IsNumeric(splitTextBox.Text) Then Exit Sub
    splitVal = CDbl(splitTextBox.Text)
    If splitVal < 0 Then Exit Sub
    baseQty = SelectedRunPaletteBaseQty()
    If baseQty <= 0 Then Exit Sub
    SetPaletteTextSilently qtyTextBox, FormatRunNumber(baseQty * splitVal / 100#)
    Set lst = ActiveRunPaletteList()
    If Not lst Is Nothing Then
        idx = lst.ListIndex
        If idx >= 0 And Not IsRunTreeParentRow(lst, idx) Then
            totalPct = RunAllocationPercentForGroup(RunAllocationGroupKeyFromList(lst, idx), RunAllocationKeyFromList(lst, idx)) + splitVal
            SetRunAllocationVisualState splitTextBox, qtyTextBox, totalPct
        End If
    End If
    mPaletteInputSource = "SPLIT"
End Sub

Private Sub PaletteQtyTextChanged(ByVal qtyTextBox As MSForms.TextBox, ByVal splitTextBox As MSForms.TextBox)
    Dim baseQty As Double
    Dim qtyVal As Double
    Dim splitVal As Double
    Dim lst As MSForms.ListBox
    Dim idx As Long
    Dim totalPct As Double

    If mLoading Or mUpdatingPaletteInputs Then Exit Sub
    mPaletteInputSource = "QTY"
    If qtyTextBox Is Nothing Or splitTextBox Is Nothing Then Exit Sub
    If Trim$(qtyTextBox.Text) = "" Then Exit Sub
    If Not IsNumeric(qtyTextBox.Text) Then Exit Sub
    qtyVal = CDbl(qtyTextBox.Text)
    If qtyVal < 0 Then Exit Sub
    baseQty = SelectedRunPaletteBaseQty()
    If baseQty <= 0 Then Exit Sub
    splitVal = qtyVal / baseQty * 100#
    SetPaletteTextSilently splitTextBox, FormatRunNumber(splitVal)
    Set lst = ActiveRunPaletteList()
    If Not lst Is Nothing Then
        idx = lst.ListIndex
        If idx >= 0 And Not IsRunTreeParentRow(lst, idx) Then
            totalPct = RunAllocationPercentForGroup(RunAllocationGroupKeyFromList(lst, idx), RunAllocationKeyFromList(lst, idx)) + splitVal
            SetRunAllocationVisualState splitTextBox, qtyTextBox, totalPct
        End If
    End If
    mPaletteInputSource = "QTY"
End Sub

Private Sub mBtnAssignRefresh_Click()
    ResetInventoryCache
    RefreshReusableDesignLists
    RefreshReusableAssignmentState
    ShowStatus "Ingredients Assignment refreshed."
End Sub

Private Sub mBtnAssignRecipe_Click()
    SelectReusableAssignmentProcess
End Sub

Private Sub mBtnAssignIngredient_Click()
    SelectReusableAssignmentRequirement
End Sub

Private Sub mBtnAssignAdd_Click()
    AddReusableInventoryAlternative
End Sub

Private Sub mBtnAssignRemove_Click()
    RemoveReusableInventoryAlternative
End Sub

Private Sub mBtnAssignSave_Click()
    SaveReusableAssignments
End Sub

Private Sub mBtnAssignClear_Click()
    mLstAssignIngredients.Clear
    mLstAssignAllowed.Clear
    Set mProcessAlternatives = New Collection
    ShowStatus "Ingredients Assignment cleared."
End Sub

Private Sub mBtnLoaderRefresh_Click()
    Dim refreshReport As String
    Dim refreshed As Boolean

    If modProductionReusableRun.ReusableRunIsLoaded() Then
        ResetInventoryCache
        RefreshRecipeLists
        RefreshReusableRunControls True
        ShowStatus "Reusable Production Run inventory refreshed from the exact entity projection."
        Exit Sub
    End If
    refreshed = RefreshProductionInventoryReadModel(refreshReport)
    ResetInventoryCache
    RefreshRecipeLists
    RefreshLoaderState
    RefreshManagerState
    If refreshed Then
        ShowStatus "Production Run inventory refreshed. " & refreshReport
    Else
        ShowStatus refreshReport
        MsgBox refreshReport, vbExclamation, "Production Inventory Refresh"
    End If
End Sub

Private Sub mBtnLoaderLoad_Click()
    If mLstLoaderRecipes.ListIndex >= 0 Then
        If Not LoadReusableRecipeIntoRun( _
                NzStr(mLstLoaderRecipes.List(mLstLoaderRecipes.ListIndex, 0)), _
                NzStr(mLstLoaderRecipes.List(mLstLoaderRecipes.ListIndex, 1))) Then Exit Sub
    Else
        ShowStatus "Select a released Recipe version first."
    End If
End Sub

Private Function LoadReusableRecipeIntoRun(ByVal recipeId As String, _
                                           ByVal recipeVersion As String) As Boolean
    Dim scalePercent As Double
    Dim report As String

    If Not TryParseBatchScalePercent(mTxtBatchScalePercent.Text, scalePercent, report) Then
        ShowStatus report
        Exit Function
    End If
    If Not modProductionReusableRun.LoadReleasedReusableRecipe( _
            recipeId, recipeVersion, scalePercent, report) Then
        ShowStatus report
        Exit Function
    End If
    RefreshReusableRunControls False
    ShowStatus report
    LoadReusableRecipeIntoRun = True
End Function

Private Sub RefreshReusableRunControls(ByVal refreshInventory As Boolean)
    Dim loaderRows As Variant
    Dim paletteRows As Variant
    Dim checkRows As Variant
    Dim outputRows As Variant
    Dim i As Long
    Dim processName As String
    Dim locationName As String
    Dim activeOutputRow As Long
    Dim selectedProcess As String

    If Not modProductionReusableRun.ReusableRunIsLoaded() Then Exit Sub
    If refreshInventory Then ResetInventoryCache
    selectedProcess = ActiveRunProcess()
    mLoading = True
    loaderRows = modProductionReusableRun.ReusableRunLoaderRows(ActiveRunLocation())
    paletteRows = modProductionReusableRun.ReusableRunPaletteRows(ActiveRunLocation())
    checkRows = modProductionReusableRun.ReusableRunManagerCheckRows()
    outputRows = modProductionReusableRun.ReusableRunOutputRows()
    FillListFromArray mLstLoaderLines, loaderRows
    FillListFromArray mLstRunPalette, paletteRows
    FillListFromArray mLstManagerCheck, checkRows
    FillListFromArray mLstManagerOutput, outputRows
    activeOutputRow = FirstReusableActiveOutputRow()
    If activeOutputRow >= 0 Then mLstManagerOutput.ListIndex = activeOutputRow

    mCmbRunProcess.Clear
    mCmbTreeRunProcess.Clear
    mCmbRunProcess.AddItem ""
    mCmbTreeRunProcess.AddItem ""
    mCmbRunTargetOutput.Clear
    For i = 0 To mLstLoaderLines.ListCount - 1
        processName = NzStr(mLstLoaderLines.List(i, 0))
        AddUniqueComboItem mCmbRunProcess, processName
        AddUniqueComboItem mCmbTreeRunProcess, processName
        If StrComp(NzStr(mLstLoaderLines.List(i, 2)), "OUTPUT", vbTextCompare) = 0 Then _
            AddUniqueComboItem mCmbRunTargetOutput, NzStr(mLstLoaderLines.List(i, 3))
    Next i
    mCmbRunProcess.ListIndex = 0
    mCmbTreeRunProcess.ListIndex = 0
    If selectedProcess <> "" Then
        SelectComboText mCmbRunProcess, selectedProcess
        SelectComboText mCmbTreeRunProcess, selectedProcess
    End If
    If mCmbRunTargetOutput.ListCount > 0 Then _
        mCmbRunTargetOutput.ListIndex = mCmbRunTargetOutput.ListCount - 1
    ApplyReusablePaletteProcessProjection selectedProcess
    RefreshReusableRunInstructions selectedProcess
    For i = 0 To mLstRunPalette.ListCount - 1
        locationName = NzStr(mLstRunPalette.List(i, 9))
        AddUniqueComboItem mCmbRunLocation, locationName
        AddUniqueComboItem mCmbTreeRunLocation, locationName
    Next i
    mLoading = False
    If activeOutputRow >= 0 Then LoadSelectedReusableActualOutput
    If Not mTxtRunBatchNote Is Nothing Then
        mTxtRunBatchNote.Text = modProductionReusableRun.ReusableRunBatchNote()
        mTxtRunBatchNote.Enabled = Not modProductionReusableRun.ReusableRunBatchNoteFrozen()
    End If
End Sub

Private Function SyncReusableRunBatchNote(ByRef report As String) As Boolean
    If mTxtRunBatchNote Is Nothing Then
        SyncReusableRunBatchNote = True
        Exit Function
    End If
    SyncReusableRunBatchNote = modProductionReusableRun.SetReusableRunBatchNote( _
        mTxtRunBatchNote.Text, report)
End Function

Private Sub mTxtRunBatchNote_Change()
    Dim report As String

    If mLoading Then Exit Sub
    If Not modProductionReusableRun.ReusableRunIsLoaded() Then Exit Sub
    If Not modProductionReusableRun.SetReusableRunBatchNote(mTxtRunBatchNote.Text, report) Then
        ShowStatus report
    End If
End Sub

Private Sub RefreshReusableRunInstructions(ByVal processName As String)
    Dim instructionRows As Variant
    If mLstRunInstructions Is Nothing Then Exit Sub
    instructionRows = modProductionReusableRun.ReusableRunInstructionRows(processName)
    FillListFromArray mLstRunInstructions, instructionRows
End Sub

Private Sub ApplyReusablePaletteProcessProjection(ByVal filterProcess As String)
    Dim i As Long
    Dim processName As String
    Dim ingredientName As String

    filterProcess = Trim$(filterProcess)
    For i = mLstRunPalette.ListCount - 1 To 0 Step -1
        processName = ReusableRunProcessNameForNode(NzStr(mLstRunPalette.List(i, 0)))
        If filterProcess <> "" And _
           StrComp(processName, filterProcess, vbTextCompare) <> 0 Then
            mLstRunPalette.RemoveItem i
        Else
            ingredientName = NzStr(mLstRunPalette.List(i, 2))
            If processName <> "" Then _
                mLstRunPalette.List(i, 2) = processName & " / " & ingredientName
        End If
    Next i
End Sub

Private Function ReusableRunProcessNameForNode(ByVal nodeId As String) As String
    Dim i As Long

    For i = 0 To mLstLoaderLines.ListCount - 1
        If StrComp(NzStr(mLstLoaderLines.List(i, 1)), nodeId, vbTextCompare) = 0 Then
            ReusableRunProcessNameForNode = NzStr(mLstLoaderLines.List(i, 0))
            Exit Function
        End If
    Next i
End Function

Private Sub RefreshReusableRunPaletteProjection()
    Dim paletteRows As Variant
    Dim selectedProcess As String
    Dim wasLoading As Boolean

    selectedProcess = ActiveRunProcess()
    wasLoading = mLoading
    mLoading = True
    paletteRows = modProductionReusableRun.ReusableRunPaletteRows(ActiveRunLocation())
    FillListFromArray mLstRunPalette, paletteRows
    ApplyReusablePaletteProcessProjection selectedProcess
    mLoading = wasLoading
End Sub

Private Function FirstReusableActiveOutputRow() As Long
    Dim i As Long

    FirstReusableActiveOutputRow = -1
    If mLstManagerOutput Is Nothing Then Exit Function
    For i = 0 To mLstManagerOutput.ListCount - 1
        If modProductionReusableRun.ReusableRunOutputDefinitionIndex(i + 1) > 0 Then
            FirstReusableActiveOutputRow = i
            Exit Function
        End If
    Next i
End Function

Private Sub AddUniqueComboItem(ByVal comboControl As MSForms.ComboBox, ByVal itemText As String)
    Dim i As Long
    itemText = Trim$(itemText)
    If comboControl Is Nothing Or itemText = "" Then Exit Sub
    For i = 0 To comboControl.ListCount - 1
        If StrComp(NzStr(comboControl.List(i)), itemText, vbTextCompare) = 0 Then Exit Sub
    Next i
    comboControl.AddItem itemText
End Sub

Private Sub mLstLoaderLines_Click()
    If mLoading Then Exit Sub
    If mLstLoaderLines.ListIndex < 0 Then Exit Sub
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        ShowStatus "Loaded Process line: " & NzStr(mLstLoaderLines.List(mLstLoaderLines.ListIndex, 3))
        Exit Sub
    End If
    RefreshRunPaletteState
    ShowStatus "Acceptable inventory filtered for: " & NzStr(mLstLoaderLines.List(mLstLoaderLines.ListIndex, 3))
End Sub

Private Sub mBtnLoaderClear_Click()
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        modProductionReusableRun.ClearReusableRun
        mLstLoaderLines.Clear
        mLstRunPalette.Clear
        mLstManagerCheck.Clear
        mLstManagerOutput.Clear
        mLstRunInstructions.Clear
        ShowStatus "Reusable Production Run cleared."
        Exit Sub
    End If
    BindOperatorWorkbookForRun
    mProduction.BtnClearRecipeChooser
    RefreshLoaderState
    RefreshManagerState
    ShowStatus "Production Run cleared."
End Sub

Private Sub mLstRunPalette_Click()
    If mLoading Then Exit Sub
    LoadSelectedRunPaletteRow
End Sub

Private Sub mLstRunTree_Click()
    If mLoading Then Exit Sub
    If IsRunTreeParentRow(mLstRunTree, mLstRunTree.ListIndex) Then
        ToggleSelectedRunTreeParent
        Exit Sub
    End If
    LoadSelectedRunPaletteRow
End Sub

Private Sub mBtnRunTreeExpandAll_Click()
    SetAllRunTreeGroupsCollapsed False
    ShowStatus "All ingredient choices shown."
End Sub

Private Sub mBtnRunTreeCollapseAll_Click()
    SetAllRunTreeGroupsCollapsed True
    ShowStatus "All ingredient choices hidden."
End Sub

Private Sub mBtnRunApplyPalette_Click()
    If mLstRunPalette.ListCount = 0 Then
        ResetInventoryCache
        RefreshRunPaletteState
        If mLstRunPalette.ListCount = 0 Then
            ShowStatus "No acceptable inventory assignments were found for this recipe. Use Ingredients Assignment to select each USED ingredient, add acceptable inventory, and Save Assignment."
            Exit Sub
        End If
    End If
    If mLstRunPalette.ListIndex < 0 And mLstRunPalette.ListCount = 1 Then
        mLstRunPalette.ListIndex = 0
        LoadSelectedRunPaletteRow
    End If
    ApplySelectedRunPaletteSplit
End Sub

Private Sub mBtnRunTreeApplyPalette_Click()
    ApplySelectedRunPaletteSplit
End Sub

Private Sub mBtnManagerRefresh_Click()
    Dim refreshReport As String
    Dim refreshed As Boolean

    If modProductionReusableRun.ReusableRunIsLoaded() Then
        RefreshReusableRunControls True
        ShowStatus "Reusable Production Run inventory refreshed from the exact entity projection."
        Exit Sub
    End If

    refreshed = RefreshProductionInventoryReadModel(refreshReport)
    ResetInventoryCache
    RefreshLoaderState
    RefreshManagerState
    If refreshed Then
        ShowStatus "Production Run inventory refreshed. " & refreshReport
    Else
        ShowStatus refreshReport
        MsgBox refreshReport, vbExclamation, "Production Inventory Refresh"
    End If
End Sub

Private Sub mLstManagerOutput_Click()
    If mLoading Then Exit Sub
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        LoadSelectedReusableActualOutput
        Exit Sub
    End If
    LoadSelectedProductionOutput
End Sub

Private Sub LoadSelectedReusableActualOutput()
    Dim wasLoading As Boolean
    Dim outputIndex As Long

    If mLstManagerOutput Is Nothing Or mTxtOutputReal Is Nothing Then Exit Sub
    If mLstManagerOutput.ListIndex < 0 Then Exit Sub
    outputIndex = modProductionReusableRun.ReusableRunOutputDefinitionIndex( _
        mLstManagerOutput.ListIndex + 1)
    wasLoading = mLoading
    mLoading = True
    If outputIndex > 0 Then
        mTxtOutputReal.Enabled = True
        mTxtOutputReal.Text = modProductionReusableRun.ReusableRunActualOutput(outputIndex)
    Else
        mTxtOutputReal.Text = ""
        mTxtOutputReal.Enabled = False
    End If
    mLoading = wasLoading
End Sub

Private Function StageSelectedReusableActualOutput( _
    Optional ByVal showFailure As Boolean = True) As Boolean
    Dim report As String
    Dim outputIndex As Long

    If mLstManagerOutput Is Nothing Or mTxtOutputReal Is Nothing Then Exit Function
    If mLstManagerOutput.ListIndex < 0 And mLstManagerOutput.ListCount = 1 Then
        mLstManagerOutput.ListIndex = 0
    End If
    If mLstManagerOutput.ListIndex < 0 Then
        If showFailure Then ShowStatus "Select an output row before entering Actual Output."
        Exit Function
    End If
    outputIndex = modProductionReusableRun.ReusableRunOutputDefinitionIndex( _
        mLstManagerOutput.ListIndex + 1)
    If outputIndex = 0 Then
        If showFailure Then ShowStatus "Select an active-batch output row before entering Actual Output."
        Exit Function
    End If
    StageSelectedReusableActualOutput = _
        modProductionReusableRun.StageReusableRunActualOutput( _
            outputIndex, mTxtOutputReal.Text, report)
    If showFailure And Not StageSelectedReusableActualOutput Then ShowStatus report
End Function

Private Sub mBtnManagerCheckIn_Click()
    CheckInProductionRun
End Sub

Private Sub mBtnManagerApplyOutput_Click()
    CompleteProductionRun
End Sub

Private Sub mCmbRunProcess_Change()
    If mLoading Then Exit Sub
    SyncRunProcessCombo mCmbRunProcess, mCmbTreeRunProcess
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        RefreshReusableRunPaletteProjection
        RefreshReusableRunInstructions ActiveRunProcess()
        Exit Sub
    End If
    RefreshRunPaletteState
End Sub

Private Sub mCmbTreeRunProcess_Change()
    If mLoading Then Exit Sub
    SyncRunProcessCombo mCmbTreeRunProcess, mCmbRunProcess
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        RefreshReusableRunPaletteProjection
        RefreshReusableRunInstructions ActiveRunProcess()
        Exit Sub
    End If
    RefreshRunPaletteState
End Sub

Private Sub mBtnManagerNext_Click()
    Dim reusableReport As String

    If modProductionReusableRun.ReusableRunIsLoaded() Then
        If modProductionReusableRun.BeginNextReusableBatch(reusableReport) Then
            RefreshReusableRunControls True
        End If
        ShowStatus reusableReport
        Exit Sub
    End If
    BindOperatorWorkbookForRun
    mProduction.BtnNextBatch
    ResetInventoryCache
    RefreshLoaderState
    RefreshManagerState
    ShowStatus "Next Batch completed."
End Sub

Private Sub mBtnManagerPrint_Click()
    BindOperatorWorkbookForRun
    mProduction.BtnPrintRecallCodes
    ShowStatus "Print Recall Codes completed. " & NzStr(mProduction.GetRecallPrintDiagnostic())
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub
