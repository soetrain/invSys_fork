Attribute VB_Name = "mProduction"
' run "mProduction.InitializeProductionUI" in immediate window to clean up UI
Option Explicit

' Production system core module (wiring + helpers).

Private Const SHEET_PRODUCTION As String = "Production"
Private Const SHEET_TEMPLATES As String = "TemplatesTable"
Private Const SHEET_RUNTIME_RECIPES As String = "ProductionRecipes"
Private Const SHEET_RUNTIME_INGREDIENT_PALETTE As String = "ProductionIngredientPalette"
Private Const EVENT_TYPE_PROD_CONSUME As String = "PROD_CONSUME"
Private Const EVENT_TYPE_PROD_COMPLETE As String = "PROD_COMPLETE"
Private Const EVENT_TYPE_DESIGN_CREATE As String = "DESIGN_CREATE"
Private Const PRODUCTION_DEFAULT_ROW_BUDGET As Long = 50
Private Const PRODUCTION_MAX_ROW_BUDGET As Long = 1000

Private Const TABLE_RECIPE_CHOOSER As String = "RC_RecipeChoose"
Private Const TABLE_RECIPE_CHOOSER_GENERATED As String = "RecipeChooser_generated"
Private Const TABLE_INV_PALETTE_GENERATED As String = "InventoryPalette_generated"
Private Const TABLE_RECALL_REPORT As String = "RecallCodesReport"
Private Const TABLE_RUNTIME_RECIPES As String = "tblProductionRecipes"
Private Const TABLE_RUNTIME_INGREDIENT_PALETTE As String = "tblProductionIngredientPalette"
Private Const TABLE_TEMPLATES As String = "TemplatesTable"
' System 1: Recipe List Builder tables.
Private Const TABLE_RECIPE_BUILDER_HEADER As String = "RB_AddRecipeName"
Private Const TABLE_RECIPE_BUILDER_LINES As String = "RecipeBuilder"

Private Const BTN_HIDE_SYSTEM As String = "BTN_HIDE_SYSTEM"
Private Const BTN_SHOW_SYSTEM As String = "BTN_SHOW_SYSTEM"
Private Const BTN_LOAD_RECIPE As String = "BTN_LOAD_RECIPE"             ' System 1: Recipe List Builder
Private Const BTN_SAVE_RECIPE As String = "BTN_SAVE_RECIPE"             ' System 1: Recipe List Builder
Private Const BTN_SAVE_FORMULAS As String = "BTN_SAVE_FORMULAS"
Private Const BTN_BUILD_RECIPE_TABLES As String = "BTN_BUILD_RECIPE_TABLES" ' System 1: Recipe List Builder
Private Const BTN_REMOVE_RECIPE_TABLES As String = "BTN_REMOVE_RECIPE_TABLES" ' System 1: Recipe List Builder
Private Const BTN_CLEAR_RECIPE_BUILDER As String = "BTN_CLEAR_RECIPE_BUILDER" ' System 1: Recipe List Builder
Private Const BTN_CLEAR_RECIPE_CHOOSER As String = "BTN_CLEAR_RECIPE_CHOOSER" ' System 3: Recipe Chooser
Private Const BTN_CLEAR_PALETTE_BUILDER As String = "BTN_CLEAR_PALETTE_BUILDER" ' System 2: Inventory Palette Builder
Private Const BTN_SAVE_PALETTE As String = "BTN_SAVE_PALETTE"
Private Const BTN_TO_USED As String = "BTN_TO_USED"
Private Const BTN_TO_MADE As String = "BTN_TO_MADE"
Private Const BTN_TO_TOTALINV As String = "BTN_TO_TOTALINV"
Private Const BTN_NEXT_BATCH As String = "BTN_NEXT_BATCH"
Private Const BTN_PRINT_CODES As String = "BTN_PRINT_CODES"

Private Const CHK_PROC_PREFIX As String = "CHK_PROC_"
Private Const CHK_BATCH_PREFIX As String = "CHK_BATCH_"
Private Const CHK_RECALL_PREFIX As String = "CHK_RECALL_"

Private Const TEMPLATE_SCOPE_RECIPE_PROCESS As String = "RECIPE_PROCESS"
Private Const TEMPLATE_SCOPE_PALETTE_BUILDER As String = "PALETTE_BUILDER"
Private Const TEMPLATE_SCOPE_PROD_RUN As String = "PROD_RUN"
Private Const TEMPLATE_TABLEKEY_PALETTE As String = "proc_*_palette"
Private Const RECIPE_PROC_TABLE_SUFFIX As String = "rbuilder"
Private Const RECIPE_CHOOSER_TABLE_SUFFIX As String = "rchooser"
Private Const RECIPE_LINES_STAGING_ROW As Long = 500000 ' System 1: staging for RecipeBuilder lines during load
Private Const PALETTE_LINES_STAGING_ROW As Long = 500000 ' System 4: staging for InventoryPalette lines table
Private Const PROD_LAYOUT_RECIPE_HEADER_ADDR As String = "C3"
Private Const PROD_LAYOUT_RECIPE_LINES_ADDR As String = "C6"
Private Const PROD_LAYOUT_PALETTE_RECIPE_ADDR As String = "P3"
Private Const PROD_LAYOUT_PALETTE_ING_ADDR As String = "P6"
Private Const PROD_LAYOUT_PALETTE_ITEM_ADDR As String = "P9"
Private Const PROD_LAYOUT_CHOOSER_ADDR As String = "Z3"
Private Const PROD_LAYOUT_CHOOSER_GEN_ADDR As String = "Z6"
Private Const PROD_LAYOUT_OUTPUT_ADDR As String = "AJ4"
Private Const PROD_LAYOUT_CHECK_ADDR As String = "AR4"
Private Const PROD_RECALL_REPORT_SHEET As String = "RecallCodesPrint"
Private Const SHAPE_TYPE_FORM_CONTROL As Long = 8
Private Const SHAPE_VISIBLE_FALSE As Long = 0
Private Const SHAPE_VISIBLE_TRUE As Long = -1

Private mRowCountCache As Object
Private mPaletteTableMeta As Object
Private mHiddenSystems As Collection
Private mRecipePicker As Object
Private mProcessItemPicker As Object
Private mPickerRouter As Object
Private mSystemGroupsInit As Boolean
Private mSystemGroupNames(1 To 4) As String
Private mSystemGroupTables(1 To 4) As Variant
Private mProductionOperatorWorkbook As Workbook
Private mProductionLayoutValidationForm As frmProduction

Public Sub BindProductionOperatorWorkbook(ByVal operatorWb As Workbook)
    If operatorWb Is Nothing Then
        Set mProductionOperatorWorkbook = Nothing
    Else
        Set mProductionOperatorWorkbook = operatorWb
    End If
End Sub

Public Sub ClearProductionOperatorWorkbookBinding(Optional ByVal operatorWb As Workbook = Nothing)
    If mProductionOperatorWorkbook Is Nothing Then Exit Sub
    If operatorWb Is Nothing Then
        Set mProductionOperatorWorkbook = Nothing
    ElseIf mProductionOperatorWorkbook Is operatorWb Then
        Set mProductionOperatorWorkbook = Nothing
    End If
End Sub

Public Sub HandleProductionOperatorWorkbookClosing(ByVal operatorWb As Workbook)
    If operatorWb Is Nothing Then Exit Sub
    If mProductionOperatorWorkbook Is Nothing Then Exit Sub
    If Not mProductionOperatorWorkbook Is operatorWb Then Exit Sub

    On Error Resume Next
    Unload frmProduction
    Set mProductionOperatorWorkbook = Nothing
    On Error GoTo 0
End Sub

Public Sub InitializeProductionUI()
    InitializeProductionUiForWorkbook Application.ActiveWorkbook
End Sub

Public Sub InitializeProductionUiForWorkbook(Optional ByVal targetWb As Workbook = Nothing)
    Dim surfaceReport As String
    Dim wb As Workbook

    Set wb = ResolveProductionWorkbook(targetWb, SHEET_PRODUCTION)
    If wb Is Nothing Then Set wb = ThisWorkbook

    Call modOperationsPrimitiveBridge.EnsureProductionWorkbookSurface(wb.Name, surfaceReport)
    ArrangeProductionSurface wb
    PrimeProductionRowCountCache wb
    EnsureProductionButtons
    EnsureSystemGroups
    modOperationsPrimitiveBridge.InitializeProductionAutoSnapshot wb.Name
End Sub

Public Sub BtnOpenProductionForm()
    On Error GoTo ErrHandler

    Dim launcherStage As String
    Dim preferredWorkbookName As String
    Dim workbookName As String
    Dim report As String
    Dim wb As Workbook

    launcherStage = "capability"
    If Not modRoleUiAccess.RequireCurrentUserCapabilityCached("PROD_POST") Then Exit Sub

    launcherStage = "capture active workbook"
    If Not Application.ActiveWorkbook Is Nothing Then
        preferredWorkbookName = Application.ActiveWorkbook.Name
    End If

    launcherStage = "resolve or provision Production workbook"
    If Not modOperationsPrimitiveBridge.OpenOrCreateCurrentRoleOperatorWorkbook( _
            preferredWorkbookName, "PRODUCTION", workbookName, report) Then
        If Trim$(report) = "" Then
            report = "The station-local Production operator workbook could not be opened."
        End If
        MsgBox report, vbExclamation
        Exit Sub
    End If

    launcherStage = "capture resolved Production workbook"
    Set wb = modOperationsInit.ResolveOpenWorkbookByName(workbookName)
    If wb Is Nothing Then
        MsgBox "The resolved Production operator workbook is no longer open.", vbExclamation
        Exit Sub
    End If

    launcherStage = "show production form"
    ShowProductionForm wb
    Exit Sub

ErrHandler:
    ShowProductionLauncherError launcherStage, Err.Number, Err.Source, Err.Description
End Sub

Public Sub ShowProductionForm(Optional ByVal targetWb As Workbook = Nothing)
    On Error GoTo ErrHandler

    Dim wb As Workbook
    Dim repairReport As String
    Dim quietStarted As Boolean
    Dim launcherStage As String
    Dim preferredWorkbookName As String
    Dim workbookName As String

    launcherStage = "validate Production workbook"
    If Not targetWb Is Nothing Then preferredWorkbookName = targetWb.Name
    If Not modOperationsPrimitiveBridge.ResolveEligibleRoleOperatorWorkbookName( _
            preferredWorkbookName, "PRODUCTION", workbookName, repairReport) Then
        MsgBox repairReport, vbExclamation
        Exit Sub
    End If
    Set wb = modOperationsInit.ResolveOpenWorkbookByName(workbookName)
    If wb Is Nothing Then
        MsgBox "Open a Production operator workbook before opening the Production form.", vbExclamation
        Exit Sub
    End If

    launcherStage = "begin quiet UI"
    modOperationsPrimitiveBridge.BeginQuietUiForWorkbook wb.Name
    quietStarted = True
    launcherStage = "repair Production surface"
    If Not modOperationsPrimitiveBridge.EnsureProductionWorkbookSurface(wb.Name, repairReport) Then
        If Trim$(repairReport) = "" Then repairReport = "Production surface repair failed without detail."
        MsgBox repairReport, vbCritical
        GoTo CleanExit
    End If
    launcherStage = "initialize Production UI"
    InitializeProductionUiForWorkbook wb
    launcherStage = "bind Production form"
    frmProduction.SetOperatorWorkbook wb
    launcherStage = "initialize Production form"
    frmProduction.InitializeFromProduction
    If quietStarted Then
        launcherStage = "end quiet UI"
        modUiQuiet.EndQuietUi
        quietStarted = False
    End If
    launcherStage = "show Production form"
    frmProduction.Show vbModeless

CleanExit:
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    On Error GoTo 0
    Exit Sub

ErrHandler:
    ShowProductionLauncherError launcherStage, Err.Number, Err.Source, Err.Description
    Resume CleanExit
End Sub

Private Sub ShowProductionLauncherError(ByVal launcherStage As String, _
                                        ByVal errorNumber As Long, _
                                        ByVal errorSource As String, _
                                        ByVal errorDescription As String)
    MsgBox "Production form failed [Stage=" & Trim$(launcherStage) & _
           "; Err.Number=" & CStr(errorNumber) & _
           "; Err.Source=" & modOperationsInit.SanitizeLauncherErrorSource(errorSource) & _
           "]: " & errorDescription, vbCritical
End Sub

Public Function ProductionFormInitializeSmokeForWorkbook(ByVal operatorWb As Workbook) As String
    On Error GoTo ErrHandler

    Dim frm As frmProduction
    Dim pageCount As Long
    Dim statusText As String
    Dim windowStyle As String

    If operatorWb Is Nothing Then
        Err.Raise vbObjectError + 7310, "ProductionFormInitializeSmokeForWorkbook", _
                  "Production operator workbook is required."
    End If
    If operatorWb.IsAddin Then
        Err.Raise vbObjectError + 7311, "ProductionFormInitializeSmokeForWorkbook", _
                  "Production operator workbook cannot be an add-in."
    End If

    Set frm = New frmProduction
    pageCount = frm.TestInitializeForWorkbook(operatorWb)
    statusText = frm.TestStatusText()
    Call modProductionFormWindow.EnableResizable(frm, True, True)
    windowStyle = modProductionFormWindow.DiagnoseWindowStyle(frm)

    If pageCount <> 6 Then
        Err.Raise vbObjectError + 7312, "ProductionFormInitializeSmokeForWorkbook", _
                  "Production form page count was " & CStr(pageCount) & "; expected 6."
    End If
    If InStr(1, statusText, "failed", vbTextCompare) > 0 Then
        Err.Raise vbObjectError + 7313, "ProductionFormInitializeSmokeForWorkbook", statusText
    End If
    If InStr(1, statusText, "Inventory: ContractVersion=R1-INVENTORY-1", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 7314, "ProductionFormInitializeSmokeForWorkbook", _
                  "Production form did not report a healthy Inventory Domain. " & statusText
    End If
    If InStr(1, windowStyle, "Resizable=True", vbTextCompare) = 0 _
       Or InStr(1, windowStyle, "Minimize=True", vbTextCompare) = 0 _
       Or InStr(1, windowStyle, "Maximize=True", vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 7315, "ProductionFormInitializeSmokeForWorkbook", _
                  "Production form did not satisfy the standard window contract. " & windowStyle
    End If

    ProductionFormInitializeSmokeForWorkbook = "OK|Pages=" & CStr(pageCount) & _
                                                "|WindowStyle=" & windowStyle & _
                                                "|Status=" & statusText

CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    On Error GoTo 0
    Exit Function

ErrHandler:
    ProductionFormInitializeSmokeForWorkbook = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
    Resume CleanExit
End Function

Public Function RunProductionBatchScaleContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProductionBatchScaleContractTest = "FAIL|FormNotOpen"
    Else
        RunProductionBatchScaleContractTest = frmProduction.TestBatchScaleContract()
    End If
    Exit Function

Failed:
    RunProductionBatchScaleContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunReusableProductionSurfaceContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunReusableProductionSurfaceContractTest = "FAIL|FormNotOpen"
    Else
        RunReusableProductionSurfaceContractTest = _
            frmProduction.TestReusableProductionSurfaceContract()
    End If
    Exit Function

Failed:
    RunReusableProductionSurfaceContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunProcessWorksheetRoundTripContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProcessWorksheetRoundTripContractTest = "FAIL|FormNotOpen"
    Else
        RunProcessWorksheetRoundTripContractTest = _
            frmProduction.TestProcessWorksheetRoundTripContract()
    End If
    Exit Function

Failed:
    RunProcessWorksheetRoundTripContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunProcessWorksheetWorkbenchContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProcessWorksheetWorkbenchContractTest = "FAIL|FormNotOpen"
    Else
        RunProcessWorksheetWorkbenchContractTest = _
            frmProduction.TestProcessWorksheetWorkbenchContract()
    End If
    Exit Function

Failed:
    RunProcessWorksheetWorkbenchContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunProcessWorksheetBulkImportContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProcessWorksheetBulkImportContractTest = "FAIL|FormNotOpen"
    Else
        RunProcessWorksheetBulkImportContractTest = _
            frmProduction.TestProcessWorksheetBulkImportContract()
    End If
    Exit Function

Failed:
    RunProcessWorksheetBulkImportContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunProcessWorksheetOutputPickerContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProcessWorksheetOutputPickerContractTest = "FAIL|FormNotOpen"
    Else
        RunProcessWorksheetOutputPickerContractTest = _
            frmProduction.TestProcessWorksheetOutputPickerContract()
    End If
    Exit Function

Failed:
    RunProcessWorksheetOutputPickerContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunProcessEditExportContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProcessEditExportContractTest = "FAIL|FormNotOpen"
    Else
        RunProcessEditExportContractTest = _
            frmProduction.TestProcessEditExportContract()
    End If
    Exit Function

Failed:
    RunProcessEditExportContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunReusableProductionFormActionContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunReusableProductionFormActionContractTest = "FAIL|FormNotOpen"
    Else
        RunReusableProductionFormActionContractTest = _
            frmProduction.TestReusableProductionFormActionContract()
    End If
    Exit Function

Failed:
    RunReusableProductionFormActionContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunReusableProductionRunActionContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunReusableProductionRunActionContractTest = "FAIL|FormNotOpen"
    Else
        RunReusableProductionRunActionContractTest = _
            frmProduction.TestReusableProductionRunActionContract()
    End If
    Exit Function
Failed:
    RunReusableProductionRunActionContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunChaiForkConvergenceRunActionContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunChaiForkConvergenceRunActionContractTest = "FAIL|FormNotOpen"
    Else
        RunChaiForkConvergenceRunActionContractTest = _
            frmProduction.TestChaiForkConvergenceRunActionContract()
    End If
    Exit Function
Failed:
    RunChaiForkConvergenceRunActionContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunEaWholeUnitActionContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunEaWholeUnitActionContractTest = "FAIL|FormNotOpen"
    Else
        RunEaWholeUnitActionContractTest = frmProduction.TestEaWholeUnitActionContract()
    End If
    Exit Function

Failed:
    RunEaWholeUnitActionContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunProductionOutputRegulationActionContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProductionOutputRegulationActionContractTest = "FAIL|FormNotOpen"
    Else
        RunProductionOutputRegulationActionContractTest = _
            frmProduction.TestProductionOutputRegulationActionContract()
    End If
    Exit Function
Failed:
    RunProductionOutputRegulationActionContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunProductionVariableQuantityModeActionContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProductionVariableQuantityModeActionContractTest = "FAIL|FormNotOpen"
    Else
        RunProductionVariableQuantityModeActionContractTest = _
            frmProduction.TestProductionVariableQuantityModeActionContract()
    End If
    Exit Function
Failed:
    RunProductionVariableQuantityModeActionContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunProductionExternalStockUomConversionActionContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProductionExternalStockUomConversionActionContractTest = "FAIL|FormNotOpen"
    Else
        RunProductionExternalStockUomConversionActionContractTest = _
            frmProduction.TestProductionExternalStockUomConversionActionContract()
    End If
    Exit Function
Failed:
    RunProductionExternalStockUomConversionActionContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunProductionExternalStockUomWorksheetHandlerContractTest() As String
    On Error GoTo Failed

    RunProductionExternalStockUomWorksheetHandlerContractTest = _
        frmProduction.TestProductionExternalStockUomConversionActionContract()
    Exit Function
Failed:
    RunProductionExternalStockUomWorksheetHandlerContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunProductionBatchNoteHandlerContractTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProductionBatchNoteHandlerContractTest = "FAIL|FormNotOpen"
    Else
        RunProductionBatchNoteHandlerContractTest = _
            frmProduction.TestProductionBatchNoteHandlerContract()
    End If
    Exit Function
Failed:
    RunProductionBatchNoteHandlerContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function RunReusableProductionRestartActionContractTest( _
    ByVal recipeId As String, ByVal recipeVersion As String, _
    ByVal expectedWorkbookFullName As String) As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunReusableProductionRestartActionContractTest = "FAIL|FormNotOpen"
    Else
        RunReusableProductionRestartActionContractTest = _
            frmProduction.TestReusableProductionRestartActionContract( _
                recipeId, recipeVersion, expectedWorkbookFullName)
    End If
    Exit Function
Failed:
    RunReusableProductionRestartActionContractTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function ShowProductionLayoutForValidation(ByVal requestedWidth As Double, _
                                                  ByVal requestedHeight As Double, _
                                                  Optional ByVal pageIndex As Long = 2) As String
    On Error GoTo FailValidation

    CloseProductionLayoutValidation
    Set mProductionLayoutValidationForm = New frmProduction
    mProductionLayoutValidationForm.Show vbModeless
    DoEvents
    ShowProductionLayoutForValidation = _
        mProductionLayoutValidationForm.TestPrepareLayoutForScreenshot( _
            requestedWidth, requestedHeight, pageIndex)
    DoEvents
    Exit Function

FailValidation:
    ShowProductionLayoutForValidation = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Sub CloseProductionLayoutValidation()
    On Error Resume Next
    If Not mProductionLayoutValidationForm Is Nothing Then
        Unload mProductionLayoutValidationForm
    End If
    Set mProductionLayoutValidationForm = Nothing
    On Error GoTo 0
End Sub

Public Function CurrentProductionLayoutValidationReport( _
    Optional ByVal pageIndex As Long = 2) As String
    If mProductionLayoutValidationForm Is Nothing Then
        CurrentProductionLayoutValidationReport = "FAIL|No validation form is open."
        Exit Function
    End If
    CurrentProductionLayoutValidationReport = _
        mProductionLayoutValidationForm.TestCurrentLayoutGeometryReport(pageIndex)
End Function

Public Function RunProductionRunListResponsiveLayoutTest() As String
    On Error GoTo Failed

    BtnOpenProductionForm
    If Not frmProduction.Visible Then
        RunProductionRunListResponsiveLayoutTest = "FAIL|FormNotOpen"
    Else
        RunProductionRunListResponsiveLayoutTest = _
            frmProduction.TestRunListResponsiveLayoutReportForSize()
    End If
    Exit Function

Failed:
    RunProductionRunListResponsiveLayoutTest = _
        "FAIL|" & CStr(Err.Number) & "|" & Err.Description
End Function

Public Function ProductionFormTwoBatchActionReportForTest(ByVal operatorWb As Workbook, _
                                                          ByVal inputItemCode As String, _
                                                          ByVal inputItemName As String, _
                                                          ByVal inputQty As Double, _
                                                          ByVal inputUom As String, _
                                                          ByVal inputLocation As String, _
                                                          ByVal outputQty As Double, _
                                                          Optional ByVal activatedWb As Workbook = Nothing) As String
    Dim frm As frmProduction

    On Error GoTo FailAction
    Set frm = New frmProduction
    ProductionFormTwoBatchActionReportForTest = _
        frm.TestRunTwoConsecutiveBatchesForWorkbook(operatorWb, inputItemCode, _
            inputItemName, inputQty, inputUom, inputLocation, outputQty, activatedWb)
    Unload frm
    Exit Function

FailAction:
    ProductionFormTwoBatchActionReportForTest = "FAIL|Error=" & CStr(Err.Number) & _
        "|Source=" & Err.Source & "|Description=" & Err.Description
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    On Error GoTo 0
End Function

' ===== Worksheet event entry points =====
Public Sub HandleProductionSelectionChange(ByVal target As Range)
    If target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(target) Then Exit Sub
    EnsurePickerRouter
    mPickerRouter.HandleSelectionChange target
End Sub

Public Sub ShowProductionProcessItemSearch(ByVal target As Range)
    If target Is Nothing Then Exit Sub
    If Not modProductionProcessWorksheet.IsProcessWorksheetItemSearchTarget(target) Then Exit Sub
    If mProcessItemPicker Is Nothing Then Set mProcessItemPicker = CreateDynItemSearch()
    mProcessItemPicker.UseRoleProfile "production"
    mProcessItemPicker.ShowForCell target
End Sub

Public Function ProductionProcessItemSearchVisibleForTest() As Boolean
    If mProcessItemPicker Is Nothing Then Exit Function
    ProductionProcessItemSearchVisibleForTest = _
        CBool(mProcessItemPicker.IsSearchVisible())
End Function

Public Function ProductionProcessItemSearchResultCountForTest() As Long
    If mProcessItemPicker Is Nothing Then Exit Function
    ProductionProcessItemSearchResultCountForTest = _
        CLng(mProcessItemPicker.SearchResultCount())
End Function

Public Function CommitFirstProductionProcessItemSearchResultForTest() As Boolean
    If mProcessItemPicker Is Nothing Then Exit Function
    CommitFirstProductionProcessItemSearchResultForTest = _
        CBool(mProcessItemPicker.CommitFirstSearchResultForTest())
End Function

Public Sub CloseProductionProcessItemSearchForTest()
    If mProcessItemPicker Is Nothing Then Exit Sub
    mProcessItemPicker.CloseSearch
End Sub

Public Sub HandleProductionBeforeDoubleClick(ByVal target As Range, ByRef Cancel As Boolean)
    If target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(target) Then Exit Sub
    EnsurePickerRouter
    If mPickerRouter.HandleBeforeDoubleClick(target, Cancel) Then Exit Sub
End Sub

Private Sub EnsurePickerRouter()
    If mPickerRouter Is Nothing Then Set mPickerRouter = CreatePickerRouter()
End Sub

Public Sub HandleProductionChange(ByVal target As Range)
    If target Is Nothing Then Exit Sub
    If Not IsOnProductionSheet(target) Then Exit Sub

    Dim lo As ListObject
    On Error Resume Next
    Set lo = target.ListObject
    On Error GoTo 0
    If lo Is Nothing Then Exit Sub

    If IsBandManagedTable(lo) Then
        EnsureRowCountCache
        Dim key As String: key = RowCountCacheKey(lo)
        Dim newCount As Long: newCount = ListObjectRowCount(lo)
        If Not mRowCountCache.Exists(key) Then
            mRowCountCache(key) = newCount
            Exit Sub
        End If
        Dim oldCount As Long: oldCount = CLng(mRowCountCache(key))
        If newCount > oldCount Then
            If LCase$(lo.Name) <> "prod_invsys_check" Then
                ExpandProcessBandForTable lo, (newCount - oldCount)
            End If
        End If
        mRowCountCache(key) = newCount
    End If

    If LCase$(lo.Name) = "productionoutput" Then
        RenderOutputRowCheckboxes lo.Parent
    End If
End Sub

' ===== Band/table helpers =====
Private Sub EnsureRowCountCache()
    If mRowCountCache Is Nothing Then
        Set mRowCountCache = CreateObject("Scripting.Dictionary")
    End If
End Sub

Private Sub ExpandProcessBandForTable(ByVal lo As ListObject, ByVal rowsAdded As Long)
    ' UserForm production uses reserved table bands; do not shift worksheet rows.
End Sub

Private Function GetProductionBandTables(ByVal ws As Worksheet, ByVal bandKey As String) As Collection
    Dim result As New Collection
    Dim lo As ListObject

    If ws Is Nothing Then
        Set GetProductionBandTables = result
        Exit Function
    End If

    For Each lo In ws.ListObjects
        If StrComp(BandKeyForProductionTable(lo), bandKey, vbTextCompare) = 0 Then
            result.Add lo
        End If
    Next lo

    Set GetProductionBandTables = result
End Function

Private Sub ComputeProductionBandBounds(ByVal bandTables As Collection, ByRef bandTop As Long, ByRef bandBottom As Long, ByRef bandLeft As Long, ByRef bandRight As Long)
    Dim lo As ListObject
    Dim topSet As Boolean

    For Each lo In bandTables
        If lo Is Nothing Then GoTo NextLo
        Dim rTop As Long: rTop = lo.Range.Row
        Dim rBottom As Long: rBottom = lo.Range.Row + lo.Range.Rows.Count - 1
        Dim cLeft As Long: cLeft = lo.Range.Column
        Dim cRight As Long: cRight = lo.Range.Column + lo.Range.Columns.Count - 1
        If Not topSet Then
            bandTop = rTop
            bandBottom = rBottom
            bandLeft = cLeft
            bandRight = cRight
            topSet = True
        Else
            If rTop < bandTop Then bandTop = rTop
            If rBottom > bandBottom Then bandBottom = rBottom
            If cLeft < bandLeft Then bandLeft = cLeft
            If cRight > bandRight Then bandRight = cRight
        End If
NextLo:
    Next lo
End Sub

Private Function BandKeyForProductionTable(ByVal lo As ListObject) As String
    If lo Is Nothing Then Exit Function

    Dim nm As String
    nm = LCase$(Trim$(lo.Name))
    If nm = "productionoutput" Or nm = "prod_invsys_check" Then Exit Function

    On Error Resume Next
    If Not lo.DataBodyRange Is Nothing Then
        Dim cProc As Long
        cProc = ColumnIndex(lo, "PROCESS")
        If cProc > 0 Then
            BandKeyForProductionTable = NormalizeProcessBandKey(NzStr(lo.DataBodyRange.Cells(1, cProc).Value))
        End If
    End If
    On Error GoTo 0

    If BandKeyForProductionTable = "" Then
        BandKeyForProductionTable = NormalizeProcessBandKey(ExtractProcessKeyFromTableName(lo.Name))
    End If
End Function

Private Function RowCountCacheKey(ByVal lo As ListObject) As String
    If lo Is Nothing Then Exit Function
    RowCountCacheKey = lo.Parent.Parent.Name & "|" & lo.Parent.Name & "|" & lo.Name
End Function

Private Sub PrimeProductionRowCountCache(Optional ByVal targetWb As Workbook = Nothing)
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject

    Set wb = ResolveProductionWorkbook(targetWb, SHEET_PRODUCTION)
    If wb Is Nothing Then Exit Sub

    Set ws = WorkbookSheetExists(wb, SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub

    EnsureRowCountCache
    For Each lo In ws.ListObjects
        If IsBandManagedTable(lo) Then
            mRowCountCache(RowCountCacheKey(lo)) = ListObjectRowCount(lo)
        End If
    Next lo
End Sub

Private Sub EnsurePaletteTableMeta()
    If mPaletteTableMeta Is Nothing Then
        Set mPaletteTableMeta = CreateObject("Scripting.Dictionary")
    End If
End Sub

Private Sub ClearPaletteTableMeta()
    If Not mPaletteTableMeta Is Nothing Then mPaletteTableMeta.RemoveAll
End Sub

Private Function IsOnProductionSheet(ByVal target As Range) As Boolean
    On Error Resume Next
    IsOnProductionSheet = (target.Worksheet.Name = SHEET_PRODUCTION)
    On Error GoTo 0
End Function

Private Function IsPaletteTable(lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(lo.Name)
    If nm = LCase$(TABLE_INV_PALETTE_GENERATED) Then
        IsPaletteTable = True
    ElseIf nm Like "proc_*_palette" Then
        IsPaletteTable = True
    End If
End Function

Private Function IsBandManagedTable(lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(lo.Name)
    If nm = LCase$(TABLE_INV_PALETTE_GENERATED) Then
        IsBandManagedTable = True
    ElseIf nm Like "proc_*_palette" Then
        IsBandManagedTable = True
    ElseIf nm = "productionoutput" Then
        IsBandManagedTable = True
    ElseIf nm = "prod_invsys_check" Then
        IsBandManagedTable = True
    End If
End Function

Private Function ListObjectRowCount(lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    ListObjectRowCount = lo.DataBodyRange.rows.count
End Function

' ===== Generic helpers =====
Public Function GetProductionSheet() As Worksheet
    Set GetProductionSheet = SheetExists(SHEET_PRODUCTION)
End Function

Public Function SheetExists(nameOrCode As String) As Worksheet
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ResolveProductionWorkbook(, nameOrCode)
    If wb Is Nothing Then Set wb = ThisWorkbook

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nameOrCode, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, nameOrCode, vbTextCompare) = 0 Then
            Set SheetExists = ws
            Exit Function
        End If
    Next ws
End Function

Private Function ResolveProductionWorkbook(Optional ByVal preferredWb As Workbook = Nothing, Optional ByVal requiredSheet As String = "") As Workbook
    If Not preferredWb Is Nothing Then
        Set ResolveProductionWorkbook = preferredWb
        Exit Function
    End If

    Set ResolveProductionWorkbook = ResolveBoundProductionWorkbook(requiredSheet)
    If Not ResolveProductionWorkbook Is Nothing Then Exit Function

    If Not Application.ActiveWorkbook Is Nothing Then
        If Not Application.ActiveWorkbook.IsAddin Then
            If requiredSheet = "" Then
                Set ResolveProductionWorkbook = Application.ActiveWorkbook
                Exit Function
            ElseIf Not WorkbookSheetExists(Application.ActiveWorkbook, requiredSheet) Is Nothing Then
                Set ResolveProductionWorkbook = Application.ActiveWorkbook
                Exit Function
            End If
        End If
    End If

    If requiredSheet <> "" Then
        Set ResolveProductionWorkbook = FindOpenOperationalWorkbookWithSheet(requiredSheet)
        If Not ResolveProductionWorkbook Is Nothing Then Exit Function
    Else
        Set ResolveProductionWorkbook = FindFirstOpenOperationalWorkbook()
        If Not ResolveProductionWorkbook Is Nothing Then Exit Function
    End If

End Function

Private Function FindFirstOpenOperationalWorkbook() As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If Not wb Is Nothing Then
            If Not wb.IsAddin Then
                Set FindFirstOpenOperationalWorkbook = wb
                Exit Function
            End If
        End If
    Next wb
End Function

Private Function FindOpenOperationalWorkbookWithSheet(ByVal requiredSheet As String) As Workbook
    Dim wb As Workbook

    If Trim$(requiredSheet) = "" Then Exit Function

    For Each wb In Application.Workbooks
        If Not wb Is Nothing Then
            If Not wb.IsAddin Then
                If Not WorkbookSheetExists(wb, requiredSheet) Is Nothing Then
                    Set FindOpenOperationalWorkbookWithSheet = wb
                    Exit Function
                End If
            End If
        End If
    Next wb
End Function

Private Function WorkbookSheetExists(ByVal wb As Workbook, ByVal nameOrCode As String) As Worksheet
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, nameOrCode, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, nameOrCode, vbTextCompare) = 0 Then
            Set WorkbookSheetExists = ws
            Exit Function
        End If
    Next ws
End Function

Public Function GetListObject(ws As Worksheet, tableName As String) As ListObject
    On Error Resume Next
    Set GetListObject = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Sub ArrangeProductionSurface(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim rowBudget As Long

    If wb Is Nothing Then Exit Sub
    Set ws = WorkbookSheetExists(wb, SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub
    rowBudget = ProductionSurfaceRowBudget(ws)

    Set lo = GetListObject(ws, TABLE_RECIPE_BUILDER_HEADER)
    MoveListObjectToAddress lo, PROD_LAYOUT_RECIPE_HEADER_ADDR

    Set lo = GetListObject(ws, TABLE_RECIPE_BUILDER_LINES)
    If Not lo Is Nothing Then
        If lo.Range.Row < RECIPE_LINES_STAGING_ROW Then MoveListObjectToAddress lo, PROD_LAYOUT_RECIPE_LINES_ADDR
    End If

    Set lo = GetListObject(ws, "IP_ChooseRecipe")
    MoveListObjectToAddress lo, PROD_LAYOUT_PALETTE_RECIPE_ADDR
    Set lo = GetListObject(ws, "IP_ChooseIngredient")
    MoveListObjectToAddress lo, PROD_LAYOUT_PALETTE_ING_ADDR
    Set lo = GetListObject(ws, "IP_ChooseItem")
    MoveListObjectToAddress lo, PROD_LAYOUT_PALETTE_ITEM_ADDR

    Set lo = GetListObject(ws, TABLE_RECIPE_CHOOSER)
    MoveListObjectToAddress lo, PROD_LAYOUT_CHOOSER_ADDR
    Set lo = GetListObject(ws, TABLE_RECIPE_CHOOSER_GENERATED)
    MoveListObjectToAddress lo, PROD_LAYOUT_CHOOSER_GEN_ADDR
    EnsureProductionReservedRows lo, rowBudget

    Set lo = GetListObject(ws, "ProductionOutput")
    MoveListObjectToAddress lo, PROD_LAYOUT_OUTPUT_ADDR
    EnsureProductionReservedRows lo, rowBudget
    Set lo = GetListObject(ws, "Prod_invSys_Check")
    MoveListObjectToAddress lo, PROD_LAYOUT_CHECK_ADDR
    EnsureProductionReservedRows lo, rowBudget

    Set lo = GetListObject(ws, TABLE_INV_PALETTE_GENERATED)
    If Not lo Is Nothing Then
        MoveListObjectToRowCol lo, PALETTE_LINES_STAGING_ROW, ws.Range(PROD_LAYOUT_OUTPUT_ADDR).Column
        EnsureProductionReservedRows lo, rowBudget
    End If
End Sub

Private Function ProductionSurfaceRowBudget(ByVal ws As Worksheet) As Long
    Dim lo As ListObject
    Dim c As Long

    ProductionSurfaceRowBudget = PRODUCTION_DEFAULT_ROW_BUDGET
    If ws Is Nothing Then Exit Function
    Set lo = GetListObject(ws, TABLE_RECIPE_BUILDER_HEADER)
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    c = ColumnIndex(lo, "ROW_BUDGET")
    If c = 0 Then Exit Function
    ProductionSurfaceRowBudget = NormalizeProductionRowBudget(lo.DataBodyRange.Cells(1, c).Value)
End Function

Private Sub EnsureProductionReservedRows(ByVal lo As ListObject, ByVal rowBudget As Long)
    On Error GoTo CleanExit

    rowBudget = NormalizeProductionRowBudget(rowBudget)
    If lo Is Nothing Then Exit Sub
    Do While lo.ListRows.Count < rowBudget
        lo.ListRows.Add AlwaysInsert:=False
    Loop

CleanExit:
End Sub

Private Sub MoveListObjectToAddress(ByVal lo As ListObject, ByVal addressText As String)
    If lo Is Nothing Then Exit Sub
    MoveListObjectToRowCol lo, lo.Parent.Range(addressText).Row, lo.Parent.Range(addressText).Column
End Sub

Private Sub MoveListObjectToRowCol(ByVal lo As ListObject, ByVal targetRow As Long, ByVal targetCol As Long)
    Dim ws As Worksheet
    Dim dest As Range

    If lo Is Nothing Then Exit Sub
    If targetRow < 1 Or targetCol < 1 Then Exit Sub
    If lo.Range.Row = targetRow And lo.Range.Column = targetCol Then Exit Sub

    Set ws = lo.Parent
    Set dest = ws.Cells(targetRow, targetCol)

    On Error Resume Next
    lo.Range.Cut Destination:=dest
    ClearExcelClipboardStateProduction
    Err.Clear
    On Error GoTo 0
End Sub

Public Function LoadRecipeList() As Variant
    Dim canonicalRecipes As Variant
    Dim pendingRecipes As Variant
    Dim legacyRecipes As Variant

    If ProductionDesignsEnabled() Then
        canonicalRecipes = LoadRecipeListFromDesigns("")
        pendingRecipes = LoadPendingStagedRecipeList()
    End If
    legacyRecipes = LoadLegacyRuntimeRecipeList()
    If Not IsUsableProductionArray(legacyRecipes) Then legacyRecipes = LoadLegacyRecipeList()

    LoadRecipeList = BuildUnifiedRecipeList(canonicalRecipes, pendingRecipes, legacyRecipes)
End Function

Public Function LoadReleasedRecipeList() As Variant
    Dim designs As Variant

    If ProductionDesignsEnabled() Then
        designs = LoadRecipeListFromDesigns("RELEASED")
        LoadReleasedRecipeList = designs
    Else
        LoadReleasedRecipeList = LoadLegacyRecipeList()
    End If
End Function

Public Function ReleaseRecipeForProduction(ByVal recipeId As String) As String
    On Error GoTo FailRelease

    Dim capabilityError As String
    Dim designVersion As String
    Dim eventId As String
    Dim queueError As String
    Dim processorReport As String
    Dim appliedCount As Long
    Dim warehouseId As String
    Dim stationId As String
    Dim currentStatus As String

    recipeId = CanonicalRecipeIdProduction(recipeId)
    If recipeId = "" Then
        ReleaseRecipeForProduction = "A Recipe ID is required."
        Exit Function
    End If
    If Not ProductionDesignsEnabled() Then
        ReleaseRecipeForProduction = "Designs Domain is disabled; legacy recipes do not require release."
        Exit Function
    End If
    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", capabilityError) Then
        ReleaseRecipeForProduction = capabilityError
        Exit Function
    End If

    designVersion = LatestRecipeVersionForReleaseProduction(recipeId)
    If designVersion = "" Then
        ReleaseRecipeForProduction = "No saved Designs version was found for Recipe ID " & recipeId & "."
        Exit Function
    End If
    currentStatus = DesignStatusForVersionProduction(recipeId, designVersion)
    If StrComp(currentStatus, "RELEASED", vbTextCompare) = 0 Then
        ReleaseRecipeForProduction = "Recipe " & recipeId & " version " & designVersion & _
                                     " is already released and available for Production Run."
        Exit Function
    End If

    If Not modRoleEventWriter.QueueDesignEventCurrent( _
        "DESIGN_RELEASE", recipeId, designVersion, "", _
        "Recipe Builder release for production", "", _
        eventId, queueError) Then
        ReleaseRecipeForProduction = "Release was not queued: " & queueError
        Exit Function
    End If

    ResolveProductionTargetIdentity warehouseId, stationId
    appliedCount = modProcessor.RunBatch(warehouseId, 0, processorReport)
    If warehouseId <> "" Then Call modConfig.LoadConfig(warehouseId, stationId)
    currentStatus = DesignStatusForVersionProduction(recipeId, designVersion)
    If StrComp(currentStatus, "RELEASED", vbTextCompare) = 0 Then
        ReleaseRecipeForProduction = "Released Recipe " & recipeId & " version " & designVersion & _
                                     " for Production Run. EventID=" & eventId
    Else
        ReleaseRecipeForProduction = "Release queued but not yet applied for Recipe " & _
                                     recipeId & " version " & designVersion & _
                                     ". EventID=" & eventId & "; Processor applied=" & _
                                     CStr(appliedCount)
        If Trim$(processorReport) <> "" Then _
            ReleaseRecipeForProduction = ReleaseRecipeForProduction & "; " & processorReport
    End If
    Exit Function

FailRelease:
    ReleaseRecipeForProduction = "Release for Production failed: " & Err.Description
End Function

Private Function LatestRecipeVersionForReleaseProduction(ByVal recipeId As String) As String
    On Error GoTo CleanFail

    Dim designs As Variant
    Dim staged As Variant
    Dim warehouseId As String
    Dim r As Long
    Dim candidate As String
    Dim bestNumeric As Long
    Dim bestText As String

    designs = modOperationsPrimitiveBridge.ListDesigns("")
    AccumulateLatestRecipeVersionProduction designs, recipeId, bestNumeric, bestText

    warehouseId = CurrentProductionWarehouseId()
    staged = modRoleEventWriter.GetLocalStagedDesignIdentities(warehouseId)
    AccumulateLatestRecipeVersionProduction staged, recipeId, bestNumeric, bestText

    If bestNumeric > 0 Then
        LatestRecipeVersionForReleaseProduction = CStr(bestNumeric)
    Else
        LatestRecipeVersionForReleaseProduction = bestText
    End If
CleanFail:
End Function

Private Sub AccumulateLatestRecipeVersionProduction(ByVal rows As Variant, _
                                                     ByVal recipeId As String, _
                                                     ByRef bestNumeric As Long, _
                                                     ByRef bestText As String)
    Dim r As Long
    Dim candidate As String
    Dim numericVersion As Long

    If Not IsUsableProductionArray(rows) Then Exit Sub
    For r = LBound(rows, 1) To UBound(rows, 1)
        If RecipeIdsMatchProduction(rows(r, 1), recipeId) Then
            candidate = Trim$(NzStr(rows(r, 2)))
            If IsNumeric(candidate) Then
                numericVersion = CLng(CDbl(candidate))
                If numericVersion > bestNumeric Then bestNumeric = numericVersion
            ElseIf candidate <> "" Then
                bestText = candidate
            End If
        End If
    Next r
End Sub

Private Function DesignStatusForVersionProduction(ByVal recipeId As String, _
                                                   ByVal designVersion As String) As String
    On Error GoTo CleanFail

    Dim designs As Variant
    Dim r As Long

    designs = modOperationsPrimitiveBridge.ListDesigns("")
    If Not IsUsableProductionArray(designs) Then Exit Function
    For r = LBound(designs, 1) To UBound(designs, 1)
        If RecipeIdsMatchProduction(designs(r, 1), recipeId) _
           And StrComp(Trim$(NzStr(designs(r, 2))), Trim$(designVersion), vbTextCompare) = 0 Then
            DesignStatusForVersionProduction = Trim$(NzStr(designs(r, 6)))
        End If
    Next r
CleanFail:
End Function

Private Function LoadLegacyRecipeList() As Variant
    Dim syncReport As String
    Dim wbOps As Workbook
    Set wbOps = ResolveProductionWorkbook(, "Recipes")
    If Not LocalProductionRecipeRowsExist(wbOps) Then RefreshProductionRecipesFromRuntime wbOps, syncReport

    Dim wsRec As Worksheet: Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then Exit Function
    Dim lo As ListObject: Set lo = GetListObject(wsRec, "Recipes")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim cId As Long: cId = ColumnIndex(lo, "RECIPE_ID")
    Dim cName As Long: cName = ColumnIndex(lo, "RECIPE")
    Dim cDesc As Long: cDesc = ColumnIndex(lo, "DESCRIPTION")
    If cId = 0 Or cName = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim seenRows As Object: Set seenRows = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = lo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rid As String: rid = NzStr(arr(r, cId))
        Dim rname As String: rname = NzStr(arr(r, cName))
        If rid = "" Or rname = "" Then GoTo NextRow
        If Not dict.Exists(rid) Then
            Dim info(1 To 3) As Variant
            info(1) = rid
            info(2) = rname
            If cDesc > 0 Then info(3) = NzStr(arr(r, cDesc)) Else info(3) = ""
            dict.Add rid, info
        End If
NextRow:
    Next r

    If dict.count = 0 Then Exit Function
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 3)
    Dim i As Long: i = 1
    Dim key As Variant
    For Each key In dict.keys
        Dim infoArr As Variant
        infoArr = dict(key)
        result(i, 1) = infoArr(1)
        result(i, 2) = infoArr(2)
        result(i, 3) = infoArr(3)
        i = i + 1
    Next key
    LoadLegacyRecipeList = result
End Function

Private Function LoadLegacyRuntimeRecipeList() As Variant
    On Error GoTo CleanExit

    Dim warehouseId As String
    Dim rootPath As String
    Dim report As String
    Dim wbRuntime As Workbook
    Dim wsRuntime As Worksheet
    Dim loRuntime As ListObject
    Dim openedTransient As Boolean
    Dim cId As Long
    Dim cName As Long
    Dim cDesc As Long
    Dim arr As Variant
    Dim r As Long
    Dim recipeId As String
    Dim recipeName As String
    Dim recipes As Object
    Dim info As Variant
    Dim result() As Variant
    Dim key As Variant
    Dim i As Long

    If Not ResolveProductionRecipesStorageTarget(warehouseId, rootPath, report) Then Exit Function
    Set wbRuntime = OpenProductionRecipesWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbRuntime Is Nothing Then Exit Function
    Set wsRuntime = WorkbookSheetExists(wbRuntime, SHEET_RUNTIME_RECIPES)
    If wsRuntime Is Nothing Then GoTo CleanExit
    Set loRuntime = GetListObject(wsRuntime, TABLE_RUNTIME_RECIPES)
    If loRuntime Is Nothing Or loRuntime.DataBodyRange Is Nothing Then GoTo CleanExit

    cId = ColumnIndex(loRuntime, "RECIPE_ID")
    cName = ColumnIndex(loRuntime, "RECIPE")
    cDesc = ColumnIndex(loRuntime, "DESCRIPTION")
    If cId = 0 Or cName = 0 Then GoTo CleanExit

    Set recipes = CreateObject("Scripting.Dictionary")
    recipes.CompareMode = vbTextCompare
    arr = loRuntime.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        recipeId = CanonicalRecipeIdProduction(arr(r, cId))
        recipeName = Trim$(NzStr(arr(r, cName)))
        If recipeId <> "" And recipeName <> "" Then
            If cDesc > 0 Then
                recipes(recipeId) = Array(recipeName, NzStr(arr(r, cDesc)))
            Else
                recipes(recipeId) = Array(recipeName, vbNullString)
            End If
        End If
    Next r

    If recipes.Count = 0 Then GoTo CleanExit
    ReDim result(1 To recipes.Count, 1 To 3)
    i = 1
    For Each key In recipes.Keys
        info = recipes(key)
        result(i, 1) = CStr(key)
        result(i, 2) = info(0)
        result(i, 3) = info(1)
        i = i + 1
    Next key
    LoadLegacyRuntimeRecipeList = result

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveProduction wbRuntime
End Function

Private Function LoadPendingStagedRecipeList() As Variant
    On Error GoTo CleanFail

    Dim warehouseId As String
    Dim staged As Variant
    Dim recipes As Object
    Dim recipeId As String
    Dim recipeName As String
    Dim recipeDescription As String
    Dim result() As Variant
    Dim info As Variant
    Dim key As Variant
    Dim r As Long
    Dim i As Long

    warehouseId = CurrentProductionWarehouseId()
    staged = modRoleEventWriter.GetLocalStagedDesignIdentities(warehouseId)
    If Not IsUsableProductionArray(staged) Then Exit Function

    Set recipes = CreateObject("Scripting.Dictionary")
    recipes.CompareMode = vbTextCompare
    For r = LBound(staged, 1) To UBound(staged, 1)
        recipeId = CanonicalRecipeIdProduction(staged(r, 1))
        recipeName = vbNullString
        recipeDescription = vbNullString
        If UBound(staged, 2) >= 3 Then recipeName = Trim$(NzStr(staged(r, 3)))
        If UBound(staged, 2) >= 4 Then recipeDescription = NzStr(staged(r, 4))
        If recipeId <> "" And recipeName <> "" Then
            recipes(recipeId) = Array(recipeName, recipeDescription)
        End If
    Next r

    If recipes.Count = 0 Then Exit Function
    ReDim result(1 To recipes.Count, 1 To 3)
    i = 1
    For Each key In recipes.Keys
        info = recipes(key)
        result(i, 1) = CStr(key)
        result(i, 2) = info(0)
        result(i, 3) = info(1)
        i = i + 1
    Next key
    LoadPendingStagedRecipeList = result
CleanFail:
End Function

Private Function BuildUnifiedRecipeList(ByVal canonicalRecipes As Variant, _
                                        ByVal pendingRecipes As Variant, _
                                        ByVal legacyRecipes As Variant) As Variant
    Dim recipes As Object
    Dim result() As Variant
    Dim info As Variant
    Dim key As Variant
    Dim i As Long

    Set recipes = CreateObject("Scripting.Dictionary")
    recipes.CompareMode = vbTextCompare
    MergeRecipeListRowsProduction recipes, canonicalRecipes
    MergeRecipeListRowsProduction recipes, pendingRecipes
    MergeRecipeListRowsProduction recipes, legacyRecipes

    If recipes.Count = 0 Then Exit Function
    ReDim result(1 To recipes.Count, 1 To 3)
    i = 1
    For Each key In recipes.Keys
        info = recipes(key)
        result(i, 1) = CStr(key)
        result(i, 2) = info(0)
        result(i, 3) = info(1)
        i = i + 1
    Next key
    BuildUnifiedRecipeList = result
End Function

Private Sub MergeRecipeListRowsProduction(ByVal recipes As Object, ByVal sourceRows As Variant)
    Dim r As Long
    Dim recipeId As String
    Dim recipeName As String
    Dim recipeDescription As String

    If recipes Is Nothing Then Exit Sub
    If Not IsUsableProductionArray(sourceRows) Then Exit Sub
    For r = LBound(sourceRows, 1) To UBound(sourceRows, 1)
        recipeId = CanonicalRecipeIdProduction(sourceRows(r, 1))
        recipeName = Trim$(NzStr(sourceRows(r, 2)))
        recipeDescription = vbNullString
        If UBound(sourceRows, 2) >= 3 Then recipeDescription = NzStr(sourceRows(r, 3))
        If recipeId <> "" And recipeName <> "" Then
            ' Sources are merged in authority order. The first identity wins.
            If Not recipes.Exists(recipeId) Then _
                recipes.Add recipeId, Array(recipeName, recipeDescription)
        End If
    Next r
End Sub

'@TestOnlyBegin
Public Function TestUnifiedRecipeList(ByVal canonicalRecipes As Variant, _
                                      ByVal pendingRecipes As Variant, _
                                      ByVal legacyRecipes As Variant) As Variant
    TestUnifiedRecipeList = BuildUnifiedRecipeList(canonicalRecipes, pendingRecipes, legacyRecipes)
End Function
'@TestOnlyEnd

Public Function GenerateRecipeId(Optional ByVal preferredWb As Workbook = Nothing) As String
    Const MAX_BASE36_RECIPE_ID As Long = 46655 ' ZZZ

    Dim used As Object
    Set used = CollectUsedRecipeIdsProduction(preferredWb)

    Dim candidateValue As Long
    For candidateValue = 1 To MAX_BASE36_RECIPE_ID
        Dim candidateId As String
        candidateId = ToBase36RecipeId(candidateValue)
        If Not used.Exists(candidateId) Then
            GenerateRecipeId = candidateId
            Exit Function
        End If
    Next candidateValue

    GenerateRecipeId = ToBase36RecipeId(CLng(Timer) Mod MAX_BASE36_RECIPE_ID)
End Function

Private Function CollectUsedRecipeIdsProduction(Optional ByVal preferredWb As Workbook = Nothing) As Object
    Dim used As Object
    Dim designs As Variant
    Dim stagedDesigns As Variant
    Dim warehouseId As String
    Dim wsRec As Worksheet
    Dim lo As ListObject
    Dim cId As Long
    Dim arr As Variant
    Dim r As Long
    Dim existingId As String

    Set used = CreateObject("Scripting.Dictionary")
    used.CompareMode = vbTextCompare

    If ProductionDesignsEnabled() Then
        designs = modOperationsPrimitiveBridge.ListDesigns("")
        AddUsedRecipeIdsFromDesignRows used, designs
    End If
    AddUsedRecipeIdsFromLegacyRuntime used

    warehouseId = CurrentProductionWarehouseId()
    stagedDesigns = modRoleEventWriter.GetLocalStagedDesignIdentities(warehouseId)
    AddUsedRecipeIdsFromDesignRows used, stagedDesigns

    If preferredWb Is Nothing Then
        Set wsRec = SheetExists("Recipes")
    Else
        Set wsRec = WorkbookSheetExists(preferredWb, "Recipes")
    End If
    If Not wsRec Is Nothing Then
        Set lo = GetListObject(wsRec, "Recipes")
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                cId = ColumnIndex(lo, "RECIPE_ID")
                If cId > 0 Then
                    arr = lo.DataBodyRange.value
                    For r = 1 To UBound(arr, 1)
                        existingId = CanonicalRecipeIdProduction(arr(r, cId))
                        If existingId <> "" Then used(existingId) = True
                    Next r
                End If
            End If
        End If
    End If

    Set CollectUsedRecipeIdsProduction = used
End Function

Private Sub AddUsedRecipeIdsFromLegacyRuntime(ByVal used As Object)
    On Error GoTo CleanExit

    Dim warehouseId As String
    Dim rootPath As String
    Dim report As String
    Dim wbRuntime As Workbook
    Dim wsRuntime As Worksheet
    Dim loRuntime As ListObject
    Dim openedTransient As Boolean
    Dim cId As Long
    Dim arr As Variant
    Dim r As Long
    Dim recipeId As String

    If used Is Nothing Then Exit Sub
    If Not ResolveProductionRecipesStorageTarget(warehouseId, rootPath, report) Then Exit Sub
    Set wbRuntime = OpenProductionRecipesWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbRuntime Is Nothing Then Exit Sub
    Set wsRuntime = WorkbookSheetExists(wbRuntime, SHEET_RUNTIME_RECIPES)
    If wsRuntime Is Nothing Then GoTo CleanExit
    Set loRuntime = GetListObject(wsRuntime, TABLE_RUNTIME_RECIPES)
    If loRuntime Is Nothing Or loRuntime.DataBodyRange Is Nothing Then GoTo CleanExit
    cId = ColumnIndex(loRuntime, "RECIPE_ID")
    If cId <= 0 Then GoTo CleanExit
    arr = loRuntime.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        recipeId = CanonicalRecipeIdProduction(arr(r, cId))
        If recipeId <> "" Then used(recipeId) = True
    Next r

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveProduction wbRuntime
End Sub

Private Sub AddUsedRecipeIdsFromDesignRows(ByVal used As Object, ByVal designs As Variant)
    Dim r As Long
    Dim recipeId As String

    If used Is Nothing Then Exit Sub
    If Not IsUsableProductionArray(designs) Then Exit Sub
    For r = LBound(designs, 1) To UBound(designs, 1)
        recipeId = CanonicalRecipeIdProduction(designs(r, 1))
        If recipeId <> "" Then used(recipeId) = True
    Next r
End Sub

Private Function CanonicalRecipeIdProduction(ByVal valueIn As Variant) As String
    Dim textValue As String
    Dim numericValue As Long

    textValue = UCase$(Trim$(NzStr(valueIn)))
    If textValue = "" Then Exit Function
    If Len(textValue) <= 3 And IsNumeric(textValue) Then
        numericValue = CLng(CDbl(textValue))
        If numericValue >= 0 And numericValue <= 999 Then
            CanonicalRecipeIdProduction = Right$("000" & CStr(numericValue), 3)
            Exit Function
        End If
    End If
    CanonicalRecipeIdProduction = textValue
End Function

Private Function RecipeIdsMatchProduction(ByVal leftId As Variant, ByVal rightId As Variant) As Boolean
    RecipeIdsMatchProduction = (StrComp(CanonicalRecipeIdProduction(leftId), _
                                        CanonicalRecipeIdProduction(rightId), vbTextCompare) = 0)
End Function

Private Function RecipeIdExistsProduction(ByVal recipeId As String, _
                                          Optional ByVal preferredWb As Workbook = Nothing) As Boolean
    Dim used As Object

    recipeId = CanonicalRecipeIdProduction(recipeId)
    If recipeId = "" Then Exit Function
    Set used = CollectUsedRecipeIdsProduction(preferredWb)
    If used Is Nothing Then Exit Function
    RecipeIdExistsProduction = used.Exists(recipeId)
End Function

'@TestOnlyBegin
Public Function TestNextBase36RecipeId(ByVal usedIds As Variant) As String
    TestNextBase36RecipeId = NextBase36Identifier(usedIds)
End Function
'@TestOnlyEnd

Public Function NextBase36Identifier(ByVal usedIds As Variant) As String
    Dim used As Object
    Dim i As Long
    Dim candidateValue As Long
    Dim candidateId As String

    Set used = CreateObject("Scripting.Dictionary")
    If IsArray(usedIds) Then
        For i = LBound(usedIds) To UBound(usedIds)
            used(CanonicalRecipeIdProduction(usedIds(i))) = True
        Next i
    End If
    For candidateValue = 1 To 46655
        candidateId = ToBase36RecipeId(candidateValue)
        If Not used.Exists(candidateId) Then
            NextBase36Identifier = candidateId
            Exit Function
        End If
    Next candidateValue
End Function

Public Function IsBase36Identifier(ByVal valueText As String) As Boolean
    Dim index As Long
    Dim currentChar As String

    valueText = UCase$(Trim$(valueText))
    If Len(valueText) <> 3 Or valueText = "000" Then Exit Function
    For index = 1 To 3
        currentChar = Mid$(valueText, index, 1)
        If InStr(1, "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", currentChar, vbBinaryCompare) = 0 Then Exit Function
    Next index
    IsBase36Identifier = True
End Function

Public Function GenerateRecipeIdForCurrentWorkbook() As String
    GenerateRecipeIdForCurrentWorkbook = GenerateRecipeId()
End Function

Public Function GenerateRecipeIdForWorkbookName(ByVal workbookName As String) As String
    Dim wb As Workbook

    If Trim$(workbookName) <> "" Then
        On Error Resume Next
        Set wb = Application.Workbooks(workbookName)
        On Error GoTo 0
    End If
    GenerateRecipeIdForWorkbookName = GenerateRecipeId(wb)
End Function

Private Function ToBase36RecipeId(ByVal value As Long) As String
    Const DIGITS As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"

    If value < 0 Then value = 0
    value = value Mod 46656

    Dim result As String
    Dim i As Long
    For i = 1 To 3
        result = Mid$(DIGITS, (value Mod 36) + 1, 1) & result
        value = value \ 36
    Next i

    ToBase36RecipeId = result
End Function

' ===== System 3: Recipe Chooser =====
Public Sub LoadRecipeChooser(ByVal recipeId As String)
    On Error GoTo ErrHandler
    If Trim$(recipeId) = "" Then Exit Sub

    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim syncReport As String
    Dim stagingWb As Workbook
    Dim wsRec As Worksheet
    If ProductionDesignsEnabled() Then
        Set stagingWb = BuildReleasedDesignRecipeStagingWorkbook(recipeId, syncReport)
        If stagingWb Is Nothing Then
            Dim unavailableVersion As String
            Dim unavailableName As String
            Dim unavailableDescription As String
            Dim unavailableStatus As String
            If FindLatestDesignSummaryProduction(recipeId, "", unavailableVersion, _
                                                  unavailableName, unavailableDescription, _
                                                  unavailableStatus) Then
                syncReport = "Recipe " & ValueOrPlaceholderProduction(unavailableName, recipeId) & _
                             " (Design ID " & recipeId & ", version " & unavailableVersion & _
                             ") is " & ValueOrPlaceholderProduction(unavailableStatus, "not released") & _
                             "." & vbCrLf & vbCrLf & _
                             "Production can load only RELEASED Designs Domain versions. " & _
                             "Use Admin > Design Lifecycle to release this design."
            Else
                syncReport = "Recipe Design ID " & recipeId & _
                             " is not present in the Designs Domain." & vbCrLf & vbCrLf & _
                             "For an existing legacy recipe, use Admin > Design Lifecycle > " & _
                             "Import Legacy Recipes, then release the imported version."
            End If
            MsgBox syncReport, vbExclamation, "Production Designs"
            Exit Sub
        End If
        Set wsRec = WorkbookSheetExists(stagingWb, "Recipes")
    Else
        If Not LocalProductionRecipeRowsExist(wsProd.Parent, recipeId) Then
            RefreshProductionRecipesFromRuntime wsProd.Parent, syncReport
        End If
        Set wsRec = SheetExists("Recipes")
    End If
    If wsRec Is Nothing Then
        MsgBox "Recipes sheet not found.", vbCritical
        GoTo CleanExit
    End If

    Dim loChooser As ListObject
    Set loChooser = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If loChooser Is Nothing Then
        MsgBox "RC_RecipeChoose table not found on Production sheet.", vbExclamation
        GoTo CleanExit
    End If

    EnsureTableHasRow loChooser

    Dim recipeName As String
    Dim recipeDesc As String
    Dim recipeDept As String
    GetRecipeSummary wsRec, recipeId, recipeName, recipeDesc, recipeDept

    Dim cRec As Long: cRec = ColumnIndex(loChooser, "RECIPE")
    If cRec = 0 Then cRec = ColumnIndex(loChooser, "RECIPE_NAME")
    Dim cRecId As Long: cRecId = ColumnIndex(loChooser, "RECIPE_ID")
    Dim cDesc As Long: cDesc = ColumnIndex(loChooser, "DESCRIPTION")
    Dim cDept As Long: cDept = ColumnIndex(loChooser, "DEPARTMENT")
    Dim cProc As Long: cProc = ColumnIndex(loChooser, "PROCESS")

    If Not loChooser.DataBodyRange Is Nothing Then
        If cRec > 0 Then loChooser.DataBodyRange.Cells(1, cRec).value = recipeName
        If cRecId > 0 Then loChooser.DataBodyRange.Cells(1, cRecId).value = recipeId
        If cDesc > 0 Then loChooser.DataBodyRange.Cells(1, cDesc).value = recipeDesc
        If cDept > 0 Then loChooser.DataBodyRange.Cells(1, cDept).value = recipeDept
        If cProc > 0 Then loChooser.DataBodyRange.Cells(1, cProc).value = ""
    End If

    Dim chooserStyle As String
    Dim loStyle As ListObject
    Set loStyle = GetListObject(wsProd, TABLE_RECIPE_CHOOSER_GENERATED)
    If Not loStyle Is Nothing Then
        On Error Resume Next
        chooserStyle = loStyle.TableStyle
        On Error GoTo 0
    End If

    Dim paletteStyle As String
    Dim loPalette As ListObject
    Set loPalette = GetListObject(wsProd, TABLE_INV_PALETTE_GENERATED)
    If Not loPalette Is Nothing Then
        On Error Resume Next
        paletteStyle = loPalette.TableStyle
        On Error GoTo 0
    End If

    DeleteRecipeChooserProcessTables wsProd
    DeleteInventoryPaletteTables wsProd

    Dim procTables As Collection
    Set procTables = BuildRecipeChooserProcessTablesFromRecipes(recipeId, wsProd, wsRec, chooserStyle)
    RenderProcessSelectorCheckboxes wsProd, procTables
    BuildPaletteTablesForRecipeChooser recipeId, wsProd, wsRec, procTables, paletteStyle
    RenderPaletteKeepCheckboxes wsProd
    ApplyProductionOutputTemplates recipeId, wsProd
    RenderOutputRowCheckboxes wsProd

CleanExit:
    On Error Resume Next
    If Not stagingWb Is Nothing Then stagingWb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub
ErrHandler:
    MsgBox "Load Recipe Chooser failed: " & Err.description, vbCritical
    Resume CleanExit
End Sub

Private Function FindListObjectByNameOrHeaders(ws As Worksheet, tableName As String, headers As Variant) As ListObject
    Dim lo As ListObject
    Set lo = GetListObject(ws, tableName)
    If Not lo Is Nothing Then
        Set FindListObjectByNameOrHeaders = lo
        Exit Function
    End If
    For Each lo In ws.ListObjects
        If ListObjectHasHeaders(lo, headers) Then
            Set FindListObjectByNameOrHeaders = lo
            Exit Function
        End If
    Next lo
End Function

Private Function ListObjectHasHeaders(lo As ListObject, headers As Variant) As Boolean
    If lo Is Nothing Then Exit Function
    If lo.HeaderRowRange Is Nothing Then Exit Function
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If ColumnIndex(lo, CStr(headers(i))) = 0 Then Exit Function
    Next i
    ListObjectHasHeaders = True
End Function

Private Function TableColumnCount(lo As ListObject) As Long
    If lo Is Nothing Then Exit Function
    If lo.HeaderRowRange Is Nothing Then Exit Function
    TableColumnCount = lo.HeaderRowRange.Columns.count
End Function

Private Sub ExpandSpanForTable(lo As ListObject, ByRef firstCol As Long, ByRef lastCol As Long)
    If lo Is Nothing Then Exit Sub
    If lo.HeaderRowRange Is Nothing Then Exit Sub
    Dim startCol As Long
    Dim endCol As Long
    If Not TableEffectiveSpan(lo, startCol, endCol) Then Exit Sub
    If firstCol = 0 Or startCol < firstCol Then firstCol = startCol
    If endCol > lastCol Then lastCol = endCol
End Sub

Private Function TableEffectiveSpan(lo As ListObject, ByRef startCol As Long, ByRef endCol As Long) As Boolean
    TableEffectiveSpan = False
    If lo Is Nothing Then Exit Function
    If lo.HeaderRowRange Is Nothing Then Exit Function

    Dim hdr As Range: Set hdr = lo.HeaderRowRange
    startCol = hdr.Column

    Dim lastIdx As Long
    Dim i As Long
    For i = hdr.Columns.count To 1 Step -1
        Dim val As String
        val = Trim$(CStr(hdr.Cells(1, i).value))
        If val <> "" Then
            lastIdx = i
            Exit For
        End If
    Next i
    If lastIdx = 0 Then lastIdx = hdr.Columns.count
    endCol = startCol + lastIdx - 1
    TableEffectiveSpan = (endCol >= startCol)
End Function

Private Function ResolveListObject(ws As Worksheet, tableName As String) As ListObject
    Select Case tableName
        Case "RB_AddRecipeName", "RecipeBuilder", "IP_ChooseRecipe", "IP_ChooseIngredient", _
             "IP_ChooseItem", "RC_RecipeChoose", "RecipeChooser_generated", _
             "InventoryPalette_generated", "ProductionOutput", "Prod_invSys_Check"
            Set ResolveListObject = GetListObject(ws, tableName)
        Case Else
            Set ResolveListObject = GetListObject(ws, tableName)
    End Select
End Function

Private Sub GetSystemBounds(ws As Worksheet, ByRef startCols() As Long, ByRef endCols() As Long, ByRef topRows() As Long, ByRef bottomRows() As Long)
    Dim i As Long
    ReDim startCols(LBound(mSystemGroupNames) To UBound(mSystemGroupNames))
    ReDim endCols(LBound(mSystemGroupNames) To UBound(mSystemGroupNames))
    ReDim topRows(LBound(mSystemGroupNames) To UBound(mSystemGroupNames))
    ReDim bottomRows(LBound(mSystemGroupNames) To UBound(mSystemGroupNames))

    Dim maxEnd() As Long
    ReDim maxEnd(LBound(mSystemGroupNames) To UBound(mSystemGroupNames))

    For i = LBound(mSystemGroupNames) To UBound(mSystemGroupNames)
        Dim tablesArr As Variant
        tablesArr = mSystemGroupTables(i)
        Dim j As Long
        For j = LBound(tablesArr) To UBound(tablesArr)
            Dim lo As ListObject
            Set lo = ResolveListObject(ws, CStr(tablesArr(j)))
            If Not lo Is Nothing Then
                Dim sCol As Long
                Dim eCol As Long
                Dim rTop As Long
                Dim rBottom As Long
                If TableEffectiveSpan(lo, sCol, eCol) Then
                    rTop = lo.Range.row
                    rBottom = lo.Range.row + lo.Range.rows.count - 1
                    If startCols(i) = 0 Or sCol < startCols(i) Then startCols(i) = sCol
                    If eCol > maxEnd(i) Then maxEnd(i) = eCol
                    If topRows(i) = 0 Or rTop < topRows(i) Then topRows(i) = rTop
                    If rBottom > bottomRows(i) Then bottomRows(i) = rBottom
                End If
            End If
        Next j
    Next i

    ' Define end bounds by the next group's start (keeps bands discrete).
    For i = LBound(mSystemGroupNames) To UBound(mSystemGroupNames)
        If startCols(i) = 0 Then GoTo NextGroup
        Dim nextStart As Long
        nextStart = 0
        Dim k As Long
        For k = i + 1 To UBound(mSystemGroupNames)
            If startCols(k) > 0 Then
                nextStart = startCols(k)
                Exit For
            End If
        Next k
        If nextStart > 0 Then
            endCols(i) = nextStart - 1
        Else
            endCols(i) = maxEnd(i)
            ' Rightmost system: extend to include any checkbox shapes to the right.
            Dim maxChkCol As Long
            maxChkCol = MaxCheckboxColumn(ws, startCols(i))
            If maxChkCol > endCols(i) Then endCols(i) = maxChkCol
        End If
NextGroup:
    Next i
End Sub

Private Function MaxCheckboxColumn(ws As Worksheet, startCol As Long) As Long
    If ws Is Nothing Then Exit Function
    If startCol = 0 Then Exit Function
    Dim shp As Shape
    For Each shp In ws.shapes
        Dim isCheckbox As Boolean
        On Error Resume Next
        If shp.Type = SHAPE_TYPE_FORM_CONTROL Then
            If shp.FormControlType = xlCheckBox Then isCheckbox = True
        End If
        If Not isCheckbox Then
            If LCase$(shp.Name) Like "check box*" Then isCheckbox = True
        End If
        Dim c As Long
        c = shp.TopLeftCell.Column
        On Error GoTo 0
        If isCheckbox And c >= startCol Then
            If c > MaxCheckboxColumn Then MaxCheckboxColumn = c
        End If
    Next shp

    Dim ole As OLEObject
    For Each ole In ws.OLEObjects
        Dim isChk As Boolean
        On Error Resume Next
        Dim tName As String
        tName = TypeName(ole.Object)
        If LCase$(tName) Like "*checkbox*" Then isChk = True
        Dim cOle As Long
        cOle = ole.TopLeftCell.Column
        On Error GoTo 0
        If isChk And cOle >= startCol Then
            If cOle > MaxCheckboxColumn Then MaxCheckboxColumn = cOle
        End If
    Next ole
End Function

Private Function IsSystemVisible(ws As Worksheet, startCol As Long, endCol As Long) As Boolean
    If ws Is Nothing Then Exit Function
    If startCol = 0 Or endCol = 0 Then Exit Function
    Dim c As Long
    For c = startCol To endCol
        If Not ws.Columns(c).EntireColumn.Hidden Then
            IsSystemVisible = True
            Exit Function
        End If
    Next c
End Function

Public Function ColumnIndex(lo As ListObject, colName As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(Trim$(lc.Name), Trim$(colName), vbTextCompare) = 0 Then
            ColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    ColumnIndex = 0
End Function

Private Function ColumnIndexLoose(lo As ListObject, ParamArray names() As Variant) As Long
    If lo Is Nothing Then Exit Function
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        Dim hdr As String
        hdr = NormalizeHeaderKey(NzStr(lc.Name))
        Dim i As Long
        For i = LBound(names) To UBound(names)
            If hdr = NormalizeHeaderKey(CStr(names(i))) Then
                ColumnIndexLoose = lc.Index
                Exit Function
            End If
        Next i
    Next lc
End Function

Private Function NormalizeHeaderKey(ByVal v As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(v)
        ch = Mid$(v, i, 1)
        If ch Like "[A-Za-z0-9]" Then out = out & UCase$(ch)
    Next i
    NormalizeHeaderKey = out
End Function

Private Function NormalizeSystemKey(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then Exit Function
    NormalizeSystemKey = Trim$(CStr(v))
End Function

' ===== button scaffolding =====
Private Sub EnsureProductionButtons()
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub

    DeleteLegacyProductionButtons ws
End Sub

Private Sub RefreshProductionUiAccess(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    modOperationsPrimitiveBridge.ApplyShapeCapability _
        ws.Parent.Name, ws.Name, BTN_TO_MADE, "PROD_POST"
    modOperationsPrimitiveBridge.ApplyShapeCapability _
        ws.Parent.Name, ws.Name, BTN_TO_TOTALINV, "PROD_POST"
End Sub

Private Sub DeleteLegacyProductionButtons(ByVal ws As Worksheet)
    DeleteShapeIfExists ws, BTN_HIDE_SYSTEM
    DeleteShapeIfExists ws, BTN_SHOW_SYSTEM
    DeleteShapeIfExists ws, "BTN_TOGGLE_RECIPE_BUILDER"
    DeleteShapeIfExists ws, "BTN_TOGGLE_PALETTE_BUILDER"
    DeleteShapeIfExists ws, "BTN_TOGGLE_PRODUCTION"
    DeleteShapeIfExists ws, BTN_LOAD_RECIPE
    DeleteShapeIfExists ws, BTN_SAVE_RECIPE
    DeleteShapeIfExists ws, BTN_SAVE_FORMULAS
    DeleteShapeIfExists ws, BTN_BUILD_RECIPE_TABLES
    DeleteShapeIfExists ws, BTN_REMOVE_RECIPE_TABLES
    DeleteShapeIfExists ws, BTN_CLEAR_RECIPE_BUILDER
    DeleteShapeIfExists ws, BTN_CLEAR_RECIPE_CHOOSER
    DeleteShapeIfExists ws, BTN_CLEAR_PALETTE_BUILDER
    DeleteShapeIfExists ws, BTN_SAVE_PALETTE
    DeleteShapeIfExists ws, BTN_TO_USED
    DeleteShapeIfExists ws, BTN_TO_MADE
    DeleteShapeIfExists ws, BTN_TO_TOTALINV
    DeleteShapeIfExists ws, BTN_NEXT_BATCH
    DeleteShapeIfExists ws, BTN_PRINT_CODES
End Sub

Private Sub ClearExcelClipboardStateProduction()
    On Error Resume Next
    If Application.CutCopyMode <> False Then Application.CutCopyMode = False
    On Error GoTo 0
End Sub

Private Sub EnsureSystemGroups()
    If mSystemGroupsInit Then Exit Sub
    ' System 1: Recipe List Builder.
    mSystemGroupNames(1) = "RecipeListBuilder"
    mSystemGroupTables(1) = Array("RecipeBuilder", "RB_AddRecipeName")

    mSystemGroupNames(2) = "InventoryPaletteBuilder"
    mSystemGroupTables(2) = Array("IP_ChooseIngredient", "IP_ChooseItem", "IP_ChooseRecipe")

    mSystemGroupNames(3) = "RecipeChooser"
    mSystemGroupTables(3) = Array("RC_RecipeChoose", "RecipeChooser_generated")

    mSystemGroupNames(4) = "ProductionInputOutput"
    mSystemGroupTables(4) = Array("InventoryPalette_generated", "ProductionOutput", "Prod_invSys_Check")

    Set mHiddenSystems = New Collection
    mSystemGroupsInit = True
End Sub

' ===== show/hide system bands =====
Public Sub BtnHideSystem()
    EnsureSystemGroups
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub

    Dim starts() As Long, ends() As Long, tops() As Long, bottoms() As Long
    GetSystemBounds ws, starts, ends, tops, bottoms

    Dim i As Long, nearestIdx As Long, bestStart As Long
    For i = LBound(mSystemGroupNames) To UBound(mSystemGroupNames)
        If starts(i) > 0 And ends(i) > 0 Then
            If IsSystemVisible(ws, starts(i), ends(i)) Then
                If bestStart = 0 Or starts(i) < bestStart Then
                    bestStart = starts(i)
                    nearestIdx = i
                End If
            End If
        End If
    Next i

    If nearestIdx = 0 Then Exit Sub
    ws.Range(ws.Columns(starts(nearestIdx)), ws.Columns(ends(nearestIdx))).EntireColumn.Hidden = True
    HideGroupShapes ws, starts(nearestIdx), ends(nearestIdx), tops(nearestIdx), bottoms(nearestIdx), True
    mHiddenSystems.Add nearestIdx
End Sub

Public Sub BtnShowSystem()
    EnsureSystemGroups
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub
    Dim starts() As Long, ends() As Long, tops() As Long, bottoms() As Long
    GetSystemBounds ws, starts, ends, tops, bottoms

    Dim idx As Long
    If Not mHiddenSystems Is Nothing Then
        If mHiddenSystems.count > 0 Then
            idx = CLng(mHiddenSystems(mHiddenSystems.count))
            mHiddenSystems.Remove mHiddenSystems.count
        End If
    End If
    If idx = 0 Then
        ' Fallback: show rightmost hidden system.
        Dim i As Long, bestStart As Long
        For i = LBound(mSystemGroupNames) To UBound(mSystemGroupNames)
            If starts(i) > 0 And ends(i) > 0 Then
                If Not IsSystemVisible(ws, starts(i), ends(i)) Then
                    If starts(i) > bestStart Then
                        bestStart = starts(i)
                        idx = i
                    End If
                End If
            End If
        Next i
        If idx = 0 Then Exit Sub
    End If

    If starts(idx) = 0 Or ends(idx) = 0 Then Exit Sub
    ws.Range(ws.Columns(starts(idx)), ws.Columns(ends(idx))).EntireColumn.Hidden = False
    HideGroupShapes ws, starts(idx), ends(idx), tops(idx), bottoms(idx), False
End Sub

Private Sub EnsureButtonCustom(ws As Worksheet, shapeName As String, caption As String, onActionMacro As String, leftPos As Double, topPos As Double, Optional widthPts As Double = 118)
    Const BTN_HEIGHT As Double = 20
    If widthPts < 20 Then widthPts = 118
    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then
        Set shp = ws.shapes.AddFormControl(xlButtonControl, leftPos, topPos, widthPts, BTN_HEIGHT)
        shp.Name = shapeName
        shp.TextFrame.Characters.text = caption
        shp.OnAction = onActionMacro
    Else
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthPts
        shp.Height = BTN_HEIGHT
        shp.TextFrame.Characters.text = caption
        shp.OnAction = onActionMacro
    End If
End Sub

Private Sub DeleteShapeIfExists(ws As Worksheet, shapeName As String)
    On Error Resume Next
    ws.shapes(shapeName).Delete
    On Error GoTo 0
End Sub

Private Sub HideGroupShapes(ws As Worksheet, startCol As Long, endCol As Long, topRow As Long, bottomRow As Long, hideIt As Boolean)
    If ws Is Nothing Then Exit Sub
    If startCol = 0 Or endCol = 0 Then Exit Sub
    Dim endColAdj As Long
    endColAdj = endCol + 6 ' allow checkboxes just right of the table
    Dim shp As Shape
    For Each shp In ws.shapes
        Dim c As Long
        Dim r As Long
        On Error Resume Next
        c = shp.TopLeftCell.Column
        r = shp.TopLeftCell.row
        On Error GoTo 0
        If c >= startCol And c <= endColAdj Then
            shp.Visible = IIf(hideIt, SHAPE_VISIBLE_FALSE, SHAPE_VISIBLE_TRUE)
        End If
    Next shp

    Dim ole As OLEObject
    For Each ole In ws.OLEObjects
        Dim isChk As Boolean
        On Error Resume Next
        Dim tName As String
        tName = TypeName(ole.Object)
        If LCase$(tName) Like "*checkbox*" Then isChk = True
        Dim cOle As Long
        cOle = ole.TopLeftCell.Column
        On Error GoTo 0
        If isChk Then
            If cOle >= startCol And cOle <= endColAdj Then
                ole.Visible = Not hideIt
            End If
        End If
    Next ole
End Sub

' ===== button handlers (stubs for now) =====
' System 1: Recipe List Builder actions (Load/Save/Add/Remove/Clear).
Public Sub BtnLoadRecipe()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim loHeader As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    If loHeader Is Nothing Then
        MsgBox "Recipe Builder header table not found on Production sheet.", vbExclamation
        Exit Sub
    End If
    Dim targetCell As Range
    Set targetCell = GetHeaderDataCell(loHeader, "RECIPE_NAME")
    If targetCell Is Nothing Then
        MsgBox "Recipe Builder header missing RECIPE_NAME column.", vbCritical
        Exit Sub
    End If
    If mRecipePicker Is Nothing Then Set mRecipePicker = CreateDynItemSearch()
    mRecipePicker.UseRoleProfile "production"
    mRecipePicker.ShowForRecipeCell targetCell
End Sub

Public Sub BtnSaveRecipe()
    SaveRecipeToRecipes
End Sub

Public Sub BtnSaveFormulas()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim recipeId As String
    recipeId = ResolveActiveRecipeId(wsProd, True)
    If recipeId = "" Then
        MsgBox "Select or load a RECIPE before saving formulas.", vbExclamation
        Exit Sub
    End If

    Dim saved As Long
    saved = SaveFormulaTemplatesForRecipe(recipeId, wsProd)
    MsgBox "Saved formulas: " & saved & ".", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Save Formulas failed: " & Err.description, vbCritical
End Sub

Public Sub BtnBuildRecipeProcessTables()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim loHeader As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    Dim procTables As Collection
    Set procTables = GetRecipeBuilderProcessTables(wsProd)
    Dim recipeId As String
    If Not loHeader Is Nothing Then
        Dim idCell As Range: Set idCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
        If Not idCell Is Nothing Then recipeId = NzStr(idCell.value)
    End If
    Dim builtCount As Long
    If procTables.count = 0 Then
        builtCount = BuildRecipeProcessTablesFromLines(recipeId, True)
    End If
    If builtCount = 0 Then
        Dim newLo As ListObject
        Set newLo = CreateRecipeProcessTable(wsProd, "", 1)
        If newLo Is Nothing Then
            MsgBox "No PROCESS rows found to build process tables.", vbInformation
        Else
            FocusRecipeProcessTable newLo
            MsgBox "Created process table '" & newLo.Name & "'.", vbInformation
        End If
    End If
End Sub

Public Sub BtnRemoveRecipeProcessTables()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim sel As Range
    On Error Resume Next
    Set sel = Application.Selection
    On Error GoTo 0
    If sel Is Nothing Then
        MsgBox "Select one or more Recipe Process tables to remove.", vbInformation
        Exit Sub
    End If

    Dim targets As Object: Set targets = CreateObject("Scripting.Dictionary")
    Dim lo As ListObject
    For Each lo In wsProd.ListObjects
        If IsRecipeProcessTable(lo) Then
            If Not Intersect(lo.Range, sel) Is Nothing Then
                targets(lo.Name) = lo.Range.Address
            End If
        End If
    Next lo

    If targets.count = 0 Then
        MsgBox "No Recipe Process tables selected.", vbInformation
        Exit Sub
    End If

    Dim key As Variant
    For Each key In targets.keys
        On Error Resume Next
        wsProd.ListObjects(CStr(key)).Delete
        wsProd.Range(CStr(targets(key))).Clear
        On Error GoTo 0
    Next key

    MsgBox "Removed " & targets.count & " Recipe Process table(s).", vbInformation
End Sub

' System 1: Recipe List Builder actions (Load/Save/Add/Remove).
Public Sub BtnClearRecipeBuilder()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim loHeader As ListObject
    Dim loLines As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    Set loLines = GetRecipeBuilderLinesTable(wsProd, loHeader)

    DeleteRecipeProcessTables wsProd

    If Not loLines Is Nothing Then
        RemoveRecipeBuilderLinesTable loLines
    End If

    If Not loHeader Is Nothing Then
        ClearListObjectData loHeader
    End If

    MsgBox "Recipe List Builder cleared.", vbInformation
End Sub

' System 2+: Inventory Palette / Production actions.

Public Sub BtnSavePalette()
    SaveIngredientPalette
End Sub

Public Function GetPaletteSaveDiagnostic() As String
    On Error GoTo ErrHandler

    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then
        GetPaletteSaveDiagnostic = "ProductionSheet=missing"
        Exit Function
    End If

    Dim wbProd As Workbook: Set wbProd = wsProd.Parent
    Dim wsPal As Worksheet
    Dim wsRec As Worksheet
    Dim loRecipe As ListObject
    Dim loIng As ListObject
    Dim loItems As ListObject
    Dim loPal As ListObject
    Dim recipeId As String
    Dim ingredientId As String
    Dim itemCount As Long
    Dim paletteCount As Long
    Dim firstItem As String
    Dim firstPalRecipe As String

    Set loRecipe = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseRecipe", Array("RECIPE_NAME", "RECIPE_ID"))
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))
    recipeId = GetPaletteRecipeId()
    ingredientId = GetPaletteIngredientId()

    Set wsRec = WorkbookSheetExists(wbProd, "Recipes")
    Set wsPal = WorkbookSheetExists(wbProd, "IngredientPalette")
    If wsPal Is Nothing Then Set wsPal = WorkbookSheetExists(wbProd, "IngredientsPalette")
    If Not wsPal Is Nothing Then
        Set loPal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "ITEM"))
    End If

    If Not loItems Is Nothing Then
        itemCount = ListObjectRowCount(loItems)
        firstItem = FirstNonEmptyColumnValue(loItems, "ITEMS")
        If firstItem = "" Then firstItem = FirstNonEmptyColumnValue(loItems, "ITEM")
    End If

    If Not loPal Is Nothing Then
        paletteCount = ListObjectRowCount(loPal)
        firstPalRecipe = FirstNonEmptyColumnValue(loPal, "RECIPE_ID")
    End If

    GetPaletteSaveDiagnostic = "ProdWb=" & wbProd.Name & _
        "; RecipesSheet=" & IIf(wsRec Is Nothing, "missing", wsRec.Name) & _
        "; PaletteSheet=" & IIf(wsPal Is Nothing, "missing", wsPal.Name) & _
        "; RecipeId=" & recipeId & _
        "; IngredientId=" & ingredientId & _
        "; ChooseRecipeRows=" & ListObjectRowCount(loRecipe) & _
        "; ChooseIngredientRows=" & ListObjectRowCount(loIng) & _
        "; ChooseItemRows=" & itemCount & _
        "; FirstItem=" & firstItem & _
        "; PaletteRows=" & paletteCount & _
        "; FirstPaletteRecipe=" & firstPalRecipe
    Exit Function

ErrHandler:
    GetPaletteSaveDiagnostic = "DiagnosticError=" & Err.Number & ":" & Err.Description
End Function

Public Sub BtnClearPaletteBuilder()
    ClearInventoryPaletteBuilder
End Sub

Public Sub BtnClearRecipeChooser()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim loChooser As ListObject
    Set loChooser = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If Not loChooser Is Nothing Then
        EnsureTableHasRow loChooser
        If Not loChooser.DataBodyRange Is Nothing Then
            loChooser.DataBodyRange.ClearContents
        End If
    End If

    DeleteRecipeChooserProcessTables wsProd
    DeleteInventoryPaletteTables wsProd
    DeleteCheckboxesByPrefix wsProd, CHK_PROC_PREFIX
    DeleteCheckboxesByPrefix wsProd, CHK_BATCH_PREFIX
    DeleteCheckboxesByPrefix wsProd, CHK_RECALL_PREFIX

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If Not loOut Is Nothing Then
        ClearListObjectContents loOut
    End If

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loCheck Is Nothing Then
        ClearListObjectContents loCheck
    End If

    MsgBox "Recipe Chooser cleared.", vbInformation
End Sub

Public Sub BtnToUsed()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim usedDict As Object
    Set usedDict = BuildUsedDeltasFromPalette(wsProd)
    If usedDict Is Nothing Then
        MsgBox "No USED quantities found in palette tables.", vbInformation
        Exit Sub
    ElseIf usedDict.count = 0 Then
        MsgBox "No USED quantities found in palette tables.", vbInformation
        Exit Sub
    End If

    Dim errNotes As String
    Dim stagedTotal As Double
    stagedTotal = ValidateUsedStagingAgainstInvSys(invLo, usedDict, errNotes)
    If stagedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unknown staging failure."
        MsgBox "To USED cancelled: " & errNotes, vbCritical
        Exit Sub
    End If

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If loCheck Is Nothing Then
        MsgBox "To USED cancelled: Production staging table Prod_invSys_Check was not found.", vbCritical
        Exit Sub
    End If
    WriteProdInvSysCheck loCheck, invLo, usedDict

    Dim msg As String
    msg = "Staged production usage: " & Format$(stagedTotal, "0.###") & " units."
    If errNotes <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
        MsgBox msg, vbExclamation
    Else
        ShowProductionStatus msg
    End If
    Exit Sub
ErrHandler:
    MsgBox "BTN_TO_USED failed: " & Err.description, vbCritical
End Sub

Public Function PrepareProductionOutputForCurrentRecipe(Optional ByRef report As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then
        report = "Production sheet not found."
        Exit Function
    End If

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then
        report = "ProductionOutput table not found on Production sheet."
        Exit Function
    End If

    Dim outputEntries As Collection
    Set outputEntries = BuildOutputEntriesFromProcessTables(wsProd)
    If outputEntries Is Nothing Then
        report = "No OUTPUT items found in loaded recipe."
        Exit Function
    ElseIf outputEntries.count = 0 Then
        report = "No OUTPUT items found in loaded recipe."
        Exit Function
    End If

    Dim errNotes As String
    UpdateProductionOutputTable loOut, outputEntries, invLo, errNotes
    EnsureOutputBatchNumbers loOut
    RenderOutputRowCheckboxes wsProd
    ApplyRecallCodesForOutput wsProd, loOut, invLo, errNotes

    PrepareProductionOutputForCurrentRecipe = True
    report = "Prepared ProductionOutput rows: " & CStr(outputEntries.count)
    If errNotes <> "" Then report = report & "; " & errNotes
    Exit Function

ErrHandler:
    report = "PrepareProductionOutputForCurrentRecipe failed: " & Err.Description
End Function

Public Sub BtnPrepareProductionOutput()
    Dim report As String

    If PrepareProductionOutputForCurrentRecipe(report) Then
        ShowProductionStatus report
    Else
        MsgBox report, vbInformation
    End If
End Sub

Public Function CompleteProductionRunWithUsedPayload(ByVal usedPayloadJson As String, Optional ByRef report As String = "") As Boolean
    On Error GoTo ErrHandler

    If Not modRoleUiAccess.RequireCurrentUserCapability("PROD_POST") Then
        report = "Current user lacks PROD_POST capability."
        Exit Function
    End If

    usedPayloadJson = Trim$(usedPayloadJson)
    If usedPayloadJson = "" Or usedPayloadJson = "[]" Then
        report = "No production input payload rows were generated."
        Exit Function
    End If

    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then
        report = "Production sheet not found."
        Exit Function
    End If

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then
        report = "ProductionOutput table not found on Production sheet."
        Exit Function
    End If

    Dim madeNotes As String
    Dim madeDeltas As Collection
    Set madeDeltas = BuildMadeDeltasFromProductionOutput(loOut, invLo, madeNotes)
    If madeDeltas Is Nothing Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        report = madeNotes
        Exit Function
    ElseIf madeDeltas.Count = 0 Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        report = madeNotes
        Exit Function
    End If

    Dim errNotes As String
    Dim consumeEventId As String
    If Not modRoleEventWriter.QueuePayloadEventCurrent(EVENT_TYPE_PROD_CONSUME, _
                                                       "", _
                                                       usedPayloadJson, _
                                                       "COMPLETE_RUN_USED", _
                                                       consumeEventId, _
                                                       errNotes) Then
        If errNotes = "" Then errNotes = "Unable to queue production consume event."
        report = errNotes
        Exit Function
    End If

    Dim completeEventId As String
    If Not QueueProductionCompleteEvent(madeDeltas, errNotes, completeEventId) Then
        If errNotes = "" Then errNotes = "Unable to queue production completion event."
        report = errNotes
        Exit Function
    End If

    Dim runtimeReport As String
    If Not modOperationsPrimitiveBridge.RunBatchAndRefreshOperatorWorkbook(wsProd.Parent.Name, "", "LOCAL", runtimeReport) Then
        If runtimeReport = "" Then runtimeReport = "Production events queued, but runtime processing or read-model refresh did not complete cleanly."
        AppendNote errNotes, runtimeReport
    ElseIf runtimeReport <> "" Then
        AppendNote errNotes, runtimeReport
    End If

    CompleteProductionRunWithUsedPayload = True
    report = "ConsumeEvent=" & consumeEventId & "; CompleteEvent=" & completeEventId
    If errNotes <> "" Then report = report & "; " & errNotes
    Exit Function

ErrHandler:
    report = "CompleteProductionRunWithUsedPayload failed: " & Err.Description
End Function

Public Function CheckInProductionRunWithUsedPayload(ByVal usedPayloadJson As String, Optional ByRef report As String = "") As Boolean
    On Error GoTo ErrHandler

    If Not modRoleUiAccess.RequireCurrentUserCapability("PROD_POST") Then
        report = "Current user lacks PROD_POST capability."
        Exit Function
    End If

    usedPayloadJson = Trim$(usedPayloadJson)
    If usedPayloadJson = "" Or usedPayloadJson = "[]" Then
        report = "No production input payload rows were generated."
        Exit Function
    End If

    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then
        report = "Production sheet not found."
        Exit Function
    End If

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    Dim errNotes As String
    Dim usedDict As Object
    Set usedDict = BuildUsedDictFromPayloadJson(usedPayloadJson, errNotes)
    If usedDict Is Nothing Then
        If errNotes = "" Then errNotes = "No production input payload rows were generated."
        report = errNotes
        Exit Function
    End If
    If usedDict.count = 0 Then
        If errNotes = "" Then errNotes = "No production input payload rows were generated."
        report = errNotes
        Exit Function
    End If

    Dim stagedTotal As Double
    stagedTotal = ValidateUsedStagingAgainstInvSys(invLo, usedDict, errNotes)
    If stagedTotal < 0 Then
        If errNotes = "" Then errNotes = "Unable to stage production input rows."
        report = errNotes
        Exit Function
    End If

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If loCheck Is Nothing Then
        report = "Production staging table Prod_invSys_Check was not found."
        Exit Function
    End If
    WriteProdInvSysCheck loCheck, invLo, usedDict

    CheckInProductionRunWithUsedPayload = True
    report = "StagedUsed=" & Format$(stagedTotal, "0.###")
    If errNotes <> "" Then report = report & "; " & errNotes
    Exit Function

ErrHandler:
    report = "CheckInProductionRunWithUsedPayload failed: " & Err.Description
End Function

Public Function CheckInProductionRunWithUsedPayloadReportForAutomation(ByVal usedPayloadJson As String) As String
    Dim report As String

    If CheckInProductionRunWithUsedPayload(usedPayloadJson, report) Then
        CheckInProductionRunWithUsedPayloadReportForAutomation = "OK|" & report
    Else
        CheckInProductionRunWithUsedPayloadReportForAutomation = "FAIL|" & report
    End If
End Function

Public Function CompleteProductionRunAfterCheckIn(Optional ByRef report As String = "") As Boolean
    On Error GoTo ErrHandler

    If Not modRoleUiAccess.RequireCurrentUserCapability("PROD_POST") Then
        report = "Current user lacks PROD_POST capability."
        Exit Function
    End If

    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then
        report = "Production sheet not found."
        Exit Function
    End If

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        report = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then
        report = "ProductionOutput table not found on Production sheet."
        Exit Function
    End If

    Dim madeNotes As String
    Dim madeDeltas As Collection
    Set madeDeltas = BuildMadeDeltasFromProductionOutput(loOut, invLo, madeNotes)
    If madeDeltas Is Nothing Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        report = madeNotes
        Exit Function
    ElseIf madeDeltas.Count = 0 Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        report = madeNotes
        Exit Function
    End If

    Dim errNotes As String
    Dim completeEventId As String
    If Not QueueProductionCompleteEvent(madeDeltas, errNotes, completeEventId) Then
        If errNotes = "" Then errNotes = "Unable to queue production completion event."
        report = errNotes
        Exit Function
    End If

    ApplyRecallCodesForOutput wsProd, loOut, invLo, errNotes

    Dim pendingOutputValues As Object
    Set pendingOutputValues = CaptureProductionOutputCompletionValues(loOut)

    Dim runtimeReport As String
    If Not modOperationsPrimitiveBridge.RunBatchAndRefreshOperatorWorkbook(wsProd.Parent.Name, "", "LOCAL", runtimeReport) Then
        If runtimeReport = "" Then runtimeReport = "Production completion queued, but runtime processing or read-model refresh did not complete cleanly."
        AppendNote errNotes, runtimeReport
        report = "CompleteEvent=" & completeEventId & "; " & errNotes
        Exit Function
    ElseIf runtimeReport <> "" Then
        AppendNote errNotes, runtimeReport
    End If

    RestoreProductionOutputCompletionValues loOut, pendingOutputValues
    Dim logNotes As String
    LogProductionOutputToProductionLog wsProd, loOut, invLo, logNotes
    If logNotes <> "" Then AppendNote errNotes, logNotes

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loCheck Is Nothing Then
        If Not loCheck.DataBodyRange Is Nothing Then loCheck.DataBodyRange.ClearContents
    End If

    CompleteProductionRunAfterCheckIn = True
    report = "CompleteEvent=" & completeEventId
    If errNotes <> "" Then report = report & "; " & errNotes
    Exit Function

ErrHandler:
    report = "CompleteProductionRunAfterCheckIn failed: " & Err.Description
End Function

Public Function CompleteProductionRunAfterCheckInForOutput(ByVal outputRowNumber As Long, Optional ByRef report As String = "") As Boolean
    Dim completionStep As String
    Dim quietStarted As Boolean
    Dim failureNumber As Long
    Dim failureDescription As String

    On Error GoTo ErrHandler

    completionStep = "checking PROD_POST capability"
    If Not modRoleUiAccess.RequireCurrentUserCapability("PROD_POST") Then
        report = "Current user lacks PROD_POST capability."
        Exit Function
    End If

    completionStep = "resolving the Production worksheet"
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then
        report = "Production sheet not found."
        Exit Function
    End If

    completionStep = "resolving the operator inventory table"
    Dim invLo As ListObject
    Set invLo = GetInvSysTableFromWorkbook(wsProd.Parent)

    completionStep = "resolving the Production Output table"
    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then
        report = "ProductionOutput table not found on Production sheet."
        Exit Function
    End If
    If loOut.DataBodyRange Is Nothing Or outputRowNumber < 1 Or outputRowNumber > loOut.ListRows.Count Then
        report = "Select a valid Production Output row before completing the run."
        Exit Function
    End If

    Dim errNotes As String
    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    completionStep = "creating the typed Production session"
    Dim productionSession As cProductionRunSession
    Set productionSession = modProductionCompletionService.CreateProductionSessionFromWorkbook( _
        wsProd.Parent, outputRowNumber, errNotes)
    If productionSession Is Nothing Then
        If errNotes = "" Then errNotes = "Production session could not be prepared."
        report = errNotes
        Exit Function
    End If

    Dim consumeEventId As String
    Dim completeEventId As String
    consumeEventId = productionSession.ConsumeEventId
    completeEventId = productionSession.CompleteEventId

    completionStep = "applying recall-code metadata"
    ApplyRecallCodesForOutput wsProd, loOut, invLo, errNotes

    completionStep = "capturing pending output values"
    Dim pendingOutputValues As Object
    Set pendingOutputValues = CaptureProductionOutputCompletionValues(loOut, outputRowNumber)

    completionStep = "executing the typed Production completion service"
    Dim completionResult As cProductionCompletionResult
    modUiQuiet.BeginQuietUi wsProd.Parent
    quietStarted = True
    Set completionResult = modProductionCompletionService.ExecuteProductionSession( _
        wsProd.Parent, productionSession, errNotes)
    modUiQuiet.EndQuietUi
    quietStarted = False

    completionStep = "restoring completed output values after inventory refresh"
    RestoreProductionOutputCompletionValues loOut, pendingOutputValues
    If completionResult Is Nothing Then
        If errNotes = "" Then errNotes = "Production completion service returned no result."
        report = errNotes
        Exit Function
    End If
    If Not completionResult.Succeeded Then
        report = completionResult.ToEnvelope()
        Exit Function
    End If

    completionStep = "writing the completed output to ProductionLog"
    Dim logNotes As String
    If Not LogProductionOutputToProductionLog(wsProd, loOut, invLo, logNotes, outputRowNumber) Then
        If logNotes = "" Then logNotes = "ProductionLog did not accept the completed output."
        Err.Raise vbObjectError + 2101, "mProduction.CompleteProductionRunAfterCheckInForOutput", logNotes
    End If
    If logNotes <> "" Then AppendNote errNotes, logNotes

    completionStep = "clearing completed Production Run staging"
    If Not loCheck Is Nothing Then
        If Not loCheck.DataBodyRange Is Nothing Then loCheck.DataBodyRange.ClearContents
    End If

    CompleteProductionRunAfterCheckInForOutput = True
    report = "ConsumeEvent=" & consumeEventId & "; CompleteEvent=" & completeEventId
    If errNotes <> "" Then report = report & "; " & errNotes
    report = report & vbCrLf & _
             "Persistence summary: Production inbox events saved; processor durability saves retained."
    Exit Function

ErrHandler:
    failureNumber = Err.Number
    failureDescription = Err.Description
    On Error Resume Next
    If quietStarted Then modUiQuiet.EndQuietUi
    On Error GoTo 0
    report = "CompleteProductionRunAfterCheckInForOutput failed while " & completionStep & _
        ": " & CStr(failureNumber) & " - " & failureDescription
End Function

Public Function CompleteProductionRunAfterCheckInForOutputResult(ByVal outputRowNumber As Long) As String
    Dim report As String
    Dim succeeded As Boolean

    On Error GoTo ErrHandler
    succeeded = CompleteProductionRunAfterCheckInForOutput(outputRowNumber, report)
    If succeeded Then
        CompleteProductionRunAfterCheckInForOutputResult = "OK" & vbTab & report
    Else
        If Trim$(report) = "" Then report = "The Production completion backend returned False without a diagnostic report."
        CompleteProductionRunAfterCheckInForOutputResult = "FAIL" & vbTab & report
    End If
    Exit Function

ErrHandler:
    CompleteProductionRunAfterCheckInForOutputResult = "FAIL" & vbTab & _
        "CompleteProductionRunAfterCheckInForOutputResult failed: " & Err.Description
End Function

Public Sub RepairLastCompletedProductionRun()
    Dim report As String

    If FinalizePostedProductionRunLocal(report) Then
        MsgBox report, vbInformation, "Production Complete Run Repair"
    Else
        MsgBox report, vbExclamation, "Production Complete Run Repair"
    End If
End Sub

Public Function FinalizePostedProductionRunLocal(Optional ByRef report As String = "") As Boolean
    On Error GoTo ErrHandler

    Dim wsProd As Worksheet
    Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then
        report = "Production sheet not found."
        Exit Function
    End If

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Or loOut.DataBodyRange Is Nothing Then
        report = "ProductionOutput table not found or empty."
        Exit Function
    End If

    Dim cReal As Long
    cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    If cReal = 0 Then
        report = "ProductionOutput is missing its Real Output column."
        Exit Function
    End If

    Dim outputRowNumber As Long
    Dim positiveRows As Long
    Dim r As Long
    For r = 1 To loOut.ListRows.Count
        If NzDbl(loOut.DataBodyRange.Cells(r, cReal).Value) > 0 Then
            positiveRows = positiveRows + 1
            outputRowNumber = r
        End If
    Next r

    If positiveRows = 0 Then
        report = "No Production Output row contains a positive Real Output to repair."
        Exit Function
    End If
    If positiveRows > 1 Then
        report = "More than one Production Output row contains a Real Output. Clear unrelated entries before running this repair."
        Exit Function
    End If

    Dim invLo As ListObject
    Set invLo = GetInvSysTableFromWorkbook(wsProd.Parent)

    Dim logNotes As String
    If Not LogProductionOutputToProductionLog(wsProd, loOut, invLo, logNotes, outputRowNumber) Then
        If logNotes = "" Then logNotes = "ProductionLog did not accept the completed output."
        report = logNotes
        Exit Function
    End If

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loCheck Is Nothing Then
        If Not loCheck.DataBodyRange Is Nothing Then loCheck.DataBodyRange.ClearContents
    End If

    loOut.DataBodyRange.Cells(outputRowNumber, cReal).ClearContents
    Dim cBatch As Long
    cBatch = ColumnIndex(loOut, "BATCH")
    If cBatch > 0 Then loOut.DataBodyRange.Cells(outputRowNumber, cBatch).ClearContents

    FinalizePostedProductionRunLocal = True
    report = "The already-posted batch was recorded in ProductionLog and local run staging was cleared. Inventory was not posted again. Refresh Production Run to display Last, Batch, and Total."
    If logNotes <> "" Then report = report & " " & logNotes
    Exit Function

ErrHandler:
    report = "Local Production completion repair failed: " & CStr(Err.Number) & " - " & Err.Description
End Function

Public Sub BtnToMade()
    On Error GoTo ErrHandler
    If Not modRoleUiAccess.RequireCurrentUserCapability("PROD_POST") Then Exit Sub
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then
        MsgBox "ProductionOutput table not found on Production sheet.", vbExclamation
        Exit Sub
    End If

    Dim errNotes As String
    Dim outputEntries As Collection
    Set outputEntries = BuildOutputEntriesFromProcessTables(wsProd)
    If outputEntries Is Nothing Then
        MsgBox "No OUTPUT items found in process tables.", vbInformation
        Exit Sub
    ElseIf outputEntries.count = 0 Then
        MsgBox "No OUTPUT items found in process tables.", vbInformation
        Exit Sub
    End If

    UpdateProductionOutputTable loOut, outputEntries, invLo, errNotes
    EnsureOutputBatchNumbers loOut
    RenderOutputRowCheckboxes wsProd
    ApplyRecallCodesForOutput wsProd, loOut, invLo, errNotes

    Dim usedNotes As String
    Dim usedDeltas As Collection
    Dim loUsedCheck As ListObject
    Set loUsedCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    Set usedDeltas = BuildUsedDeltaPacketFromCheck(loUsedCheck, invLo, usedNotes)
    If usedDeltas Is Nothing Then
        If usedNotes = "" Then usedNotes = "No checked-in production input rows were found."
        MsgBox "Send to MADE cancelled: " & usedNotes, vbExclamation
        Exit Sub
    End If

    Dim madeNotes As String
    Dim madeDeltas As Collection
    Set madeDeltas = BuildMadeDeltasFromProductionOutput(loOut, invLo, madeNotes)
    If madeDeltas Is Nothing Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        MsgBox "Send to MADE cancelled: " & madeNotes, vbExclamation
        Exit Sub
    ElseIf madeDeltas.count = 0 Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        MsgBox "Send to MADE cancelled: " & madeNotes, vbExclamation
        Exit Sub
    End If

    Dim usedTotal As Double
    Dim madeTotal As Double
    Dim queuedEventId As String
    Dim runtimeReport As String

    If Not QueueProductionConsumeEvent(usedDeltas, madeDeltas, errNotes, queuedEventId) Then
        If errNotes = "" Then errNotes = "Unable to queue production consume event."
        MsgBox "Send to MADE cancelled: " & errNotes, vbExclamation
        Exit Sub
    End If

    If Not usedDeltas Is Nothing Then
        usedTotal = SumProductionDeltaQuantities(usedDeltas)
    ElseIf usedNotes <> "" Then
        AppendNote errNotes, usedNotes
    End If

    madeTotal = SumProductionDeltaQuantities(madeDeltas)

    Dim logNotes As String
    LogProductionOutputToProductionLog wsProd, loOut, invLo, logNotes
    If logNotes <> "" Then AppendNote errNotes, logNotes

    Dim rowKeys As Object
    Set rowKeys = BuildSystemKeySetFromDeltas(usedDeltas, madeDeltas)
    Dim usedSnapshot As Object
    Set usedSnapshot = BuildUsedSnapshotForSystemKeys(invLo, rowKeys)

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(wsProd, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loCheck Is Nothing Then
        If Not usedSnapshot Is Nothing Then
            WriteProdInvSysCheck loCheck, invLo, usedSnapshot
        End If
    End If

    If Not modOperationsPrimitiveBridge.RunBatchAndRefreshOperatorWorkbook(wsProd.Parent.Name, "", "LOCAL", runtimeReport) Then
        If runtimeReport = "" Then runtimeReport = "Local production post succeeded, but runtime processing or read-model refresh did not complete cleanly."
        AppendNote errNotes, runtimeReport
    ElseIf runtimeReport <> "" Then
        AppendNote errNotes, runtimeReport
    End If
    AllowExcelRefreshToSettleProduction

    Dim msg As String
    msg = "Recorded component usage: " & Format$(usedTotal, "0.###") & " units."
    msg = msg & vbCrLf & "Recorded finished goods (MADE): " & Format$(madeTotal, "0.###")
    If queuedEventId <> "" Then msg = msg & vbCrLf & "Inbox EventID: " & queuedEventId
    If errNotes <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
        If HasActionableProductionWarning(errNotes) Then
            MsgBox msg, vbExclamation
        Else
            ShowProductionStatus msg
        End If
    Else
        ShowProductionStatus msg
    End If
    Exit Sub
ErrHandler:
    MsgBox "BTN_TO_MADE failed: " & Err.description, vbCritical
End Sub

Public Sub ProductionToTotalInv()
    On Error GoTo ErrHandler
    If Not modRoleUiAccess.RequireCurrentUserCapability("PROD_POST") Then Exit Sub
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub

    Dim invLo As ListObject: Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        MsgBox "InventoryManagement!invSys table not found.", vbCritical
        Exit Sub
    End If

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then
        MsgBox "ProductionOutput table not found on Production sheet.", vbExclamation
        Exit Sub
    End If

    Dim errNotes As String
    Dim madeNotes As String
    Dim madeDeltas As Collection
    Dim queuedEventId As String
    Dim runtimeReport As String
    Set madeDeltas = BuildMadeDeltasFromProductionOutput(loOut, invLo, madeNotes)
    If madeDeltas Is Nothing Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        MsgBox "Send to TOTAL INV cancelled: " & madeNotes, vbExclamation
        Exit Sub
    ElseIf madeDeltas.count = 0 Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        MsgBox "Send to TOTAL INV cancelled: " & madeNotes, vbExclamation
        Exit Sub
    End If

    If Not QueueProductionCompleteEvent(madeDeltas, errNotes, queuedEventId) Then
        If errNotes = "" Then errNotes = "Unable to queue production completion event."
        MsgBox "Send to TOTAL INV cancelled: " & errNotes, vbExclamation
        Exit Sub
    End If

    Dim totalMoved As Double
    totalMoved = SumProductionDeltaQuantities(madeDeltas)

    If Not modOperationsPrimitiveBridge.RunBatchAndRefreshOperatorWorkbook(wsProd.Parent.Name, "", "LOCAL", runtimeReport) Then
        If runtimeReport = "" Then runtimeReport = "Local production completion succeeded, but runtime processing or read-model refresh did not complete cleanly."
        AppendNote errNotes, runtimeReport
    ElseIf runtimeReport <> "" Then
        AppendNote errNotes, runtimeReport
    End If
    AllowExcelRefreshToSettleProduction

    Dim msg As String
    msg = "Moved MADE to TOTAL INV: " & Format$(totalMoved, "0.###") & " units."
    If queuedEventId <> "" Then msg = msg & vbCrLf & "Inbox EventID: " & queuedEventId
    If errNotes <> "" Then
        msg = msg & vbCrLf & vbCrLf & "Warnings:" & vbCrLf & errNotes
        If HasActionableProductionWarning(errNotes) Then
            MsgBox msg, vbExclamation
        Else
            ShowProductionStatus msg
        End If
    Else
        ShowProductionStatus msg
    End If
    Exit Sub
ErrHandler:
    MsgBox "BTN_TO_TOTALINV failed: " & Err.description, vbCritical
End Sub

Private Sub AllowExcelRefreshToSettleProduction()
    On Error Resume Next
    DoEvents
    Application.Wait Now + TimeSerial(0, 0, 1)
    DoEvents
    On Error GoTo 0
End Sub

Private Sub ShowProductionStatus(ByVal messageText As String)
    On Error Resume Next
    Application.StatusBar = FlattenProductionStatusText(messageText)
    On Error GoTo 0
End Sub

Private Function FlattenProductionStatusText(ByVal messageText As String) As String
    Dim result As String

    result = Replace(messageText, vbCrLf, "  ")
    result = Replace(result, vbCr, "  ")
    result = Replace(result, vbLf, "  ")
    Do While InStr(result, "   ") > 0
        result = Replace(result, "   ", "  ")
    Loop
    If Len(result) > 240 Then result = Left$(result, 237) & "..."
    FlattenProductionStatusText = result
End Function

Private Function HasActionableProductionWarning(ByVal notes As String) As Boolean
    Dim lowered As String

    lowered = LCase$(Trim$(notes))
    If lowered = "" Then Exit Function

    HasActionableProductionWarning = _
        (InStr(1, lowered, "failed", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "error", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "poison", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "cancel", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "insufficient", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "unable", vbTextCompare) > 0) _
        Or (InStr(1, lowered, "did not complete", vbTextCompare) > 0)
End Function

Public Function QueueProductionCompleteEventFromCurrentWorkbook(ByRef eventIdOut As String, ByRef errNotes As String) As Boolean
    Dim wsProd As Worksheet
    Dim invLo As ListObject
    Dim loOut As ListObject
    Dim madeNotes As String
    Dim madeDeltas As Collection

    If Not modRoleUiAccess.CanCurrentUserPerformCapability("PROD_POST", "", "", "", errNotes) Then Exit Function

    Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then
        errNotes = "Production sheet not found."
        Exit Function
    End If

    Set invLo = GetInvSysTable()
    If invLo Is Nothing Then
        errNotes = "InventoryManagement!invSys table not found."
        Exit Function
    End If

    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then
        errNotes = "ProductionOutput table not found on Production sheet."
        Exit Function
    End If

    Set madeDeltas = BuildMadeDeltasFromProductionOutput(loOut, invLo, madeNotes)
    If madeDeltas Is Nothing Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        errNotes = madeNotes
        Exit Function
    End If
    If madeDeltas.Count = 0 Then
        If madeNotes = "" Then madeNotes = "No made quantities found in ProductionOutput."
        errNotes = madeNotes
        Exit Function
    End If

    QueueProductionCompleteEventFromCurrentWorkbook = QueueProductionCompleteEvent(madeDeltas, errNotes, eventIdOut)
End Function

Public Function ValidateQueueProductionCompleteEventFromCurrentWorkbook() As String
    Dim eventIdOut As String
    Dim errNotes As String

    If QueueProductionCompleteEventFromCurrentWorkbook(eventIdOut, errNotes) Then
        ValidateQueueProductionCompleteEventFromCurrentWorkbook = "OK"
    Else
        ValidateQueueProductionCompleteEventFromCurrentWorkbook = errNotes
    End If
End Function

'@TestOnlyBegin
Public Function TestCompletionDeltasFromStagedRows(ByVal loOut As ListObject, ByVal loCheck As ListObject) As String
    Dim madeNotes As String
    Dim usedNotes As String
    Dim madeDeltas As Collection
    Dim usedDeltas As Collection
    Dim madeDelta As Object
    Dim usedDelta As Object

    On Error GoTo ErrHandler
    Set madeDeltas = BuildMadeDeltasFromProductionOutputRow(loOut, Nothing, 1, madeNotes)
    Set usedDeltas = BuildUsedDeltaPacketFromCheck(loCheck, Nothing, usedNotes)
    If madeDeltas Is Nothing Then
        TestCompletionDeltasFromStagedRows = "FAIL|Made=" & madeNotes
        Exit Function
    End If
    If usedDeltas Is Nothing Then
        TestCompletionDeltasFromStagedRows = "FAIL|Used=" & usedNotes
        Exit Function
    End If
    If madeDeltas.Count <> 1 Or usedDeltas.Count <> 1 Then
        TestCompletionDeltasFromStagedRows = "FAIL|MadeCount=" & CStr(madeDeltas.Count) & ";UsedCount=" & CStr(usedDeltas.Count)
        Exit Function
    End If

    Set madeDelta = madeDeltas(1)
    Set usedDelta = usedDeltas(1)
    TestCompletionDeltasFromStagedRows = "OK|MadeSystemKey=" & CStr(madeDelta("System_Key")) & _
        ";MadeQty=" & CStr(madeDelta("QTY")) & _
        ";UsedSystemKey=" & CStr(usedDelta("System_Key")) & _
        ";UsedQty=" & CStr(usedDelta("QTY"))
    Exit Function

ErrHandler:
    TestCompletionDeltasFromStagedRows = "FAIL|" & Err.Description
End Function

Public Function TestProductionUsedStagingDoesNotMutateInvSys(ByVal invLo As ListObject, _
                                                             ByVal loCheck As ListObject, _
                                                             ByVal systemKey As String, _
                                                             ByVal qtyValue As Double) As String
    Dim usedDict As Object
    Dim rowIndex As Object
    Dim invIndex As Long
    Dim cUsed As Long
    Dim cTotal As Long
    Dim cCheckUsed As Long
    Dim beforeUsed As Double
    Dim beforeTotal As Double
    Dim stagedTotal As Double
    Dim errNotes As String

    On Error GoTo ErrHandler
    Set rowIndex = BuildInvSysSystemKeyIndex(invLo)
    If rowIndex Is Nothing Then GoTo MissingRow
    If Not rowIndex.Exists(systemKey) Then GoTo MissingRow
    invIndex = CLng(rowIndex(systemKey))
    cUsed = ColumnIndex(invLo, "USED")
    cTotal = ColumnIndex(invLo, "TOTAL INV")
    cCheckUsed = ColumnIndex(loCheck, "USED")
    If cUsed = 0 Or cTotal = 0 Or cCheckUsed = 0 Then
        TestProductionUsedStagingDoesNotMutateInvSys = "FAIL|Required columns missing."
        Exit Function
    End If

    beforeUsed = NzDbl(invLo.DataBodyRange.Cells(invIndex, cUsed).Value)
    beforeTotal = NzDbl(invLo.DataBodyRange.Cells(invIndex, cTotal).Value)
    Set usedDict = CreateObject("Scripting.Dictionary")
    usedDict(systemKey) = qtyValue
    stagedTotal = ValidateUsedStagingAgainstInvSys(invLo, usedDict, errNotes)
    If stagedTotal < 0 Then
        TestProductionUsedStagingDoesNotMutateInvSys = "FAIL|" & errNotes
        Exit Function
    End If
    WriteProdInvSysCheck loCheck, invLo, usedDict

    TestProductionUsedStagingDoesNotMutateInvSys = _
        "OK|BeforeUsed=" & Format$(beforeUsed, "0.############") & _
        ";AfterUsed=" & Format$(NzDbl(invLo.DataBodyRange.Cells(invIndex, cUsed).Value), "0.############") & _
        ";BeforeTotal=" & Format$(beforeTotal, "0.############") & _
        ";AfterTotal=" & Format$(NzDbl(invLo.DataBodyRange.Cells(invIndex, cTotal).Value), "0.############") & _
        ";CheckUsed=" & Format$(NzDbl(loCheck.DataBodyRange.Cells(1, cCheckUsed).Value), "0.############")
    Exit Function

MissingRow:
    TestProductionUsedStagingDoesNotMutateInvSys = "FAIL|Inventory System_Key was not found."
    Exit Function

ErrHandler:
    TestProductionUsedStagingDoesNotMutateInvSys = "FAIL|" & Err.Description
End Function

Public Function TestProductionSystemKeyPayloadStagesWithoutInventoryMutation( _
        ByVal invLo As ListObject, _
        ByVal loCheck As ListObject, _
        ByVal systemKey As String, _
        ByVal sku As String, _
        ByVal qtyValue As Double) As String
    Dim usedDict As Object
    Dim payloadJson As String
    Dim errNotes As String
    Dim stagedTotal As Double
    Dim cUsedInv As Long
    Dim cUsedCheck As Long
    Dim cSystemKeyCheck As Long
    Dim beforeUsed As Double

    On Error GoTo ErrHandler
    cUsedInv = ColumnIndex(invLo, "USED")
    cUsedCheck = ColumnIndex(loCheck, "USED")
    cSystemKeyCheck = ColumnIndex(loCheck, "System_Key")
    If cUsedInv = 0 Or cUsedCheck = 0 Or cSystemKeyCheck = 0 Then
        TestProductionSystemKeyPayloadStagesWithoutInventoryMutation = "FAIL|Required columns missing."
        Exit Function
    End If
    beforeUsed = NzDbl(invLo.DataBodyRange.Cells(1, cUsedInv).Value)
    payloadJson = "[{""System_Key"":""" & systemKey & _
                  """,""SKU"":""" & sku & _
                  """,""Qty"":" & Replace$(CStr(qtyValue), Application.International(xlDecimalSeparator), ".") & _
                  ",""IoType"":""USED""}]"
    Set usedDict = BuildUsedDictFromPayloadJson(payloadJson, errNotes)
    If usedDict Is Nothing Then
        TestProductionSystemKeyPayloadStagesWithoutInventoryMutation = "FAIL|" & errNotes
        Exit Function
    End If
    stagedTotal = ValidateUsedStagingAgainstInvSys(invLo, usedDict, errNotes)
    If stagedTotal < 0 Then
        TestProductionSystemKeyPayloadStagesWithoutInventoryMutation = "FAIL|" & errNotes
        Exit Function
    End If
    WriteProdInvSysCheck loCheck, invLo, usedDict
    TestProductionSystemKeyPayloadStagesWithoutInventoryMutation = _
        "OK|SystemKey=" & NzStr(loCheck.DataBodyRange.Cells(1, cSystemKeyCheck).Value) & _
        "|Staged=" & Format$(NzDbl(loCheck.DataBodyRange.Cells(1, cUsedCheck).Value), "0.###") & _
        "|InventoryUsedUnchanged=" & IIf( _
            Abs(NzDbl(invLo.DataBodyRange.Cells(1, cUsedInv).Value) - beforeUsed) < 0.0000001, _
            "TRUE", "FALSE")
    Exit Function

ErrHandler:
    TestProductionSystemKeyPayloadStagesWithoutInventoryMutation = "FAIL|" & Err.Description
End Function

Public Function TestLookupOutputSystemKeyFromPicker(ByVal pickerItems As Variant, ByVal outputName As String) As String
    Dim notes As String
    Dim systemKey As String

    systemKey = LookupOutputSystemKeyFromPicker(pickerItems, outputName, notes)
    TestLookupOutputSystemKeyFromPicker = systemKey & "|" & notes
End Function

Public Function TestSelectedMadeDeltaSystemKey(ByVal loOut As ListObject, ByVal invLo As ListObject) As String
    Dim notes As String
    Dim deltas As Collection
    Dim delta As Object

    Set deltas = BuildMadeDeltasFromProductionOutputRow(loOut, invLo, 1, notes)
    If deltas Is Nothing Then
        TestSelectedMadeDeltaSystemKey = "FAIL|" & notes
        Exit Function
    End If
    If deltas.Count <> 1 Then
        TestSelectedMadeDeltaSystemKey = "FAIL|Count=" & CStr(deltas.Count) & "|" & notes
        Exit Function
    End If
    Set delta = deltas(1)
    TestSelectedMadeDeltaSystemKey = "OK|" & CStr(delta("System_Key")) & "|" & notes
End Function

Public Function TestOutputIdentityFromPicker(ByVal pickerItems As Variant, _
                                             ByVal systemKey As String, _
                                             ByVal outputName As String) As String
    Dim delta As Object
    Set delta = CreateObject("Scripting.Dictionary")
    delta("System_Key") = systemKey
    delta("ITEM_CODE") = ""
    delta("ITEM_NAME") = outputName
    EnrichOutputDeltaFromPickerBySystemKey delta, pickerItems, systemKey
    TestOutputIdentityFromPicker = CStr(delta("System_Key")) & "|" & NzStr(delta("ITEM_CODE")) & "|" & NzStr(delta("ITEM_NAME"))
End Function

Public Function TestSelectedMadeDeltaSkuIdentity(ByVal loOut As ListObject) As String
    Dim notes As String
    Dim deltas As Collection
    Dim delta As Object

    Set deltas = BuildMadeDeltasFromProductionOutputRow(loOut, Nothing, 1, notes)
    If deltas Is Nothing Then
        TestSelectedMadeDeltaSkuIdentity = "FAIL|" & notes
        Exit Function
    End If
    If deltas.Count <> 1 Then
        TestSelectedMadeDeltaSkuIdentity = "FAIL|Count=" & CStr(deltas.Count) & "|" & notes
        Exit Function
    End If
    Set delta = deltas(1)
    TestSelectedMadeDeltaSkuIdentity = "OK|" & CStr(delta("System_Key")) & "|" & _
        NzStr(delta("ITEM_CODE")) & "|" & Format$(NzDbl(delta("QTY")), "0.###")
End Function

Public Function TestLogProductionOutputRow(ByVal wsProd As Worksheet, ByVal loOut As ListObject, ByVal outputRowNumber As Long) As String
    Dim firstNotes As String
    Dim secondNotes As String

    If Not LogProductionOutputToProductionLog(wsProd, loOut, Nothing, firstNotes, outputRowNumber) Then
        TestLogProductionOutputRow = "FAIL|First=" & firstNotes
        Exit Function
    End If
    If Not LogProductionOutputToProductionLog(wsProd, loOut, Nothing, secondNotes, outputRowNumber) Then
        TestLogProductionOutputRow = "FAIL|Second=" & secondNotes
        Exit Function
    End If
    TestLogProductionOutputRow = "OK|" & firstNotes & "|" & secondNotes
End Function
'@TestOnlyEnd

Public Sub BtnNextBatch()
    On Error GoTo ErrHandler
    Dim ws As Worksheet: Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub

    EnsurePaletteTableMetaForExistingTables ws

    Dim invLo As ListObject
    Set invLo = GetInvSysTable()

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(ws, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If Not loOut Is Nothing Then
        EnsureOutputBatchNumbers loOut
        ClearProductionOutputForNextBatch ws, loOut
    End If

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row >= PALETTE_LINES_STAGING_ROW Then GoTo NextLo
            Dim procName As String
            Dim recipeId As String
            Dim ingId As String
            Dim amtVal As Variant
            Dim ioVal As String
            If GetPaletteTableContext(lo, recipeId, ingId, amtVal, procName, ioVal) = False Then
                procName = ProcessNameFromTable(lo)
            End If
            If Trim$(procName) = "" Then procName = lo.Name

            If Not IsPaletteKeepSelected(ws, procName) Then
                ClearPaletteTableSelection lo
            End If
        End If
NextLo:
    Next lo

    MsgBox "Next Batch ready. Inventory selections cleared for unchecked processes.", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "BTN_NEXT_BATCH failed: " & Err.description, vbCritical
End Sub

Public Sub BtnPrintRecallCodes()
    On Error GoTo ErrHandler

    Dim wsReport As Worksheet
    Dim rowCount As Long
    Dim detail As String

    If Not BuildRecallCodesReportFromCurrentWorkbook(wsReport, rowCount, detail) Then
        MsgBox detail, vbInformation
        Exit Sub
    End If

    wsReport.Activate
    wsReport.PrintOut Preview:=True
    Exit Sub
ErrHandler:
    MsgBox "BTN_PRINT_CODES failed: " & Err.Description, vbCritical
End Sub

Public Function GetRecallPrintDiagnostic() As String
    Dim wsReport As Worksheet
    Dim rowCount As Long
    Dim detail As String

    If BuildRecallCodesReportFromCurrentWorkbook(wsReport, rowCount, detail) Then
        GetRecallPrintDiagnostic = "OK; Sheet=" & wsReport.Name & "; Rows=" & rowCount
    Else
        GetRecallPrintDiagnostic = detail
    End If
End Function

' ===== System 2: Inventory Palette Builder =====
Private Sub SaveIngredientPalette()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim wbProd As Workbook: Set wbProd = wsProd.Parent

    Dim loRecipe As ListObject
    Dim loIng As ListObject
    Dim loItems As ListObject
    Set loRecipe = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseRecipe", Array("RECIPE_NAME", "RECIPE_ID"))
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))
    If loRecipe Is Nothing Or loIng Is Nothing Or loItems Is Nothing Then
        MsgBox "Inventory Palette Builder tables not found on Production sheet.", vbExclamation
        Exit Sub
    End If

    Dim recipeId As String: recipeId = GetPaletteRecipeId()
    If recipeId = "" Then
        MsgBox "Select a RECIPE in IP_ChooseRecipe before saving.", vbInformation
        Exit Sub
    End If

    Dim ingredientId As String: ingredientId = GetPaletteIngredientId()
    If ingredientId = "" Then
        MsgBox "Select an INGREDIENT in IP_ChooseIngredient before saving.", vbInformation
        Exit Sub
    End If

    If loItems.DataBodyRange Is Nothing Then
        MsgBox "Add at least one acceptable item before saving.", vbInformation
        Exit Sub
    End If

    Dim wsPal As Worksheet
    Set wsPal = WorkbookSheetExists(wbProd, "IngredientPalette")
    If wsPal Is Nothing Then Set wsPal = WorkbookSheetExists(wbProd, "IngredientsPalette")
    If wsPal Is Nothing Then
        MsgBox "IngredientPalette sheet not found.", vbCritical
        Exit Sub
    End If

    Dim loPal As ListObject
    Set loPal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "ITEM"))
    If loPal Is Nothing Then
        Set loPal = FindListObjectByNameOrHeaders(wsPal, "Table40", Array("RECIPE_ID", "INGREDIENT_ID", "ITEM"))
    End If
    If loPal Is Nothing Then
        MsgBox "IngredientPalette table not found on IngredientPalette sheet.", vbCritical
        Exit Sub
    End If

    Dim ioVal As String
    Dim pctVal As Variant
    Dim uomVal As String
    Dim amtVal As Variant
    FindRecipeIngredientInfo recipeId, ingredientId, ioVal, pctVal, uomVal, amtVal, wbProd
    If PaletteIoValueIsOutput(ioVal) Then
        MsgBox "Outputs do not accept ingredients. Assign acceptable inventory only to recipe inputs.", vbExclamation
        Exit Sub
    End If

    ' Remove existing palette rows for this recipe + ingredient.
    If Not loPal.DataBodyRange Is Nothing Then
        Dim cPalRec As Long: cPalRec = ColumnIndex(loPal, "RECIPE_ID")
        Dim cPalIng As Long: cPalIng = ColumnIndex(loPal, "INGREDIENT_ID")
        If cPalRec > 0 And cPalIng > 0 Then
            Dim r As Long
            For r = loPal.DataBodyRange.rows.count To 1 Step -1
                If NzStr(loPal.DataBodyRange.Cells(r, cPalRec).value) = recipeId _
                   And NzStr(loPal.DataBodyRange.Cells(r, cPalIng).value) = ingredientId Then
                    loPal.ListRows(r).Delete
                End If
            Next r
        End If
    End If

    Dim cItem As Long: cItem = ColumnIndex(loItems, "ITEMS")
    If cItem = 0 Then cItem = ColumnIndex(loItems, "ITEM")
    Dim cUom As Long: cUom = ColumnIndex(loItems, "UOM")
    Dim cRow As Long: cRow = ColumnIndex(loItems, "System_Key")
    Dim cItemCode As Long: cItemCode = ColumnIndex(loItems, "ITEM_CODE")

    Dim cOutRec As Long: cOutRec = ColumnIndex(loPal, "RECIPE_ID")
    Dim cOutIng As Long: cOutIng = ColumnIndex(loPal, "INGREDIENT_ID")
    Dim cOutIO As Long: cOutIO = ColumnIndex(loPal, "INPUT/OUTPUT")
    Dim cOutItem As Long: cOutItem = ColumnIndex(loPal, "ITEM")
    Dim cOutPct As Long: cOutPct = ColumnIndex(loPal, "PERCENT")
    Dim cOutUom As Long: cOutUom = ColumnIndex(loPal, "UOM")
    Dim cOutAmt As Long: cOutAmt = ColumnIndex(loPal, "AMOUNT")
    Dim cOutRow As Long: cOutRow = ColumnIndex(loPal, "System_Key")
    Dim cOutItemCode As Long: cOutItemCode = ColumnIndex(loPal, "ITEM_CODE")
    Dim cOutGuid As Long: cOutGuid = ColumnIndex(loPal, "GUID")

    Dim added As Long
    Dim arr As Variant: arr = loItems.DataBodyRange.value
    Dim i As Long
    For i = 1 To UBound(arr, 1)
        Dim itemVal As String
        If cItem > 0 Then itemVal = NzStr(arr(i, cItem))
        If Trim$(itemVal) = "" Then GoTo NextItem

        Dim lr As ListRow: Set lr = loPal.ListRows.Add
        If cOutRec > 0 Then lr.Range.Cells(1, cOutRec).value = recipeId
        If cOutIng > 0 Then lr.Range.Cells(1, cOutIng).value = ingredientId
        If cOutIO > 0 Then lr.Range.Cells(1, cOutIO).value = ioVal
        If cOutItem > 0 Then lr.Range.Cells(1, cOutItem).value = itemVal
        If cOutPct > 0 Then lr.Range.Cells(1, cOutPct).value = pctVal
        If cOutUom > 0 Then
            Dim itemUom As String
            If cUom > 0 Then itemUom = NzStr(arr(i, cUom))
            If itemUom <> "" Then
                lr.Range.Cells(1, cOutUom).value = itemUom
            Else
                lr.Range.Cells(1, cOutUom).value = uomVal
            End If
        End If
        If cOutAmt > 0 Then lr.Range.Cells(1, cOutAmt).value = amtVal
        If cOutRow > 0 And cRow > 0 Then lr.Range.Cells(1, cOutRow).value = arr(i, cRow)
        If cOutItemCode > 0 And cItemCode > 0 Then _
            lr.Range.Cells(1, cOutItemCode).Value = arr(i, cItemCode)
        If cOutGuid > 0 Then lr.Range.Cells(1, cOutGuid).value = CreateProductionGuid()
        added = added + 1
NextItem:
    Next i

    Dim runtimeReport As String
    Dim runtimeSaved As Boolean
    Dim msg As String

    runtimeSaved = PublishIngredientPaletteRowsToRuntime(wbProd, loPal, recipeId, ingredientId, runtimeReport)
    msg = "Saved IngredientPalette rows: " & added & "."
    If runtimeSaved Then
        msg = msg & vbCrLf & "Server/NAS saved: yes."
    ElseIf Trim$(runtimeReport) <> "" Then
        msg = msg & vbCrLf & "Server/NAS save did not complete: " & runtimeReport
    End If

    MsgBox msg, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Save IngredientPalette failed: " & Err.description, vbCritical
End Sub

Public Function PaletteIngredientIsOutput(ByVal recipeId As String, ByVal ingredientId As String) As Boolean
    Dim ioVal As String
    Dim pctVal As Variant
    Dim uomVal As String
    Dim amtVal As Variant

    FindRecipeIngredientInfo recipeId, ingredientId, ioVal, pctVal, uomVal, amtVal
    PaletteIngredientIsOutput = PaletteIoValueIsOutput(ioVal)
End Function

Private Function PaletteIoValueIsOutput(ByVal ioVal As String) As Boolean
    ioVal = UCase$(Trim$(ioVal))
    PaletteIoValueIsOutput = (ioVal = "OUTPUT" Or ioVal = "MADE")
End Function

Private Sub ClearInventoryPaletteBuilder()
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim loRecipe As ListObject
    Dim loIng As ListObject
    Dim loItems As ListObject
    Set loRecipe = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseRecipe", Array("RECIPE_NAME", "RECIPE_ID"))
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))

    ResetPaletteTable loItems
    ResetPaletteTable loIng
    ResetPaletteTable loRecipe

    MsgBox "Inventory Palette Builder cleared.", vbInformation
End Sub

Public Sub HandlePaletteRecipeSelected(ByVal recipeId As String)
    On Error GoTo ErrHandler

    HandlePaletteRecipeSelectedCore recipeId
    Exit Sub

ErrHandler:
    MsgBox "Select assignment recipe failed: " & Err.Description, vbCritical
End Sub

'@TestOnlyBegin
Public Function TestHandlePaletteRecipeSelected(ByVal recipeId As String) As String
    On Error GoTo ErrHandler

    HandlePaletteRecipeSelectedCore recipeId
    TestHandlePaletteRecipeSelected = "OK"
    Exit Function

ErrHandler:
    TestHandlePaletteRecipeSelected = "ERROR " & CStr(Err.Number) & ": " & Err.Description
End Function

Public Function TestHandlePaletteRecipeSelectedStage(ByVal recipeId As String, ByVal maxStage As Long) As String
    On Error GoTo ErrHandler

    Dim wsProd As Worksheet
    Dim loIng As ListObject
    Dim loItems As ListObject

    TestHandlePaletteRecipeSelectedStage = "stage0"
    If Trim$(recipeId) = "" Then Exit Function
    Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then
        TestHandlePaletteRecipeSelectedStage = "stage1:no production sheet"
        Exit Function
    End If
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))
    TestHandlePaletteRecipeSelectedStage = "stage1:tables"
    If maxStage <= 1 Then Exit Function

    ResetPaletteTable loItems
    TestHandlePaletteRecipeSelectedStage = "stage2:items reset"
    If maxStage <= 2 Then Exit Function

    If Not loIng Is Nothing Then
        ResetPaletteTable loIng
        Dim cRec As Long: cRec = ColumnIndex(loIng, "RECIPE_ID")
        If cRec > 0 Then
            Dim recCell As Range
            Set recCell = GetHeaderDataCell(loIng, "RECIPE_ID")
            If Not recCell Is Nothing Then recCell.value = recipeId
        End If
    End If
    TestHandlePaletteRecipeSelectedStage = "stage3:ingredient reset"
    If maxStage <= 3 Then Exit Function

    If Not loItems Is Nothing Then ClearListObjectFormulas loItems
    If Not loIng Is Nothing Then ClearListObjectFormulas loIng
    TestHandlePaletteRecipeSelectedStage = "stage4:formulas clear"
    If maxStage <= 4 Then Exit Function

    TestHandlePaletteRecipeSelectedStage = "stage5:templates skipped"
    Exit Function

ErrHandler:
    TestHandlePaletteRecipeSelectedStage = "ERROR stage=" & TestHandlePaletteRecipeSelectedStage & " " & CStr(Err.Number) & ": " & Err.Description
End Function
'@TestOnlyEnd

Private Sub HandlePaletteRecipeSelectedCore(ByVal recipeId As String)
    ' System 2: Inventory Palette Builder - clear ingredient/items when recipe changes.
    If Trim$(recipeId) = "" Then Exit Sub
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim loIng As ListObject
    Dim loItems As ListObject
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))
    ResetPaletteTable loItems
    If Not loIng Is Nothing Then
        ResetPaletteTable loIng
        Dim cRec As Long: cRec = ColumnIndex(loIng, "RECIPE_ID")
        If cRec > 0 Then
            Dim recCell As Range
            Set recCell = GetHeaderDataCell(loIng, "RECIPE_ID")
            If Not recCell Is Nothing Then recCell.value = recipeId
        End If
    End If

    If Not loItems Is Nothing Then
        ClearListObjectFormulas loItems
    End If
    If Not loIng Is Nothing Then
        ClearListObjectFormulas loIng
    End If

End Sub

Public Sub HandlePaletteIngredientSelected(ByVal recipeId As String, ByVal ingredientId As String)
    ' System 2: Inventory Palette Builder - clear items when ingredient changes.
    If Trim$(ingredientId) = "" Then Exit Sub
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim syncReport As String
    RefreshProductionIngredientPaletteFromRuntime wsProd.Parent, syncReport
    Dim loItems As ListObject
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))
    If Not loItems Is Nothing Then
        ResetPaletteTable loItems
        Dim cRec As Long: cRec = ColumnIndex(loItems, "RECIPE_ID")
        Dim cIng As Long: cIng = ColumnIndex(loItems, "INGREDIENT_ID")
        If cRec > 0 Then
            Dim recCell As Range
            Set recCell = GetHeaderDataCell(loItems, "RECIPE_ID")
            If Not recCell Is Nothing Then recCell.value = recipeId
        End If
        If cIng > 0 Then
            Dim ingCell As Range
            Set ingCell = GetHeaderDataCell(loItems, "INGREDIENT_ID")
            If Not ingCell Is Nothing Then ingCell.value = ingredientId
        End If
        PopulateChooseItemFromIngredientPalette recipeId, ingredientId, loItems
    End If
End Sub

Private Sub PopulateChooseItemFromIngredientPalette(ByVal recipeId As String, ByVal ingredientId As String, ByVal loItems As ListObject)
    If Trim$(recipeId) = "" Or Trim$(ingredientId) = "" Then Exit Sub
    If loItems Is Nothing Then Exit Sub

    Dim wsPal As Worksheet: Set wsPal = SheetExists("IngredientPalette")
    If wsPal Is Nothing Then Set wsPal = SheetExists("IngredientsPalette")
    If wsPal Is Nothing Then Exit Sub

    Dim loPal As ListObject
    Set loPal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "System_Key"))
    If loPal Is Nothing Then
        Set loPal = FindListObjectByNameOrHeaders(wsPal, "Table40", Array("RECIPE_ID", "INGREDIENT_ID", "System_Key"))
    End If
    If loPal Is Nothing Then Exit Sub
    If loPal.DataBodyRange Is Nothing Then Exit Sub

    Dim cRec As Long: cRec = ColumnIndex(loPal, "RECIPE_ID")
    Dim cIng As Long: cIng = ColumnIndex(loPal, "INGREDIENT_ID")
    Dim cSystemKey As Long: cSystemKey = ColumnIndex(loPal, "System_Key")
    Dim cItem As Long: cItem = ColumnIndex(loPal, "ITEM")
    Dim cUom As Long: cUom = ColumnIndex(loPal, "UOM")
    Dim cItemCode As Long: cItemCode = ColumnIndex(loPal, "ITEM_CODE")
    If cRec = 0 Or cIng = 0 Then Exit Sub

    Dim oItem As Long: oItem = ColumnIndex(loItems, "ITEMS")
    If oItem = 0 Then oItem = ColumnIndex(loItems, "ITEM")
    Dim oUom As Long: oUom = ColumnIndex(loItems, "UOM")
    Dim oDesc As Long: oDesc = ColumnIndex(loItems, "DESCRIPTION")
    Dim oSystemKey As Long: oSystemKey = ColumnIndex(loItems, "System_Key")
    Dim oRec As Long: oRec = ColumnIndex(loItems, "RECIPE_ID")
    Dim oIng As Long: oIng = ColumnIndex(loItems, "INGREDIENT_ID")
    Dim oItemCode As Long: oItemCode = ColumnIndex(loItems, "ITEM_CODE")

    Dim wsInv As Worksheet: Set wsInv = SheetExists("InventoryManagement")
    Dim loInv As ListObject
    If Not wsInv Is Nothing Then Set loInv = GetListObject(wsInv, "invSys")

    Dim arr As Variant: arr = loPal.DataBodyRange.value
    Dim r As Long, writeRow As Long
    Dim normRec As String: normRec = NormalizeIdFirst(recipeId)
    Dim normIng As String: normIng = NormalizeIdLast(ingredientId)
    For r = 1 To UBound(arr, 1)
        If NormalizeIdFirst(NzStr(arr(r, cRec))) = normRec And NormalizeIdLast(NzStr(arr(r, cIng))) = normIng Then
            Dim systemKey As String
            If cSystemKey > 0 Then systemKey = Trim$(NzStr(arr(r, cSystemKey)))
            Dim itemName As String: itemName = IIf(cItem > 0, NzStr(arr(r, cItem)), "")
            Dim uomVal As String: uomVal = IIf(cUom > 0, NzStr(arr(r, cUom)), "")
            Dim descVal As String: descVal = ""
            Dim itemCode As String: itemCode = IIf(cItemCode > 0, NzStr(arr(r, cItemCode)), "")

            If systemKey <> "" And Not loInv Is Nothing Then
                ResolveInvSysDetailsBySystemKey loInv, systemKey, itemName, uomVal, descVal
            End If

            writeRow = writeRow + 1
            EnsureListObjectRowCount loItems, writeRow
            If oItem > 0 Then loItems.DataBodyRange.Cells(writeRow, oItem).value = itemName
            If oUom > 0 Then loItems.DataBodyRange.Cells(writeRow, oUom).value = uomVal
            If oDesc > 0 Then loItems.DataBodyRange.Cells(writeRow, oDesc).value = descVal
            If oSystemKey > 0 Then loItems.DataBodyRange.Cells(writeRow, oSystemKey).value = systemKey
            If oItemCode > 0 Then loItems.DataBodyRange.Cells(writeRow, oItemCode).Value = itemCode
            If oRec > 0 Then loItems.DataBodyRange.Cells(writeRow, oRec).value = recipeId
            If oIng > 0 Then loItems.DataBodyRange.Cells(writeRow, oIng).value = ingredientId
        End If
    Next r
End Sub

Private Sub EnsureListObjectRowCount(ByVal lo As ListObject, ByVal needed As Long)
    If lo Is Nothing Then Exit Sub
    If needed < 1 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add AlwaysInsert:=True
    End If
    Do While lo.ListRows.count < needed
        lo.ListRows.Add AlwaysInsert:=True
    Loop
End Sub

Private Function EnsureListObjectRowCountSafe(ByVal lo As ListObject, ByVal needed As Long) As Boolean
    EnsureListObjectRowCountSafe = True
    If lo Is Nothing Then Exit Function
    If needed < 1 Then Exit Function

    On Error Resume Next
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add AlwaysInsert:=True
        If Err.Number <> 0 Then
            EnsureListObjectRowCountSafe = False
            Err.Clear
            Exit Function
        End If
    End If
    Do While lo.ListRows.count < needed
        lo.ListRows.Add AlwaysInsert:=True
        If Err.Number <> 0 Then
            EnsureListObjectRowCountSafe = False
            Err.Clear
            Exit Function
        End If
    Loop
    On Error GoTo 0
End Function

Private Function ExpandProductionInputOutputBand(ByVal ws As Worksheet, ByVal loCheck As ListObject, ByVal rowsAdded As Long) As Boolean
    ExpandProductionInputOutputBand = False
    If ws Is Nothing Then Exit Function
    If loCheck Is Nothing Then Exit Function
    If rowsAdded <= 0 Then Exit Function

    Dim bandLeft As Long
    Dim bandRight As Long
    Dim lo As ListObject

    Dim sCol As Long
    Dim eCol As Long
    If TableEffectiveSpan(loCheck, sCol, eCol) Then
        bandLeft = sCol
        bandRight = eCol
    Else
        bandLeft = loCheck.Range.Column
        bandRight = loCheck.Range.Column + loCheck.Range.Columns.count - 1
    End If

    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(ws, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If Not loOut Is Nothing Then
        If TableEffectiveSpan(loOut, sCol, eCol) Then
            If sCol < bandLeft Then bandLeft = sCol
            If eCol > bandRight Then bandRight = eCol
        End If
    End If

    For Each lo In ws.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
                If TableEffectiveSpan(lo, sCol, eCol) Then
                    If sCol < bandLeft Then bandLeft = sCol
                    If eCol > bandRight Then bandRight = eCol
                End If
            End If
        End If
    Next lo

    If bandLeft = 0 Or bandRight = 0 Then Exit Function

    Dim insertTop As Long
    insertTop = loCheck.Range.row + loCheck.Range.rows.count
    If insertTop <= 0 Then Exit Function
    If insertTop + rowsAdded - 1 > ws.rows.count Then Exit Function

    On Error Resume Next
    ws.rows(insertTop).Resize(rowsAdded).Insert Shift:=xlShiftDown
    If Err.Number = 0 Then ExpandProductionInputOutputBand = True
    Err.Clear
    On Error GoTo 0
End Function

Private Function ExpandProductionOutputBand(ByVal ws As Worksheet, ByVal loOut As ListObject, ByVal rowsAdded As Long) As Boolean
    ExpandProductionOutputBand = False
    If ws Is Nothing Then Exit Function
    If loOut Is Nothing Then Exit Function
    If rowsAdded <= 0 Then Exit Function

    Dim bandLeft As Long
    Dim bandRight As Long
    Dim lo As ListObject

    Dim sCol As Long
    Dim eCol As Long
    If TableEffectiveSpan(loOut, sCol, eCol) Then
        bandLeft = sCol
        bandRight = eCol
    Else
        bandLeft = loOut.Range.Column
        bandRight = loOut.Range.Column + loOut.Range.Columns.count - 1
    End If

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(ws, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loCheck Is Nothing Then
        If TableEffectiveSpan(loCheck, sCol, eCol) Then
            If sCol < bandLeft Then bandLeft = sCol
            If eCol > bandRight Then bandRight = eCol
        End If
    End If

    For Each lo In ws.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
                If TableEffectiveSpan(lo, sCol, eCol) Then
                    If sCol < bandLeft Then bandLeft = sCol
                    If eCol > bandRight Then bandRight = eCol
                End If
            End If
        End If
    Next lo

    If bandLeft = 0 Or bandRight = 0 Then Exit Function

    Dim insertTop As Long
    insertTop = loOut.Range.row + loOut.Range.rows.count
    If insertTop <= 0 Then Exit Function
    If insertTop + rowsAdded - 1 > ws.rows.count Then Exit Function

    On Error Resume Next
    ws.rows(insertTop).Resize(rowsAdded).Insert Shift:=xlShiftDown
    If Err.Number = 0 Then ExpandProductionOutputBand = True
    Err.Clear
    On Error GoTo 0
End Function

Private Function ExpandListObjectRows(ByVal lo As ListObject, ByVal addRows As Long) As Boolean
    ExpandListObjectRows = False
    If lo Is Nothing Then Exit Function
    If addRows <= 0 Then Exit Function

    Dim ws As Worksheet
    Set ws = lo.Parent

    Dim baseRange As Range
    Set baseRange = lo.Range
    If baseRange Is Nothing Then Exit Function

    Dim newRowCount As Long
    newRowCount = baseRange.rows.count + addRows
    If baseRange.row + newRowCount - 1 > ws.rows.count Then Exit Function

    Dim newRange As Range
    Set newRange = baseRange.Resize(newRowCount, baseRange.Columns.count)

    On Error Resume Next
    lo.Resize newRange
    If Err.Number = 0 Then ExpandListObjectRows = True
    Err.Clear
    On Error GoTo 0
End Function

Private Sub EnsureListObjectRowCountFullRow(ByVal lo As ListObject, ByVal needed As Long)
    ' Expand table capacity without inserting worksheet rows; Production surfaces use reserved bands.
    If lo Is Nothing Then Exit Sub
    If needed < 1 Then Exit Sub

    Dim currentRows As Long
    If lo.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = lo.DataBodyRange.rows.count
    End If
    If currentRows >= needed Then Exit Sub

    Dim addRows As Long
    addRows = needed - currentRows

    If Not ExpandListObjectRows(lo, addRows) Then
        On Error Resume Next
        Do While lo.ListRows.Count < needed
            lo.ListRows.Add AlwaysInsert:=False
            If Err.Number <> 0 Then Exit Do
        Loop
        Err.Clear
        On Error GoTo 0
    End If

    Dim newRange As Range
    If lo.ListRows.Count < needed Then
        Set newRange = lo.Range.Resize(needed + 1, lo.Range.Columns.count)
        On Error Resume Next
        lo.Resize newRange
        Err.Clear
        On Error GoTo 0
    End If
End Sub

Private Sub ResolveInvSysDetailsBySystemKey(ByVal loInv As ListObject, _
                                           ByVal systemKey As String, _
                                           ByRef itemName As String, _
                                           ByRef uomVal As String, _
                                           ByRef descVal As String)

    If loInv Is Nothing Then Exit Sub
    systemKey = Trim$(systemKey)
    If systemKey = "" Then Exit Sub
    If loInv.DataBodyRange Is Nothing Then Exit Sub

    Dim cSystemKey As Long: cSystemKey = ColumnIndex(loInv, "System_Key")
    Dim cItem As Long: cItem = ColumnIndex(loInv, "ITEM")
    Dim cUom As Long: cUom = ColumnIndex(loInv, "UOM")
    Dim cDesc As Long: cDesc = ColumnIndex(loInv, "DESCRIPTION")
    If cSystemKey = 0 Then Exit Sub

    Dim cel As Range
    For Each cel In loInv.ListColumns(cSystemKey).DataBodyRange.Cells
        If StrComp(Trim$(NzStr(cel.value)), systemKey, vbTextCompare) = 0 Then
            If itemName = "" And cItem > 0 Then itemName = NzStr(cel.Offset(0, cItem - cel.Column).value)
            If uomVal = "" And cUom > 0 Then uomVal = NzStr(cel.Offset(0, cUom - cel.Column).value)
            If descVal = "" And cDesc > 0 Then descVal = NzStr(cel.Offset(0, cDesc - cel.Column).value)
            Exit Sub
        End If
    Next cel
End Sub

Public Function GetPaletteRecipeId() As String
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Function
    Dim loRecipe As ListObject
    Set loRecipe = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseRecipe", Array("RECIPE_NAME", "RECIPE_ID"))
    If loRecipe Is Nothing Then Exit Function
    GetPaletteRecipeId = NormalizeIdFirst(FirstNonEmptyColumnValue(loRecipe, "RECIPE_ID"))
End Function

Private Function GetRecipeBuilderRecipeId(ByVal wsProd As Worksheet, Optional ByVal allowGenerate As Boolean = False) As String
    If wsProd Is Nothing Then Exit Function
    Dim loHeader As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    If loHeader Is Nothing Then Exit Function

    Dim idCell As Range: Set idCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
    Dim nameCell As Range: Set nameCell = GetHeaderDataCell(loHeader, "RECIPE_NAME")
    Dim recipeId As String
    If Not idCell Is Nothing Then recipeId = NzStr(idCell.value)

    If recipeId = "" And allowGenerate Then
        If Not nameCell Is Nothing Then
            If Trim$(NzStr(nameCell.value)) <> "" Then
                recipeId = GenerateRecipeId(wsProd.Parent)
                If Not idCell Is Nothing Then idCell.value = recipeId
            End If
        End If
    End If

    GetRecipeBuilderRecipeId = NormalizeIdFirst(recipeId)
End Function

Private Function ResolveActiveRecipeId(ByVal wsProd As Worksheet, Optional ByVal allowGenerate As Boolean = False) As String
    Dim recipeId As String
    recipeId = GetRecipeBuilderRecipeId(wsProd, allowGenerate)
    If recipeId <> "" Then
        ResolveActiveRecipeId = recipeId
        Exit Function
    End If

    recipeId = GetPaletteRecipeId()
    If recipeId <> "" Then
        ResolveActiveRecipeId = recipeId
        Exit Function
    End If

    recipeId = GetRecipeChooserRecipeId(wsProd)
    ResolveActiveRecipeId = recipeId
End Function

Public Function GetPaletteIngredientId() As String
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Function
    Dim loIng As ListObject
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    If loIng Is Nothing Then Exit Function
    GetPaletteIngredientId = NormalizeIdLast(FirstNonEmptyColumnValue(loIng, "INGREDIENT_ID"))
End Function

Public Function LoadIngredientListForRecipe(ByVal recipeId As String) As Variant
    Dim ingredients As Variant

    If ProductionDesignsEnabled() Then
        ingredients = LoadIngredientListFromDesigns(recipeId)
        If IsUsableProductionArray(ingredients) Then _
            LoadIngredientListForRecipe = ingredients
        Exit Function
    End If
    LoadIngredientListForRecipe = LoadLegacyIngredientListForRecipe(recipeId)
End Function

Private Function LoadLegacyIngredientListForRecipe(ByVal recipeId As String) As Variant
    Dim syncReport As String
    Dim wbOps As Workbook
    Set wbOps = ResolveProductionWorkbook(, "Recipes")
    If Not LocalProductionRecipeRowsExist(wbOps, recipeId) Then RefreshProductionRecipesFromRuntime wbOps, syncReport

    Dim wsRec As Worksheet: Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then Exit Function
    Dim lo As ListObject: Set lo = GetListObject(wsRec, "Recipes")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If Trim$(recipeId) = "" Then Exit Function

    Dim cRecId As Long: cRecId = ColumnIndex(lo, "RECIPE_ID")
    Dim cIngId As Long: cIngId = ColumnIndex(lo, "INGREDIENT_ID")
    Dim cIng As Long: cIng = ColumnIndex(lo, "INGREDIENT")
    Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
    Dim cProc As Long: cProc = ColumnIndex(lo, "PROCESS")
    Dim cIO As Long: cIO = ColumnIndex(lo, "INPUT/OUTPUT")
    Dim cAmt As Long: cAmt = ColumnIndex(lo, "AMOUNT")
    Dim cPct As Long: cPct = ColumnIndex(lo, "PERCENT")
    If cRecId = 0 Or cIngId = 0 Or cIng = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = lo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId Then
            If cIO > 0 Then
                If StrComp(Trim$(NzStr(arr(r, cIO))), "INSTRUCTION", vbTextCompare) = 0 Then GoTo NextLegacyIngredient
            End If
            Dim key As String
            key = NzStr(arr(r, cIngId)) & "|" & NzStr(arr(r, cProc))
            If Not dict.Exists(key) Then
                Dim info(1 To 7) As Variant
                info(1) = NzStr(arr(r, cIngId))
                info(2) = NzStr(arr(r, cIng))
                If cUom > 0 Then info(3) = NzStr(arr(r, cUom)) Else info(3) = ""
                If cProc > 0 Then info(4) = NzStr(arr(r, cProc)) Else info(4) = ""
                If cIO > 0 Then info(5) = NzStr(arr(r, cIO)) Else info(5) = ""
                If cAmt > 0 Then info(6) = arr(r, cAmt) Else info(6) = ""
                If cPct > 0 Then info(7) = arr(r, cPct) Else info(7) = ""
                dict.Add key, info
            End If
        End If
NextLegacyIngredient:
    Next r

    If dict.count = 0 Then Exit Function
    Dim result() As Variant
    ReDim result(1 To dict.count, 1 To 7)
    Dim i As Long: i = 1
    Dim k As Variant
    For Each k In dict.keys
        Dim infoArr As Variant
        infoArr = dict(k)
        result(i, 1) = infoArr(1)
        result(i, 2) = infoArr(2)
        result(i, 3) = infoArr(3)
        result(i, 4) = infoArr(4)
        result(i, 5) = infoArr(5)
        result(i, 6) = infoArr(6)
        result(i, 7) = infoArr(7)
        i = i + 1
    Next k
    LoadLegacyIngredientListForRecipe = result
End Function

Private Function LoadRecipeListFromDesigns(ByVal statusFilter As String) As Variant
    On Error GoTo CleanFail

    Dim designs As Variant
    designs = modOperationsPrimitiveBridge.ListDesigns(statusFilter)
    LoadRecipeListFromDesigns = BuildRecipeListFromDesignRows(designs)
CleanFail:
End Function

Private Function BuildRecipeListFromDesignRows(ByVal designs As Variant) As Variant
    On Error GoTo CleanFail

    Dim latest As Object
    Dim info As Variant
    Dim result() As Variant
    Dim r As Long
    Dim i As Long
    Dim designId As String
    Dim designVersion As String
    Dim designType As String
    Dim key As Variant

    If Not IsUsableProductionArray(designs) Then Exit Function

    Set latest = CreateObject("Scripting.Dictionary")
    latest.CompareMode = vbTextCompare
    For r = LBound(designs, 1) To UBound(designs, 1)
        designId = CanonicalRecipeIdProduction(designs(r, 1))
        designVersion = Trim$(NzStr(designs(r, 2)))
        designType = UCase$(Trim$(NzStr(designs(r, 3))))
        If designId = "" Or designVersion = "" Then GoTo NextDesign
        If designType <> "" And designType <> "RECIPE" Then GoTo NextDesign

        ' Designs rows are returned in authoritative replay/apply order.
        ' The last row for an ID is the latest immutable version; version text
        ' itself is not an ordering contract (for example v1 versus timestamps).
        If latest.Exists(designId) Then
            latest(designId) = Array(designVersion, NzStr(designs(r, 4)), NzStr(designs(r, 5)))
        Else
            latest.Add designId, Array(designVersion, NzStr(designs(r, 4)), NzStr(designs(r, 5)))
        End If
NextDesign:
    Next r

    If latest.Count = 0 Then Exit Function
    ReDim result(1 To latest.Count, 1 To 3)
    i = 1
    For Each key In latest.Keys
        info = latest(key)
        result(i, 1) = CStr(key)
        result(i, 2) = info(1)
        result(i, 3) = info(2)
        i = i + 1
    Next key
    BuildRecipeListFromDesignRows = result
CleanFail:
End Function

'@TestOnlyBegin
Public Function TestLatestRecipeNameFromDesignRows(ByVal designs As Variant, _
                                                   ByVal recipeId As String) As String
    Dim recipes As Variant
    Dim r As Long

    recipes = BuildRecipeListFromDesignRows(designs)
    If Not IsUsableProductionArray(recipes) Then Exit Function
    For r = LBound(recipes, 1) To UBound(recipes, 1)
        If StrComp(Trim$(NzStr(recipes(r, 1))), Trim$(recipeId), vbTextCompare) = 0 Then
            TestLatestRecipeNameFromDesignRows = NzStr(recipes(r, 2))
            Exit Function
        End If
    Next r
End Function
'@TestOnlyEnd

Private Function ValueOrPlaceholderProduction(ByVal preferredValue As String, _
                                              ByVal fallbackValue As String) As String
    preferredValue = Trim$(preferredValue)
    If preferredValue <> "" Then
        ValueOrPlaceholderProduction = preferredValue
    Else
        ValueOrPlaceholderProduction = Trim$(fallbackValue)
    End If
End Function

Private Function LoadIngredientListFromDesigns(ByVal recipeId As String) As Variant
    On Error GoTo CleanFail

    Dim designVersion As String
    Dim designName As String
    Dim designDescription As String
    Dim designStatus As String
    Dim bom As Variant
    Dim result() As Variant
    Dim r As Long
    Dim ingredientId As String
    Dim ingredientName As String
    Dim outRow As Long
    Dim trimmed() As Variant
    Dim c As Long

    recipeId = Trim$(recipeId)
    If recipeId = "" Then Exit Function
    If Not ProductionDesignsEnabled() Then Exit Function
    If Not FindLatestDesignSummaryProduction( _
        recipeId, "RELEASED", designVersion, designName, _
        designDescription, designStatus) Then Exit Function

    bom = modOperationsPrimitiveBridge.GetDesignBom(recipeId, designVersion)
    If Not IsUsableProductionArray(bom) Then Exit Function

    ReDim result(1 To UBound(bom, 1) - LBound(bom, 1) + 1, 1 To 7)
    For r = LBound(bom, 1) To UBound(bom, 1)
        If StrComp(Trim$(NzStr(bom(r, 3))), "INSTRUCTION", vbTextCompare) = 0 Then GoTo NextBomRow
        outRow = outRow + 1
        ingredientId = Trim$(NzStr(bom(r, 5)))
        If ingredientId = "" Then ingredientId = Trim$(NzStr(bom(r, 4)))
        ingredientName = Trim$(NzStr(bom(r, 10)))
        If ingredientName = "" Then ingredientName = ingredientId

        result(outRow, 1) = ingredientId
        result(outRow, 2) = ingredientName
        result(outRow, 3) = bom(r, 8)
        result(outRow, 4) = bom(r, 2)
        result(outRow, 5) = bom(r, 3)
        result(outRow, 6) = bom(r, 7)
        result(outRow, 7) = bom(r, 9)
NextBomRow:
    Next r
    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 7)
    For r = 1 To outRow
        For c = 1 To 7
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    LoadIngredientListFromDesigns = trimmed
CleanFail:
End Function

Private Function FindLatestDesignSummaryProduction(ByVal designId As String, _
                                                   ByVal statusFilter As String, _
                                                   ByRef designVersion As String, _
                                                   ByRef designName As String, _
                                                   ByRef designDescription As String, _
                                                   ByRef designStatus As String) As Boolean
    On Error GoTo CleanFail

    Dim designs As Variant
    Dim r As Long

    If Not ProductionDesignsEnabled() Then Exit Function
    designs = modOperationsPrimitiveBridge.ListDesigns(statusFilter)
    If Not IsUsableProductionArray(designs) Then Exit Function
    For r = LBound(designs, 1) To UBound(designs, 1)
        If RecipeIdsMatchProduction(designs(r, 1), designId) Then
            designVersion = Trim$(NzStr(designs(r, 2)))
            designName = NzStr(designs(r, 4))
            designDescription = NzStr(designs(r, 5))
            designStatus = NzStr(designs(r, 6))
        End If
    Next r
    FindLatestDesignSummaryProduction = (designVersion <> "")
CleanFail:
End Function

Private Function IsUsableProductionArray(ByVal values As Variant) As Boolean
    On Error GoTo CleanFail
    If IsEmpty(values) Then Exit Function
    If Not IsArray(values) Then Exit Function
    IsUsableProductionArray = (UBound(values, 1) >= LBound(values, 1))
CleanFail:
End Function

Private Sub ResolveProductionTargetIdentity(ByRef warehouseId As String, _
                                            ByRef stationId As String)
    warehouseId = Trim$(modNasConnection.GetCurrentTargetWarehouseId())
    stationId = Trim$(modNasConnection.GetCurrentTargetStationId())
End Sub

Private Function CurrentProductionWarehouseId() As String
    CurrentProductionWarehouseId = Trim$(modNasConnection.GetCurrentTargetWarehouseId())
End Function

Private Function ProductionDesignsEnabled() As Boolean
    On Error GoTo CleanFail
    Dim targetWarehouseId As String
    Dim targetStationId As String

    ResolveProductionTargetIdentity targetWarehouseId, targetStationId
    If Not modConfig.IsLoaded() _
       Or (targetWarehouseId <> "" And StrComp(modConfig.GetWarehouseId(), targetWarehouseId, vbTextCompare) <> 0) Then
        If Not modConfig.LoadConfig(targetWarehouseId, targetStationId) Then Exit Function
    End If
    ProductionDesignsEnabled = modConfig.GetBool("DesignsEnabled", False) _
                               Or modConfig.GetBool("FF_DesignsEnabled", False)
CleanFail:
End Function

Public Function GetProductionInventoryModeStatus() As String
    On Error GoTo FailStatus
    GetProductionInventoryModeStatus = "Inventory: " & modInventoryDomainBridge.DiagnoseInventoryDomainBridge()
    Exit Function
FailStatus:
    GetProductionInventoryModeStatus = "Inventory: unavailable (" & Err.Description & ")."
End Function

Public Function GetProductionDesignsModeStatus() As String
    On Error GoTo FailStatus
    If Not ProductionDesignsEnabled() Then
        GetProductionDesignsModeStatus = "Designs: legacy recipe fallback (disabled in warehouse config)."
        Exit Function
    End If
    GetProductionDesignsModeStatus = "Designs: " & modDesignsDomainBridge.DiagnoseDesignsDomainBridge()
    Exit Function
FailStatus:
    GetProductionDesignsModeStatus = "Designs: unavailable (" & Err.Description & ")."
End Function

Private Function QueueSavedRecipeDesignCreate(ByVal loRecipes As ListObject, _
                                              ByVal recipeId As String, _
                                              ByVal recipeName As String, _
                                              ByVal recipeDescription As String, _
                                              ByRef designVersion As String, _
                                              ByRef report As String) As Boolean
    On Error GoTo FailQueue

    Dim payloadRows As Collection
    Dim payloadItem As Object
    Dim payloadJson As String
    Dim eventId As String
    Dim queueError As String
    Dim processorReport As String
    Dim appliedCount As Long
    Dim r As Long
    Dim cRecipeId As Long
    Dim cProcess As Long
    Dim cIo As Long
    Dim cIngredient As Long
    Dim cIngredientId As Long
    Dim cAmount As Long
    Dim cUom As Long
    Dim cPercent As Long
    Dim cLineNo As Long
    Dim lineNo As Long
    Dim targetWarehouseId As String
    Dim targetStationId As String

    If loRecipes Is Nothing Or loRecipes.DataBodyRange Is Nothing Then
        report = "No saved recipe rows were available for the Designs event."
        Exit Function
    End If

    designVersion = NextRecipeDesignVersionProduction(recipeId)
    If designVersion = "" Then designVersion = "1"

    cRecipeId = ColumnIndex(loRecipes, "RECIPE_ID")
    cProcess = ColumnIndex(loRecipes, "PROCESS")
    cIo = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    cIngredient = ColumnIndex(loRecipes, "INGREDIENT")
    cIngredientId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    cAmount = ColumnIndex(loRecipes, "AMOUNT")
    cUom = ColumnIndex(loRecipes, "UOM")
    cPercent = ColumnIndex(loRecipes, "PERCENT")
    cLineNo = ColumnIndex(loRecipes, "RECIPE_LIST_ROW")
    If cRecipeId = 0 Then
        report = "Recipes table is missing RECIPE_ID."
        Exit Function
    End If

    Set payloadRows = New Collection
    For r = 1 To loRecipes.ListRows.Count
        If StrComp(Trim$(NzStr(loRecipes.DataBodyRange.Cells(r, cRecipeId).Value)), _
                   Trim$(recipeId), vbTextCompare) = 0 Then
            Set payloadItem = CreateObject("Scripting.Dictionary")
            payloadItem.CompareMode = vbTextCompare
            If payloadRows.Count = 0 Then
                payloadItem("DesignType") = "RECIPE"
                payloadItem("DesignName") = recipeName
                payloadItem("Description") = recipeDescription
            End If

            lineNo = payloadRows.Count + 1
            If cLineNo > 0 Then
                If IsNumeric(loRecipes.DataBodyRange.Cells(r, cLineNo).Value) Then
                    lineNo = CLng(loRecipes.DataBodyRange.Cells(r, cLineNo).Value)
                End If
            End If
            payloadItem("LineNo") = lineNo
            If cProcess > 0 Then payloadItem("Process") = NzStr(loRecipes.DataBodyRange.Cells(r, cProcess).Value)
            If cIo > 0 Then payloadItem("IOType") = NzStr(loRecipes.DataBodyRange.Cells(r, cIo).Value)
            If cIngredientId > 0 Then payloadItem("ComponentDesignId") = NzStr(loRecipes.DataBodyRange.Cells(r, cIngredientId).Value)
            If cAmount > 0 Then payloadItem("Qty") = NzDbl(loRecipes.DataBodyRange.Cells(r, cAmount).Value)
            If cUom > 0 Then payloadItem("UOM") = NzStr(loRecipes.DataBodyRange.Cells(r, cUom).Value)
            If cPercent > 0 Then payloadItem("Percent") = NzDbl(loRecipes.DataBodyRange.Cells(r, cPercent).Value)
            If cIngredient > 0 Then payloadItem("Instruction") = NzStr(loRecipes.DataBodyRange.Cells(r, cIngredient).Value)
            payloadRows.Add payloadItem
        End If
    Next r

    If payloadRows.Count = 0 Then
        report = "No saved rows matched recipe " & recipeId & "."
        Exit Function
    End If

    payloadJson = modProductionJson.BuildJsonArray(payloadRows)
    If Not modRoleEventWriter.QueueDesignEventCurrent(EVENT_TYPE_DESIGN_CREATE, _
                                                      recipeId, _
                                                      designVersion, _
                                                      payloadJson, _
                                                      "Production Recipe Builder save", _
                                                      "", _
                                                      eventId, _
                                                      queueError) Then
        report = queueError
        Exit Function
    End If

    ResolveProductionTargetIdentity targetWarehouseId, targetStationId
    appliedCount = modProcessor.RunBatch(targetWarehouseId, 0, processorReport)
    If targetWarehouseId <> "" Then
        Call modConfig.LoadConfig(targetWarehouseId, targetStationId)
    End If
    If DesignVersionExistsProduction(recipeId, designVersion) Then
        report = "applied; EventID=" & eventId
        QueueSavedRecipeDesignCreate = True
    Else
        report = "queued but not yet applied; EventID=" & eventId
        If Trim$(processorReport) <> "" Then report = report & "; Processor=" & processorReport
    End If
    Exit Function

FailQueue:
    report = "QueueSavedRecipeDesignCreate failed: " & Err.Description
End Function

Private Function NextRecipeDesignVersionProduction(ByVal recipeId As String) As String
    On Error GoTo CleanFail

    Dim designs As Variant
    Dim r As Long
    Dim maxNumericVersion As Long
    Dim candidate As String
    Dim foundNonNumeric As Boolean
    Dim stagedDesigns As Variant
    Dim warehouseId As String

    designs = modOperationsPrimitiveBridge.ListDesigns("")
    If IsUsableProductionArray(designs) Then
        For r = LBound(designs, 1) To UBound(designs, 1)
            If RecipeIdsMatchProduction(designs(r, 1), recipeId) Then
                candidate = Trim$(NzStr(designs(r, 2)))
                If IsNumeric(candidate) Then
                    If CLng(CDbl(candidate)) > maxNumericVersion Then maxNumericVersion = CLng(CDbl(candidate))
                ElseIf candidate <> "" Then
                    foundNonNumeric = True
                End If
            End If
        Next r
    End If

    warehouseId = CurrentProductionWarehouseId()
    stagedDesigns = modRoleEventWriter.GetLocalStagedDesignIdentities(warehouseId)
    If IsUsableProductionArray(stagedDesigns) Then
        For r = LBound(stagedDesigns, 1) To UBound(stagedDesigns, 1)
            If RecipeIdsMatchProduction(stagedDesigns(r, 1), recipeId) Then
                candidate = Trim$(NzStr(stagedDesigns(r, 2)))
                If IsNumeric(candidate) Then
                    If CLng(CDbl(candidate)) > maxNumericVersion Then maxNumericVersion = CLng(CDbl(candidate))
                ElseIf candidate <> "" Then
                    foundNonNumeric = True
                End If
            End If
        Next r
    End If

    If foundNonNumeric And maxNumericVersion = 0 Then
        NextRecipeDesignVersionProduction = Format$(Now, "yyyymmddhhnnss")
    Else
        NextRecipeDesignVersionProduction = CStr(maxNumericVersion + 1)
    End If
    Exit Function
CleanFail:
    NextRecipeDesignVersionProduction = "1"
End Function

Private Function DesignVersionExistsProduction(ByVal designId As String, ByVal designVersion As String) As Boolean
    On Error GoTo CleanFail

    Dim designs As Variant
    Dim r As Long

    designs = modOperationsPrimitiveBridge.ListDesigns("")
    If Not IsUsableProductionArray(designs) Then Exit Function
    For r = LBound(designs, 1) To UBound(designs, 1)
        If RecipeIdsMatchProduction(designs(r, 1), designId) _
           And StrComp(Trim$(NzStr(designs(r, 2))), Trim$(designVersion), vbTextCompare) = 0 Then
            DesignVersionExistsProduction = True
            Exit Function
        End If
    Next r
CleanFail:
End Function

Private Function BuildReleasedDesignRecipeStagingWorkbook(ByVal recipeId As String, _
                                                          ByRef report As String) As Workbook
    Set BuildReleasedDesignRecipeStagingWorkbook = _
        BuildDesignRecipeStagingWorkbook(recipeId, "RELEASED", report)
End Function

Private Function BuildDesignRecipeStagingWorkbook(ByVal recipeId As String, _
                                                   ByVal statusFilter As String, _
                                                   ByRef report As String) As Workbook
    On Error GoTo FailBuild

    Dim designVersion As String
    Dim recipeName As String
    Dim recipeDescription As String
    Dim designStatus As String
    Dim bom As Variant

    If Not FindLatestDesignSummaryProduction(recipeId, statusFilter, designVersion, _
                                             recipeName, recipeDescription, designStatus) Then
        report = "No matching Designs Domain recipe was found for " & recipeId & "."
        Exit Function
    End If
    If Trim$(statusFilter) = "" Then
        bom = modOperationsPrimitiveBridge.GetDesignBom(recipeId, designVersion)
    Else
        bom = modOperationsPrimitiveBridge.GetDesignBomForStatus( _
            recipeId, designVersion, statusFilter)
    End If
    If Not IsUsableProductionArray(bom) Then
        report = "Design " & recipeId & " version " & designVersion & " has no available BOM lines."
        Exit Function
    End If

    Set BuildDesignRecipeStagingWorkbook = _
        BuildDesignRecipeStagingWorkbookFromData(recipeId, recipeName, recipeDescription, bom, report)
    Exit Function

FailBuild:
    report = "BuildDesignRecipeStagingWorkbook failed: " & Err.Description
End Function

Private Function BuildPendingDesignRecipeStagingWorkbook(ByVal recipeId As String, _
                                                         ByRef report As String) As Workbook
    On Error GoTo FailBuild

    Dim warehouseId As String
    Dim staged As Variant
    Dim payloadJson As String
    Dim recipeName As String
    Dim recipeDescription As String
    Dim rx As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim bom() As Variant
    Dim objectText As String
    Dim r As Long
    Dim candidateId As String

    warehouseId = CurrentProductionWarehouseId()
    staged = modRoleEventWriter.GetLocalStagedDesignIdentities(warehouseId)
    If Not IsUsableProductionArray(staged) Then Exit Function

    For r = LBound(staged, 1) To UBound(staged, 1)
        candidateId = CanonicalRecipeIdProduction(staged(r, 1))
        If StrComp(candidateId, CanonicalRecipeIdProduction(recipeId), vbTextCompare) = 0 Then
            If UBound(staged, 2) >= 3 Then recipeName = NzStr(staged(r, 3))
            If UBound(staged, 2) >= 4 Then recipeDescription = NzStr(staged(r, 4))
            If UBound(staged, 2) >= 5 Then payloadJson = NzStr(staged(r, 5))
        End If
    Next r
    If Trim$(payloadJson) = "" Then Exit Function

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "\{[^{}]*\}"
    Set matches = rx.Execute(payloadJson)
    If matches.Count = 0 Then Exit Function

    ReDim bom(1 To matches.Count, 1 To 10)
    r = 1
    For Each matchItem In matches
        objectText = CStr(matchItem.Value)
        bom(r, 1) = JsonPayloadNumberField(objectText, "LineNo")
        bom(r, 2) = JsonPayloadNumberField(objectText, "Process")
        bom(r, 3) = JsonPayloadStringField(objectText, "IOType")
        bom(r, 4) = JsonPayloadStringField(objectText, "ComponentDesignId")
        bom(r, 5) = bom(r, 4)
        bom(r, 7) = JsonPayloadNumberField(objectText, "Qty")
        bom(r, 8) = JsonPayloadStringField(objectText, "UOM")
        bom(r, 9) = JsonPayloadNumberField(objectText, "Percent")
        bom(r, 10) = JsonPayloadStringField(objectText, "Instruction")
        r = r + 1
    Next matchItem

    Set BuildPendingDesignRecipeStagingWorkbook = _
        BuildDesignRecipeStagingWorkbookFromData( _
            CanonicalRecipeIdProduction(recipeId), recipeName, recipeDescription, bom, report)
    Exit Function

FailBuild:
    report = "BuildPendingDesignRecipeStagingWorkbook failed: " & Err.Description
End Function

Private Function BuildLegacyRuntimeRecipeStagingWorkbook(ByVal recipeId As String, _
                                                         ByRef report As String) As Workbook
    On Error GoTo FailBuild

    Dim warehouseId As String
    Dim rootPath As String
    Dim wbRuntime As Workbook
    Dim wsRuntime As Worksheet
    Dim loRuntime As ListObject
    Dim openedTransient As Boolean
    Dim wbStaging As Workbook
    Dim wsStaging As Worksheet
    Dim loStaging As ListObject
    Dim copied As Long

    If Not ResolveProductionRecipesStorageTarget(warehouseId, rootPath, report) Then Exit Function
    Set wbRuntime = OpenProductionRecipesWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbRuntime Is Nothing Then Exit Function
    Set wsRuntime = WorkbookSheetExists(wbRuntime, SHEET_RUNTIME_RECIPES)
    If wsRuntime Is Nothing Then GoTo CleanExit
    Set loRuntime = GetListObject(wsRuntime, TABLE_RUNTIME_RECIPES)
    If loRuntime Is Nothing Then GoTo CleanExit

    Set wbStaging = CreateEmptyDesignRecipeStagingWorkbook(report)
    If wbStaging Is Nothing Then GoTo CleanExit
    Set wsStaging = WorkbookSheetExists(wbStaging, "Recipes")
    If wsStaging Is Nothing Then GoTo CleanExit
    Set loStaging = GetListObject(wsStaging, "Recipes")
    If loStaging Is Nothing Then GoTo CleanExit
    copied = CopyProductionRecipeRowsById( _
        loRuntime, loStaging, CanonicalRecipeIdProduction(recipeId), False)
    If copied = 0 Then
        CloseWorkbookNoSaveProduction wbStaging
        Set wbStaging = Nothing
        GoTo CleanExit
    End If

    report = "Prepared transient legacy runtime staging for " & recipeId & "."
    Set BuildLegacyRuntimeRecipeStagingWorkbook = wbStaging

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveProduction wbRuntime
    Exit Function

FailBuild:
    report = "BuildLegacyRuntimeRecipeStagingWorkbook failed: " & Err.Description
    On Error Resume Next
    If Not wbStaging Is Nothing Then wbStaging.Close SaveChanges:=False
    If openedTransient Then CloseWorkbookNoSaveProduction wbRuntime
    On Error GoTo 0
End Function

Public Function BuildDesignRecipeStagingWorkbookFromData(ByVal recipeId As String, _
                                                         ByVal recipeName As String, _
                                                         ByVal recipeDescription As String, _
                                                         ByVal bom As Variant, _
                                                         Optional ByRef report As String = "") As Workbook
    On Error GoTo FailBuild

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim tableRange As Range
    Dim r As Long
    Dim c As Long
    Dim ingredientId As String
    Dim ingredientName As String

    If Trim$(recipeId) = "" Then
        report = "RecipeId is required for Designs staging."
        Exit Function
    End If
    If Not IsUsableProductionArray(bom) Then
        report = "A Designs BOM array is required for staging."
        Exit Function
    End If

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Recipes"
    headers = ProductionRecipeRuntimeHeaders()
    Set tableRange = ws.Range(ws.Cells(1, 1), _
        ws.Cells(UBound(bom, 1) - LBound(bom, 1) + 2, UBound(headers) - LBound(headers) + 1))
    For c = LBound(headers) To UBound(headers)
        tableRange.Cells(1, c - LBound(headers) + 1).Value = headers(c)
    Next c
    Set lo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    lo.Name = "Recipes"

    For r = LBound(bom, 1) To UBound(bom, 1)
        ingredientId = Trim$(NzStr(bom(r, 5)))
        If ingredientId = "" Then ingredientId = Trim$(NzStr(bom(r, 4)))
        ingredientName = Trim$(NzStr(bom(r, 10)))
        If ingredientName = "" Then ingredientName = ingredientId

        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "RECIPE_ID", recipeId
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "RECIPE", recipeName
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "DESCRIPTION", recipeDescription
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "ROW_BUDGET", PRODUCTION_DEFAULT_ROW_BUDGET
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "PROCESS", bom(r, 2)
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "INPUT/OUTPUT", bom(r, 3)
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "INGREDIENT", ingredientName
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "PERCENT", bom(r, 9)
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "UOM", bom(r, 8)
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "AMOUNT", bom(r, 7)
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "RECIPE_LIST_ROW", bom(r, 1)
        SetProductionTableCellByHeader lo, r - LBound(bom, 1) + 1, "INGREDIENT_ID", ingredientId
    Next r

    On Error Resume Next
    wb.Windows(1).Visible = False
    On Error GoTo FailBuild
    report = "Prepared transient Designs staging for " & recipeId & "."
    Set BuildDesignRecipeStagingWorkbookFromData = wb
    Exit Function

FailBuild:
    report = "BuildDesignRecipeStagingWorkbookFromData failed: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Private Function CreateEmptyDesignRecipeStagingWorkbook(ByRef report As String) As Workbook
    On Error GoTo FailBuild

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim c As Long

    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    Set ws = wb.Worksheets(1)
    ws.Name = "Recipes"
    headers = ProductionRecipeRuntimeHeaders()
    For c = LBound(headers) To UBound(headers)
        ws.Cells(1, c - LBound(headers) + 1).Value = headers(c)
    Next c
    Set lo = ws.ListObjects.Add(xlSrcRange, _
        ws.Range(ws.Cells(1, 1), ws.Cells(2, UBound(headers) - LBound(headers) + 1)), , xlYes)
    lo.Name = "Recipes"
    If Not lo.DataBodyRange Is Nothing Then lo.ListRows(1).Delete
    On Error Resume Next
    wb.Windows(1).Visible = False
    On Error GoTo FailBuild
    Set CreateEmptyDesignRecipeStagingWorkbook = wb
    Exit Function

FailBuild:
    report = "CreateEmptyDesignRecipeStagingWorkbook failed: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Public Function GetCurrentProductionRunRecipeId() As String
    Dim wsProd As Worksheet
    Set wsProd = SheetExists(SHEET_PRODUCTION)
    GetCurrentProductionRunRecipeId = GetRecipeChooserRecipeId(wsProd)
End Function

Public Function LoadProductionRunIngredientChoices(Optional ByVal recipeId As String = "") As Variant
    Dim ingredients As Variant
    Dim paletteRows As Variant
    Dim result() As Variant
    Dim ingRow As Long
    Dim palRow As Long
    Dim outRow As Long
    Dim matched As Boolean
    Dim ingredientId As String
    Dim ingredientName As String
    Dim processName As String
    Dim baseQty As Variant
    Dim uomVal As String
    Dim ioVal As String

    If Trim$(recipeId) = "" Then recipeId = GetCurrentProductionRunRecipeId()
    recipeId = NormalizeIdFirst(recipeId)
    If recipeId = "" Then Exit Function

    ingredients = LoadIngredientListForRecipe(recipeId)
    If IsEmpty(ingredients) Then Exit Function
    paletteRows = LoadIngredientPaletteRowsForRecipe(recipeId)

    ReDim result(1 To 5000, 1 To 11)
    For ingRow = LBound(ingredients, 1) To UBound(ingredients, 1)
        ingredientId = NzStr(ingredients(ingRow, 1))
        ingredientName = NzStr(ingredients(ingRow, 2))
        uomVal = NzStr(ingredients(ingRow, 3))
        processName = NzStr(ingredients(ingRow, 4))
        ioVal = UCase$(Trim$(NzStr(ingredients(ingRow, 5))))
        If ioVal <> "" And ioVal <> "USED" Then GoTo NextIngredient
        baseQty = ingredients(ingRow, 6)
        matched = False

        If Not IsEmpty(paletteRows) Then
            For palRow = LBound(paletteRows, 1) To UBound(paletteRows, 1)
                If ProductionIngredientIdsMatch(NzStr(paletteRows(palRow, 2)), ingredientId) Then
                    outRow = outRow + 1
                    result(outRow, 1) = processName
                    result(outRow, 2) = ingredientId
                    result(outRow, 3) = ingredientName
                    result(outRow, 4) = NzStr(paletteRows(palRow, 5))
                    result(outRow, 5) = NzStr(paletteRows(palRow, 4))
                    result(outRow, 6) = ""
                    result(outRow, 7) = baseQty
                    result(outRow, 8) = IIf(NzStr(paletteRows(palRow, 6)) <> "", NzStr(paletteRows(palRow, 6)), uomVal)
                    result(outRow, 9) = ""
                    result(outRow, 10) = baseQty
                    If UBound(paletteRows, 2) >= 7 Then result(outRow, 11) = NzStr(paletteRows(palRow, 7))
                    matched = True
                End If
            Next palRow
        End If

        If Not matched Then
            outRow = outRow + 1
            result(outRow, 1) = processName
            result(outRow, 2) = ingredientId
            result(outRow, 3) = ingredientName
            result(outRow, 4) = ""
            result(outRow, 5) = ""
            result(outRow, 6) = ""
            result(outRow, 7) = baseQty
            result(outRow, 8) = uomVal
            result(outRow, 9) = ""
            result(outRow, 10) = baseQty
            result(outRow, 11) = ""
        End If
NextIngredient:
    Next ingRow

    If outRow = 0 Then Exit Function
    LoadProductionRunIngredientChoices = TrimRunChoiceRows(result, outRow)
End Function

Private Function ProductionIngredientIdsMatch(ByVal savedIngredientId As String, ByVal recipeIngredientId As String) As Boolean
    savedIngredientId = Trim$(savedIngredientId)
    recipeIngredientId = Trim$(recipeIngredientId)
    If savedIngredientId = "" Or recipeIngredientId = "" Then Exit Function
    ProductionIngredientIdsMatch = (StrComp(savedIngredientId, recipeIngredientId, vbTextCompare) = 0)
End Function

Private Function LoadIngredientPaletteRowsForRecipe(ByVal recipeId As String) As Variant
    Dim syncReport As String
    RefreshProductionIngredientPaletteFromRuntime ResolveProductionWorkbook(, "IngredientPalette"), syncReport

    Dim wsPal As Worksheet
    Set wsPal = SheetExists("IngredientPalette")
    If wsPal Is Nothing Then Set wsPal = SheetExists("IngredientsPalette")
    If wsPal Is Nothing Then Exit Function

    Dim loPal As ListObject
    Set loPal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "System_Key"))
    If loPal Is Nothing Then Exit Function
    If loPal.DataBodyRange Is Nothing Then Exit Function

    Dim cRecipe As Long: cRecipe = ColumnIndex(loPal, "RECIPE_ID")
    Dim cIngredient As Long: cIngredient = ColumnIndex(loPal, "INGREDIENT_ID")
    Dim cItem As Long: cItem = ColumnIndex(loPal, "ITEM")
    Dim cRow As Long: cRow = ColumnIndex(loPal, "System_Key")
    Dim cUom As Long: cUom = ColumnIndex(loPal, "UOM")
    Dim cItemCode As Long: cItemCode = ColumnIndex(loPal, "ITEM_CODE")
    If cRecipe = 0 Or cIngredient = 0 Then Exit Function

    Dim arr As Variant: arr = loPal.DataBodyRange.Value
    Dim result() As Variant
    Dim r As Long
    Dim outRow As Long
    ReDim result(1 To UBound(arr, 1), 1 To 7)
    For r = 1 To UBound(arr, 1)
        If NormalizeIdFirst(NzStr(arr(r, cRecipe))) = NormalizeIdFirst(recipeId) Then
            outRow = outRow + 1
            result(outRow, 1) = NormalizeIdFirst(NzStr(arr(r, cRecipe)))
            result(outRow, 2) = NzStr(arr(r, cIngredient))
            If cRow > 0 Then result(outRow, 3) = NzStr(arr(r, cRow))
            If cItem > 0 Then result(outRow, 4) = NzStr(arr(r, cItem))
            If cRow > 0 Then result(outRow, 5) = NzStr(arr(r, cRow))
            If cUom > 0 Then result(outRow, 6) = NzStr(arr(r, cUom))
            If cItemCode > 0 Then result(outRow, 7) = NzStr(arr(r, cItemCode))
        End If
    Next r

    If outRow = 0 Then Exit Function
    HydrateIngredientPaletteIdentityRowsProduction result, outRow
    LoadIngredientPaletteRowsForRecipe = TrimIngredientPaletteRows(result, outRow)
End Function

Private Sub HydrateIngredientPaletteIdentityRowsProduction(ByRef paletteRows() As Variant, _
                                                           ByVal rowCount As Long, _
                                                           Optional ByVal inventoryRowsOverride As Variant)
    Dim inventoryRows As Variant
    Dim pr As Long
    Dim ir As Long
    Dim paletteItem As String
    Dim paletteUom As String
    Dim candidateItem As String
    Dim candidateUom As String
    Dim candidateCode As String
    Dim matchedCode As String
    Dim matchedRow As String
    Dim ambiguous As Boolean

    If rowCount <= 0 Then Exit Sub
    If IsMissing(inventoryRowsOverride) Then
        inventoryRows = LoadProductionRunInventoryPickerItems("")
    Else
        inventoryRows = inventoryRowsOverride
    End If
    If Not IsUsableProductionArray(inventoryRows) Then Exit Sub
    If UBound(inventoryRows, 2) < 7 Then Exit Sub

    For pr = 1 To rowCount
        If Trim$(NzStr(paletteRows(pr, 7))) <> "" Then GoTo NextPaletteRow
        paletteItem = Trim$(NzStr(paletteRows(pr, 4)))
        paletteUom = Trim$(NzStr(paletteRows(pr, 6)))
        If paletteItem = "" Or paletteUom = "" Then GoTo NextPaletteRow

        matchedCode = vbNullString
        matchedRow = vbNullString
        ambiguous = False
        For ir = LBound(inventoryRows, 1) To UBound(inventoryRows, 1)
            candidateItem = Trim$(NzStr(inventoryRows(ir, 2)))
            candidateUom = Trim$(NzStr(inventoryRows(ir, 3)))
            candidateCode = Trim$(NzStr(inventoryRows(ir, 7)))
            If candidateCode <> "" _
               And StrComp(candidateItem, paletteItem, vbTextCompare) = 0 _
               And StrComp(candidateUom, paletteUom, vbTextCompare) = 0 Then
                If matchedCode = "" Then
                    matchedCode = candidateCode
                    matchedRow = NzStr(inventoryRows(ir, 1))
                ElseIf StrComp(matchedCode, candidateCode, vbTextCompare) <> 0 Then
                    ambiguous = True
                    Exit For
                End If
            End If
        Next ir
        If Not ambiguous And matchedCode <> "" Then
            paletteRows(pr, 7) = matchedCode
            If Trim$(NzStr(paletteRows(pr, 5))) = "" Then paletteRows(pr, 5) = matchedRow
        End If
NextPaletteRow:
    Next pr
End Sub

Private Function TrimRunChoiceRows(ByVal result As Variant, ByVal rowCount As Long) As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    ReDim trimmed(1 To rowCount, 1 To 11)
    For r = 1 To rowCount
        For c = 1 To 11
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    TrimRunChoiceRows = trimmed
End Function

Private Function TrimIngredientPaletteRows(ByVal result As Variant, ByVal rowCount As Long) As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    ReDim trimmed(1 To rowCount, 1 To 7)
    For r = 1 To rowCount
        For c = 1 To 7
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    TrimIngredientPaletteRows = trimmed
End Function

' ===== System 3: Recipe Chooser - data helpers =====
Private Sub GetRecipeSummary(ByVal wsRec As Worksheet, ByVal recipeId As String, _
    ByRef recipeName As String, ByRef recipeDesc As String, ByRef recipeDept As String)

    recipeName = ""
    recipeDesc = ""
    recipeDept = ""
    If wsRec Is Nothing Then Exit Sub

    Dim lo As ListObject: Set lo = GetListObject(wsRec, "Recipes")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cRecId As Long: cRecId = ColumnIndex(lo, "RECIPE_ID")
    Dim cRec As Long: cRec = ColumnIndex(lo, "RECIPE")
    Dim cDesc As Long: cDesc = ColumnIndex(lo, "DESCRIPTION")
    Dim cDept As Long: cDept = ColumnIndex(lo, "DEPARTMENT")
    If cRecId = 0 Or cRec = 0 Then Exit Sub

    Dim arr As Variant: arr = lo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId Then
            recipeName = NzStr(arr(r, cRec))
            If cDesc > 0 Then recipeDesc = NzStr(arr(r, cDesc))
            If cDept > 0 Then recipeDept = NzStr(arr(r, cDept))
            Exit Sub
        End If
    Next r
End Sub

Private Function BuildRecipeChooserProcessTablesFromRecipes(ByVal recipeId As String, _
    ByVal wsProd As Worksheet, ByVal wsRec As Worksheet, Optional ByVal baseStyle As String = "") As Collection

    Dim created As New Collection
    If wsProd Is Nothing Or wsRec Is Nothing Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If
    If Trim$(recipeId) = "" Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If

    Dim loRecipes As ListObject: Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If
    If loRecipes.DataBodyRange Is Nothing Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If

    Dim cRecId As Long: cRecId = ColumnIndex(loRecipes, "RECIPE_ID")
    Dim cProc As Long: cProc = ColumnIndex(loRecipes, "PROCESS")
    Dim cDiag As Long: cDiag = ColumnIndex(loRecipes, "DIAGRAM_ID")
    Dim cIO As Long: cIO = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    Dim cIng As Long: cIng = ColumnIndex(loRecipes, "INGREDIENT")
    Dim cPct As Long: cPct = ColumnIndex(loRecipes, "PERCENT")
    Dim cUom As Long: cUom = ColumnIndex(loRecipes, "UOM")
    Dim cAmt As Long: cAmt = ColumnIndex(loRecipes, "AMOUNT")
    Dim cIngId As Long: cIngId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    Dim cListRow As Long: cListRow = ColumnIndex(loRecipes, "RECIPE_LIST_ROW")

    If cRecId = 0 Or cProc = 0 Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If

    Dim arr As Variant: arr = loRecipes.DataBodyRange.value
    Dim procMap As Object: Set procMap = CreateObject("Scripting.Dictionary")
    Dim procOrder As Collection: Set procOrder = New Collection

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId Then
            Dim procName As String: procName = Trim$(NzStr(arr(r, cProc)))
            If procName <> "" Then
                If Not procMap.Exists(procName) Then
                    procMap.Add procName, New Collection
                    procOrder.Add procName
                End If
                procMap(procName).Add r
            End If
        End If
    Next r

    If procOrder.count = 0 Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If

    Dim startRow As Long, startCol As Long
    If Not GetRecipeChooserAnchor(wsProd, startRow, startCol) Then
        Set BuildRecipeChooserProcessTablesFromRecipes = created
        Exit Function
    End If

    Dim headerNames As Variant
    headerNames = RecipeChooserHeaderList()
    Dim colCount As Long: colCount = UBound(headerNames) - LBound(headerNames) + 1

    Dim procKey As Variant
    Dim nextSeq As Long
    nextSeq = NextRecipeChooserSequence(wsProd)
    Dim idxProc As Long
    idxProc = 0
    For Each procKey In procOrder
        idxProc = idxProc + 1
        Dim rowsColl As Collection: Set rowsColl = procMap(procKey)
        Dim dataCount As Long: dataCount = rowsColl.count
        If dataCount = 0 Then GoTo NextProc

        Dim tableRange As Range
        Set tableRange = wsProd.Range(wsProd.Cells(startRow, startCol), wsProd.Cells(startRow + dataCount, startCol + colCount - 1))
        If RangeHasListObjectCollisionStrict(wsProd, tableRange) Then
            Set tableRange = FindAvailableRecipeChooserRange(wsProd, startRow, startCol, dataCount + 1, colCount)
            If tableRange Is Nothing Then Exit For
        End If

        tableRange.Clear
        tableRange.rows(1).value = HeaderRowArray(headerNames)

        Dim dataArr() As Variant
        ReDim dataArr(1 To dataCount, 1 To colCount)
        Dim i As Long, c As Long
        For i = 1 To dataCount
            Dim srcRow As Long: srcRow = rowsColl(i)
            For c = 1 To colCount
                Dim hdr As String
                hdr = CStr(headerNames(LBound(headerNames) + c - 1))
                Select Case UCase$(hdr)
                    Case "PROCESS"
                        dataArr(i, c) = procKey
                    Case "DIAGRAM_ID"
                        If cDiag > 0 Then dataArr(i, c) = arr(srcRow, cDiag)
                    Case "INPUT/OUTPUT"
                        If cIO > 0 Then dataArr(i, c) = arr(srcRow, cIO)
                    Case "INGREDIENT"
                        If cIng > 0 Then dataArr(i, c) = arr(srcRow, cIng)
                    Case "PERCENT"
                        If cPct > 0 Then dataArr(i, c) = arr(srcRow, cPct)
                    Case "UOM"
                        If cUom > 0 Then dataArr(i, c) = arr(srcRow, cUom)
                    Case "AMOUNT NEEDED"
                        If cAmt > 0 Then dataArr(i, c) = arr(srcRow, cAmt)
                    Case "INGREDIENT_ID"
                        If cIngId > 0 Then dataArr(i, c) = arr(srcRow, cIngId)
                    Case "RECIPE_LIST_ROW"
                        If cListRow > 0 Then dataArr(i, c) = arr(srcRow, cListRow)
                End Select
            Next c
        Next i

        tableRange.Offset(1, 0).Resize(dataCount, colCount).value = dataArr

        Dim newLo As ListObject
        Set newLo = wsProd.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        If idxProc = 1 Then
            On Error Resume Next
            newLo.Name = TABLE_RECIPE_CHOOSER_GENERATED
            If Err.Number <> 0 Then
                newLo.Name = UniqueListObjectName(wsProd, TABLE_RECIPE_CHOOSER_GENERATED)
            End If
            Err.Clear
            On Error GoTo 0
        Else
            newLo.Name = UniqueListObjectName(wsProd, BuildRecipeChooserProcessTableName(CStr(nextSeq)))
            nextSeq = nextSeq + 1
        End If
        If baseStyle <> "" Then
            On Error Resume Next
            newLo.TableStyle = baseStyle
            On Error GoTo 0
        End If
        created.Add newLo

        startRow = tableRange.row + tableRange.rows.count + 3 ' keep 2 blank rows
NextProc:
    Next procKey

    If created.count > 0 Then
        Dim loProc As ListObject
        For Each loProc In created
            Dim procNameTpl As String: procNameTpl = ProcessNameFromTable(loProc)
            ApplyProductionTemplates loProc, TEMPLATE_SCOPE_RECIPE_PROCESS, procNameTpl, "", recipeId
        Next loProc
    End If

    Set BuildRecipeChooserProcessTablesFromRecipes = created
End Function

Private Sub BuildPaletteTablesForRecipeChooser(ByVal recipeId As String, ByVal wsProd As Worksheet, ByVal wsRec As Worksheet, _
    ByVal procTables As Collection, Optional ByVal baseStyle As String = "")

    If wsProd Is Nothing Or wsRec Is Nothing Then Exit Sub
    If Trim$(recipeId) = "" Then Exit Sub

    Dim loRecipes As ListObject: Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then Exit Sub
    If loRecipes.DataBodyRange Is Nothing Then Exit Sub

    Dim cRecId As Long: cRecId = ColumnIndex(loRecipes, "RECIPE_ID")
    Dim cProc As Long: cProc = ColumnIndex(loRecipes, "PROCESS")
    Dim cIO As Long: cIO = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    Dim cIng As Long: cIng = ColumnIndex(loRecipes, "INGREDIENT")
    Dim cIngId As Long: cIngId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    Dim cAmt As Long: cAmt = ColumnIndex(loRecipes, "AMOUNT")
    Dim cPct As Long: cPct = ColumnIndex(loRecipes, "PERCENT")
    Dim cUom As Long: cUom = ColumnIndex(loRecipes, "UOM")
    If cRecId = 0 Or cProc = 0 Or cIO = 0 Or cIngId = 0 Then Exit Sub

    Dim arr As Variant: arr = loRecipes.DataBodyRange.value
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim entries As Collection: Set entries = New Collection

    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId Then
            Dim ioVal As String: ioVal = UCase$(Trim$(NzStr(arr(r, cIO))))
            If IsInputIoValue(ioVal) Then
                Dim ingId As String: ingId = NzStr(arr(r, cIngId))
                Dim procName As String: procName = NzStr(arr(r, cProc))
                If ingId <> "" And procName <> "" Then
                    If Not IsProcessSelected(procName, wsProd) Then GoTo NextRecipeRow
                    Dim key As String: key = procName & "|" & ingId
                    Dim amtVal As Variant
                    Dim pctVal As Variant
                    Dim uomVal As String
                    Dim ingName As String
                    If cAmt > 0 Then amtVal = arr(r, cAmt)
                    If cPct > 0 Then pctVal = arr(r, cPct)
                    If cUom > 0 Then uomVal = NzStr(arr(r, cUom))
                    If cIng > 0 Then ingName = NzStr(arr(r, cIng))
                    If Not seen.Exists(key) Then
                        Dim info(0 To 7) As Variant
                        info(0) = recipeId
                        info(1) = ingId
                        info(2) = amtVal
                        info(3) = procName
                        info(4) = "USED"
                        info(5) = pctVal
                        info(6) = uomVal
                        info(7) = ingName
                        seen.Add key, info
                        entries.Add info
                    Else
                        If IsNumeric(amtVal) Then
                            Dim curInfo As Variant
                            curInfo = seen(key)
                            If IsNumeric(curInfo(2)) Then
                                curInfo(2) = CDbl(curInfo(2)) + CDbl(amtVal)
                                seen(key) = curInfo
                            End If
                        End If
                    End If
                End If
            End If
        End If
NextRecipeRow:
    Next r

    If entries.count = 0 Then Exit Sub

    Dim startRow As Long, startCol As Long
    Dim anchorStyle As String
    If Not GetInventoryPaletteAnchor(wsProd, startRow, startCol, anchorStyle) Then Exit Sub
    If baseStyle = "" Then baseStyle = anchorStyle

    EnsurePaletteTableMeta
    ClearPaletteTableMeta

    EnsureInventoryPaletteLinesTable wsProd, baseStyle

    Dim headerNames As Variant
    headerNames = InventoryPaletteHeaderList()
    Dim colCount As Long: colCount = UBound(headerNames) - LBound(headerNames) + 1
    Dim invRowMap As Object
    Set invRowMap = BuildInvSysSystemKeyMap()

    Dim procTableMap As Object: Set procTableMap = CreateObject("Scripting.Dictionary")
    Dim procOrder As New Collection
    Dim procLo As ListObject
    If Not procTables Is Nothing Then
        For Each procLo In procTables
            If Not procLo Is Nothing Then
                Dim procTableName As String
                procTableName = Trim$(ProcessNameFromTable(procLo))
                If procTableName <> "" Then
                    Dim procTableKey As String
                    procTableKey = NormalizeProcessBandKey(procTableName)
                    If Not procTableMap.Exists(procTableKey) Then
                        procTableMap.Add procTableKey, procLo
                        procOrder.Add procTableKey
                    End If
                End If
            End If
        Next procLo
    End If

    Dim procEntries As Object: Set procEntries = CreateObject("Scripting.Dictionary")
    Dim procLabels As Object: Set procLabels = CreateObject("Scripting.Dictionary")

    Dim idx As Long
    Dim nextSeq As Long: nextSeq = 1
    Dim hdrProc As Long: hdrProc = HeaderIndex(headerNames, "PROCESS")
    Dim hdrIO As Long: hdrIO = HeaderIndex(headerNames, "INPUT/OUTPUT")
    Dim hdrQty As Long: hdrQty = HeaderIndex(headerNames, "QUANTITY")
    Dim hdrBaseQty As Long: hdrBaseQty = HeaderIndex(headerNames, "BASE QUANTITY")
    Dim hdrSplit As Long: hdrSplit = HeaderIndex(headerNames, "SPLIT %")
    Dim hdrIngredient As Long: hdrIngredient = HeaderIndex(headerNames, "INGREDIENT")
    Dim hdrIngredientId As Long: hdrIngredientId = HeaderIndex(headerNames, "INGREDIENT_ID")
    Dim hdrItem As Long: hdrItem = HeaderIndex(headerNames, "ITEM")
    Dim hdrUom As Long: hdrUom = HeaderIndex(headerNames, "UOM")
    Dim hdrPct As Long: hdrPct = HeaderIndex(headerNames, "PERCENT")
    Dim hdrRow As Long: hdrRow = HeaderIndex(headerNames, "System_Key")

    For idx = 1 To entries.count
        Dim infoArr As Variant
        infoArr = entries(idx)

        Dim procNameForEntry As String
        procNameForEntry = Trim$(NzStr(infoArr(3)))
        If procNameForEntry = "" Then GoTo NextEntry

        Dim procKey As String
        procKey = NormalizeProcessBandKey(procNameForEntry)
        If procKey = "" Then GoTo NextEntry

        If Not procEntries.Exists(procKey) Then
            procEntries.Add procKey, New Collection
            procLabels(procKey) = procNameForEntry
            If Not procTableMap.Exists(procKey) Then
                procOrder.Add procKey
            End If
        End If
        procEntries(procKey).Add infoArr
NextEntry:
    Next idx

    If procEntries.count = 0 Then Exit Sub

    Dim orderIdx As Long
    For orderIdx = 1 To procOrder.count
        Dim procOrderKey As String
        procOrderKey = CStr(procOrder(orderIdx))
        If Not procEntries.Exists(procOrderKey) Then GoTo NextProcOrder

        Dim currentRow As Long
        Dim nextBandTop As Long
        Dim procLabel As String
        procLabel = CStr(procLabels(procOrderKey))

        If procTableMap.Exists(procOrderKey) Then
            Set procLo = procTableMap(procOrderKey)
            currentRow = procLo.Range.Row
        Else
            currentRow = startRow
        End If

        nextBandTop = 0
        If orderIdx < procOrder.count Then
            Dim nextOrderIdx As Long
            For nextOrderIdx = orderIdx + 1 To procOrder.count
                Dim nextProcKey As String
                nextProcKey = CStr(procOrder(nextOrderIdx))
                If procTableMap.Exists(nextProcKey) Then
                    Dim nextProcLo As ListObject
                    Set nextProcLo = procTableMap(nextProcKey)
                    nextBandTop = nextProcLo.Range.Row
                    Exit For
                End If
            Next nextOrderIdx
        End If

        Dim bandEntries As Collection
        Set bandEntries = procEntries(procOrderKey)
        Dim bandIdx As Long
        For bandIdx = 1 To bandEntries.count
            infoArr = bandEntries(bandIdx)

            Dim rowList As Collection
            Set rowList = GetIngredientPaletteRows(infoArr(0), infoArr(1))

            Dim dataCount As Long
            If rowList Is Nothing Then
                dataCount = 1
            ElseIf rowList.count = 0 Then
                dataCount = 1
            Else
                dataCount = rowList.count
            End If

            Dim tableRange As Range
            Set tableRange = wsProd.Range(wsProd.Cells(currentRow, startCol), wsProd.Cells(currentRow + dataCount, startCol + colCount - 1))
            If RangeHasListObjectCollisionStrict(wsProd, tableRange) Then
                Set tableRange = FindAvailablePaletteRange(wsProd, currentRow, startCol, dataCount + 1, colCount)
                If tableRange Is Nothing Then Exit For
            End If

            tableRange.Clear
            tableRange.Rows(1).Value = HeaderRowArray(headerNames)

            Dim dataArr() As Variant
            ReDim dataArr(1 To dataCount, 1 To colCount)
            Dim r2 As Long
            Dim infoUpper As Long
            Dim splitVal As Double
            infoUpper = UBound(infoArr)
            If dataCount > 0 Then splitVal = 100# / CDbl(dataCount) Else splitVal = 100#
            For r2 = 1 To dataCount
                If hdrProc > 0 Then dataArr(r2, hdrProc) = procLabel
                If hdrIO > 0 Then dataArr(r2, hdrIO) = NzStr(infoArr(4))
                If hdrIngredientId > 0 Then dataArr(r2, hdrIngredientId) = NzStr(infoArr(1))
                If hdrIngredient > 0 And infoUpper >= 7 Then dataArr(r2, hdrIngredient) = NzStr(infoArr(7))
                If hdrBaseQty > 0 Then dataArr(r2, hdrBaseQty) = infoArr(2)
                If hdrSplit > 0 Then dataArr(r2, hdrSplit) = splitVal
                If hdrQty > 0 Then
                    If IsNumeric(infoArr(2)) Then
                        dataArr(r2, hdrQty) = CDbl(infoArr(2)) * splitVal / 100#
                    Else
                        dataArr(r2, hdrQty) = infoArr(2)
                    End If
                End If
                If hdrPct > 0 And infoUpper >= 5 Then dataArr(r2, hdrPct) = infoArr(5)
                If hdrUom > 0 And infoUpper >= 6 Then dataArr(r2, hdrUom) = infoArr(6)
                If hdrRow > 0 Then
                    If Not rowList Is Nothing Then
                        If rowList.count > 0 Then
                            Dim palEntry As Variant
                            palEntry = rowList(r2)
                            dataArr(r2, hdrRow) = PaletteEntryField(palEntry, 0)
                            If hdrItem > 0 Then dataArr(r2, hdrItem) = PaletteEntryField(palEntry, 1)
                            If hdrUom > 0 And NzStr(dataArr(r2, hdrUom)) = "" Then dataArr(r2, hdrUom) = PaletteEntryField(palEntry, 2)
                        End If
                    End If
                End If
            Next r2
            tableRange.Offset(1, 0).Resize(dataCount, colCount).Value = dataArr

            Dim newLo As ListObject
            Set newLo = wsProd.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
            newLo.Name = UniqueListObjectName(wsProd, "proc_" & CStr(nextSeq) & "_palette")
            nextSeq = nextSeq + 1
            If baseStyle <> "" Then
                On Error Resume Next
                newLo.TableStyle = baseStyle
                On Error GoTo 0
            End If

            mPaletteTableMeta(newLo.Name) = infoArr

            ApplyProcessHeaderColor newLo, procLabel

            FillPaletteTableFromInvSys newLo, invRowMap

            ApplyProductionTemplates newLo, TEMPLATE_SCOPE_PROD_RUN, procLabel, TEMPLATE_TABLEKEY_PALETTE, recipeId

            currentRow = tableRange.Row + tableRange.Rows.Count + 3
        Next bandIdx
NextProcOrder:
    Next orderIdx
End Sub

Private Sub ApplyProductionOutputTemplates(ByVal recipeId As String, ByVal wsProd As Worksheet)
    If wsProd Is Nothing Then Exit Sub
    If Trim$(recipeId) = "" Then Exit Sub
    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then Exit Sub
    ClearListObjectFormulas loOut
    ApplyProductionTemplates loOut, TEMPLATE_SCOPE_PROD_RUN, "", "ProductionOutput", recipeId
End Sub

Private Sub DeleteRecipeChooserProcessTables(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = ws.ListObjects.count To 1 Step -1
        Dim lo As ListObject
        Set lo = ws.ListObjects(i)
        If IsRecipeChooserProcessTable(lo) Or LCase$(lo.Name) = LCase$(TABLE_RECIPE_CHOOSER_GENERATED) Then
            Dim addr As String: addr = lo.Range.Address
            On Error Resume Next
            lo.Delete
            ws.Range(addr).Clear
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function IsRecipeChooserProcessTable(ByVal lo As ListObject) As Boolean
    If lo Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(lo.Name)
    If Left$(nm, 5) <> "proc_" Then Exit Function
    If Right$(nm, Len(RECIPE_CHOOSER_TABLE_SUFFIX) + 1) = "_" & LCase$(RECIPE_CHOOSER_TABLE_SUFFIX) Then
        IsRecipeChooserProcessTable = True
    End If
End Function

Private Function BuildRecipeChooserProcessTableName(ByVal processKey As String) As String
    Dim key As String: key = Trim$(processKey)
    If key <> "" And IsNumeric(key) Then
        BuildRecipeChooserProcessTableName = "proc_" & CLng(key) & "_" & RECIPE_CHOOSER_TABLE_SUFFIX
    Else
        key = SafeProcessKey(processKey)
        BuildRecipeChooserProcessTableName = "proc_" & key & "_" & RECIPE_CHOOSER_TABLE_SUFFIX
    End If
End Function

Private Function NextRecipeChooserSequence(ByVal ws As Worksheet) As Long
    Dim maxSeq As Long
    If ws Is Nothing Then
        NextRecipeChooserSequence = 1
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsRecipeChooserProcessTable(lo) Then
            Dim seq As Long
            seq = RecipeChooserSequenceFromName(lo.Name)
            If seq > maxSeq Then maxSeq = seq
        End If
    Next lo
    NextRecipeChooserSequence = maxSeq + 1
End Function

Private Function RecipeChooserSequenceFromName(ByVal tableName As String) As Long
    Dim nm As String: nm = LCase$(tableName)
    If Left$(nm, 5) <> "proc_" Then Exit Function
    If Right$(nm, Len(RECIPE_CHOOSER_TABLE_SUFFIX) + 1) <> "_" & LCase$(RECIPE_CHOOSER_TABLE_SUFFIX) Then Exit Function
    Dim core As String
    core = Mid$(nm, 6, Len(nm) - 5 - (Len(RECIPE_CHOOSER_TABLE_SUFFIX) + 1))
    If core = "" Then Exit Function
    If Left$(core, 2) = "p_" Then core = Mid$(core, 3)
    RecipeChooserSequenceFromName = CLng(val(core))
End Function

Private Function RecipeChooserHeaderList() As Variant
    RecipeChooserHeaderList = Array( _
        "PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", _
        "AMOUNT NEEDED", "INGREDIENT_ID", "RECIPE_LIST_ROW")
End Function

Private Function GetRecipeChooserAnchor(ByVal ws As Worksheet, ByRef startRow As Long, ByRef startCol As Long) As Boolean
    GetRecipeChooserAnchor = False
    If ws Is Nothing Then Exit Function
    Dim loChooser As ListObject
    Set loChooser = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If loChooser Is Nothing Then Exit Function

    startCol = loChooser.Range.Column
    startRow = loChooser.Range.row + loChooser.Range.rows.count + 2 ' one blank row
    If startRow > 0 And startCol > 0 Then GetRecipeChooserAnchor = True
End Function

Private Function FindAvailableRecipeChooserRange(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, _
    ByVal totalRows As Long, ByVal totalCols As Long) As Range

    If ws Is Nothing Then Exit Function
    If totalRows < 1 Or totalCols < 1 Then Exit Function
    If startRow < 1 Then startRow = 1
    If startCol < 1 Then startCol = 1

    Dim maxRow As Long: maxRow = ws.rows.count
    Dim tryRow As Long: tryRow = startRow
    Dim candidate As Range
    Do While tryRow + totalRows - 1 <= maxRow
        Set candidate = ws.Range(ws.Cells(tryRow, startCol), ws.Cells(tryRow + totalRows - 1, startCol + totalCols - 1))
        If Not RangeHasListObjectCollisionStrict(ws, candidate) Then
            Set FindAvailableRecipeChooserRange = candidate
            Exit Function
        End If
        tryRow = tryRow + totalRows + 3
    Loop
End Function

Private Sub DeleteInventoryPaletteTables(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = ws.ListObjects.count To 1 Step -1
        Dim lo As ListObject
        Set lo = ws.ListObjects(i)
        If lo Is Nothing Then GoTo NextLo
        If LCase$(lo.Name) = LCase$(TABLE_INV_PALETTE_GENERATED) Or LCase$(lo.Name) Like "proc_*_palette" Then
            Dim addr As String: addr = lo.Range.Address
            On Error Resume Next
            lo.Delete
            ws.Range(addr).Clear
            On Error GoTo 0
        End If
NextLo:
    Next i
    ClearPaletteTableMeta
End Sub

Private Function InventoryPaletteHeaderList() As Variant
    InventoryPaletteHeaderList = Array( _
        "INGREDIENT", "INGREDIENT_ID", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "PERCENT", _
        "SPLIT %", "UOM", "QUANTITY", "BASE QUANTITY", "PROCESS", "LOCATION", "System_Key", "INPUT/OUTPUT")
End Function

Private Function GetInventoryPaletteAnchor(ByVal ws As Worksheet, ByRef startRow As Long, ByRef startCol As Long, ByRef baseStyle As String) As Boolean
    GetInventoryPaletteAnchor = False
    If ws Is Nothing Then Exit Function
    Dim lo As ListObject
    Set lo = GetListObject(ws, TABLE_INV_PALETTE_GENERATED)
    If lo Is Nothing Then
        Set lo = FindListObjectByNameOrHeaders(ws, TABLE_INV_PALETTE_GENERATED, Array("ITEM_CODE", "ITEM", "System_Key"))
    End If
    If Not lo Is Nothing Then
        On Error Resume Next
        baseStyle = lo.TableStyle
        On Error GoTo 0
        If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
            startRow = lo.Range.row
            startCol = lo.Range.Column
            GetInventoryPaletteAnchor = True
            Exit Function
        End If
    End If

    Dim loProd As ListObject
    Dim loCheck As ListObject
    Set loProd = FindListObjectByNameOrHeaders(ws, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    Set loCheck = FindListObjectByNameOrHeaders(ws, "Prod_invSys_Check", Array("USED", "TOTAL INV"))
    If Not loProd Is Nothing Then
        startCol = loProd.Range.Column
        Dim bottom As Long
        bottom = loProd.Range.row + loProd.Range.rows.count - 1
        If Not loCheck Is Nothing Then
            Dim chkBottom As Long
            chkBottom = loCheck.Range.row + loCheck.Range.rows.count - 1
            If chkBottom > bottom Then bottom = chkBottom
        End If
        startRow = bottom + 2
        GetInventoryPaletteAnchor = True
    End If
End Function

Private Function FindAvailablePaletteRange(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, _
    ByVal totalRows As Long, ByVal totalCols As Long) As Range

    If ws Is Nothing Then Exit Function
    If totalRows < 1 Or totalCols < 1 Then Exit Function
    If startRow < 1 Then startRow = 1
    If startCol < 1 Then startCol = 1

    Dim maxRow As Long: maxRow = ws.rows.count
    Dim tryRow As Long: tryRow = startRow
    Dim candidate As Range
    Do While tryRow + totalRows - 1 <= maxRow
        Set candidate = ws.Range(ws.Cells(tryRow, startCol), ws.Cells(tryRow + totalRows - 1, startCol + totalCols - 1))
        If Not RangeHasListObjectCollisionStrict(ws, candidate) Then
            Set FindAvailablePaletteRange = candidate
            Exit Function
        End If
        tryRow = tryRow + totalRows + 3
    Loop
End Function

Private Function NormalizeProcessBandKey(ByVal value As String) As String
    NormalizeProcessBandKey = LCase$(Trim$(value))
End Function

Private Function BuildInvSysSystemKeyMap() As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim loInv As ListObject
    Set loInv = GetInvSysTable()
    AddInvSysTableRowsToSystemKeyMap dict, loInv
    If dict.count = 0 Then AddInventoryPickerRowsToSystemKeyMap dict, LoadProductionInventoryPickerItems("")
    If dict.count > 0 Then Set BuildInvSysSystemKeyMap = dict
End Function

Private Sub AddInvSysTableRowsToSystemKeyMap(ByVal dict As Object, ByVal loInv As ListObject)
    If dict Is Nothing Then Exit Sub
    If loInv Is Nothing Or loInv.DataBodyRange Is Nothing Then Exit Sub

    Dim cRow As Long: cRow = ColumnIndex(loInv, "System_Key")
    If cRow = 0 Then Exit Sub
    Dim cCode As Long: cCode = ColumnIndex(loInv, "ITEM_CODE")
    If cCode = 0 Then cCode = ColumnIndexLoose(loInv, "ITEM_CODE", "ITEMCODE", "ITEM CODE")
    Dim cVend As Long: cVend = ColumnIndex(loInv, "VENDOR(s)")
    If cVend = 0 Then cVend = ColumnIndexLoose(loInv, "VENDORS", "VENDOR", "VENDOR(S)")
    Dim cVendCode As Long: cVendCode = ColumnIndex(loInv, "VENDOR_CODE")
    If cVendCode = 0 Then cVendCode = ColumnIndexLoose(loInv, "VENDOR_CODE", "VENDORCODE", "VENDOR CODE")
    Dim cDesc As Long: cDesc = ColumnIndex(loInv, "DESCRIPTION")
    If cDesc = 0 Then cDesc = ColumnIndexLoose(loInv, "DESCRIPTION", "DESC")
    Dim cItem As Long: cItem = ColumnIndex(loInv, "ITEM")
    If cItem = 0 Then cItem = ColumnIndexLoose(loInv, "ITEM", "ITEMS", "ITEMNAME", "ITEM NAME")
    Dim cUom As Long: cUom = ColumnIndex(loInv, "UOM")
    If cUom = 0 Then cUom = ColumnIndexLoose(loInv, "UOM", "UNIT", "UNITOFMEASURE", "UNITOFMEASUREMENT")
    Dim cLoc As Long: cLoc = ColumnIndex(loInv, "LOCATION")
    If cLoc = 0 Then cLoc = ColumnIndexLoose(loInv, "LOCATION", "LOC")

    Dim arr As Variant: arr = loInv.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowKey As String
        rowKey = NormalizeSystemKey(arr(r, cRow))
        If rowKey <> "" Then
            If Not dict.Exists(rowKey) Then
                Dim info(1 To 7) As Variant
                If cCode > 0 Then info(1) = NzStr(arr(r, cCode)) Else info(1) = ""
                If cVend > 0 Then info(2) = NzStr(arr(r, cVend)) Else info(2) = ""
                If cVendCode > 0 Then info(3) = NzStr(arr(r, cVendCode)) Else info(3) = ""
                If cDesc > 0 Then info(4) = NzStr(arr(r, cDesc)) Else info(4) = ""
                If cItem > 0 Then info(5) = NzStr(arr(r, cItem)) Else info(5) = ""
                If cUom > 0 Then info(6) = NzStr(arr(r, cUom)) Else info(6) = ""
                If cLoc > 0 Then info(7) = NzStr(arr(r, cLoc)) Else info(7) = ""
                dict.Add rowKey, info
            End If
        End If
    Next r
End Sub

Private Sub AddInventoryPickerRowsToSystemKeyMap(ByVal dict As Object, ByVal pickerRows As Variant)
    Dim r As Long
    Dim rowKey As String
    Dim info(1 To 7) As Variant

    If dict Is Nothing Then Exit Sub
    If IsEmpty(pickerRows) Then Exit Sub
    If Not IsArray(pickerRows) Then Exit Sub

    For r = LBound(pickerRows, 1) To UBound(pickerRows, 1)
        rowKey = NormalizeSystemKey(pickerRows(r, 1))
        If rowKey <> "" Then
            If Not dict.Exists(rowKey) Then
                info(1) = NzStr(pickerRows(r, 7))
                info(2) = ""
                info(3) = ""
                info(4) = NzStr(pickerRows(r, 6))
                info(5) = NzStr(pickerRows(r, 2))
                info(6) = NzStr(pickerRows(r, 3))
                info(7) = NzStr(pickerRows(r, 5))
                dict.Add rowKey, info
            End If
        End If
    Next r
End Sub


Private Function GetInvSysTable() As ListObject
    Dim wsInv As Worksheet: Set wsInv = SheetExists("InventoryManagement")
    If wsInv Is Nothing Then Set wsInv = SheetExists("Inventory Management")
    If wsInv Is Nothing Then Set wsInv = SheetExists("INVENTORY MANAGEMENT")
    If wsInv Is Nothing Then Exit Function

    Dim loInv As ListObject: Set loInv = GetListObject(wsInv, "invSys")
    If Not loInv Is Nothing Then
        Set GetInvSysTable = loInv
        Exit Function
    End If

    Dim lo As ListObject
    For Each lo In wsInv.ListObjects
        If ColumnIndex(lo, "System_Key") > 0 Then
            If ColumnIndexLoose(lo, "ITEM", "ITEMS", "ITEMNAME", "ITEM NAME") > 0 _
                Or ColumnIndexLoose(lo, "ITEM_CODE", "ITEMCODE", "ITEM CODE") > 0 Then
                Set GetInvSysTable = lo
                Exit Function
            End If
        End If
    Next lo
End Function

Private Function GetInvSysTableFromWorkbook(ByVal wb As Workbook) As ListObject
    Dim wsInv As Worksheet
    Dim loInv As ListObject
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    Set wsInv = WorkbookSheetExists(wb, "InventoryManagement")
    If wsInv Is Nothing Then Set wsInv = WorkbookSheetExists(wb, "Inventory Management")
    If wsInv Is Nothing Then Set wsInv = WorkbookSheetExists(wb, "INVENTORY MANAGEMENT")
    If wsInv Is Nothing Then Exit Function

    Set loInv = GetListObject(wsInv, "invSys")
    If Not loInv Is Nothing Then
        Set GetInvSysTableFromWorkbook = loInv
        Exit Function
    End If

    For Each lo In wsInv.ListObjects
        If ColumnIndex(lo, "System_Key") > 0 Then
            If ColumnIndexLoose(lo, "ITEM", "ITEMS", "ITEMNAME", "ITEM NAME") > 0 _
                Or ColumnIndexLoose(lo, "ITEM_CODE", "ITEMCODE", "ITEM CODE") > 0 Then
                Set GetInvSysTableFromWorkbook = lo
                Exit Function
            End If
        End If
    Next lo
End Function

Public Function LoadProductionInventoryPickerItems(Optional ByVal filterText As String = "") As Variant
    On Error GoTo FailSoft

    Dim lo As ListObject
    Dim result As Variant
    Dim wbRuntime As Workbook
    Dim openedTransient As Boolean
    Dim warehouseId As String
    Dim rootPath As String
    Dim report As String
    Dim inventoryPath As String
    Dim wbOps As Workbook

    filterText = Trim$(filterText)

    If ResolveProductionRecipesStorageTarget(warehouseId, rootPath, report) Then
        inventoryPath = NormalizeFolderPathProduction(rootPath) & "\" & warehouseId & ".invSys.Data.Inventory.xlsb"
        If Len(Dir$(inventoryPath, vbNormal)) > 0 Then
            Set wbRuntime = OpenWorkbookHiddenProduction(inventoryPath, True, openedTransient)
            If Not wbRuntime Is Nothing Then
                result = BuildProductionCanonicalInventoryPickerItems(wbRuntime, filterText)
                If Not IsEmpty(result) Then
                    LoadProductionInventoryPickerItems = result
                    GoTo CleanExit
                End If

                Set lo = GetInvSysTableFromWorkbook(wbRuntime)
                result = BuildProductionInventoryPickerItems(lo, filterText)
                If Not IsEmpty(result) Then
                    LoadProductionInventoryPickerItems = result
                    GoTo CleanExit
                End If
            End If
        End If
    End If

    Set lo = GetInvSysTable()
    result = BuildProductionInventoryPickerItems(lo, filterText)
    If Not IsEmpty(result) Then
        LoadProductionInventoryPickerItems = result
        GoTo CleanExit
    End If

    Set wbOps = ResolveProductionWorkbook(, SHEET_PRODUCTION)
    If Not wbOps Is Nothing Then
        Set lo = GetInvSysTableFromWorkbook(wbOps)
        result = BuildProductionInventoryPickerItems(lo, filterText)
        If Not IsEmpty(result) Then LoadProductionInventoryPickerItems = result
    End If

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveProduction wbRuntime
    Exit Function

FailSoft:
    LoadProductionInventoryPickerItems = Empty
    Resume CleanExit
End Function

Public Function LoadProductionRunInventoryPickerItems(Optional ByVal filterText As String = "") As Variant
    On Error GoTo FailSoft

    Dim result As Variant

    filterText = Trim$(filterText)
    result = modInventoryDomainBridge.ListInventoryPickerItemsBridge(filterText)
    If Not IsEmpty(result) Then
        If IsArray(result) Then
            If ProductionPickerRowsHaveSystemKeys(result) Then
                LoadProductionRunInventoryPickerItems = result
                Exit Function
            End If
        End If
    End If

FailSoft:
    LoadProductionRunInventoryPickerItems = Empty
End Function

Private Function ProductionPickerRowsHaveSystemKeys(ByVal pickerRows As Variant) As Boolean
    Dim r As Long

    If IsEmpty(pickerRows) Or Not IsArray(pickerRows) Then Exit Function
    If UBound(pickerRows, 2) < 1 Then Exit Function
    For r = LBound(pickerRows, 1) To UBound(pickerRows, 1)
        If Trim$(NzStr(pickerRows(r, 1))) = "" Then Exit Function
    Next r
    ProductionPickerRowsHaveSystemKeys = True
End Function

Public Function RefreshProductionInventoryReadModelForWorkbookResult(ByVal targetWb As Workbook) As String
    On Error GoTo FailRefresh

    Dim report As String
    Dim refreshed As Boolean

    refreshed = modOperationsPrimitiveBridge.RefreshInventoryReadModel(targetWb.Name, "", "LOCAL", report)
    If report = "" Then report = IIf(refreshed, "OK", "Inventory read-model refresh failed.")
    RefreshProductionInventoryReadModelForWorkbookResult = IIf(refreshed, "OK", "FAIL") & vbTab & report
    Exit Function

FailRefresh:
    RefreshProductionInventoryReadModelForWorkbookResult = "FAIL" & vbTab & _
        "Production inventory read-model refresh failed: " & Err.Description
End Function

Public Function DiagnoseProductionInventoryReadModelForWorkbook(ByVal targetWb As Workbook) As String
    On Error GoTo FailDiagnose
    DiagnoseProductionInventoryReadModelForWorkbook = _
        modOperationsPrimitiveBridge.DiagnoseInventoryReadModel(targetWb.Name, "", "LOCAL")
    Exit Function

FailDiagnose:
    DiagnoseProductionInventoryReadModelForWorkbook = _
        "Production inventory read-model diagnostic failed: " & Err.Description
End Function

Public Function GetProductionRunDefaultLocation() As String
    On Error GoTo CleanFail
    GetProductionRunDefaultLocation = Trim$(modConfig.GetString("DefaultLocation", ""))
CleanFail:
End Function

Private Function BuildProductionInventoryPickerItems(ByVal loInv As ListObject, _
                                                     Optional ByVal filterText As String = "") As Variant
    On Error GoTo FailSoft

    Dim arr As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim rowVal As String
    Dim itemVal As String
    Dim uomVal As String
    Dim totalVal As String
    Dim locVal As String
    Dim descVal As String
    Dim itemCode As String
    Dim categoryVal As String
    Dim trackQtyVal As String
    Dim itemKindVal As String
    Dim haystack As String

    If loInv Is Nothing Then Exit Function
    If loInv.DataBodyRange Is Nothing Then Exit Function

    Dim cRow As Long: cRow = ColumnIndex(loInv, "System_Key")
    Dim cItem As Long: cItem = ColumnIndex(loInv, "ITEM")
    If cItem = 0 Then cItem = ColumnIndexLoose(loInv, "ITEM", "ITEMS", "ITEMNAME", "ITEM NAME")
    Dim cUom As Long: cUom = ColumnIndex(loInv, "UOM")
    If cUom = 0 Then cUom = ColumnIndexLoose(loInv, "UOM", "UNIT", "UNITOFMEASURE", "UNITOFMEASUREMENT")
    Dim cTotal As Long: cTotal = ColumnIndex(loInv, "TOTAL INV")
    If cTotal = 0 Then cTotal = ColumnIndexLoose(loInv, "TOTALINV", "TOTAL_INV", "TOTALINVENTORY", "QTY", "QUANTITY")
    Dim cLoc As Long: cLoc = ColumnIndex(loInv, "LOCATION")
    If cLoc = 0 Then cLoc = ColumnIndexLoose(loInv, "LOCATION", "LOC")
    Dim cDesc As Long: cDesc = ColumnIndex(loInv, "DESCRIPTION")
    If cDesc = 0 Then cDesc = ColumnIndexLoose(loInv, "DESCRIPTION", "DESC")
    Dim cCode As Long: cCode = ColumnIndex(loInv, "ITEM_CODE")
    If cCode = 0 Then cCode = ColumnIndexLoose(loInv, "ITEM_CODE", "ITEMCODE", "ITEM CODE", "SKU")
    Dim cCategory As Long: cCategory = ColumnIndex(loInv, "CATEGORY")
    Dim cTrackQty As Long: cTrackQty = ColumnIndex(loInv, "TRACK_QTY")
    Dim cItemKind As Long: cItemKind = ColumnIndex(loInv, "ITEM_KIND")
    If cRow = 0 And cItem = 0 And cCode = 0 Then Exit Function

    filterText = LCase$(Trim$(filterText))
    arr = loInv.DataBodyRange.Value
    ReDim result(1 To UBound(arr, 1), 1 To 7)
    For r = 1 To UBound(arr, 1)
        rowVal = ""
        itemVal = ""
        uomVal = ""
        totalVal = ""
        locVal = ""
        descVal = ""
        itemCode = ""
        categoryVal = ""
        trackQtyVal = ""
        itemKindVal = ""

        If cRow > 0 Then rowVal = NzStr(arr(r, cRow))
        If cItem > 0 Then itemVal = NzStr(arr(r, cItem))
        If cUom > 0 Then uomVal = NzStr(arr(r, cUom))
        If cTotal > 0 Then totalVal = NzStr(arr(r, cTotal))
        If cLoc > 0 Then locVal = NzStr(arr(r, cLoc))
        If cDesc > 0 Then descVal = NzStr(arr(r, cDesc))
        If cCode > 0 Then itemCode = NzStr(arr(r, cCode))
        If cCategory > 0 Then categoryVal = NzStr(arr(r, cCategory))
        If cTrackQty > 0 Then trackQtyVal = NzStr(arr(r, cTrackQty))
        If cItemKind > 0 Then itemKindVal = NzStr(arr(r, cItemKind))
        If itemVal = "" And itemCode = "" Then GoTo NextRow

        haystack = LCase$(rowVal & " " & itemVal & " " & itemCode & " " & descVal & " " & locVal & " " & uomVal & " " & categoryVal & " " & itemKindVal)
        If filterText <> "" Then
            If InStr(1, haystack, filterText, vbTextCompare) = 0 Then GoTo NextRow
        End If

        outRow = outRow + 1
        result(outRow, 1) = rowVal
        result(outRow, 2) = itemVal
        result(outRow, 3) = uomVal
        If ProductionInventoryItemIsNonCounted(trackQtyVal, itemKindVal, categoryVal) Then
            result(outRow, 4) = "utility"
        Else
            result(outRow, 4) = totalVal
        End If
        result(outRow, 5) = locVal
        result(outRow, 6) = descVal
        result(outRow, 7) = itemCode
NextRow:
    Next r

    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 7)
    For r = 1 To outRow
        For c = 1 To 7
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    BuildProductionInventoryPickerItems = trimmed
    Exit Function

FailSoft:
    BuildProductionInventoryPickerItems = Empty
End Function

Private Function BuildProductionCanonicalInventoryPickerItems(ByVal wbInv As Workbook, _
                                                              Optional ByVal filterText As String = "") As Variant
    On Error GoTo FailSoft

    Dim loCatalog As ListObject
    Dim loBalance As ListObject
    Dim balances As Object
    Dim locations As Object
    Dim src As Variant
    Dim result() As Variant
    Dim trimmed() As Variant
    Dim r As Long
    Dim c As Long
    Dim outRow As Long
    Dim rowVal As String
    Dim itemVal As String
    Dim uomVal As String
    Dim locVal As String
    Dim descVal As String
    Dim itemCode As String
    Dim sku As String
    Dim categoryVal As String
    Dim trackQtyVal As String
    Dim itemKindVal As String
    Dim haystack As String

    Set loCatalog = FindProductionListObjectByName(wbInv, "tblSkuCatalog")
    If loCatalog Is Nothing Then Exit Function
    If loCatalog.DataBodyRange Is Nothing Then Exit Function
    Set loBalance = FindProductionListObjectByName(wbInv, "tblSkuBalance")
    Set balances = BuildProductionSkuBalanceDictionary(loBalance)
    Set locations = BuildProductionLocationBalanceDictionary(FindProductionListObjectByName(wbInv, "tblLocationBalance"))

    Dim cSku As Long: cSku = ColumnIndex(loCatalog, "SKU")
    Dim cRow As Long: cRow = ColumnIndex(loCatalog, "System_Key")
    Dim cCode As Long: cCode = ColumnIndex(loCatalog, "ITEM_CODE")
    If cCode = 0 Then cCode = ColumnIndexLoose(loCatalog, "ITEM_CODE", "ITEMCODE", "ITEM CODE", "SKU")
    Dim cItem As Long: cItem = ColumnIndex(loCatalog, "ITEM")
    If cItem = 0 Then cItem = ColumnIndexLoose(loCatalog, "ITEM", "ITEMS", "ITEMNAME", "ITEM NAME")
    Dim cUom As Long: cUom = ColumnIndex(loCatalog, "UOM")
    If cUom = 0 Then cUom = ColumnIndexLoose(loCatalog, "UOM", "UNIT", "UNITOFMEASURE", "UNITOFMEASUREMENT")
    Dim cLoc As Long: cLoc = ColumnIndex(loCatalog, "LOCATION")
    If cLoc = 0 Then cLoc = ColumnIndexLoose(loCatalog, "LOCATION", "LOC")
    Dim cDesc As Long: cDesc = ColumnIndex(loCatalog, "DESCRIPTION")
    If cDesc = 0 Then cDesc = ColumnIndexLoose(loCatalog, "DESCRIPTION", "DESC")
    Dim cCategory As Long: cCategory = ColumnIndex(loCatalog, "CATEGORY")
    Dim cTrackQty As Long: cTrackQty = ColumnIndex(loCatalog, "TRACK_QTY")
    Dim cItemKind As Long: cItemKind = ColumnIndex(loCatalog, "ITEM_KIND")
    If cItem = 0 And cCode = 0 Then Exit Function

    filterText = LCase$(Trim$(filterText))
    src = loCatalog.DataBodyRange.Value
    ReDim result(1 To UBound(src, 1), 1 To 7)
    For r = 1 To UBound(src, 1)
        rowVal = ""
        itemVal = ""
        uomVal = ""
        locVal = ""
        descVal = ""
        itemCode = ""
        sku = ""
        categoryVal = ""
        trackQtyVal = ""
        itemKindVal = ""

        If cSku > 0 Then sku = NzStr(src(r, cSku))
        If cRow > 0 Then rowVal = NzStr(src(r, cRow))
        If cCode > 0 Then itemCode = NzStr(src(r, cCode))
        If itemCode = "" Then itemCode = sku
        If cItem > 0 Then itemVal = NzStr(src(r, cItem))
        If cUom > 0 Then uomVal = NzStr(src(r, cUom))
        If cLoc > 0 Then locVal = NzStr(src(r, cLoc))
        If cDesc > 0 Then descVal = NzStr(src(r, cDesc))
        If cCategory > 0 Then categoryVal = NzStr(src(r, cCategory))
        If cTrackQty > 0 Then trackQtyVal = NzStr(src(r, cTrackQty))
        If cItemKind > 0 Then itemKindVal = NzStr(src(r, cItemKind))
        If itemVal = "" And itemCode = "" Then GoTo NextRow

        haystack = LCase$(rowVal & " " & itemVal & " " & itemCode & " " & descVal & " " & locVal & " " & uomVal & " " & categoryVal & " " & itemKindVal)
        If filterText <> "" Then
            If InStr(1, haystack, filterText, vbTextCompare) = 0 Then GoTo NextRow
        End If

        outRow = outRow + 1
        result(outRow, 1) = rowVal
        result(outRow, 2) = itemVal
        result(outRow, 3) = uomVal
        If ProductionInventoryItemIsNonCounted(trackQtyVal, itemKindVal, categoryVal) Then
            result(outRow, 4) = "utility"
        ElseIf Not balances Is Nothing Then
            If sku <> "" And balances.Exists("SKU:" & UCase$(sku)) Then result(outRow, 4) = balances("SKU:" & UCase$(sku))
            If NzStr(result(outRow, 4)) = "" And sku <> "" And balances.Exists(UCase$(sku)) Then result(outRow, 4) = balances(UCase$(sku))
            If NzStr(result(outRow, 4)) = "" And itemCode <> "" And balances.Exists("CODE:" & UCase$(itemCode)) Then result(outRow, 4) = balances("CODE:" & UCase$(itemCode))
            If NzStr(result(outRow, 4)) = "" And itemCode <> "" And balances.Exists(UCase$(itemCode)) Then result(outRow, 4) = balances(UCase$(itemCode))
            If NzStr(result(outRow, 4)) = "" And rowVal <> "" And balances.Exists("SYSTEM_KEY:" & UCase$(rowVal)) Then result(outRow, 4) = balances("SYSTEM_KEY:" & UCase$(rowVal))
        End If
        If locVal = "" Then locVal = ProductionLocationForIdentity(locations, sku, itemCode)
        result(outRow, 5) = locVal
        result(outRow, 6) = descVal
        result(outRow, 7) = itemCode
NextRow:
    Next r

    If outRow = 0 Then Exit Function
    ReDim trimmed(1 To outRow, 1 To 7)
    For r = 1 To outRow
        For c = 1 To 7
            trimmed(r, c) = result(r, c)
        Next c
    Next r
    BuildProductionCanonicalInventoryPickerItems = trimmed
    Exit Function

FailSoft:
    BuildProductionCanonicalInventoryPickerItems = Empty
End Function

Private Function BuildProductionLocationBalanceDictionary(ByVal loLocation As ListObject) As Object
    On Error GoTo FailSoft

    If loLocation Is Nothing Then Exit Function
    If loLocation.DataBodyRange Is Nothing Then Exit Function

    Dim cSku As Long: cSku = ColumnIndex(loLocation, "SKU")
    Dim cLocation As Long: cLocation = ColumnIndex(loLocation, "Location")
    Dim cQty As Long: cQty = ColumnIndex(loLocation, "QtyOnHand")
    If cSku = 0 Or cLocation = 0 Then Exit Function

    Dim best As Object
    Set best = CreateObject("Scripting.Dictionary")
    best.CompareMode = vbTextCompare

    Dim values As Variant
    values = loLocation.DataBodyRange.Value
    Dim r As Long
    Dim sku As String
    Dim locVal As String
    Dim qtyVal As Double
    Dim current As Variant
    For r = 1 To UBound(values, 1)
        sku = UCase$(Trim$(NzStr(values(r, cSku))))
        locVal = Trim$(NzStr(values(r, cLocation)))
        If sku = "" Or locVal = "" Then GoTo NextRow
        If cQty > 0 Then qtyVal = NzDbl(values(r, cQty)) Else qtyVal = 0#
        If Not best.Exists(sku) Then
            best.Add sku, Array(locVal, qtyVal)
        Else
            current = best(sku)
            If qtyVal > CDbl(current(1)) Then best(sku) = Array(locVal, qtyVal)
        End If
NextRow:
    Next r

    If best.Count > 0 Then Set BuildProductionLocationBalanceDictionary = best
    Exit Function

FailSoft:
    Set BuildProductionLocationBalanceDictionary = Nothing
End Function

Private Function ProductionLocationForIdentity(ByVal locations As Object, ByVal sku As String, ByVal itemCode As String) As String
    If locations Is Nothing Then Exit Function
    sku = UCase$(Trim$(sku))
    itemCode = UCase$(Trim$(itemCode))

    Dim info As Variant
    If sku <> "" And locations.Exists(sku) Then
        info = locations(sku)
        ProductionLocationForIdentity = NzStr(info(0))
        Exit Function
    End If
    If itemCode <> "" And locations.Exists(itemCode) Then
        info = locations(itemCode)
        ProductionLocationForIdentity = NzStr(info(0))
    End If
End Function

Private Function ProductionInventoryItemIsNonCounted(ByVal trackQtyVal As String, _
                                                     ByVal itemKindVal As String, _
                                                     ByVal categoryVal As String) As Boolean
    trackQtyVal = UCase$(Trim$(trackQtyVal))
    itemKindVal = UCase$(Trim$(itemKindVal))
    categoryVal = UCase$(Trim$(categoryVal))
    ProductionInventoryItemIsNonCounted = (trackQtyVal = "FALSE" Or trackQtyVal = "NO" Or trackQtyVal = "0" _
                                           Or itemKindVal = "UTILITY" Or itemKindVal = "SERVICE" Or itemKindVal = "NON_COUNTED" _
                                           Or categoryVal = "UTILITY" Or categoryVal = "SERVICE")
End Function

Private Function BuildProductionSkuBalanceDictionary(ByVal loBalance As ListObject) As Object
    On Error GoTo FailSoft

    Dim dict As Object
    Dim arr As Variant
    Dim r As Long
    Dim sku As String
    Dim itemCode As String
    Dim rowVal As String
    Dim qtyVal As Variant

    If loBalance Is Nothing Then Exit Function
    If loBalance.DataBodyRange Is Nothing Then Exit Function

    Dim cSku As Long: cSku = ColumnIndex(loBalance, "SKU")
    If cSku = 0 Then cSku = ColumnIndexLoose(loBalance, "SKU")
    Dim cCode As Long: cCode = ColumnIndex(loBalance, "ITEM_CODE")
    If cCode = 0 Then cCode = ColumnIndexLoose(loBalance, "ITEM_CODE", "ITEMCODE", "ITEM CODE")
    Dim cRow As Long: cRow = ColumnIndex(loBalance, "System_Key")
    Dim cQty As Long: cQty = ColumnIndex(loBalance, "QtyOnHand")
    If cQty = 0 Then cQty = ColumnIndex(loBalance, "Qty")
    If cQty = 0 Then cQty = ColumnIndexLoose(loBalance, "QTYONHAND", "QTY_ON_HAND", "QTY ON HAND", "ONHAND", "ON HAND", "BALANCE", "QTY", "QUANTITY", "TOTALINV", "TOTAL INV")
    If (cSku = 0 And cCode = 0 And cRow = 0) Or cQty = 0 Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    arr = loBalance.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        sku = ""
        itemCode = ""
        rowVal = ""
        If cSku > 0 Then sku = UCase$(Trim$(NzStr(arr(r, cSku))))
        If cCode > 0 Then itemCode = UCase$(Trim$(NzStr(arr(r, cCode))))
        If cRow > 0 Then rowVal = UCase$(Trim$(NzStr(arr(r, cRow))))
        qtyVal = arr(r, cQty)
        AddProductionBalanceValue dict, "SKU:" & sku, qtyVal
        AddProductionBalanceValue dict, sku, qtyVal
        AddProductionBalanceValue dict, "CODE:" & itemCode, qtyVal
        AddProductionBalanceValue dict, itemCode, qtyVal
        AddProductionBalanceValue dict, "SYSTEM_KEY:" & rowVal, qtyVal
    Next r
    Set BuildProductionSkuBalanceDictionary = dict
    Exit Function

FailSoft:
    Set BuildProductionSkuBalanceDictionary = Nothing
End Function

Private Sub AddProductionBalanceValue(ByVal dict As Object, ByVal key As String, ByVal qtyVal As Variant)
    key = UCase$(Trim$(key))
    If dict Is Nothing Then Exit Sub
    If key = "" Or Right$(key, 1) = ":" Then Exit Sub
    If dict.Exists(key) Then
        If IsNumeric(qtyVal) Then dict(key) = NzDbl(dict(key)) + CDbl(qtyVal)
    Else
        dict.Add key, qtyVal
    End If
End Sub

Private Function FindProductionListObjectByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindProductionListObjectByName = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindProductionListObjectByName Is Nothing Then Exit Function
    Next ws
End Function

Private Function BuildUsedDeltasFromPalette(ByVal wsProd As Worksheet) As Object
    If wsProd Is Nothing Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim seenRows As Object: Set seenRows = CreateObject("Scripting.Dictionary")
    Dim lo As ListObject
    For Each lo In wsProd.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row >= PALETTE_LINES_STAGING_ROW Then GoTo NextLo
            If lo.DataBodyRange Is Nothing Then GoTo NextLo

            Dim cRow As Long: cRow = ColumnIndex(lo, "System_Key")
            If cRow = 0 Then GoTo NextLo
            Dim cQty As Long: cQty = ColumnIndex(lo, "QUANTITY")
            If cQty = 0 Then GoTo NextLo
            Dim cIO As Long: cIO = ColumnIndex(lo, "INPUT/OUTPUT")
            Dim cProc As Long: cProc = ColumnIndex(lo, "PROCESS")

            Dim arr As Variant: arr = lo.DataBodyRange.value
            Dim r As Long
            For r = 1 To UBound(arr, 1)
                Dim rowKey As String
                rowKey = NormalizeSystemKey(arr(r, cRow))
                If rowKey = "" Then GoTo NextRow

                If cProc > 0 Then
                    Dim procName As String
                    procName = NzStr(arr(r, cProc))
                    If procName <> "" Then
                        If Not IsProcessSelected(procName, wsProd) Then GoTo NextRow
                    End If
                End If

                If cIO > 0 Then
                    Dim ioVal As String
                    ioVal = LCase$(Trim$(NzStr(arr(r, cIO))))
                    If ioVal <> "" And ioVal <> "used" Then GoTo NextRow
                End If

                Dim qty As Double
                qty = NzDbl(arr(r, cQty))
                If qty = 0 Then GoTo NextRow

                If seenRows.Exists(rowKey) Then GoTo NextRow
                seenRows.Add rowKey, True
                dict.Add rowKey, qty
NextRow:
            Next r
        End If
NextLo:
    Next lo

    If dict.count = 0 Then Exit Function
    Set BuildUsedDeltasFromPalette = dict
End Function

Private Function BuildUsedSnapshotFromCheck(ByVal loCheck As ListObject) As Object
    If loCheck Is Nothing Then Exit Function
    If loCheck.DataBodyRange Is Nothing Then Exit Function

    Dim cUsed As Long: cUsed = ColumnIndex(loCheck, "USED")
    If cUsed = 0 Then cUsed = ColumnIndexLoose(loCheck, "USED")
    Dim cRow As Long: cRow = ColumnIndex(loCheck, "System_Key")
    If cUsed = 0 Or cRow = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = loCheck.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowKey As String
        rowKey = NormalizeSystemKey(arr(r, cRow))
        If rowKey <> "" Then
            dict(rowKey) = NzDbl(arr(r, cUsed))
        End If
    Next r

    If dict.count = 0 Then Exit Function
    Set BuildUsedSnapshotFromCheck = dict
End Function

Private Function ValidateUsedStagingAgainstInvSys(ByVal invLo As ListObject, ByVal usedDict As Object, _
                                                  ByRef errNotes As String) As Double
    ValidateUsedStagingAgainstInvSys = -1
    If invLo Is Nothing Then
        AppendNote errNotes, "invSys table not found."
        Exit Function
    End If
    If usedDict Is Nothing Then
        AppendNote errNotes, "No USED quantities to stage."
        Exit Function
    ElseIf usedDict.count = 0 Then
        AppendNote errNotes, "No USED quantities to stage."
        Exit Function
    End If
    If invLo.DataBodyRange Is Nothing Then
        AppendNote errNotes, "invSys table has no data rows."
        Exit Function
    End If

    Dim systemKeyIndex As Object
    Set systemKeyIndex = BuildInvSysSystemKeyIndex(invLo)
    If systemKeyIndex Is Nothing Or systemKeyIndex.count = 0 Then
        AppendNote errNotes, "invSys System_Key identity index is not available."
        Exit Function
    End If

    Dim cTotal As Long: cTotal = ColumnIndex(invLo, "TOTAL INV")
    If cTotal = 0 Then cTotal = ColumnIndexLoose(invLo, "TOTALINV", "TOTAL_INV", "TOTALINVENTORY")

    Dim key As Variant
    For Each key In usedDict.keys
        Dim idx As Long
        Dim qtyValue As Double
        Dim identityLabel As String
        idx = ResolveUsedInventoryIndex(systemKeyIndex, CStr(key), usedDict(key), identityLabel)
        qtyValue = UsedEntryQty(usedDict(key))
        If idx <= 0 Then
            AppendNote errNotes, "invSys " & identityLabel & " not found; staging cancelled."
        ElseIf cTotal > 0 Then
            If qtyValue > NzDbl(invLo.DataBodyRange.Cells(idx, cTotal).Value) + 0.0000001 Then
                AppendNote errNotes, "invSys " & identityLabel & " does not have enough available inventory."
            End If
        End If
    Next key
    If errNotes <> "" Then Exit Function

    Dim total As Double
    For Each key In usedDict.keys
        total = total + UsedEntryQty(usedDict(key))
    Next key

    ValidateUsedStagingAgainstInvSys = total
End Function

Private Function SumProductionDeltaQuantities(ByVal deltas As Collection) As Double
    Dim delta As Variant
    If deltas Is Nothing Then Exit Function
    For Each delta In deltas
        SumProductionDeltaQuantities = SumProductionDeltaQuantities + NzDbl(delta("QTY"))
    Next delta
End Function

Private Sub WriteProdInvSysCheck(ByVal loCheck As ListObject, ByVal invLo As ListObject, ByVal usedDict As Object)
    If loCheck Is Nothing Then Exit Sub
    If usedDict Is Nothing Then
        ClearListObjectContents loCheck
        Exit Sub
    ElseIf usedDict.count = 0 Then
        ClearListObjectContents loCheck
        Exit Sub
    End If
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim systemKeyIndex As Object
    Set systemKeyIndex = BuildInvSysSystemKeyIndex(invLo)
    If systemKeyIndex Is Nothing Or systemKeyIndex.count = 0 Then Exit Sub

    Dim cUsedChk As Long: cUsedChk = ColumnIndex(loCheck, "USED")
    If cUsedChk = 0 Then cUsedChk = ColumnIndexLoose(loCheck, "USED")
    Dim cMadeChk As Long: cMadeChk = ColumnIndex(loCheck, "MADE")
    If cMadeChk = 0 Then cMadeChk = ColumnIndexLoose(loCheck, "MADE")
    Dim cTotalChk As Long: cTotalChk = ColumnIndex(loCheck, "TOTAL INV")
    If cTotalChk = 0 Then cTotalChk = ColumnIndexLoose(loCheck, "TOTALINV", "TOTAL_INV", "TOTALINVENTORY")
    Dim cSystemKeyChk As Long: cSystemKeyChk = ColumnIndex(loCheck, "System_Key")
    Dim cItemCodeChk As Long: cItemCodeChk = ColumnIndex(loCheck, "ITEM_CODE")
    Dim cItemChk As Long: cItemChk = ColumnIndex(loCheck, "ITEM")
    Dim cUomChk As Long: cUomChk = ColumnIndex(loCheck, "UOM")

    Dim cUsedInv As Long: cUsedInv = ColumnIndex(invLo, "USED")
    If cUsedInv = 0 Then cUsedInv = ColumnIndexLoose(invLo, "USED")
    Dim cMadeInv As Long: cMadeInv = ColumnIndex(invLo, "MADE")
    If cMadeInv = 0 Then cMadeInv = ColumnIndexLoose(invLo, "MADE")
    Dim cTotalInv As Long: cTotalInv = ColumnIndex(invLo, "TOTAL INV")
    If cTotalInv = 0 Then cTotalInv = ColumnIndexLoose(invLo, "TOTALINV", "TOTAL_INV", "TOTALINVENTORY")
    Dim cSystemKeyInv As Long: cSystemKeyInv = ColumnIndex(invLo, "System_Key")
    Dim cItemCodeInv As Long: cItemCodeInv = ColumnIndex(invLo, "ITEM_CODE")
    If cItemCodeInv = 0 Then cItemCodeInv = ColumnIndex(invLo, "SKU")
    Dim cItemInv As Long: cItemInv = ColumnIndex(invLo, "ITEM")
    If cItemInv = 0 Then cItemInv = ColumnIndex(invLo, "ItemName")
    Dim cUomInv As Long: cUomInv = ColumnIndex(invLo, "UOM")

    Dim keys As Variant
    keys = SortedKeys(usedDict)
    If IsEmpty(keys) Then Exit Sub

    Dim rowsNeeded As Long
    If IsArray(keys) Then
        rowsNeeded = UBound(keys) - LBound(keys) + 1
    Else
        rowsNeeded = 1
    End If
    If rowsNeeded <= 0 Then
        ClearListObjectData loCheck
        Exit Sub
    End If

    Dim cols As Long
    cols = TableColumnCount(loCheck)
    If cols <= 0 Then Exit Sub

    Dim currentRows As Long
    If loCheck.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = loCheck.DataBodyRange.rows.count
    End If
    If currentRows < rowsNeeded Then
        Call EnsureListObjectRowCountFullRow(loCheck, rowsNeeded)
    End If
    If loCheck.DataBodyRange Is Nothing Then Exit Sub
    currentRows = loCheck.DataBodyRange.rows.count
    If rowsNeeded > currentRows Then rowsNeeded = currentRows
    If rowsNeeded <= 0 Then Exit Sub

    Dim i As Long
    For i = 1 To loCheck.DataBodyRange.rows.count
        If i > rowsNeeded Then
            loCheck.DataBodyRange.rows(i).ClearContents
        Else
            Dim rowKey As String
            If IsArray(keys) Then
                rowKey = CStr(keys(LBound(keys) + i - 1))
            Else
                rowKey = CStr(keys)
            End If

            Dim invIdx As Long
            Dim identityLabel As String
            invIdx = ResolveUsedInventoryIndex(systemKeyIndex, rowKey, usedDict(rowKey), identityLabel)
            If invIdx > 0 Then
                If cUsedChk > 0 Then loCheck.DataBodyRange.Cells(i, cUsedChk).value = UsedEntryQty(usedDict(rowKey))
                If cMadeChk > 0 And cMadeInv > 0 Then loCheck.DataBodyRange.Cells(i, cMadeChk).value = NzDbl(invLo.DataBodyRange.Cells(invIdx, cMadeInv).value)
                If cTotalChk > 0 And cTotalInv > 0 Then loCheck.DataBodyRange.Cells(i, cTotalChk).value = NzDbl(invLo.DataBodyRange.Cells(invIdx, cTotalInv).value)
                If cSystemKeyChk > 0 And cSystemKeyInv > 0 Then loCheck.DataBodyRange.Cells(i, cSystemKeyChk).Value = NzStr(invLo.DataBodyRange.Cells(invIdx, cSystemKeyInv).Value)
                If cItemCodeChk > 0 And cItemCodeInv > 0 Then loCheck.DataBodyRange.Cells(i, cItemCodeChk).Value = NzStr(invLo.DataBodyRange.Cells(invIdx, cItemCodeInv).Value)
                If cItemChk > 0 And cItemInv > 0 Then loCheck.DataBodyRange.Cells(i, cItemChk).Value = NzStr(invLo.DataBodyRange.Cells(invIdx, cItemInv).Value)
                If cUomChk > 0 And cUomInv > 0 Then loCheck.DataBodyRange.Cells(i, cUomChk).Value = NzStr(invLo.DataBodyRange.Cells(invIdx, cUomInv).Value)
            End If
        End If
    Next i
End Sub

Private Function BuildInvSysSystemKeyIndex(ByVal invLo As ListObject) As Object
    Dim dict As Object
    Dim cSystemKey As Long
    Dim r As Long
    Dim systemKey As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then
        Set BuildInvSysSystemKeyIndex = dict
        Exit Function
    End If
    cSystemKey = ColumnIndex(invLo, "System_Key")
    If cSystemKey = 0 Then
        Set BuildInvSysSystemKeyIndex = dict
        Exit Function
    End If

    For r = 1 To invLo.ListRows.count
        systemKey = Trim$(NzStr(invLo.DataBodyRange.Cells(r, cSystemKey).Value))
        If systemKey <> "" Then
            If Not dict.Exists(systemKey) Then dict.Add systemKey, r
        End If
    Next r
    Set BuildInvSysSystemKeyIndex = dict
End Function

Private Function ResolveUsedInventoryIndex(ByVal systemKeyIndex As Object, _
                                           ByVal dictKey As String, _
                                           ByVal usedEntry As Variant, _
                                           ByRef identityLabel As String) As Long
    Dim systemKey As String

    If IsObject(usedEntry) Then
        On Error Resume Next
        systemKey = Trim$(NzStr(usedEntry("System_Key")))
        On Error GoTo 0
    End If

    If systemKey = "" Then
        systemKey = Trim$(dictKey)
        If Left$(systemKey, 4) = "SYS|" Then systemKey = Mid$(systemKey, 5)
    End If
    identityLabel = "System_Key " & systemKey
    If Not systemKeyIndex Is Nothing Then
        If systemKeyIndex.Exists(systemKey) Then ResolveUsedInventoryIndex = CLng(systemKeyIndex(systemKey))
    End If
End Function

Private Function ResolveBoundProductionWorkbook(Optional ByVal requiredSheet As String = "") As Workbook
    On Error GoTo InvalidBinding

    If mProductionOperatorWorkbook Is Nothing Then Exit Function
    If mProductionOperatorWorkbook.IsAddin Then GoTo InvalidBinding
    If requiredSheet <> "" Then
        If WorkbookSheetExists(mProductionOperatorWorkbook, requiredSheet) Is Nothing Then Exit Function
    End If
    Set ResolveBoundProductionWorkbook = mProductionOperatorWorkbook
    Exit Function

InvalidBinding:
    Set mProductionOperatorWorkbook = Nothing
End Function

Private Function UsedEntryQty(ByVal usedEntry As Variant) As Double
    If IsObject(usedEntry) Then
        On Error Resume Next
        UsedEntryQty = NzDbl(usedEntry("QTY"))
        On Error GoTo 0
    Else
        UsedEntryQty = NzDbl(usedEntry)
    End If
End Function

Private Function BuildOutputEntriesFromProcessTables(ByVal wsProd As Worksheet) As Collection
    If wsProd Is Nothing Then Exit Function

    Dim recipeId As String
    recipeId = GetRecipeChooserRecipeId(wsProd)

    Dim procTables As Collection
    Set procTables = GetRecipeChooserProcessTables(wsProd)
    If procTables Is Nothing Then Exit Function
    If procTables.count = 0 Then Exit Function

    Dim entryMap As Object: Set entryMap = CreateObject("Scripting.Dictionary")
    Dim order As New Collection

    Dim lo As ListObject
    For Each lo In procTables
        If lo Is Nothing Then GoTo NextLo
        If lo.DataBodyRange Is Nothing Then GoTo NextLo

        Dim cIO As Long: cIO = ColumnIndex(lo, "INPUT/OUTPUT")
        Dim cIng As Long: cIng = ColumnIndex(lo, "INGREDIENT")
        If cIO = 0 Or cIng = 0 Then GoTo NextLo

        Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
        Dim cAmt As Long: cAmt = ColumnIndex(lo, "AMOUNT NEEDED")
        If cAmt = 0 Then cAmt = ColumnIndex(lo, "AMOUNT")
        Dim cProc As Long: cProc = ColumnIndex(lo, "PROCESS")
        Dim cIngId As Long: cIngId = ColumnIndex(lo, "INGREDIENT_ID")

        Dim arr As Variant: arr = lo.DataBodyRange.value
        Dim r As Long
        For r = 1 To UBound(arr, 1)
            Dim ioVal As String
            ioVal = NzStr(arr(r, cIO))
            If Not IsOutputIoValue(ioVal) Then GoTo NextRow

            Dim procName As String
            If cProc > 0 Then procName = NzStr(arr(r, cProc))
            If procName = "" Then procName = ProcessNameFromTable(lo)
            If procName <> "" Then
                If Not IsProcessSelected(procName, wsProd) Then GoTo NextRow
            End If

            Dim outputName As String
            outputName = NzStr(arr(r, cIng))
            If outputName = "" Then GoTo NextRow

            Dim uomVal As String
            If cUom > 0 Then uomVal = NzStr(arr(r, cUom))
            Dim qtyVal As Double
            If cAmt > 0 Then qtyVal = NzDbl(arr(r, cAmt))
            Dim ingId As String
            If cIngId > 0 Then ingId = NzStr(arr(r, cIngId))

            Dim key As String
            key = BuildOutputKey(procName, outputName)
            If Not entryMap.Exists(key) Then
                Dim entry As Object: Set entry = CreateObject("Scripting.Dictionary")
                entry("PROCESS") = procName
                entry("OUTPUT") = outputName
                entry("UOM") = uomVal
                entry("QTY") = qtyVal
                entry("INGREDIENT_ID") = ingId
                entry("RECIPE_ID") = recipeId
                entryMap.Add key, entry
                order.Add key
            Else
                Dim existing As Object
                Set existing = entryMap(key)
                existing("QTY") = NzDbl(existing("QTY")) + qtyVal
                If NzStr(existing("UOM")) = "" Then existing("UOM") = uomVal
                If NzStr(existing("INGREDIENT_ID")) = "" Then existing("INGREDIENT_ID") = ingId
                If NzStr(existing("RECIPE_ID")) = "" Then existing("RECIPE_ID") = recipeId
            End If
NextRow:
        Next r
NextLo:
    Next lo

    If order.count = 0 Then Exit Function

    Dim result As New Collection
    Dim k As Variant
    For Each k In order
        result.Add entryMap(k)
    Next k
    Set BuildOutputEntriesFromProcessTables = result
End Function

Private Sub EnsureProductionOutputHeaderOrder(ByVal loOut As ListObject)
    If loOut Is Nothing Then Exit Sub
    If loOut.HeaderRowRange Is Nothing Then Exit Sub

    Dim cUom As Long
    cUom = ColumnIndex(loOut, "UOM")
    If cUom = 0 Then Exit Sub
    If cUom + 2 > loOut.ListColumns.count Then Exit Sub

    Dim h1 As String
    Dim h2 As String
    h1 = Trim$(NzStr(loOut.HeaderRowRange.Cells(1, cUom + 1).value))
    h2 = Trim$(NzStr(loOut.HeaderRowRange.Cells(1, cUom + 2).value))

    If StrComp(h1, "BATCH", vbTextCompare) = 0 And StrComp(h2, "REAL OUTPUT", vbTextCompare) = 0 Then
        On Error Resume Next
        loOut.ListColumns(cUom + 1).Name = "REAL OUTPUT"
        loOut.ListColumns(cUom + 2).Name = "BATCH"
        On Error GoTo 0
    End If
End Sub

Private Sub UpdateProductionOutputTable(ByVal loOut As ListObject, ByVal entries As Collection, ByVal invLo As ListObject, ByRef errNotes As String)
    If loOut Is Nothing Then Exit Sub
    If entries Is Nothing Then Exit Sub
    If entries.count = 0 Then Exit Sub

    EnsureProductionOutputHeaderOrder loOut

    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    Dim cOutput As Long: cOutput = ColumnIndex(loOut, "OUTPUT")
    Dim cUom As Long: cUom = ColumnIndex(loOut, "UOM")
    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cRecall As Long: cRecall = ColumnIndex(loOut, "RECALL CODE")

    If cProc = 0 Or cOutput = 0 Then
        AppendNote errNotes, "ProductionOutput missing PROCESS/OUTPUT columns."
        Exit Sub
    End If

    Dim cSystemKey As Long
    cSystemKey = EnsureProductionOutputSystemKeyColumn(loOut)
    Dim cItemCode As Long
    cItemCode = EnsureProductionOutputItemCodeColumn(loOut)

    Dim existing As Object: Set existing = CreateObject("Scripting.Dictionary")
    If Not loOut.DataBodyRange Is Nothing Then
        Dim arr As Variant: arr = loOut.DataBodyRange.value
        Dim r As Long
        For r = 1 To UBound(arr, 1)
            Dim key As String
            key = BuildOutputKey(NzStr(arr(r, cProc)), NzStr(arr(r, cOutput)))
            If key <> "|" Then
                If Not existing.Exists(key) Then existing.Add key, r
            End If
        Next r
    End If

    Dim i As Long
    Dim currentRows As Long
    If loOut.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = loOut.DataBodyRange.rows.count
    End If

    Dim addCount As Long
    For i = 1 To entries.count
        Dim entryCount As Object
        Set entryCount = entries(i)
        Dim addKey As String
        addKey = BuildOutputKey(NzStr(entryCount("PROCESS")), NzStr(entryCount("OUTPUT")))
        If addKey <> "|" Then
            If Not existing.Exists(addKey) Then addCount = addCount + 1
        End If
    Next i
    If addCount > 0 Then
        Dim emptySlots As Long
        If Not loOut.DataBodyRange Is Nothing Then
            Dim rEmpty As Long
            For rEmpty = 1 To loOut.DataBodyRange.rows.count
                Dim procVal As String
                Dim outVal As String
                If cProc > 0 Then procVal = NzStr(loOut.DataBodyRange.Cells(rEmpty, cProc).value)
                If cOutput > 0 Then outVal = NzStr(loOut.DataBodyRange.Cells(rEmpty, cOutput).value)
                If Trim$(procVal) = "" And Trim$(outVal) = "" Then emptySlots = emptySlots + 1
            Next rEmpty
        End If
        Dim needRows As Long
        needRows = addCount - emptySlots
        If needRows > 0 Then
            Call EnsureListObjectRowCountFullRow(loOut, currentRows + needRows)
        End If
    End If

    Dim NextRow As Long
    NextRow = currentRows + 1

    For i = 1 To entries.count
        Dim entry As Object
        Set entry = entries(i)
        Dim procName As String: procName = NzStr(entry("PROCESS"))
        Dim outputName As String: outputName = NzStr(entry("OUTPUT"))
        Dim uomVal As String: uomVal = NzStr(entry("UOM"))

        Dim outKey As String
        outKey = BuildOutputKey(procName, outputName)
        If outKey = "|" Then GoTo NextEntry

        Dim targetRow As Long

        If existing.Exists(outKey) Then
            targetRow = CLng(existing(outKey))
            If cProc > 0 Then loOut.DataBodyRange.Cells(targetRow, cProc).value = procName
            If cOutput > 0 Then loOut.DataBodyRange.Cells(targetRow, cOutput).value = outputName
            If cUom > 0 Then
                If NzStr(loOut.DataBodyRange.Cells(targetRow, cUom).value) = "" Then
                    loOut.DataBodyRange.Cells(targetRow, cUom).value = uomVal
                End If
            End If
        Else
            targetRow = FindFirstEmptyOutputRow(loOut, cProc, cOutput)
            If targetRow = 0 Then
                targetRow = NextRow
                NextRow = NextRow + 1
            End If
            If cProc > 0 Then loOut.DataBodyRange.Cells(targetRow, cProc).value = procName
            If cOutput > 0 Then loOut.DataBodyRange.Cells(targetRow, cOutput).value = outputName
            If cUom > 0 Then loOut.DataBodyRange.Cells(targetRow, cUom).value = uomVal
            existing.Add outKey, targetRow
        End If

        If cSystemKey > 0 Then
            Dim systemKey As String
            Dim candidateSystemKey As String
            Dim itemCodeVal As String
            Dim resolvedItemName As String
            If Not loOut.DataBodyRange Is Nothing Then
                systemKey = Trim$(NzStr(loOut.DataBodyRange.Cells(targetRow, cSystemKey).value))
                If cItemCode > 0 Then itemCodeVal = NzStr(loOut.DataBodyRange.Cells(targetRow, cItemCode).Value)
            End If
            candidateSystemKey = systemKey
            ResolveProductionOutputIdentity outputName, candidateSystemKey, itemCodeVal, resolvedItemName, errNotes
            ' A prepared output is a new durable entity.  The completion service
            ' allocates its System_Key once; picker lookup supplies SKU metadata only.
            If cItemCode > 0 And Not loOut.DataBodyRange Is Nothing Then
                loOut.DataBodyRange.Cells(targetRow, cItemCode).Value = itemCodeVal
            End If
        End If
NextEntry:
    Next i
End Sub

Private Function FindFirstEmptyOutputRow(ByVal loOut As ListObject, ByVal cProc As Long, ByVal cOutput As Long) As Long
    FindFirstEmptyOutputRow = 0
    If loOut Is Nothing Then Exit Function
    If loOut.DataBodyRange Is Nothing Then
        FindFirstEmptyOutputRow = 1
        Exit Function
    End If
    If cProc = 0 And cOutput = 0 Then Exit Function

    Dim r As Long
    For r = 1 To loOut.DataBodyRange.rows.count
        Dim procVal As String
        Dim outVal As String
        If cProc > 0 Then procVal = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
        If cOutput > 0 Then outVal = NzStr(loOut.DataBodyRange.Cells(r, cOutput).value)
        If Trim$(procVal) = "" And Trim$(outVal) = "" Then
            FindFirstEmptyOutputRow = r
            Exit Function
        End If
    Next r
End Function

Private Function EnsureProductionOutputSystemKeyColumn(ByVal loOut As ListObject) As Long
    If loOut Is Nothing Then Exit Function
    Dim cSystemKey As Long: cSystemKey = ColumnIndex(loOut, "System_Key")
    If cSystemKey = 0 Then
        On Error Resume Next
        Dim newCol As ListColumn
        Set newCol = loOut.ListColumns.Add
        If Not newCol Is Nothing Then
            newCol.Name = "System_Key"
            cSystemKey = newCol.Index
        End If
        On Error GoTo 0
    End If
    EnsureProductionOutputSystemKeyColumn = cSystemKey
End Function

Private Function EnsureProductionOutputItemCodeColumn(ByVal loOut As ListObject) As Long
    If loOut Is Nothing Then Exit Function
    Dim cItemCode As Long: cItemCode = ColumnIndex(loOut, "ITEM_CODE")
    If cItemCode = 0 Then
        On Error Resume Next
        Dim newCol As ListColumn
        Set newCol = loOut.ListColumns.Add
        If Not newCol Is Nothing Then
            newCol.Name = "ITEM_CODE"
            cItemCode = newCol.Index
        End If
        On Error GoTo 0
    End If
    EnsureProductionOutputItemCodeColumn = cItemCode
End Function

Private Function BuildOutputKey(ByVal procName As String, ByVal outputName As String) As String
    BuildOutputKey = NormalizeOutputKey(procName) & "|" & NormalizeOutputKey(outputName)
End Function

Private Function NormalizeOutputKey(ByVal v As String) As String
    NormalizeOutputKey = LCase$(Trim$(v))
End Function

Private Function NormalizeLookupKey(ByVal v As String) As String
    Dim s As String
    s = Trim$(v)
    If s = "" Then Exit Function
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    On Error Resume Next
    s = Application.WorksheetFunction.Trim(s)
    On Error GoTo 0
    NormalizeLookupKey = LCase$(s)
End Function

Private Function IsOutputIoValue(ByVal ioVal As String) As Boolean
    Dim v As String
    v = LCase$(Trim$(ioVal))
    If v = "" Then Exit Function
    If v = "made" Or v = "output" Or v = "out" Then IsOutputIoValue = True
End Function

Private Function IsInputIoValue(ByVal ioVal As String) As Boolean
    Dim v As String
    v = LCase$(Trim$(ioVal))
    If v = "" Then Exit Function
    If v = "used" Or v = "input" Or v = "in" Then IsInputIoValue = True
End Function

Private Function BuildInventoryOutputIdentityLookup(ByVal invLo As ListObject) As Object
    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function

    Dim cSystemKey As Long: cSystemKey = ColumnIndex(invLo, "System_Key")
    Dim cItem As Long: cItem = ColumnIndex(invLo, "ITEM")
    Dim cCode As Long: cCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim cDesc As Long: cDesc = ColumnIndex(invLo, "DESCRIPTION")
    If cSystemKey = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim arr As Variant: arr = invLo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim systemKey As String: systemKey = Trim$(NzStr(arr(r, cSystemKey)))
        If systemKey = "" Then GoTo NextInventory
        Dim itemName As String
        Dim itemCode As String
        Dim descVal As String
        If cItem > 0 Then itemName = NzStr(arr(r, cItem))
        If cCode > 0 Then itemCode = NzStr(arr(r, cCode))
        If cDesc > 0 Then descVal = NzStr(arr(r, cDesc))
        If itemName <> "" Then
            Dim keyName As String: keyName = NormalizeLookupKey(itemName)
            If keyName <> "" Then
                If Not dict.Exists(keyName) Then dict.Add keyName, systemKey
            End If
        End If
        If itemCode <> "" Then
            Dim keyCode As String: keyCode = NormalizeLookupKey(itemCode)
            If keyCode <> "" Then
                If Not dict.Exists(keyCode) Then dict.Add keyCode, systemKey
            End If
        End If
        If descVal <> "" Then
            Dim keyDesc As String: keyDesc = NormalizeLookupKey(descVal)
            If keyDesc <> "" Then
                If Not dict.Exists(keyDesc) Then dict.Add keyDesc, systemKey
            End If
        End If
NextInventory:
    Next r

    If dict.count = 0 Then Exit Function
    Set BuildInventoryOutputIdentityLookup = dict
End Function

Private Function LookupOutputSystemKey(ByVal outputLookup As Object, _
                                       ByVal outputName As String) As String
    If outputLookup Is Nothing Then Exit Function
    Dim key As String: key = NormalizeLookupKey(outputName)
    If key = "" Then Exit Function
    If outputLookup.Exists(key) Then LookupOutputSystemKey = Trim$(NzStr(outputLookup(key)))
End Function

Private Function ResolveProductionOutputSystemKey(ByVal invLo As ListObject, _
                                                  ByVal outputLookup As Object, _
                                                  ByVal existingSystemKey As String, _
                                                  ByVal outputName As String, _
                                                  ByVal recipeId As String, _
                                                  ByVal ingredientId As String, _
                                                  ByRef errNotes As String) As String
    Dim allowed As Object
    Dim allowedKey As Variant
    Dim firstAllowed As String
    Dim lookupSystemKey As String

    existingSystemKey = Trim$(existingSystemKey)
    If existingSystemKey <> "" Then
        If InventorySystemKeyMatchesOutputName(invLo, existingSystemKey, outputName) Then
            ResolveProductionOutputSystemKey = existingSystemKey
            Exit Function
        End If
    End If

    If Trim$(recipeId) <> "" And Trim$(ingredientId) <> "" Then
        Set allowed = GetAllowedInventorySystemKeysForIngredient(recipeId, ingredientId)
        If Not allowed Is Nothing Then
            For Each allowedKey In allowed.keys
                If Trim$(NzStr(allowedKey)) <> "" Then
                    If firstAllowed = "" Then firstAllowed = Trim$(NzStr(allowedKey))
                    If InventorySystemKeyMatchesOutputName(invLo, Trim$(NzStr(allowedKey)), outputName) Then
                        ResolveProductionOutputSystemKey = Trim$(NzStr(allowedKey))
                        Exit Function
                    End If
                End If
            Next allowedKey
            If firstAllowed <> "" Then
                ResolveProductionOutputSystemKey = firstAllowed
                Exit Function
            End If
        End If
    End If

    lookupSystemKey = LookupOutputSystemKey(outputLookup, outputName)
    If lookupSystemKey <> "" Then
        ResolveProductionOutputSystemKey = lookupSystemKey
        Exit Function
    End If
    lookupSystemKey = LookupOutputSystemKeyLoose(invLo, outputName, errNotes)
    If lookupSystemKey <> "" Then
        ResolveProductionOutputSystemKey = lookupSystemKey
        Exit Function
    End If
    Dim pickerItems As Variant
    pickerItems = LoadProductionInventoryPickerItems("")
    lookupSystemKey = LookupOutputSystemKeyFromPicker(pickerItems, outputName, errNotes)
    If lookupSystemKey <> "" Then
        ResolveProductionOutputSystemKey = lookupSystemKey
        Exit Function
    End If

    If existingSystemKey <> "" Then
        AppendNote errNotes, "Production output System_Key '" & existingSystemKey & _
            "' did not match output '" & outputName & "' and no replacement inventory entity was found."
    Else
        AppendNote errNotes, "No inventory entity was found for production output '" & outputName & "'."
    End If
End Function

Private Function ResolveProductionOutputIdentity(ByVal outputName As String, _
                                                 ByRef systemKey As String, _
                                                 ByRef itemCode As String, _
                                                 ByRef itemName As String, _
                                                 ByRef errNotes As String) As Boolean
    Dim pickerItems As Variant

    pickerItems = LoadProductionInventoryPickerItems("")
    ResolveProductionOutputIdentity = ResolveProductionOutputIdentityFromRows( _
        pickerItems, outputName, systemKey, itemCode, itemName, errNotes)
End Function

Private Function ResolveProductionOutputIdentityFromRows(ByVal pickerItems As Variant, _
                                                         ByVal outputName As String, _
                                                         ByRef systemKey As String, _
                                                         ByRef itemCode As String, _
                                                         ByRef itemName As String, _
                                                         ByRef errNotes As String) As Boolean
    Dim wantedTokens As Object
    Dim identities As Object
    Dim candidateData As Variant
    Dim pass As Long
    Dim r As Long
    Dim candidateSystemKey As String
    Dim candidateCode As String
    Dim candidateName As String
    Dim candidateText As String
    Dim identityKey As String
    Dim matched As Boolean
    Dim key As Variant

    outputName = Trim$(outputName)
    If outputName = "" Then Exit Function
    If IsEmpty(pickerItems) Or Not IsArray(pickerItems) Then Exit Function
    If UBound(pickerItems, 2) < 7 Then Exit Function
    Set wantedTokens = LookupTokens(outputName)

    For pass = 1 To 2
        Set identities = CreateObject("Scripting.Dictionary")
        identities.CompareMode = vbTextCompare
        For r = LBound(pickerItems, 1) To UBound(pickerItems, 1)
            candidateSystemKey = Trim$(NzStr(pickerItems(r, 1)))
            candidateName = Trim$(NzStr(pickerItems(r, 2)))
            candidateCode = Trim$(NzStr(pickerItems(r, 7)))
            If candidateCode = "" And candidateSystemKey = "" Then GoTo NextPickerRow

            If pass = 1 Then
                matched = (StrComp(candidateName, outputName, vbTextCompare) = 0)
            Else
                candidateText = candidateName & " " & NzStr(pickerItems(r, 6)) & " " & candidateCode
                matched = LookupCandidateMatchesTokens(candidateText, wantedTokens)
            End If
            If Not matched Then GoTo NextPickerRow

            If candidateCode <> "" Then
                identityKey = "SKU|" & candidateCode
            Else
                identityKey = "SYSTEM_KEY|" & candidateSystemKey
            End If
            If Not identities.Exists(identityKey) Then
                identities.Add identityKey, Array(candidateSystemKey, candidateCode, candidateName)
            End If
NextPickerRow:
        Next r

        If identities.Count = 1 Then
            For Each key In identities.Keys
                candidateData = identities(key)
                If Trim$(systemKey) = "" Then systemKey = Trim$(NzStr(candidateData(0)))
                If Trim$(itemCode) = "" Then itemCode = NzStr(candidateData(1))
                If Trim$(itemName) = "" Then itemName = NzStr(candidateData(2))
                ResolveProductionOutputIdentityFromRows = True
                Exit Function
            Next key
        ElseIf identities.Count > 1 Then
            AppendNote errNotes, "Production output '" & outputName & _
                "' matched more than one canonical SKU. Select a unique inventory output."
            Exit Function
        End If
    Next pass
End Function

Private Function LookupOutputSystemKeyFromPicker(ByVal pickerItems As Variant, _
                                                 ByVal outputName As String, _
                                                 ByRef errNotes As String) As String
    Dim wantedTokens As Object
    Dim r As Long
    Dim systemKey As String
    Dim candidateText As String
    Dim candidateName As String
    Dim matchedSystemKey As String
    Dim matchedName As String

    Set wantedTokens = LookupTokens(outputName)
    If wantedTokens Is Nothing Then Exit Function
    If wantedTokens.count = 0 Then Exit Function
    If IsEmpty(pickerItems) Then Exit Function
    If Not IsArray(pickerItems) Then Exit Function

    For r = LBound(pickerItems, 1) To UBound(pickerItems, 1)
        systemKey = Trim$(NzStr(pickerItems(r, 1)))
        If systemKey = "" Then GoTo NextPicker
        candidateName = NzStr(pickerItems(r, 2))
        candidateText = candidateName & " " & NzStr(pickerItems(r, 6)) & " " & NzStr(pickerItems(r, 7))
        If LookupCandidateMatchesTokens(candidateText, wantedTokens) Then
            If matchedSystemKey = "" Then
                matchedSystemKey = systemKey
                matchedName = candidateName
            ElseIf StrComp(matchedSystemKey, systemKey, vbTextCompare) <> 0 Then
                AppendNote errNotes, "Production output '" & outputName & _
                    "' matched more than one canonical inventory entity. Select a unique System_Key before completing the run."
                Exit Function
            End If
        End If
NextPicker:
    Next r

    If matchedSystemKey <> "" Then
        AppendNote errNotes, "Production output '" & outputName & _
            "' resolved to canonical System_Key '" & matchedSystemKey & "'" & _
            IIf(matchedName <> "", " (" & matchedName & ")", "") & "."
        LookupOutputSystemKeyFromPicker = matchedSystemKey
    End If
End Function

Private Function LookupOutputSystemKeyLoose(ByVal invLo As ListObject, _
                                            ByVal outputName As String, _
                                            ByRef errNotes As String) As String
    Dim wantedTokens As Object
    Dim cSystemKey As Long
    Dim cItem As Long
    Dim cCode As Long
    Dim cDesc As Long
    Dim arr As Variant
    Dim r As Long
    Dim matchedSystemKey As String
    Dim matchedName As String

    Set wantedTokens = LookupTokens(outputName)
    If wantedTokens Is Nothing Then Exit Function
    If wantedTokens.count = 0 Then Exit Function
    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function

    cSystemKey = ColumnIndex(invLo, "System_Key")
    cItem = ColumnIndex(invLo, "ITEM")
    cCode = ColumnIndex(invLo, "ITEM_CODE")
    cDesc = ColumnIndex(invLo, "DESCRIPTION")
    If cSystemKey = 0 Then Exit Function

    arr = invLo.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        Dim systemKey As String
        Dim candidateText As String
        Dim candidateName As String

        systemKey = Trim$(NzStr(arr(r, cSystemKey)))
        If systemKey = "" Then GoTo NextInventory
        If cItem > 0 Then candidateName = NzStr(arr(r, cItem))
        candidateText = candidateName
        If cDesc > 0 Then candidateText = candidateText & " " & NzStr(arr(r, cDesc))
        If cCode > 0 Then candidateText = candidateText & " " & NzStr(arr(r, cCode))
        If LookupCandidateMatchesTokens(candidateText, wantedTokens) Then
            If matchedSystemKey = "" Then
                matchedSystemKey = systemKey
                matchedName = candidateName
            ElseIf StrComp(matchedSystemKey, systemKey, vbTextCompare) <> 0 Then
                AppendNote errNotes, "Production output '" & outputName & _
                    "' matched more than one inventory entity. Select a unique System_Key before completing the run."
                Exit Function
            End If
        End If
NextInventory:
    Next r

    If matchedSystemKey <> "" Then
        AppendNote errNotes, "Production output '" & outputName & _
            "' resolved to System_Key '" & matchedSystemKey & "'" & _
            IIf(matchedName <> "", " (" & matchedName & ")", "") & "."
        LookupOutputSystemKeyLoose = matchedSystemKey
    End If
End Function

Private Function LookupCandidateMatchesTokens(ByVal candidateText As String, ByVal wantedTokens As Object) As Boolean
    Dim candidateTokens As Object
    Dim wanted As Variant

    Set candidateTokens = LookupTokens(candidateText)
    If candidateTokens Is Nothing Then Exit Function
    If candidateTokens.count = 0 Then Exit Function

    For Each wanted In wantedTokens.keys
        If Not LookupTokenExists(CStr(wanted), candidateTokens) Then Exit Function
    Next wanted
    LookupCandidateMatchesTokens = True
End Function

Private Function LookupTokenExists(ByVal wanted As String, ByVal candidateTokens As Object) As Boolean
    Dim candidate As Variant

    If candidateTokens Is Nothing Then Exit Function
    wanted = LCase$(Trim$(wanted))
    If wanted = "" Then Exit Function

    For Each candidate In candidateTokens.keys
        If CStr(candidate) = wanted Then
            LookupTokenExists = True
            Exit Function
        End If
        If Len(wanted) >= 4 And Len(CStr(candidate)) >= Len(wanted) Then
            If Left$(CStr(candidate), Len(wanted)) = wanted Then
                LookupTokenExists = True
                Exit Function
            End If
        End If
        If Len(CStr(candidate)) >= 4 And Len(wanted) >= Len(CStr(candidate)) Then
            If Left$(wanted, Len(CStr(candidate))) = CStr(candidate) Then
                LookupTokenExists = True
                Exit Function
            End If
        End If
    Next candidate
End Function

Private Function LookupTokens(ByVal value As String) As Object
    Dim dict As Object
    Dim cleaned As String
    Dim i As Long
    Dim ch As String
    Dim parts As Variant
    Dim part As Variant

    value = LCase$(Trim$(value))
    If value = "" Then Exit Function
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            cleaned = cleaned & ch
        Else
            cleaned = cleaned & " "
        End If
    Next i
    On Error Resume Next
    cleaned = Application.WorksheetFunction.Trim(cleaned)
    On Error GoTo 0
    If cleaned = "" Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    parts = Split(cleaned, " ")
    For Each part In parts
        Dim token As String
        token = Trim$(CStr(part))
        If Len(token) >= 3 Then
            If Not dict.Exists(token) Then dict.Add token, True
        End If
    Next part
    Set LookupTokens = dict
End Function

Private Function InventorySystemKeyMatchesOutputName(ByVal invLo As ListObject, _
                                                     ByVal systemKey As String, _
                                                     ByVal outputName As String) As Boolean
    Dim cSystemKey As Long
    Dim cItem As Long
    Dim cCode As Long
    Dim cDesc As Long
    Dim arr As Variant
    Dim r As Long
    Dim wanted As String

    If invLo Is Nothing Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function
    systemKey = Trim$(systemKey)
    If systemKey = "" Then Exit Function
    wanted = NormalizeLookupKey(outputName)
    If wanted = "" Then Exit Function

    cSystemKey = ColumnIndex(invLo, "System_Key")
    cItem = ColumnIndex(invLo, "ITEM")
    cCode = ColumnIndex(invLo, "ITEM_CODE")
    cDesc = ColumnIndex(invLo, "DESCRIPTION")
    If cSystemKey = 0 Then Exit Function

    arr = invLo.DataBodyRange.value
    For r = 1 To UBound(arr, 1)
        If StrComp(Trim$(NzStr(arr(r, cSystemKey))), systemKey, vbTextCompare) = 0 Then
            If cItem > 0 Then
                If NormalizeLookupKey(NzStr(arr(r, cItem))) = wanted Then InventorySystemKeyMatchesOutputName = True
            End If
            If Not InventorySystemKeyMatchesOutputName And cCode > 0 Then
                If NormalizeLookupKey(NzStr(arr(r, cCode))) = wanted Then InventorySystemKeyMatchesOutputName = True
            End If
            If Not InventorySystemKeyMatchesOutputName And cDesc > 0 Then
                If NormalizeLookupKey(NzStr(arr(r, cDesc))) = wanted Then InventorySystemKeyMatchesOutputName = True
            End If
            Exit Function
        End If
    Next r
End Function

Private Function BuildUsedDeltaPacketFromInvSys(ByVal invLo As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function

    Dim colUsed As Long: colUsed = ColumnIndex(invLo, "USED")
    Dim colSystemKey As Long: colSystemKey = ColumnIndex(invLo, "System_Key")
    Dim colItemCode As Long: colItemCode = ColumnIndex(invLo, "ITEM_CODE")
    Dim colItemName As Long: colItemName = ColumnIndex(invLo, "ITEM")
    If colUsed = 0 Or colSystemKey = 0 Then
        errNotes = "invSys table missing USED/System_Key columns."
        Exit Function
    End If

    Dim result As New Collection
    Dim arr As Variant: arr = invLo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim usedVal As Double: usedVal = NzDbl(arr(r, colUsed))
        Dim systemKey As String: systemKey = Trim$(NzStr(arr(r, colSystemKey)))
        If systemKey = "" Or usedVal <= 0 Then GoTo NextRow
        Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
        delta("System_Key") = systemKey
        delta("QTY") = usedVal
        If colItemCode > 0 Then delta("ITEM_CODE") = NzStr(arr(r, colItemCode))
        If colItemName > 0 Then delta("ITEM_NAME") = NzStr(arr(r, colItemName))
        result.Add delta
NextRow:
    Next r

    If result.count = 0 Then
        errNotes = "No staged usage found in invSys.USED."
        Exit Function
    End If
    Set BuildUsedDeltaPacketFromInvSys = result
End Function

Private Function BuildUsedDeltaPacketFromCheck(ByVal loCheck As ListObject, ByVal invLo As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If loCheck Is Nothing Then
        errNotes = "No checked-in production input rows were found."
        Exit Function
    End If
    If loCheck.DataBodyRange Is Nothing Then
        errNotes = "No checked-in production input rows were found."
        Exit Function
    End If
    Dim cSystemKeyChk As Long: cSystemKeyChk = ColumnIndex(loCheck, "System_Key")
    Dim cUsedChk As Long: cUsedChk = ColumnIndex(loCheck, "USED")
    If cUsedChk = 0 Then cUsedChk = ColumnIndexLoose(loCheck, "USED")
    Dim cItemCodeChk As Long: cItemCodeChk = ColumnIndex(loCheck, "ITEM_CODE")
    Dim cItemNameChk As Long: cItemNameChk = ColumnIndex(loCheck, "ITEM")
    If cSystemKeyChk = 0 Or cUsedChk = 0 Then
        errNotes = "Prod_invSys_Check table missing System_Key/USED columns."
        Exit Function
    End If

    Dim systemKeyIndex As Object
    If Not invLo Is Nothing Then
        If Not invLo.DataBodyRange Is Nothing Then Set systemKeyIndex = BuildInvSysSystemKeyIndex(invLo)
    End If

    Dim cItemCode As Long
    Dim cItemName As Long
    If Not invLo Is Nothing Then
        cItemCode = ColumnIndex(invLo, "ITEM_CODE")
        cItemName = ColumnIndex(invLo, "ITEM")
    End If
    Dim result As New Collection
    Dim r As Long
    For r = 1 To loCheck.ListRows.count
        Dim systemKey As String
        systemKey = Trim$(NzStr(loCheck.DataBodyRange.Cells(r, cSystemKeyChk).value))
        Dim itemCodeVal As String
        If cItemCodeChk > 0 Then itemCodeVal = Trim$(NzStr(loCheck.DataBodyRange.Cells(r, cItemCodeChk).Value))
        Dim qtyVal As Double
        qtyVal = NzDbl(loCheck.DataBodyRange.Cells(r, cUsedChk).value)
        If systemKey = "" Or qtyVal <= 0 Then GoTo NextRow

        Dim delta As Object
        Set delta = CreateObject("Scripting.Dictionary")
        delta("System_Key") = systemKey
        delta("QTY") = qtyVal
        delta("ITEM_CODE") = itemCodeVal
        If cItemNameChk > 0 Then delta("ITEM_NAME") = NzStr(loCheck.DataBodyRange.Cells(r, cItemNameChk).value)
        If Not systemKeyIndex Is Nothing Then
            If systemKeyIndex.Exists(systemKey) Then
                Dim invIdx As Long
                invIdx = CLng(systemKeyIndex(systemKey))
                If cItemCode > 0 Then delta("ITEM_CODE") = NzStr(invLo.DataBodyRange.Cells(invIdx, cItemCode).value)
                If cItemName > 0 Then delta("ITEM_NAME") = NzStr(invLo.DataBodyRange.Cells(invIdx, cItemName).value)
            End If
        End If
        result.Add delta
NextRow:
    Next r

    If result.count = 0 Then
        If errNotes = "" Then errNotes = "No checked-in production input quantities were found."
        Exit Function
    End If
    Set BuildUsedDeltaPacketFromCheck = result
End Function

Private Function BuildUsedDictFromPayloadJson(ByVal payloadJson As String, ByRef errNotes As String) As Object
    Dim rx As Object
    Dim matches As Object
    Dim match As Object
    Dim objectText As String
    Dim systemKey As String
    Dim sku As String
    Dim locationValue As String
    Dim qtyVal As Double
    Dim ioType As String
    Dim key As String
    Dim dict As Object

    errNotes = ""
    payloadJson = Trim$(payloadJson)
    If payloadJson = "" Or payloadJson = "[]" Then
        errNotes = "No production input payload rows were generated."
        Exit Function
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "\{[^}]*\}"
    Set matches = rx.Execute(payloadJson)

    For Each match In matches
        objectText = CStr(match.value)
        ioType = UCase$(JsonPayloadStringField(objectText, "IoType"))
        If ioType <> "" And ioType <> "USED" Then GoTo NextMatch
        systemKey = Trim$(JsonPayloadStringField(objectText, "System_Key"))
        sku = Trim$(JsonPayloadStringField(objectText, "SKU"))
        If sku = "" Then sku = Trim$(JsonPayloadStringField(objectText, "ITEM_CODE"))
        locationValue = Trim$(JsonPayloadStringField(objectText, "Location"))
        qtyVal = JsonPayloadNumberField(objectText, "Qty")
        If systemKey = "" Or qtyVal <= 0 Then GoTo NextMatch
        key = "SYS|" & systemKey
        If dict.Exists(key) Then
            Dim existingAllocation As Object
            Set existingAllocation = dict(key)
            existingAllocation("QTY") = NzDbl(existingAllocation("QTY")) + qtyVal
        Else
            Dim allocation As Object
            Set allocation = CreateObject("Scripting.Dictionary")
            allocation.CompareMode = vbTextCompare
            allocation("System_Key") = systemKey
            allocation("SKU") = sku
            allocation("Location") = locationValue
            allocation("QTY") = qtyVal
            dict.Add key, allocation
        End If
NextMatch:
    Next match

    If dict.count = 0 Then
        errNotes = "No USED rows were found in the production input payload."
        Exit Function
    End If
    Set BuildUsedDictFromPayloadJson = dict
End Function

Private Function JsonPayloadStringField(ByVal objectText As String, ByVal fieldName As String) As String
    Dim rx As Object
    Dim matches As Object

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = """" & fieldName & """\s*:\s*""([^""]*)"""
    Set matches = rx.Execute(objectText)
    If matches.count > 0 Then JsonPayloadStringField = JsonPayloadUnescape(CStr(matches(0).SubMatches(0)))
End Function

Private Function JsonPayloadNumberField(ByVal objectText As String, ByVal fieldName As String) As Double
    Dim rx As Object
    Dim matches As Object
    Dim rawValue As String

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = """" & fieldName & """\s*:\s*(-?\d+(?:\.\d+)?)"
    Set matches = rx.Execute(objectText)
    If matches.count > 0 Then
        rawValue = CStr(matches(0).SubMatches(0))
        If IsNumeric(rawValue) Then JsonPayloadNumberField = CDbl(rawValue)
    End If
End Function

Private Function JsonPayloadUnescape(ByVal textIn As String) As String
    JsonPayloadUnescape = textIn
    JsonPayloadUnescape = Replace$(JsonPayloadUnescape, "\t", vbTab)
    JsonPayloadUnescape = Replace$(JsonPayloadUnescape, "\n", vbLf)
    JsonPayloadUnescape = Replace$(JsonPayloadUnescape, "\r", vbCr)
    JsonPayloadUnescape = Replace$(JsonPayloadUnescape, "\" & Chr$(34), Chr$(34))
    JsonPayloadUnescape = Replace$(JsonPayloadUnescape, "\\", "\")
End Function

Private Function BuildMadeDeltasFromProductionOutput(ByVal loOut As ListObject, ByVal invLo As ListObject, ByRef errNotes As String) As Collection
    errNotes = ""
    If loOut Is Nothing Or loOut.DataBodyRange Is Nothing Then Exit Function

    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cOutput As Long: cOutput = ColumnIndex(loOut, "OUTPUT")
    Dim cSystemKeyOut As Long: cSystemKeyOut = ColumnIndex(loOut, "System_Key")
    Dim cItemCodeOut As Long: cItemCodeOut = ColumnIndex(loOut, "ITEM_CODE")

    If cReal = 0 Then
        errNotes = "ProductionOutput missing REAL OUTPUT column."
        Exit Function
    End If
    If cSystemKeyOut = 0 Or (cItemCodeOut = 0 And cOutput = 0) Then
        errNotes = "ProductionOutput requires System_Key and ITEM_CODE/OUTPUT columns."
        Exit Function
    End If

    Dim agg As Object: Set agg = CreateObject("Scripting.Dictionary")
    agg.CompareMode = vbTextCompare
    Dim arr As Variant: arr = loOut.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim qtyVal As Double: qtyVal = NzDbl(arr(r, cReal))
        If qtyVal <= 0 Then GoTo NextRow

        Dim systemKey As String
        Dim itemCodeVal As String
        Dim outputName As String
        Dim resolvedItemName As String
        systemKey = Trim$(NzStr(arr(r, cSystemKeyOut)))
        If cItemCodeOut > 0 Then itemCodeVal = Trim$(NzStr(arr(r, cItemCodeOut)))
        If cOutput > 0 Then outputName = Trim$(NzStr(arr(r, cOutput)))
        ResolveProductionOutputIdentity outputName, systemKey, itemCodeVal, resolvedItemName, errNotes
        If itemCodeVal = "" Then
            AppendNote errNotes, "Production output '" & outputName & "' is missing ITEM_CODE/SKU."
            GoTo NextRow
        End If
        If systemKey = "" Then
            systemKey = modRoleEventWriter.CreateSystemKey()
        End If
        loOut.DataBodyRange.Cells(r, cSystemKeyOut).Value = systemKey
        If cItemCodeOut > 0 Then loOut.DataBodyRange.Cells(r, cItemCodeOut).Value = itemCodeVal

        Dim key As String: key = systemKey
        If agg.Exists(key) Then
            Dim existing As Object
            Set existing = agg(key)
            existing("QTY") = NzDbl(existing("QTY")) + qtyVal
        Else
            Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
            delta("System_Key") = systemKey
            delta("QTY") = qtyVal
            delta("ITEM_CODE") = itemCodeVal
            delta("ITEM_NAME") = IIf(resolvedItemName <> "", resolvedItemName, outputName)
            agg.Add key, delta
        End If
NextRow:
    Next r

    If agg.count = 0 Then
        If errNotes = "" Then errNotes = "No made quantities found in ProductionOutput."
        Exit Function
    End If

    Dim result As New Collection
    Dim k As Variant
    For Each k In agg.keys
        result.Add agg(k)
    Next k
    Set BuildMadeDeltasFromProductionOutput = result
End Function

Private Function BuildMadeDeltasFromProductionOutputRow(ByVal loOut As ListObject, ByVal invLo As ListObject, ByVal outputRowNumber As Long, ByRef errNotes As String) As Collection
    errNotes = ""
    If loOut Is Nothing Or loOut.DataBodyRange Is Nothing Then Exit Function
    If outputRowNumber < 1 Or outputRowNumber > loOut.ListRows.Count Then
        errNotes = "Selected ProductionOutput row is out of range."
        Exit Function
    End If

    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cOutput As Long: cOutput = ColumnIndex(loOut, "OUTPUT")
    Dim cSystemKeyOut As Long: cSystemKeyOut = ColumnIndex(loOut, "System_Key")
    Dim cItemCodeOut As Long: cItemCodeOut = ColumnIndex(loOut, "ITEM_CODE")

    If cReal = 0 Then
        errNotes = "ProductionOutput missing REAL OUTPUT column."
        Exit Function
    End If
    If cSystemKeyOut = 0 Or (cItemCodeOut = 0 And cOutput = 0) Then
        errNotes = "ProductionOutput requires System_Key and ITEM_CODE/OUTPUT columns."
        Exit Function
    End If

    Dim qtyVal As Double
    qtyVal = NzDbl(loOut.DataBodyRange.Cells(outputRowNumber, cReal).value)
    If qtyVal <= 0 Then
        errNotes = "No made quantity found for the selected Production Output row."
        Exit Function
    End If

    Dim outputName As String
    If cOutput > 0 Then outputName = NzStr(loOut.DataBodyRange.Cells(outputRowNumber, cOutput).value)

    Dim systemKey As String
    Dim itemCodeVal As String
    Dim resolvedItemName As String
    If cSystemKeyOut > 0 Then systemKey = Trim$(NzStr(loOut.DataBodyRange.Cells(outputRowNumber, cSystemKeyOut).value))
    If cItemCodeOut > 0 Then itemCodeVal = Trim$(NzStr(loOut.DataBodyRange.Cells(outputRowNumber, cItemCodeOut).Value))
    ResolveProductionOutputIdentity outputName, systemKey, itemCodeVal, resolvedItemName, errNotes
    If cItemCodeOut > 0 And itemCodeVal <> "" Then _
        loOut.DataBodyRange.Cells(outputRowNumber, cItemCodeOut).Value = itemCodeVal
    If itemCodeVal = "" Then
        If errNotes = "" Then
            errNotes = "Selected Production Output is missing ITEM_CODE/SKU: " & outputName
        Else
            AppendNote errNotes, "Selected Production Output is missing ITEM_CODE/SKU: " & outputName
        End If
        Exit Function
    End If
    If systemKey = "" Then
        systemKey = modRoleEventWriter.CreateSystemKey()
    End If
    loOut.DataBodyRange.Cells(outputRowNumber, cSystemKeyOut).Value = systemKey

    Dim delta As Object: Set delta = CreateObject("Scripting.Dictionary")
    delta("System_Key") = systemKey
    delta("QTY") = qtyVal
    delta("ITEM_CODE") = itemCodeVal
    delta("ITEM_NAME") = IIf(resolvedItemName <> "", resolvedItemName, outputName)

    Dim result As New Collection
    result.Add delta
    Set BuildMadeDeltasFromProductionOutputRow = result
End Function

Private Sub EnrichOutputDeltaFromPickerBySystemKey(ByVal delta As Object, _
                                                   ByVal pickerItems As Variant, _
                                                   ByVal systemKey As String)
    If delta Is Nothing Then Exit Sub
    systemKey = Trim$(systemKey)
    If systemKey = "" Then Exit Sub
    If IsEmpty(pickerItems) Then Exit Sub
    If Not IsArray(pickerItems) Then Exit Sub

    Dim r As Long
    For r = LBound(pickerItems, 1) To UBound(pickerItems, 1)
        If StrComp(Trim$(NzStr(pickerItems(r, 1))), systemKey, vbTextCompare) = 0 Then
            If NzStr(pickerItems(r, 7)) <> "" Then delta("ITEM_CODE") = NzStr(pickerItems(r, 7))
            If NzStr(pickerItems(r, 2)) <> "" Then delta("ITEM_NAME") = NzStr(pickerItems(r, 2))
            Exit Sub
        End If
    Next r
End Sub

Private Sub ClearUsedStageColumns(ByVal invLo As ListObject, ByVal deltas As Collection)
    Err.Raise vbObjectError + 2194, "mProduction.ClearUsedStageColumns", _
              "Retired by the v4.10 domain boundary. USED staging belongs in Prod_invSys_Check."
    If invLo Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub
    If deltas.Count = 0 Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim cUsed As Long: cUsed = ColumnIndex(invLo, "USED")
    If cUsed = 0 Then cUsed = ColumnIndexLoose(invLo, "USED")
    Dim cItemCode As Long: cItemCode = ColumnIndex(invLo, "ITEM_CODE")
    If cItemCode = 0 Then cItemCode = ColumnIndexLoose(invLo, "ITEM_CODE", "ITEMCODE", "ITEM CODE")
    If cUsed = 0 Then Exit Sub

    Dim rowIndex As Object
    Set rowIndex = BuildInvSysSystemKeyIndex(invLo)

    Dim delta As Variant
    For Each delta In deltas
        Dim rowKey As String: rowKey = CStr(delta("System_Key"))
        If Not rowIndex Is Nothing Then
            If rowIndex.Exists(rowKey) Then
                invLo.DataBodyRange.Cells(CLng(rowIndex(rowKey)), cUsed).Value = 0
            Else
                ClearUsedStageByItemCode invLo, cUsed, cItemCode, NzStr(delta("ITEM_CODE"))
            End If
        Else
            ClearUsedStageByItemCode invLo, cUsed, cItemCode, NzStr(delta("ITEM_CODE"))
        End If
    Next delta
End Sub

Private Sub ClearUsedStageByItemCode(ByVal invLo As ListObject, ByVal usedColumn As Long, ByVal itemCodeColumn As Long, ByVal itemCode As String)
    If invLo Is Nothing Then Exit Sub
    If usedColumn <= 0 Or itemCodeColumn <= 0 Then Exit Sub
    If Len(Trim$(itemCode)) = 0 Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If StrComp(NzStr(invLo.DataBodyRange.Cells(r, itemCodeColumn).Value), itemCode, vbTextCompare) = 0 Then
            invLo.DataBodyRange.Cells(r, usedColumn).Value = 0
            Exit Sub
        End If
    Next r
End Sub

Private Sub RestoreMadeStageColumns(ByVal invLo As ListObject, ByVal deltas As Collection)
    Err.Raise vbObjectError + 2195, "mProduction.RestoreMadeStageColumns", _
              "Retired by the v4.10 domain boundary. MADE staging must not mutate the invSys read model."
    If invLo Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub
    If deltas.Count = 0 Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim cMade As Long: cMade = ColumnIndex(invLo, "MADE")
    If cMade = 0 Then cMade = ColumnIndexLoose(invLo, "MADE")
    Dim cItemCode As Long: cItemCode = ColumnIndex(invLo, "ITEM_CODE")
    If cItemCode = 0 Then cItemCode = ColumnIndexLoose(invLo, "ITEM_CODE", "ITEMCODE", "ITEM CODE")
    If cMade = 0 Then Exit Sub

    Dim rowIndex As Object
    Set rowIndex = BuildInvSysSystemKeyIndex(invLo)

    Dim delta As Variant
    For Each delta In deltas
        Dim rowKey As String: rowKey = CStr(delta("System_Key"))
        Dim qtyVal As Double: qtyVal = NzDbl(delta("QTY"))
        If qtyVal <= 0 Then GoTo NextDelta
        If Not rowIndex Is Nothing Then
            If rowIndex.Exists(rowKey) Then
                invLo.DataBodyRange.Cells(CLng(rowIndex(rowKey)), cMade).Value = qtyVal
            Else
                RestoreMadeStageByItemCode invLo, cMade, cItemCode, NzStr(delta("ITEM_CODE")), qtyVal
            End If
        Else
            RestoreMadeStageByItemCode invLo, cMade, cItemCode, NzStr(delta("ITEM_CODE")), qtyVal
        End If
NextDelta:
    Next delta
End Sub

Private Sub RestoreMadeStageByItemCode(ByVal invLo As ListObject, ByVal madeColumn As Long, ByVal itemCodeColumn As Long, ByVal itemCode As String, ByVal qtyVal As Double)
    If invLo Is Nothing Then Exit Sub
    If madeColumn <= 0 Or itemCodeColumn <= 0 Then Exit Sub
    If Len(Trim$(itemCode)) = 0 Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim r As Long
    For r = 1 To invLo.DataBodyRange.Rows.Count
        If StrComp(NzStr(invLo.DataBodyRange.Cells(r, itemCodeColumn).Value), itemCode, vbTextCompare) = 0 Then
            invLo.DataBodyRange.Cells(r, madeColumn).Value = qtyVal
            Exit Sub
        End If
    Next r
End Sub

Private Sub ClearMadeStageColumns(ByVal invLo As ListObject, ByVal deltas As Collection)
    Err.Raise vbObjectError + 2196, "mProduction.ClearMadeStageColumns", _
              "Retired by the v4.10 domain boundary. MADE staging must not mutate the invSys read model."
    If invLo Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub
    If deltas.Count = 0 Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub

    Dim cMade As Long: cMade = ColumnIndex(invLo, "MADE")
    If cMade = 0 Then cMade = ColumnIndexLoose(invLo, "MADE")
    Dim cItemCode As Long: cItemCode = ColumnIndex(invLo, "ITEM_CODE")
    If cItemCode = 0 Then cItemCode = ColumnIndexLoose(invLo, "ITEM_CODE", "ITEMCODE", "ITEM CODE")
    If cMade = 0 Then Exit Sub

    Dim rowIndex As Object
    Set rowIndex = BuildInvSysSystemKeyIndex(invLo)

    Dim delta As Variant
    For Each delta In deltas
        Dim rowKey As String: rowKey = CStr(delta("System_Key"))
        If Not rowIndex Is Nothing And rowIndex.Exists(rowKey) Then
            invLo.DataBodyRange.Cells(CLng(rowIndex(rowKey)), cMade).Value = 0
        ElseIf cItemCode > 0 Then
            ClearMadeStageByItemCode invLo, cMade, cItemCode, NzStr(delta("ITEM_CODE"))
        End If
    Next delta
End Sub

Private Sub ClearMadeStageByItemCode(ByVal invLo As ListObject, ByVal madeColumn As Long, ByVal itemCodeColumn As Long, ByVal itemCode As String)
    If invLo Is Nothing Then Exit Sub
    If invLo.DataBodyRange Is Nothing Then Exit Sub
    itemCode = Trim$(itemCode)
    If itemCode = "" Then Exit Sub

    Dim r As Long
    For r = 1 To invLo.ListRows.Count
        If StrComp(Trim$(NzStr(invLo.DataBodyRange.Cells(r, itemCodeColumn).Value)), itemCode, vbTextCompare) = 0 Then
            invLo.DataBodyRange.Cells(r, madeColumn).Value = 0
        End If
    Next r
End Sub

Private Function QueueProductionConsumeEvent(ByVal usedDeltas As Collection, ByVal madeDeltas As Collection, ByRef errNotes As String, ByRef eventIdOut As String) As Boolean
    Dim payloadItems As Collection
    Dim payloadJson As String

    Set payloadItems = New Collection
    AddPayloadItemsFromDeltas payloadItems, usedDeltas, "USED"
    AddPayloadItemsFromDeltas payloadItems, madeDeltas, "MADE"
    If payloadItems.Count = 0 Then
        If errNotes = "" Then errNotes = "No production consume payload rows were generated."
        Exit Function
    End If

    payloadJson = modProductionJson.BuildJsonArray(payloadItems)
    QueueProductionConsumeEvent = modRoleEventWriter.QueuePayloadEventCurrent( _
        EVENT_TYPE_PROD_CONSUME, _
        "", _
        payloadJson, _
        "BTN_TO_MADE", _
        eventIdOut, _
        errNotes)
End Function

Private Function QueueProductionCompleteEvent(ByVal madeDeltas As Collection, ByRef errNotes As String, ByRef eventIdOut As String) As Boolean
    Dim payloadItems As Collection
    Dim payloadJson As String

    Set payloadItems = New Collection
    AddPayloadItemsFromDeltas payloadItems, madeDeltas, "MADE"
    If payloadItems.Count = 0 Then
        If errNotes = "" Then errNotes = "No production completion payload rows were generated."
        Exit Function
    End If

    payloadJson = modProductionJson.BuildJsonArray(payloadItems)
    QueueProductionCompleteEvent = modRoleEventWriter.QueuePayloadEventCurrent( _
        EVENT_TYPE_PROD_COMPLETE, _
        "", _
        payloadJson, _
        "BTN_TO_TOTALINV", _
        eventIdOut, _
        errNotes)
End Function

Private Sub AddPayloadItemsFromDeltas(ByVal payloadItems As Collection, ByVal deltas As Collection, ByVal ioType As String)
    Dim delta As Variant
    Dim payloadItem As Object

    If payloadItems Is Nothing Then Exit Sub
    If deltas Is Nothing Then Exit Sub

    For Each delta In deltas
        Set payloadItem = modProductionJson.CreateProductionDeltaPayloadItem( _
            CStr(delta("System_Key")), _
            NzStr(delta("ITEM_CODE")), _
            NzDbl(delta("QTY")), _
            "", _
            NzStr(delta("ITEM_NAME")), _
            ioType)
        payloadItem("ITEM") = NzStr(delta("ITEM_NAME"))
        payloadItems.Add payloadItem
    Next delta
End Sub

Private Function BuildSystemKeySetFromDeltas(ByVal usedDeltas As Collection, ByVal madeDeltas As Collection) As Object
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim delta As Variant

    If Not usedDeltas Is Nothing Then
        For Each delta In usedDeltas
            On Error Resume Next
            If Trim$(CStr(delta("System_Key"))) <> "" Then dict(CStr(delta("System_Key"))) = True
            On Error GoTo 0
        Next delta
    End If

    If Not madeDeltas Is Nothing Then
        For Each delta In madeDeltas
            On Error Resume Next
            If Trim$(CStr(delta("System_Key"))) <> "" Then dict(CStr(delta("System_Key"))) = True
            On Error GoTo 0
        Next delta
    End If

    If dict.count = 0 Then Exit Function
    Set BuildSystemKeySetFromDeltas = dict
End Function

Private Function BuildUsedSnapshotForSystemKeys(ByVal invLo As ListObject, ByVal rowKeys As Object) As Object
    If invLo Is Nothing Or invLo.DataBodyRange Is Nothing Then Exit Function
    If rowKeys Is Nothing Then Exit Function
    If rowKeys.count = 0 Then Exit Function

    Dim cRow As Long: cRow = ColumnIndex(invLo, "System_Key")
    Dim cUsed As Long: cUsed = ColumnIndex(invLo, "USED")
    If cRow = 0 Or cUsed = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = invLo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        Dim rowVal As String: rowVal = NzStr(arr(r, cRow))
        If rowVal <> "" Then
            If rowKeys.Exists(rowVal) Then dict(rowVal) = NzDbl(arr(r, cUsed))
        End If
    Next r

    If dict.count = 0 Then Exit Function
    Set BuildUsedSnapshotForSystemKeys = dict
End Function

Private Sub WriteArrayToTable(lo As ListObject, arr As Variant)
    If lo Is Nothing Then Exit Sub
    If IsEmpty(arr) Then Exit Sub
    Dim rowsNeeded As Long
    On Error Resume Next
    rowsNeeded = UBound(arr, 1)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    If rowsNeeded <= 0 Then
        ClearListObjectData lo
        Exit Sub
    End If
    Dim currentRows As Long
    If lo.DataBodyRange Is Nothing Then
        currentRows = 0
    Else
        currentRows = lo.DataBodyRange.rows.count
    End If
    Dim diff As Long
    If currentRows < rowsNeeded Then
        For diff = 1 To rowsNeeded - currentRows
            lo.ListRows.Add
        Next diff
    ElseIf currentRows > rowsNeeded Then
        For diff = rowsNeeded + 1 To currentRows
            lo.ListRows(diff).Range.ClearContents
        Next diff
    End If
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.value = arr
End Sub

Private Sub ClearListObjectContents(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    EnsureTableHasRow lo
    If lo.DataBodyRange Is Nothing Then Exit Sub
    lo.DataBodyRange.ClearContents
End Sub

Private Function SortedKeys(dict As Object) As Variant
    If dict Is Nothing Then Exit Function
    Dim keys As Variant: keys = dict.keys
    If Not IsArray(keys) Then
        SortedKeys = keys
        Exit Function
    End If
    Dim i As Long, j As Long
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CLng(val(keys(j))) < CLng(val(keys(i))) Then
                Dim tmp As Variant
                tmp = keys(i)
                keys(i) = keys(j)
                keys(j) = tmp
            End If
        Next j
    Next i
    SortedKeys = keys
End Function

Private Sub AppendNote(ByRef target As String, ByVal text As String)
    If Len(text) = 0 Then Exit Sub
    If Len(target) > 0 Then
        target = target & vbCrLf & text
    Else
        target = text
    End If
End Sub

Private Sub ApplyProcessHeaderColor(ByVal lo As ListObject, ByVal procName As String)
    If lo Is Nothing Then Exit Sub
    If lo.HeaderRowRange Is Nothing Then Exit Sub
    procName = Trim$(procName)
    If procName = "" Then Exit Sub

    Dim colorVal As Long
    colorVal = ProcessColorFromName(procName)
    On Error Resume Next
    lo.HeaderRowRange.Interior.Color = colorVal
    If IsColorDark(colorVal) Then
        lo.HeaderRowRange.Font.Color = vbWhite
    Else
        lo.HeaderRowRange.Font.Color = vbBlack
    End If
    On Error GoTo 0
End Sub

Private Function ProcessColorFromName(ByVal procName As String) As Long
    Static colorMap As Object
    Static usedMap As Object
    If colorMap Is Nothing Then Set colorMap = CreateObject("Scripting.Dictionary")
    If usedMap Is Nothing Then Set usedMap = CreateObject("Scripting.Dictionary")

    Dim key As String
    key = LCase$(Trim$(procName))
    If key = "" Then Exit Function
    If colorMap.Exists(key) Then
        ProcessColorFromName = colorMap(key)
        Exit Function
    End If

    Dim palette As Variant
    palette = ProcessColorPalette()
    Dim n As Long
    n = UBound(palette) - LBound(palette) + 1

    Dim startIdx As Long
    startIdx = HashProcessName(key) Mod n
    If startIdx < 0 Then startIdx = startIdx + n

    Dim idx As Long
    idx = startIdx
    Dim attempts As Long
    Do
        Dim c As Long
        c = palette(idx)
        If Not usedMap.Exists(CStr(c)) Then
            usedMap.Add CStr(c), True
            colorMap.Add key, c
            ProcessColorFromName = c
            Exit Function
        End If
        idx = idx + 1
        If idx >= n Then idx = 0
        attempts = attempts + 1
    Loop While attempts < n

    colorMap.Add key, palette(startIdx)
    ProcessColorFromName = palette(startIdx)
End Function

Private Function ProcessColorPalette() As Variant
    ProcessColorPalette = Array( _
        RGB(33, 150, 243), _
        RGB(233, 30, 99), _
        RGB(0, 150, 136), _
        RGB(255, 152, 0), _
        RGB(156, 39, 176), _
        RGB(76, 175, 80), _
        RGB(121, 85, 72), _
        RGB(63, 81, 181), _
        RGB(205, 220, 57), _
        RGB(0, 188, 212), _
        RGB(244, 67, 54), _
        RGB(255, 193, 7))
End Function

Private Function HashProcessName(ByVal procName As String) As Long
    Dim h As Double
    Dim i As Long
    For i = 1 To Len(procName)
        Dim ch As Long
        ch = AscW(Mid$(procName, i, 1))
        If ch < 0 Then ch = ch + 65536
        h = (h * 131#) + (ch * i)
        If h >= 2147483647# Then
            h = h - 2147483647# * Fix(h / 2147483647#)
        End If
    Next i
    HashProcessName = CLng(h)
End Function

Private Function HsvToRgb(ByVal h As Double, ByVal s As Double, ByVal v As Double) As Long
    Dim r As Double, g As Double, b As Double
    Dim i As Long
    Dim f As Double, p As Double, q As Double, t As Double

    i = Int(h * 6)
    f = h * 6 - i
    p = v * (1 - s)
    q = v * (1 - f * s)
    t = v * (1 - (1 - f) * s)

    Select Case (i Mod 6)
        Case 0
            r = v: g = t: b = p
        Case 1
            r = q: g = v: b = p
        Case 2
            r = p: g = v: b = t
        Case 3
            r = p: g = q: b = v
        Case 4
            r = t: g = p: b = v
        Case 5
            r = v: g = p: b = q
    End Select

    HsvToRgb = RGB(CLng(r * 255), CLng(g * 255), CLng(b * 255))
End Function

Private Function IsColorDark(ByVal colorVal As Long) As Boolean
    Dim r As Long, g As Long, b As Long
    r = colorVal Mod 256
    g = (colorVal \ 256) Mod 256
    b = (colorVal \ 65536) Mod 256

    Dim luma As Double
    luma = (0.299 * r) + (0.587 * g) + (0.114 * b)
    IsColorDark = (luma < 140)
End Function

Private Sub RenderProcessSelectorCheckboxes(ByVal ws As Worksheet, ByVal procTables As Collection)
    If ws Is Nothing Then Exit Sub
    If procTables Is Nothing Then Exit Sub

    Dim prevStates As Object
    Set prevStates = CreateObject("Scripting.Dictionary")
    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(CHK_PROC_PREFIX) & "*" Then
                Dim cap As String
                cap = LCase$(Trim$(GetCheckboxCaption(shp)))
                If cap <> "" Then
                    prevStates(cap) = (shp.ControlFormat.value = 1)
                End If
            End If
        End If
    Next shp

    DeleteCheckboxesByPrefix ws, CHK_PROC_PREFIX

    If procTables.count = 0 Then Exit Sub

    Dim maxCol As Long
    Dim lo As ListObject
    For Each lo In procTables
        If Not lo Is Nothing Then
            Dim endCol As Long
            endCol = lo.Range.Column + lo.Range.Columns.count - 1
            If endCol > maxCol Then maxCol = endCol
        End If
    Next lo
    If maxCol = 0 Then Exit Sub

    Dim leftPos As Double
    leftPos = ws.Columns(maxCol + 1).Left + 2
    Const CHK_HEIGHT As Double = 16
    Const CHK_WIDTH As Double = 140

    For Each lo In procTables
        If lo Is Nothing Then GoTo NextProc
        Dim procName As String
        procName = ProcessNameFromTable(lo)
        If Trim$(procName) = "" Then procName = lo.Name

        Dim topPos As Double
        topPos = lo.HeaderRowRange.Top + 2

        Dim baseName As String
        baseName = CHK_PROC_PREFIX & SafeProcessKey(procName)
        Dim shapeName As String
        shapeName = UniqueShapeName(ws, baseName)

        Dim chk As Shape
        Set chk = EnsureCheckboxShape(ws, shapeName, procName, "mProduction.ProcessCheckboxChanged", leftPos, topPos, CHK_WIDTH, CHK_HEIGHT)
        If Not chk Is Nothing Then
            chk.AlternativeText = procName
            Dim key As String
            key = LCase$(Trim$(procName))
            If prevStates.Exists(key) Then
                chk.ControlFormat.value = IIf(prevStates(key), 1, 0)
            Else
                chk.ControlFormat.value = 1
            End If
        End If
NextProc:
    Next lo
End Sub

Private Sub RenderPaletteKeepCheckboxes(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    Dim prevStates As Object
    Set prevStates = CreateObject("Scripting.Dictionary")
    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(CHK_BATCH_PREFIX) & "*" Then
                Dim cap As String
                cap = LCase$(Trim$(GetCheckboxCaption(shp)))
                If cap = "" Then cap = LCase$(Trim$(shp.AlternativeText))
                If cap <> "" Then
                    prevStates(cap) = (shp.ControlFormat.value = 1)
                End If
            End If
        End If
    Next shp

    DeleteCheckboxesByPrefix ws, CHK_BATCH_PREFIX

    Dim maxCol As Long
    Dim lo As ListObject
    Dim paletteTables As New Collection
    Dim firstPaletteByProc As Object
    Set firstPaletteByProc = CreateObject("Scripting.Dictionary")
    For Each lo In ws.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
                paletteTables.Add lo
                Dim endCol As Long
                endCol = lo.Range.Column + lo.Range.Columns.count - 1
                If endCol > maxCol Then maxCol = endCol

                Dim procNameCollect As String
                Dim recipeIdCollect As String
                Dim ingIdCollect As String
                Dim amtValCollect As Variant
                Dim ioValCollect As String
                If GetPaletteTableContext(lo, recipeIdCollect, ingIdCollect, amtValCollect, procNameCollect, ioValCollect) = False Then
                    procNameCollect = ProcessNameFromTable(lo)
                End If
                procNameCollect = Trim$(procNameCollect)
                If procNameCollect <> "" Then
                    Dim procKeyCollect As String
                    procKeyCollect = NormalizeProcessBandKey(procNameCollect)
                    If Not firstPaletteByProc.Exists(procKeyCollect) Then
                        firstPaletteByProc.Add procKeyCollect, lo
                    ElseIf firstPaletteByProc(procKeyCollect).Range.Row > lo.Range.Row Then
                        Set firstPaletteByProc(procKeyCollect) = lo
                    End If
                End If
            End If
        End If
    Next lo
    If paletteTables.count = 0 Then Exit Sub
    If maxCol = 0 Then Exit Sub

    Dim leftPos As Double
    leftPos = ws.Columns(maxCol + 1).Left + 2
    Const CHK_HEIGHT As Double = 14
    Const CHK_WIDTH As Double = 14

    Dim procKey As Variant
    For Each procKey In firstPaletteByProc.keys
        Set lo = firstPaletteByProc(procKey)
        If lo Is Nothing Then GoTo NextPal
        Dim procName As String
        Dim recipeId As String
        Dim ingId As String
        Dim amtVal As Variant
        Dim ioVal As String
        If GetPaletteTableContext(lo, recipeId, ingId, amtVal, procName, ioVal) = False Then
            procName = ProcessNameFromTable(lo)
        End If
        If Trim$(procName) = "" Then procName = lo.Name

        Dim topPos As Double
        topPos = lo.HeaderRowRange.Top + 2

        Dim shapeName As String
        shapeName = CHK_BATCH_PREFIX & SafeProcessKey(procName)

        Dim chk As Shape
        Set chk = EnsureCheckboxShape(ws, shapeName, "", "mProduction.OutputCheckboxChanged", leftPos, topPos, CHK_WIDTH, CHK_HEIGHT)
        If Not chk Is Nothing Then
            chk.AlternativeText = procName
            Dim key As String
            key = LCase$(Trim$(procName))
            If prevStates.Exists(key) Then
                chk.ControlFormat.value = IIf(prevStates(key), 1, 0)
            Else
                chk.ControlFormat.value = 0
            End If
        End If
NextPal:
    Next procKey
End Sub

Private Function IsPaletteKeepSelected(ByVal ws As Worksheet, ByVal procName As String) As Boolean
    If ws Is Nothing Then Exit Function
    procName = Trim$(procName)
    If procName = "" Then Exit Function

    Dim shapeName As String
    shapeName = CHK_BATCH_PREFIX & SafeProcessKey(procName)

    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.shapes(shapeName)
    On Error GoTo 0
    If shp Is Nothing Then Exit Function
    If Not IsCheckboxShape(shp) Then Exit Function

    On Error Resume Next
    IsPaletteKeepSelected = (shp.ControlFormat.value = 1)
    On Error GoTo 0
End Function

Private Sub ClearPaletteTableSelection(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cCode As Long: cCode = ColumnIndex(lo, "ITEM_CODE")
    Dim cVend As Long: cVend = ColumnIndex(lo, "VENDORS")
    Dim cVendCode As Long: cVendCode = ColumnIndex(lo, "VENDOR_CODE")
    Dim cDesc As Long: cDesc = ColumnIndex(lo, "DESCRIPTION")
    Dim cItem As Long: cItem = ColumnIndex(lo, "ITEM")
    Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
    Dim cLoc As Long: cLoc = ColumnIndex(lo, "LOCATION")
    Dim cRow As Long: cRow = ColumnIndex(lo, "System_Key")

    Dim r As Long
    For r = 1 To lo.DataBodyRange.rows.count
        If cCode > 0 Then lo.DataBodyRange.Cells(r, cCode).ClearContents
        If cVend > 0 Then lo.DataBodyRange.Cells(r, cVend).ClearContents
        If cVendCode > 0 Then lo.DataBodyRange.Cells(r, cVendCode).ClearContents
        If cDesc > 0 Then lo.DataBodyRange.Cells(r, cDesc).ClearContents
        If cItem > 0 Then lo.DataBodyRange.Cells(r, cItem).ClearContents
        If cUom > 0 Then lo.DataBodyRange.Cells(r, cUom).ClearContents
        If cLoc > 0 Then lo.DataBodyRange.Cells(r, cLoc).ClearContents
        If cRow > 0 Then lo.DataBodyRange.Cells(r, cRow).ClearContents
    Next r
End Sub

Private Sub EnsurePaletteTableMetaForExistingTables(ByVal wsProd As Worksheet)
    If wsProd Is Nothing Then Exit Sub
    EnsurePaletteTableMeta

    Dim palTables As Collection
    Set palTables = GetPaletteTablesInOrder(wsProd)
    If palTables Is Nothing Then Exit Sub
    If palTables.count = 0 Then Exit Sub

    Dim needsRebuild As Boolean
    Dim lo As ListObject
    If mPaletteTableMeta Is Nothing Then
        needsRebuild = True
    Else
        For Each lo In palTables
            If Not mPaletteTableMeta.Exists(lo.Name) Then
                needsRebuild = True
                Exit For
            End If
        Next lo
    End If

    If Not needsRebuild Then Exit Sub

    Dim entries As Collection
    Set entries = BuildPaletteMetaEntries(wsProd)
    If entries Is Nothing Then Exit Sub
    If entries.count = 0 Then Exit Sub

    ClearPaletteTableMeta
    Dim used() As Boolean
    ReDim used(1 To entries.count)

    For Each lo In palTables
        Dim procName As String
        procName = ProcessNameFromTable(lo)
        If Trim$(procName) = "" Then procName = lo.Name

        Dim matchIdx As Long
        matchIdx = FindPaletteEntryIndex(entries, used, procName)
        If matchIdx = 0 Then
            matchIdx = FindFirstUnusedEntryIndex(used)
        End If

        If matchIdx > 0 Then
            mPaletteTableMeta(lo.Name) = entries(matchIdx)
            used(matchIdx) = True
        End If
    Next lo
End Sub

Private Function GetPaletteTablesInOrder(ByVal wsProd As Worksheet) As Collection
    Dim result As New Collection
    If wsProd Is Nothing Then
        Set GetPaletteTablesInOrder = result
        Exit Function
    End If

    Dim countPal As Long
    Dim lo As ListObject
    For Each lo In wsProd.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
                countPal = countPal + 1
            End If
        End If
    Next lo
    If countPal = 0 Then
        Set GetPaletteTablesInOrder = result
        Exit Function
    End If

    Dim arrLo() As ListObject
    Dim arrRow() As Long
    ReDim arrLo(1 To countPal)
    ReDim arrRow(1 To countPal)

    Dim i As Long
    i = 0
    For Each lo In wsProd.ListObjects
        If IsPaletteTable(lo) Then
            If lo.Range.row < PALETTE_LINES_STAGING_ROW Then
                i = i + 1
                Set arrLo(i) = lo
                arrRow(i) = lo.Range.row
            End If
        End If
    Next lo

    Dim j As Long, k As Long
    For j = 1 To countPal - 1
        For k = j + 1 To countPal
            If arrRow(k) < arrRow(j) Then
                Dim tmpRow As Long
                Dim tmpLo As ListObject
                tmpRow = arrRow(j)
                arrRow(j) = arrRow(k)
                arrRow(k) = tmpRow
                Set tmpLo = arrLo(j)
                Set arrLo(j) = arrLo(k)
                Set arrLo(k) = tmpLo
            End If
        Next k
    Next j

    For i = 1 To countPal
        result.Add arrLo(i)
    Next i

    Set GetPaletteTablesInOrder = result
End Function

Private Function BuildPaletteMetaEntries(ByVal wsProd As Worksheet) As Collection
    Dim result As New Collection
    If wsProd Is Nothing Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If

    Dim recipeId As String
    recipeId = GetRecipeChooserRecipeId(wsProd)
    If Trim$(recipeId) = "" Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If

    Dim wsRec As Worksheet
    Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If

    Dim loRecipes As ListObject: Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If
    If loRecipes.DataBodyRange Is Nothing Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If

    Dim cRecId As Long: cRecId = ColumnIndex(loRecipes, "RECIPE_ID")
    Dim cProc As Long: cProc = ColumnIndex(loRecipes, "PROCESS")
    Dim cIO As Long: cIO = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    Dim cIngId As Long: cIngId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    Dim cAmt As Long: cAmt = ColumnIndex(loRecipes, "AMOUNT")
    Dim cPct As Long: cPct = ColumnIndex(loRecipes, "PERCENT")
    Dim cUom As Long: cUom = ColumnIndex(loRecipes, "UOM")
    If cRecId = 0 Or cProc = 0 Or cIO = 0 Or cIngId = 0 Then
        Set BuildPaletteMetaEntries = result
        Exit Function
    End If

    Dim arr As Variant: arr = loRecipes.DataBodyRange.value
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId Then
            Dim ioVal As String: ioVal = UCase$(Trim$(NzStr(arr(r, cIO))))
            If IsInputIoValue(ioVal) Then
                Dim ingId As String: ingId = NzStr(arr(r, cIngId))
                Dim procName As String: procName = NzStr(arr(r, cProc))
                If ingId <> "" And procName <> "" Then
                    If Not IsProcessSelected(procName, wsProd) Then GoTo NextRow
                    Dim key As String: key = procName & "|" & ingId
                    Dim amtVal As Variant
                    Dim pctVal As Variant
                    Dim uomVal As String
                    If cAmt > 0 Then amtVal = arr(r, cAmt)
                    If cPct > 0 Then pctVal = arr(r, cPct)
                    If cUom > 0 Then uomVal = NzStr(arr(r, cUom))
                    If Not seen.Exists(key) Then
                        Dim info(0 To 6) As Variant
                        info(0) = recipeId
                        info(1) = ingId
                        info(2) = amtVal
                        info(3) = procName
                        info(4) = "USED"
                        info(5) = pctVal
                        info(6) = uomVal
                        seen.Add key, info
                        result.Add info
                    Else
                        If IsNumeric(amtVal) Then
                            Dim curInfo As Variant
                            curInfo = seen(key)
                            If IsNumeric(curInfo(2)) Then
                                curInfo(2) = CDbl(curInfo(2)) + CDbl(amtVal)
                                seen(key) = curInfo
                            End If
                        End If
                    End If
                End If
            End If
        End If
NextRow:
    Next r

    Set BuildPaletteMetaEntries = result
End Function

Private Function FindPaletteEntryIndex(ByVal entries As Collection, ByRef used() As Boolean, ByVal procName As String) As Long
    If entries Is Nothing Then Exit Function
    If procName = "" Then Exit Function

    Dim i As Long
    For i = 1 To entries.count
        If Not used(i) Then
            Dim info As Variant
            info = entries(i)
            If StrComp(NzStr(info(3)), procName, vbTextCompare) = 0 Then
                FindPaletteEntryIndex = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function FindFirstUnusedEntryIndex(ByRef used() As Boolean) As Long
    Dim i As Long
    For i = LBound(used) To UBound(used)
        If Not used(i) Then
            FindFirstUnusedEntryIndex = i
            Exit Function
        End If
    Next i
End Function

Private Function IsProcessSelected(ByVal procName As String, ByVal ws As Worksheet) As Boolean
    If ws Is Nothing Then
        IsProcessSelected = True
        Exit Function
    End If
    procName = Trim$(procName)
    If procName = "" Then
        IsProcessSelected = True
        Exit Function
    End If

    Dim hasAny As Boolean
    Dim hasChecked As Boolean
    Dim hasMatch As Boolean
    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(CHK_PROC_PREFIX) & "*" Then
                hasAny = True
                If shp.ControlFormat.value = 1 Then hasChecked = True
                Dim cap As String
                cap = Trim$(GetCheckboxCaption(shp))
                If cap = "" Then cap = Trim$(shp.AlternativeText)
                If StrComp(cap, procName, vbTextCompare) = 0 Then
                    hasMatch = True
                    If shp.ControlFormat.value = 1 Then
                        IsProcessSelected = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next shp

    If Not hasAny Then
        IsProcessSelected = True
    ElseIf hasMatch Then
        IsProcessSelected = False
    ElseIf hasChecked Then
        IsProcessSelected = False
    Else
        IsProcessSelected = False
    End If
End Function

Public Sub ProcessCheckboxChanged()
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Sub

    Dim recipeId As String
    recipeId = GetRecipeChooserRecipeId(ws)
    If Trim$(recipeId) = "" Then Exit Sub

    Dim wsRec As Worksheet
    Set wsRec = SheetExists("Recipes")
    If wsRec Is Nothing Then Exit Sub

    DeleteInventoryPaletteTables ws

    Dim procTables As Collection
    Set procTables = GetRecipeChooserProcessTables(ws)

    BuildPaletteTablesForRecipeChooser recipeId, ws, wsRec, procTables, ""
    RenderPaletteKeepCheckboxes ws
    Exit Sub
ErrHandler:
    MsgBox "Process checkbox update failed: " & Err.description, vbExclamation
End Sub

Private Function GetRecipeChooserRecipeId(ByVal ws As Worksheet) As String
    If ws Is Nothing Then Set ws = SheetExists(SHEET_PRODUCTION)
    If ws Is Nothing Then Exit Function
    Dim lo As ListObject
    Set lo = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If lo Is Nothing Then Exit Function
    GetRecipeChooserRecipeId = NormalizeIdFirst(FirstNonEmptyColumnValue(lo, "RECIPE_ID"))
End Function

Private Function GetRecipeChooserProcessTables(ByVal ws As Worksheet) As Collection
    Dim result As New Collection
    If ws Is Nothing Then
        Set GetRecipeChooserProcessTables = result
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsRecipeChooserProcessTable(lo) Or LCase$(lo.Name) = LCase$(TABLE_RECIPE_CHOOSER_GENERATED) Then
            result.Add lo
        End If
    Next lo
    Set GetRecipeChooserProcessTables = result
End Function

Private Sub RenderOutputRowCheckboxes(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(ws, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then Exit Sub
    If loOut.DataBodyRange Is Nothing Then Exit Sub

    Dim loCheck As ListObject
    Set loCheck = FindListObjectByNameOrHeaders(ws, "Prod_invSys_Check", Array("USED", "TOTAL INV"))

    Dim prevRecall As Object
    Set prevRecall = CreateObject("Scripting.Dictionary")

    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            Dim nm As String
            nm = LCase$(shp.Name)
            If nm Like LCase$(CHK_RECALL_PREFIX) & "*" Then
                Dim idxR As Long
                idxR = ParseCheckboxIndex(shp.Name, CHK_RECALL_PREFIX)
                If idxR > 0 Then prevRecall(idxR) = (shp.ControlFormat.value = 1)
            End If
        End If
    Next shp

    DeleteCheckboxesByPrefix ws, CHK_RECALL_PREFIX

    Dim rightCol As Long
    rightCol = loOut.Range.Column + loOut.Range.Columns.count - 1
    Dim baseCol As Long
    baseCol = rightCol + 1
    Dim gapCols As Long
    If Not loCheck Is Nothing Then
        gapCols = loCheck.Range.Column - rightCol - 1
    Else
        gapCols = 2
    End If
    If gapCols < 1 Then gapCols = 1

    Dim leftRecall As Double
    Dim chkWidth As Double
    leftRecall = ws.Columns(baseCol).Left + 2
    chkWidth = ws.Columns(baseCol).Width - 4
    If chkWidth < 12 Then chkWidth = 12

    Dim r As Long
    For r = 1 To loOut.DataBodyRange.rows.count
        Dim topPos As Double
        Dim heightPts As Double
        topPos = loOut.DataBodyRange.rows(r).Top + 1
        heightPts = loOut.DataBodyRange.rows(r).Height - 2
        If heightPts < 12 Then heightPts = 12

        Dim shpRecall As Shape
        Set shpRecall = EnsureCheckboxShape(ws, CHK_RECALL_PREFIX & CStr(r), "", "mProduction.OutputCheckboxChanged", leftRecall, topPos, chkWidth, heightPts)
        If Not shpRecall Is Nothing Then
            shpRecall.AlternativeText = CStr(r)
            If prevRecall.Exists(r) Then shpRecall.ControlFormat.value = IIf(prevRecall(r), 1, 0)
        End If
    Next r
End Sub

Private Sub ClearProductionOutputForNextBatch(ByVal ws As Worksheet, ByVal loOut As ListObject)
    If ws Is Nothing Then Exit Sub
    If loOut Is Nothing Then Exit Sub
    If loOut.DataBodyRange Is Nothing Then Exit Sub

    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cRecall As Long: cRecall = ColumnIndex(loOut, "RECALL CODE")
    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")

    Dim nextBatchMap As Object
    Dim maxBatchMap As Object
    If cBatch > 0 And cProc > 0 Then
        Set nextBatchMap = CreateObject("Scripting.Dictionary")
        Set maxBatchMap = CreateObject("Scripting.Dictionary")

        Dim batchVal As String
        Dim procName As String
        Dim key As String
        Dim curBatch As Long
        Dim r As Long
        For r = 1 To loOut.DataBodyRange.rows.count
            procName = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
            If procName <> "" Then
                batchVal = NzStr(loOut.DataBodyRange.Cells(r, cBatch).value)
                If IsNumeric(batchVal) Then
                    curBatch = CLng(val(batchVal))
                    key = LCase$(procName)
                    If Not maxBatchMap.Exists(key) Then
                        maxBatchMap.Add key, curBatch
                    ElseIf curBatch > CLng(maxBatchMap(key)) Then
                        maxBatchMap(key) = curBatch
                    End If
                End If
            End If
        Next r

        For r = 1 To loOut.DataBodyRange.rows.count
            procName = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
            If procName <> "" Then
                key = LCase$(procName)
                If Not nextBatchMap.Exists(key) Then
                    Dim nextBatch As Long
                    If Not maxBatchMap Is Nothing Then
                        If maxBatchMap.Exists(key) Then nextBatch = CLng(maxBatchMap(key)) + 1
                    End If
                    If nextBatch = 0 Then
                        nextBatch = NextBatchSequenceForProcess(ws, loOut, procName)
                    End If
                    If nextBatch > 0 Then nextBatchMap.Add key, nextBatch
                End If
            End If
        Next r
    End If

    For r = 1 To loOut.DataBodyRange.rows.count
        If cReal > 0 Then loOut.DataBodyRange.Cells(r, cReal).ClearContents
        If cBatch > 0 Then loOut.DataBodyRange.Cells(r, cBatch).ClearContents
        If cRecall > 0 Then loOut.DataBodyRange.Cells(r, cRecall).ClearContents
    Next r

    If Not nextBatchMap Is Nothing Then
        For r = 1 To loOut.DataBodyRange.rows.count
            procName = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
            If procName <> "" Then
                key = LCase$(procName)
                If nextBatchMap.Exists(key) Then
                    loOut.DataBodyRange.Cells(r, cBatch).value = nextBatchMap(key)
                End If
            End If
        Next r
    End If

    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(CHK_RECALL_PREFIX) & "*" Then
                On Error Resume Next
                shp.ControlFormat.value = 0
                On Error GoTo 0
            End If
        End If
    Next shp

    EnsureOutputBatchNumbers loOut
End Sub

Private Function CaptureProductionOutputCompletionValues(ByVal loOut As ListObject, Optional ByVal onlyOutputRow As Long = 0) As Object
    Dim result As Object
    Dim cReal As Long
    Dim cBatch As Long
    Dim cRecall As Long
    Dim cRow As Long
    Dim cItemCode As Long
    Dim cOutput As Long
    Dim r As Long
    Dim rowData As Object

    If loOut Is Nothing Then Exit Function
    If loOut.DataBodyRange Is Nothing Then Exit Function

    cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    cBatch = ColumnIndex(loOut, "BATCH")
    cRecall = ColumnIndex(loOut, "RECALL CODE")
    cRow = ColumnIndex(loOut, "System_Key")
    cItemCode = ColumnIndex(loOut, "ITEM_CODE")
    cOutput = ColumnIndex(loOut, "OUTPUT")

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    For r = 1 To loOut.DataBodyRange.rows.count
        If onlyOutputRow > 0 And r <> onlyOutputRow Then GoTo NextRow
        If cReal > 0 Then
            If NzDbl(loOut.DataBodyRange.Cells(r, cReal).value) <= 0 Then GoTo NextRow
        End If

        Set rowData = CreateObject("Scripting.Dictionary")
        rowData.CompareMode = vbTextCompare
        If cReal > 0 Then rowData("REAL OUTPUT") = loOut.DataBodyRange.Cells(r, cReal).value
        If cBatch > 0 Then rowData("BATCH") = loOut.DataBodyRange.Cells(r, cBatch).value
        If cRecall > 0 Then rowData("RECALL CODE") = loOut.DataBodyRange.Cells(r, cRecall).value
        If cRow > 0 Then rowData("System_Key") = loOut.DataBodyRange.Cells(r, cRow).value
        If cItemCode > 0 Then rowData("ITEM_CODE") = loOut.DataBodyRange.Cells(r, cItemCode).Value
        If cOutput > 0 Then rowData("OUTPUT") = loOut.DataBodyRange.Cells(r, cOutput).value
        result.Add CStr(r), rowData
NextRow:
    Next r

    If result.count > 0 Then Set CaptureProductionOutputCompletionValues = result
End Function

Private Sub RestoreProductionOutputCompletionValues(ByVal loOut As ListObject, ByVal pendingValues As Object)
    Dim key As Variant
    Dim r As Long
    Dim rowData As Object

    If loOut Is Nothing Then Exit Sub
    If loOut.DataBodyRange Is Nothing Then Exit Sub
    If pendingValues Is Nothing Then Exit Sub
    If pendingValues.count = 0 Then Exit Sub

    For Each key In pendingValues.keys
        r = CLng(val(CStr(key)))
        If r < 1 Or r > loOut.DataBodyRange.rows.count Then GoTo NextKey
        Set rowData = pendingValues(key)
        RestoreProductionOutputCell loOut, r, "REAL OUTPUT", rowData
        RestoreProductionOutputCell loOut, r, "BATCH", rowData
        RestoreProductionOutputCell loOut, r, "RECALL CODE", rowData
        RestoreProductionOutputCell loOut, r, "System_Key", rowData
        RestoreProductionOutputCell loOut, r, "ITEM_CODE", rowData
        RestoreProductionOutputCell loOut, r, "OUTPUT", rowData
NextKey:
    Next key
End Sub

Private Sub RestoreProductionOutputCell(ByVal loOut As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal rowData As Object)
    Dim c As Long

    If rowData Is Nothing Then Exit Sub
    If Not rowData.Exists(columnName) Then Exit Sub
    c = ColumnIndex(loOut, columnName)
    If c = 0 And columnName = "REAL OUTPUT" Then c = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    If c = 0 Then Exit Sub
    If rowIndex < 1 Or rowIndex > loOut.DataBodyRange.rows.count Then Exit Sub
    If Trim$(NzStr(loOut.DataBodyRange.Cells(rowIndex, c).value)) = "" Then
        loOut.DataBodyRange.Cells(rowIndex, c).value = rowData(columnName)
    End If
End Sub

Private Function LogProductionOutputToProductionLog(ByVal wsProd As Worksheet, ByVal loOut As ListObject, ByVal invLo As ListObject, ByRef errNotes As String, Optional ByVal onlyOutputRow As Long = 0) As Boolean
    Dim logStep As String
    Dim lr As ListRow
    Dim addedRow As Boolean

    On Error GoTo ErrHandler

    logStep = "validating Production output rows"
    If wsProd Is Nothing Then Exit Function
    If loOut Is Nothing Then Exit Function
    If loOut.DataBodyRange Is Nothing Then Exit Function

    Dim wsLog As Worksheet
    logStep = "finding the ProductionLog worksheet"
    Set wsLog = WorksheetFromWorkbook(wsProd.Parent, "ProductionLog")
    If wsLog Is Nothing Then
        AppendNote errNotes, "ProductionLog sheet not found."
        Exit Function
    End If

    Dim loLog As ListObject
    logStep = "finding the ProductionLog table"
    Set loLog = FindListObjectByNameOrHeaders(wsLog, "ProductionLog", Array("PROCESS", "BATCH", "TIMESTAMP"))
    If loLog Is Nothing Then
        Set loLog = FindListObjectByNameOrHeaders(wsLog, "Table46", Array("PROCESS", "BATCH", "TIMESTAMP"))
    End If
    If loLog Is Nothing Then
        AppendNote errNotes, "ProductionLog table not found."
        Exit Function
    End If

    Dim cLogRecipe As Long: cLogRecipe = ColumnIndex(loLog, "RECIPE")
    Dim cLogRecipeId As Long: cLogRecipeId = ColumnIndex(loLog, "RECIPE_ID")
    Dim cLogDept As Long: cLogDept = ColumnIndex(loLog, "DEPARTMENT")
    Dim cLogDesc As Long: cLogDesc = ColumnIndex(loLog, "DESCRIPTION")
    Dim cLogPred As Long: cLogPred = ColumnIndex(loLog, "PREDICTED OUTPUT")
    Dim cLogProc As Long: cLogProc = ColumnIndex(loLog, "PROCESS")
    Dim cLogOutput As Long: cLogOutput = ColumnIndex(loLog, "OUTPUT")
    Dim cLogReal As Long: cLogReal = ColumnIndex(loLog, "REAL OUTPUT")
    If cLogReal = 0 Then cLogReal = ColumnIndexLoose(loLog, "REALOUTPUT", "REAL_OUTPUT")
    Dim cLogBatch As Long: cLogBatch = ColumnIndex(loLog, "BATCH")
    Dim cLogBatchId As Long: cLogBatchId = ColumnIndex(loLog, "BATCH_ID")
    Dim cLogItemCode As Long: cLogItemCode = ColumnIndex(loLog, "ITEM_CODE")
    Dim cLogVendors As Long: cLogVendors = ColumnIndex(loLog, "VENDORS")
    Dim cLogVendCode As Long: cLogVendCode = ColumnIndex(loLog, "VENDOR_CODE")
    Dim cLogItem As Long: cLogItem = ColumnIndex(loLog, "ITEM")
    Dim cLogUom As Long: cLogUom = ColumnIndex(loLog, "UOM")
    Dim cLogQty As Long: cLogQty = ColumnIndex(loLog, "QUANTITY")
    Dim cLogLoc As Long: cLogLoc = ColumnIndex(loLog, "LOCATION")
    Dim cLogSystemKey As Long: cLogSystemKey = ColumnIndex(loLog, "System_Key")
    Dim cLogIO As Long: cLogIO = ColumnIndex(loLog, "INPUT/OUTPUT")
    Dim cLogTime As Long: cLogTime = ColumnIndex(loLog, "TIMESTAMP")
    Dim cLogIngId As Long: cLogIngId = ColumnIndex(loLog, "INGREDIENT_ID")
    Dim cLogGuid As Long: cLogGuid = ColumnIndex(loLog, "GUID")

    Dim recipeName As String
    Dim recipeId As String
    Dim recipeDept As String
    Dim recipeDesc As String
    Dim recipePred As String
    Dim loChooser As ListObject
    Set loChooser = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If Not loChooser Is Nothing Then
        recipeName = FirstNonEmptyColumnValue(loChooser, "RECIPE")
        recipeId = FirstNonEmptyColumnValue(loChooser, "RECIPE_ID")
        recipeDept = FirstNonEmptyColumnValue(loChooser, "DEPARTMENT")
        recipeDesc = FirstNonEmptyColumnValue(loChooser, "DESCRIPTION")
        recipePred = FirstNonEmptyColumnValue(loChooser, "PREDICTED OUTPUT")
    End If

    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    Dim cOutput As Long: cOutput = ColumnIndex(loOut, "OUTPUT")
    Dim cUom As Long: cUom = ColumnIndex(loOut, "UOM")
    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cSystemKey As Long: cSystemKey = ColumnIndex(loOut, "System_Key")
    Dim cOutputItemCode As Long: cOutputItemCode = ColumnIndex(loOut, "ITEM_CODE")

    If cReal = 0 Or cProc = 0 Then
        AppendNote errNotes, "ProductionOutput is missing PROCESS or REAL OUTPUT."
        Exit Function
    End If

    Dim systemKeyIndex As Object
    If Not invLo Is Nothing Then
        Set systemKeyIndex = BuildInvSysSystemKeyIndex(invLo)
    End If

    Dim cInvItemCode As Long, cInvVendors As Long, cInvVendCode As Long
    Dim cInvItem As Long, cInvUom As Long, cInvLoc As Long
    If Not invLo Is Nothing Then
        cInvItemCode = ColumnIndex(invLo, "ITEM_CODE")
        cInvVendors = ColumnIndexLoose(invLo, "VENDORS", "VENDOR", "VENDOR(S)")
        cInvVendCode = ColumnIndex(invLo, "VENDOR_CODE")
        cInvItem = ColumnIndex(invLo, "ITEM")
        cInvUom = ColumnIndex(invLo, "UOM")
        cInvLoc = ColumnIndex(invLo, "LOCATION")
    End If

    Dim r As Long
    For r = 1 To loOut.DataBodyRange.rows.count
        If onlyOutputRow > 0 And r <> onlyOutputRow Then GoTo NextRow

        Dim realVal As Double
        realVal = NzDbl(loOut.DataBodyRange.Cells(r, cReal).value)
        If realVal <= 0 Then GoTo NextRow

        Dim procName As String
        procName = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
        If procName = "" Then GoTo NextRow

        Dim outputName As String
        If cOutput > 0 Then outputName = NzStr(loOut.DataBodyRange.Cells(r, cOutput).value)

        Dim batchVal As String
        If cBatch > 0 Then batchVal = NzStr(loOut.DataBodyRange.Cells(r, cBatch).value)

        Dim systemKey As String
        If cSystemKey > 0 Then systemKey = Trim$(NzStr(loOut.DataBodyRange.Cells(r, cSystemKey).value))
        If systemKey = "" Then
            AppendNote errNotes, "Completed output '" & outputName & "' is missing immutable System_Key identity."
            GoTo NextRow
        End If
        Dim itemCode As String
        If cOutputItemCode > 0 Then itemCode = Trim$(NzStr(loOut.DataBodyRange.Cells(r, cOutputItemCode).Value))
        If itemCode = "" Then
            Dim resolvedItemName As String
            Dim candidateSystemKey As String
            ResolveProductionOutputIdentity outputName, candidateSystemKey, itemCode, resolvedItemName, errNotes
        End If

        Dim vendors As String
        Dim vendCode As String
        Dim itemName As String
        Dim uomVal As String
        Dim locVal As String

        If Not systemKeyIndex Is Nothing Then
            If systemKeyIndex.Exists(systemKey) Then
                Dim invIdx As Long
                invIdx = CLng(systemKeyIndex(systemKey))
                If cInvItemCode > 0 Then itemCode = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvItemCode).value)
                If cInvVendors > 0 Then vendors = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvVendors).value)
                If cInvVendCode > 0 Then vendCode = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvVendCode).value)
                If cInvItem > 0 Then itemName = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvItem).value)
                If cInvUom > 0 Then uomVal = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvUom).value)
                If cInvLoc > 0 Then locVal = NzStr(invLo.DataBodyRange.Cells(invIdx, cInvLoc).value)
            End If
        End If
        If itemName = "" Then itemName = outputName
        If uomVal = "" And cUom > 0 Then uomVal = NzStr(loOut.DataBodyRange.Cells(r, cUom).value)

        logStep = "checking for an existing ProductionLog row"
        If ProductionLogEntryExists(loLog, procName, outputName, batchVal, systemKey, realVal) Then
            AppendNote errNotes, "The completed output was already present in ProductionLog; no duplicate row was added."
            GoTo NextRow
        End If

        logStep = "adding a ProductionLog row"
        Set lr = loLog.ListRows.Add
        addedRow = True
        logStep = "writing ProductionLog fields"
        If cLogRecipe > 0 Then lr.Range.Cells(1, cLogRecipe).value = recipeName
        If cLogRecipeId > 0 Then lr.Range.Cells(1, cLogRecipeId).value = recipeId
        If cLogDept > 0 Then lr.Range.Cells(1, cLogDept).value = recipeDept
        If cLogDesc > 0 Then lr.Range.Cells(1, cLogDesc).value = recipeDesc
        If cLogPred > 0 Then lr.Range.Cells(1, cLogPred).value = recipePred
        If cLogProc > 0 Then lr.Range.Cells(1, cLogProc).value = procName
        If cLogOutput > 0 Then lr.Range.Cells(1, cLogOutput).value = outputName
        If cLogReal > 0 Then lr.Range.Cells(1, cLogReal).value = realVal
        If cLogBatch > 0 Then lr.Range.Cells(1, cLogBatch).value = batchVal
        If cLogBatchId > 0 Then lr.Range.Cells(1, cLogBatchId).value = Format$(Date, "yyyymmdd") & "-" & batchVal
        If cLogItemCode > 0 Then lr.Range.Cells(1, cLogItemCode).value = itemCode
        If cLogVendors > 0 Then lr.Range.Cells(1, cLogVendors).value = vendors
        If cLogVendCode > 0 Then lr.Range.Cells(1, cLogVendCode).value = vendCode
        If cLogItem > 0 Then lr.Range.Cells(1, cLogItem).value = itemName
        If cLogUom > 0 Then lr.Range.Cells(1, cLogUom).value = uomVal
        If cLogQty > 0 Then lr.Range.Cells(1, cLogQty).value = realVal
        If cLogLoc > 0 Then lr.Range.Cells(1, cLogLoc).value = locVal
        If cLogSystemKey > 0 Then lr.Range.Cells(1, cLogSystemKey).value = systemKey
        If cLogIO > 0 Then lr.Range.Cells(1, cLogIO).value = "MADE"
        If cLogTime > 0 Then lr.Range.Cells(1, cLogTime).value = Now
        If cLogIngId > 0 Then lr.Range.Cells(1, cLogIngId).value = ""
        If cLogGuid > 0 Then lr.Range.Cells(1, cLogGuid).value = CreateProductionGuid()
        addedRow = False
        Set lr = Nothing
NextRow:
    Next r
    LogProductionOutputToProductionLog = True
    Exit Function

ErrHandler:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.Description
    If addedRow Then
        On Error Resume Next
        If Not lr Is Nothing Then lr.Delete
        On Error GoTo 0
    End If
    AppendNote errNotes, "ProductionLog failed while " & logStep & ": " & _
        CStr(failureNumber) & " - " & failureDescription
End Function

Private Function WorksheetFromWorkbook(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 _
           Or StrComp(ws.CodeName, sheetName, vbTextCompare) = 0 Then
            Set WorksheetFromWorkbook = ws
            Exit Function
        End If
    Next ws
End Function

Private Function ProductionLogEntryExists(ByVal loLog As ListObject, _
                                          ByVal procName As String, _
                                          ByVal outputName As String, _
                                          ByVal batchVal As String, _
                                          ByVal systemKey As String, _
                                          ByVal realVal As Double) As Boolean
    If loLog Is Nothing Then Exit Function
    If loLog.DataBodyRange Is Nothing Then Exit Function

    Dim cProc As Long: cProc = ColumnIndex(loLog, "PROCESS")
    Dim cOutput As Long: cOutput = ColumnIndex(loLog, "OUTPUT")
    Dim cBatch As Long: cBatch = ColumnIndex(loLog, "BATCH")
    Dim cSystemKey As Long: cSystemKey = ColumnIndex(loLog, "System_Key")
    Dim cReal As Long: cReal = ColumnIndex(loLog, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loLog, "REALOUTPUT", "REAL_OUTPUT")
    If cProc = 0 Or cBatch = 0 Or cReal = 0 Then Exit Function

    Dim values As Variant
    values = loLog.DataBodyRange.Value

    Dim r As Long
    For r = 1 To UBound(values, 1)
        If StrComp(Trim$(NzStr(values(r, cProc))), Trim$(procName), vbTextCompare) <> 0 Then GoTo NextRow
        If StrComp(Trim$(NzStr(values(r, cBatch))), Trim$(batchVal), vbTextCompare) <> 0 Then GoTo NextRow
        If Abs(NzDbl(values(r, cReal)) - realVal) > 0.0000001 Then GoTo NextRow
        If systemKey <> "" And cSystemKey > 0 Then
            If StrComp(Trim$(NzStr(values(r, cSystemKey))), systemKey, vbTextCompare) <> 0 Then GoTo NextRow
        ElseIf outputName <> "" And cOutput > 0 Then
            If StrComp(Trim$(NzStr(values(r, cOutput))), Trim$(outputName), vbTextCompare) <> 0 Then GoTo NextRow
        End If
        ProductionLogEntryExists = True
        Exit Function
NextRow:
    Next r
End Function

Private Sub ApplyRecallCodesForOutput(ByVal wsProd As Worksheet, ByVal loOut As ListObject, ByVal invLo As ListObject, ByRef errNotes As String)
    If wsProd Is Nothing Then Exit Sub
    If loOut Is Nothing Then Exit Sub
    If loOut.DataBodyRange Is Nothing Then Exit Sub

    Dim recallRows As Object
    Set recallRows = GetRecallCheckedRows(wsProd)
    If recallRows Is Nothing Then Exit Sub
    If recallRows.count = 0 Then Exit Sub

    Dim cRecall As Long: cRecall = ColumnIndex(loOut, "RECALL CODE")
    If cRecall = 0 Then Exit Sub

    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    Dim cOutput As Long: cOutput = ColumnIndex(loOut, "OUTPUT")
    Dim cUom As Long: cUom = ColumnIndex(loOut, "UOM")
    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cSystemKey As Long: cSystemKey = ColumnIndex(loOut, "System_Key")

    Dim recipeName As String
    Dim recipeId As String
    GetRecipeChooserInfo wsProd, recipeName, recipeId

    Dim wsLog As Worksheet
    Set wsLog = SheetExists("BatchCodesLog")
    If wsLog Is Nothing Then Set wsLog = SheetExists("BatchCodeLogs")

    Dim loLog As ListObject
    If Not wsLog Is Nothing Then
        Set loLog = FindListObjectByNameOrHeaders(wsLog, "Table48", Array("RECIPE", "RECIPE_ID", "PROCESS", "OUTPUT"))
    End If

    Dim cLogRec As Long, cLogRecId As Long, cLogProc As Long, cLogOut As Long
    Dim cLogUom As Long, cLogReal As Long, cLogBatch As Long, cLogRecall As Long
    Dim cLogTime As Long, cLogLoc As Long, cLogUser As Long, cLogGuid As Long
    If Not loLog Is Nothing Then
        cLogRec = ColumnIndex(loLog, "RECIPE")
        cLogRecId = ColumnIndex(loLog, "RECIPE_ID")
        cLogProc = ColumnIndex(loLog, "PROCESS")
        cLogOut = ColumnIndex(loLog, "OUTPUT")
        cLogUom = ColumnIndex(loLog, "UOM")
        cLogReal = ColumnIndex(loLog, "REAL OUTPUT")
        If cLogReal = 0 Then cLogReal = ColumnIndexLoose(loLog, "REALOUTPUT", "REAL_OUTPUT")
        cLogBatch = ColumnIndex(loLog, "BATCH")
        cLogRecall = ColumnIndex(loLog, "RECALL CODE")
        cLogTime = ColumnIndex(loLog, "TIMESTAMP")
        cLogLoc = ColumnIndex(loLog, "LOCATION")
        cLogUser = ColumnIndex(loLog, "USER")
        cLogGuid = ColumnIndex(loLog, "GUID")
    End If

    Dim key As Variant
    For Each key In recallRows.keys
        Dim idx As Long: idx = CLng(key)
        If idx < 1 Or idx > loOut.DataBodyRange.rows.count Then GoTo NextRow

        Dim codeVal As String
        codeVal = NzStr(loOut.DataBodyRange.Cells(idx, cRecall).value)
        If Trim$(codeVal) = "" Then
            codeVal = GenerateRecallCode()
            loOut.DataBodyRange.Cells(idx, cRecall).value = codeVal

            If Not loLog Is Nothing Then
                Dim lr As ListRow: Set lr = loLog.ListRows.Add
                If cLogRec > 0 Then lr.Range.Cells(1, cLogRec).value = recipeName
                If cLogRecId > 0 Then lr.Range.Cells(1, cLogRecId).value = recipeId
                If cLogProc > 0 And cProc > 0 Then lr.Range.Cells(1, cLogProc).value = loOut.DataBodyRange.Cells(idx, cProc).value
                If cLogOut > 0 And cOutput > 0 Then lr.Range.Cells(1, cLogOut).value = loOut.DataBodyRange.Cells(idx, cOutput).value
                If cLogUom > 0 And cUom > 0 Then lr.Range.Cells(1, cLogUom).value = loOut.DataBodyRange.Cells(idx, cUom).value
                If cLogReal > 0 And cReal > 0 Then lr.Range.Cells(1, cLogReal).value = loOut.DataBodyRange.Cells(idx, cReal).value
                If cLogBatch > 0 And cBatch > 0 Then lr.Range.Cells(1, cLogBatch).value = loOut.DataBodyRange.Cells(idx, cBatch).value
                If cLogRecall > 0 Then lr.Range.Cells(1, cLogRecall).value = codeVal
                If cLogTime > 0 Then lr.Range.Cells(1, cLogTime).value = Now
                If cLogUser > 0 Then lr.Range.Cells(1, cLogUser).value = modRoleEventWriter.ResolveCurrentUserId()
                If cLogGuid > 0 Then lr.Range.Cells(1, cLogGuid).value = CreateProductionGuid()
                If cLogLoc > 0 Then
                    Dim locVal As String
                    If cSystemKey > 0 Then
                        locVal = ResolveInvSysLocationBySystemKey( _
                            invLo, Trim$(NzStr(loOut.DataBodyRange.Cells(idx, cSystemKey).value)))
                    End If
                    If locVal <> "" Then lr.Range.Cells(1, cLogLoc).value = locVal
                End If
            End If
        End If
NextRow:
    Next key
End Sub

Private Function GetRecallCheckedRows(ByVal ws As Worksheet) As Object
    If ws Is Nothing Then Exit Function
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(CHK_RECALL_PREFIX) & "*" Then
                Dim idx As Long
                idx = ParseCheckboxIndex(shp.Name, CHK_RECALL_PREFIX)
                If idx > 0 Then
                    If shp.ControlFormat.value = 1 Then dict(CStr(idx)) = True
                End If
            End If
        End If
    Next shp
    If dict.count = 0 Then Exit Function
    Set GetRecallCheckedRows = dict
End Function

Private Function GenerateRecallCode() As String
    Dim guidVal As String
    guidVal = Replace(CreateProductionGuid(), "-", "")
    GenerateRecallCode = "RC-" & Left$(guidVal, 12)
End Function

Public Function CreateProductionGuid() As String
    On Error Resume Next
    CreateProductionGuid = CStr(CreateObject("Scriptlet.TypeLib").GUID)
    CreateProductionGuid = Replace(Replace(Trim$(CreateProductionGuid), "{", ""), "}", "")
    If Trim$(CreateProductionGuid) = "" Then
        Randomize
        CreateProductionGuid = Format$(Now, "yyyymmddhhnnss") & "-" & CStr(Int(Rnd() * 1000000#))
    End If
    On Error GoTo 0
End Function

Private Function BuildRecallCodesReportFromCurrentWorkbook(ByRef wsReportOut As Worksheet, ByRef rowCountOut As Long, ByRef detailOut As String) As Boolean
    Dim wsProd As Worksheet
    Dim loOut As ListObject
    Dim invLo As ListObject

    Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then
        detailOut = "Production sheet not found."
        Exit Function
    End If

    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If loOut Is Nothing Then
        detailOut = "ProductionOutput table not found on Production sheet."
        Exit Function
    End If
    If loOut.DataBodyRange Is Nothing Then
        detailOut = "ProductionOutput has no rows to print."
        Exit Function
    End If

    Set invLo = GetInvSysTable()
    Set wsReportOut = EnsureRecallCodesReportSheet(wsProd.Parent)
    If wsReportOut Is Nothing Then
        detailOut = "Unable to create RecallCodesPrint worksheet."
        Exit Function
    End If

    rowCountOut = RenderRecallCodesReport(wsProd, loOut, invLo, wsReportOut)
    If rowCountOut <= 0 Then
        detailOut = "No recall-coded ProductionOutput rows found. Generate recall codes from checked output rows before printing."
        Exit Function
    End If

    detailOut = "Sheet=" & wsReportOut.Name & ";Rows=" & rowCountOut
    BuildRecallCodesReportFromCurrentWorkbook = True
End Function

Private Function EnsureRecallCodesReportSheet(ByVal wb As Workbook) As Worksheet
    If wb Is Nothing Then Exit Function

    Set EnsureRecallCodesReportSheet = WorkbookSheetExists(wb, PROD_RECALL_REPORT_SHEET)
    If EnsureRecallCodesReportSheet Is Nothing Then
        Set EnsureRecallCodesReportSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureRecallCodesReportSheet.Name = PROD_RECALL_REPORT_SHEET
    End If

    Dim lo As ListObject
    For Each lo In EnsureRecallCodesReportSheet.ListObjects
        lo.Delete
    Next lo

    EnsureRecallCodesReportSheet.Cells.Clear
End Function

Private Function RenderRecallCodesReport(ByVal wsProd As Worksheet, ByVal loOut As ListObject, ByVal invLo As ListObject, ByVal wsReport As Worksheet) As Long
    If wsProd Is Nothing Then Exit Function
    If loOut Is Nothing Then Exit Function
    If wsReport Is Nothing Then Exit Function
    If loOut.DataBodyRange Is Nothing Then Exit Function

    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    Dim cOutput As Long: cOutput = ColumnIndex(loOut, "OUTPUT")
    Dim cReal As Long: cReal = ColumnIndex(loOut, "REAL OUTPUT")
    If cReal = 0 Then cReal = ColumnIndexLoose(loOut, "REALOUTPUT", "REAL_OUTPUT")
    Dim cUom As Long: cUom = ColumnIndex(loOut, "UOM")
    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cRecall As Long: cRecall = ColumnIndex(loOut, "RECALL CODE")
    Dim cSystemKey As Long: cSystemKey = ColumnIndex(loOut, "System_Key")
    If cRecall = 0 Then Exit Function

    Dim src As Variant
    src = loOut.DataBodyRange.Value

    Dim rowCount As Long
    Dim r As Long
    For r = 1 To UBound(src, 1)
        If Trim$(NzStr(src(r, cRecall))) <> "" Then rowCount = rowCount + 1
    Next r
    If rowCount = 0 Then Exit Function

    Dim reportData() As Variant
    ReDim reportData(1 To rowCount + 1, 1 To 9)
    reportData(1, 1) = "RECIPE"
    reportData(1, 2) = "RECIPE_ID"
    reportData(1, 3) = "PROCESS"
    reportData(1, 4) = "OUTPUT"
    reportData(1, 5) = "REAL OUTPUT"
    reportData(1, 6) = "UOM"
    reportData(1, 7) = "BATCH"
    reportData(1, 8) = "RECALL CODE"
    reportData(1, 9) = "LOCATION"

    Dim recipeName As String
    Dim recipeId As String
    GetRecipeChooserInfo wsProd, recipeName, recipeId

    Dim outRow As Long
    outRow = 2
    For r = 1 To UBound(src, 1)
        Dim recallCode As String
        recallCode = Trim$(NzStr(src(r, cRecall)))
        If recallCode = "" Then GoTo NextSourceRow

        reportData(outRow, 1) = recipeName
        reportData(outRow, 2) = recipeId
        If cProc > 0 Then reportData(outRow, 3) = NzStr(src(r, cProc))
        If cOutput > 0 Then reportData(outRow, 4) = NzStr(src(r, cOutput))
        If cReal > 0 Then reportData(outRow, 5) = src(r, cReal)
        If cUom > 0 Then reportData(outRow, 6) = NzStr(src(r, cUom))
        If cBatch > 0 Then reportData(outRow, 7) = NzStr(src(r, cBatch))
        reportData(outRow, 8) = recallCode
        If cSystemKey > 0 Then
            reportData(outRow, 9) = ResolveInvSysLocationBySystemKey( _
                invLo, Trim$(NzStr(src(r, cSystemKey))))
        End If
        outRow = outRow + 1
NextSourceRow:
    Next r

    With wsReport
        .Range("A1").Value = "Production Recall Codes"
        .Range("A2").Value = "Workbook"
        .Range("B2").Value = wsProd.Parent.Name
        .Range("C2").Value = "Generated"
        .Range("D2").Value = Now
        .Range("A3").Value = "Rows"
        .Range("B3").Value = rowCount
        .Range("C3").Value = "User"
        .Range("D3").Value = modRoleEventWriter.ResolveCurrentUserId()
        .Range("A1:D3").Font.Bold = True

        Dim tableRange As Range
        Set tableRange = .Range("A5").Resize(rowCount + 1, 9)
        tableRange.Value = reportData

        Dim loReport As ListObject
        Set loReport = .ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        loReport.Name = TABLE_RECALL_REPORT
        loReport.TableStyle = "TableStyleMedium2"

        .Columns("A:I").AutoFit
        .Range("D2").NumberFormat = "yyyy-mm-dd hh:mm:ss"
        ' PageSetup can raise 1004 when the workstation has no usable default
        ' printer. The recall table remains valid; only print-page preferences
        ' are optional in that environment.
        On Error Resume Next
        .PageSetup.Orientation = xlLandscape
        .PageSetup.Zoom = False
        .PageSetup.FitToPagesWide = 1
        .PageSetup.FitToPagesTall = False
        .PageSetup.PrintArea = .Range("A1").Resize(tableRange.Rows.Count + 4, tableRange.Columns.Count).Address
        On Error GoTo 0
    End With

    RenderRecallCodesReport = rowCount
End Function

Private Function GenerateBatchNumber(ByVal wsProd As Worksheet, ByVal loOut As ListObject, ByVal procName As String) As Long
    GenerateBatchNumber = NextBatchSequenceForProcess(wsProd, loOut, procName)
End Function

Private Sub EnsureOutputBatchNumbers(ByVal loOut As ListObject)
    If loOut Is Nothing Then Exit Sub
    If loOut.DataBodyRange Is Nothing Then Exit Sub

    Dim wsProd As Worksheet
    Set wsProd = loOut.Parent

    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    If cBatch = 0 Or cProc = 0 Then Exit Sub

    Dim batchMap As Object
    Set batchMap = CreateObject("Scripting.Dictionary")

    Dim r As Long
    For r = 1 To loOut.DataBodyRange.rows.count
        Dim procName As String
        procName = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
        If procName <> "" Then
            Dim existingBatch As String
            existingBatch = NzStr(loOut.DataBodyRange.Cells(r, cBatch).value)
            If IsNumeric(existingBatch) Then
                batchMap(LCase$(procName)) = CStr(CLng(val(existingBatch)))
            End If
        End If
    Next r

    For r = 1 To loOut.DataBodyRange.rows.count
        Dim batchVal As String
        batchVal = NzStr(loOut.DataBodyRange.Cells(r, cBatch).value)
        If batchVal = "" Or Not IsNumeric(batchVal) Then
            Dim procName2 As String
            procName2 = NzStr(loOut.DataBodyRange.Cells(r, cProc).value)
            If procName2 <> "" Then
                Dim key As String
                key = LCase$(procName2)
                If batchMap.Exists(key) Then
                    loOut.DataBodyRange.Cells(r, cBatch).value = batchMap(key)
                Else
                    Dim newBatch As Long
                    newBatch = GenerateBatchNumber(wsProd, loOut, procName2)
                    If newBatch > 0 Then
                        loOut.DataBodyRange.Cells(r, cBatch).value = newBatch
                        batchMap(key) = CStr(newBatch)
                    End If
                End If
            End If
        End If
    Next r
End Sub

Private Function NextBatchSequenceForProcess(ByVal wsProd As Worksheet, ByVal loOut As ListObject, ByVal procName As String) As Long
    Dim maxBatch As Long
    maxBatch = MaxBatchFromOutput(loOut, procName)

    Dim wsLog As Worksheet
    Dim loLog As ListObject

    Set wsLog = SheetExists("BatchCodesLog")
    If wsLog Is Nothing Then Set wsLog = SheetExists("BatchCodeLogs")
    If Not wsLog Is Nothing Then
        Set loLog = FindListObjectByNameOrHeaders(wsLog, "Table48", Array("PROCESS", "BATCH", "TIMESTAMP"))
        If Not loLog Is Nothing Then
            AccumulateBatchMaxFromLog loLog, procName, maxBatch
        End If
    End If

    Dim wsProdLog As Worksheet
    Set wsProdLog = SheetExists("ProductionLog")
    If Not wsProdLog Is Nothing Then
        Dim loProdLog As ListObject
        Set loProdLog = FindListObjectByNameOrHeaders(wsProdLog, "ProductionLog", Array("PROCESS", "BATCH", "TIMESTAMP"))
        If loProdLog Is Nothing Then
            Set loProdLog = FindListObjectByNameOrHeaders(wsProdLog, "Table46", Array("PROCESS", "BATCH", "TIMESTAMP"))
        End If
        If Not loProdLog Is Nothing Then
            AccumulateBatchMaxFromLog loProdLog, procName, maxBatch
        End If
    End If

    NextBatchSequenceForProcess = maxBatch + 1
End Function

Private Function MaxBatchFromOutput(ByVal loOut As ListObject, ByVal procName As String) As Long
    If loOut Is Nothing Then Exit Function
    If loOut.DataBodyRange Is Nothing Then Exit Function

    Dim cBatch As Long: cBatch = ColumnIndex(loOut, "BATCH")
    Dim cProc As Long: cProc = ColumnIndex(loOut, "PROCESS")
    If cBatch = 0 Or cProc = 0 Then Exit Function

    Dim arr As Variant: arr = loOut.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If StrComp(NzStr(arr(r, cProc)), procName, vbTextCompare) = 0 Then
            Dim b As Long
            b = CLng(val(arr(r, cBatch)))
            If b > MaxBatchFromOutput Then MaxBatchFromOutput = b
        End If
    Next r
End Function

Private Sub AccumulateBatchMaxFromLog(ByVal loLog As ListObject, ByVal procName As String, ByRef maxBatch As Long)
    If loLog Is Nothing Then Exit Sub
    If loLog.DataBodyRange Is Nothing Then Exit Sub

    Dim cBatch As Long: cBatch = ColumnIndex(loLog, "BATCH")
    Dim cProc As Long: cProc = ColumnIndex(loLog, "PROCESS")
    Dim cTime As Long: cTime = ColumnIndex(loLog, "TIMESTAMP")
    If cBatch = 0 Or cTime = 0 Then Exit Sub

    Dim arr As Variant: arr = loLog.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If cProc > 0 Then
            If StrComp(NzStr(arr(r, cProc)), procName, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        Dim tVal As Variant
        tVal = arr(r, cTime)
        If Not IsDate(tVal) Then GoTo NextRow
        If DateValue(tVal) <> Date Then GoTo NextRow

        Dim b As Long
        b = CLng(val(arr(r, cBatch)))
        If b > maxBatch Then maxBatch = b
NextRow:
    Next r
End Sub

Private Sub GetRecipeChooserInfo(ByVal ws As Worksheet, ByRef recipeName As String, ByRef recipeId As String)
    recipeName = ""
    recipeId = ""
    If ws Is Nothing Then Exit Sub
    Dim lo As ListObject
    Set lo = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_CHOOSER, Array("RECIPE", "RECIPE_ID"))
    If lo Is Nothing Then Exit Sub
    recipeName = FirstNonEmptyColumnValue(lo, "RECIPE")
    recipeId = FirstNonEmptyColumnValue(lo, "RECIPE_ID")
End Sub

Private Function ResolveInvSysLocationBySystemKey(ByVal invLo As ListObject, _
                                                  ByVal systemKey As String) As String
    If invLo Is Nothing Then Exit Function
    systemKey = Trim$(systemKey)
    If systemKey = "" Then Exit Function
    If invLo.DataBodyRange Is Nothing Then Exit Function

    Dim cSystemKey As Long: cSystemKey = ColumnIndex(invLo, "System_Key")
    Dim cLoc As Long: cLoc = ColumnIndex(invLo, "LOCATION")
    If cSystemKey = 0 Or cLoc = 0 Then Exit Function

    Dim cel As Range
    For Each cel In invLo.ListColumns(cSystemKey).DataBodyRange.Cells
        If StrComp(Trim$(NzStr(cel.value)), systemKey, vbTextCompare) = 0 Then
            ResolveInvSysLocationBySystemKey = NzStr(cel.Offset(0, cLoc - cel.Column).value)
            Exit Function
        End If
    Next cel
End Function

Public Sub OutputCheckboxChanged()
    ' Placeholder for batch/recall checkbox behavior.
End Sub

Private Function ParseCheckboxIndex(ByVal shapeName As String, ByVal prefix As String) As Long
    If LCase$(Left$(shapeName, Len(prefix))) <> LCase$(prefix) Then Exit Function
    Dim tail As String
    tail = Mid$(shapeName, Len(prefix) + 1)
    If tail = "" Then Exit Function
    If IsNumeric(tail) Then ParseCheckboxIndex = CLng(val(tail))
End Function

Private Sub DeleteCheckboxesByPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    If ws Is Nothing Then Exit Sub
    Dim toDelete As Collection
    Set toDelete = New Collection
    Dim shp As Shape
    For Each shp In ws.shapes
        If IsCheckboxShape(shp) Then
            If LCase$(shp.Name) Like LCase$(prefix) & "*" Then
                toDelete.Add shp.Name
            End If
        End If
    Next shp
    Dim nameVal As Variant
    For Each nameVal In toDelete
        On Error Resume Next
        ws.shapes(CStr(nameVal)).Delete
        On Error GoTo 0
    Next nameVal
End Sub

Private Function EnsureCheckboxShape(ByVal ws As Worksheet, ByVal shapeName As String, ByVal caption As String, ByVal onActionMacro As String, _
    ByVal leftPos As Double, ByVal topPos As Double, ByVal widthPts As Double, ByVal heightPts As Double) As Shape

    If ws Is Nothing Then Exit Function
    If widthPts < 10 Then widthPts = 10
    If heightPts < 10 Then heightPts = 10

    Dim shp As Shape
    On Error Resume Next
    Set shp = ws.shapes(shapeName)
    On Error GoTo 0
    If Not shp Is Nothing Then
        If Not IsCheckboxShape(shp) Then Set shp = Nothing
    End If

    If shp Is Nothing Then
        Set shp = ws.shapes.AddFormControl(xlCheckBox, leftPos, topPos, widthPts, heightPts)
        shp.Name = shapeName
    Else
        shp.Name = shapeName
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthPts
        shp.Height = heightPts
    End If

    If onActionMacro <> "" Then shp.OnAction = onActionMacro
    ForceCheckboxCaption shp, caption
    Set EnsureCheckboxShape = shp
End Function

Private Sub ForceCheckboxCaption(ByVal shp As Shape, ByVal caption As String)
    If shp Is Nothing Then Exit Sub
    On Error Resume Next
    shp.ControlFormat.caption = caption
    shp.TextFrame.Characters.text = caption
    On Error GoTo 0
End Sub

Private Function IsCheckboxShape(ByVal shp As Shape) As Boolean
    If shp Is Nothing Then Exit Function
    If shp.Type <> SHAPE_TYPE_FORM_CONTROL Then Exit Function
    On Error Resume Next
    If shp.FormControlType = xlCheckBox Then IsCheckboxShape = True
    On Error GoTo 0
End Function

Private Function GetCheckboxCaption(ByVal shp As Shape) As String
    If shp Is Nothing Then Exit Function
    On Error Resume Next
    GetCheckboxCaption = shp.ControlFormat.caption
    If GetCheckboxCaption = "" Then GetCheckboxCaption = shp.TextFrame.Characters.text
    If GetCheckboxCaption = "" Then GetCheckboxCaption = shp.AlternativeText
    On Error GoTo 0
End Function

Private Function UniqueShapeName(ByVal ws As Worksheet, ByVal baseName As String) As String
    Dim nameTry As String
    nameTry = baseName
    Dim idx As Long
    idx = 1
    Do While ShapeExists(ws, nameTry)
        nameTry = baseName & "_" & CStr(idx)
        idx = idx + 1
    Loop
    UniqueShapeName = nameTry
End Function

Private Function ShapeExists(ByVal ws As Worksheet, ByVal shapeName As String) As Boolean
    On Error Resume Next
    Dim shp As Shape
    Set shp = ws.shapes(shapeName)
    ShapeExists = Not shp Is Nothing
    On Error GoTo 0
End Function


Private Sub FillPaletteTableFromInvSys(ByVal lo As ListObject, ByVal rowMap As Object)
    If lo Is Nothing Then Exit Sub
    If rowMap Is Nothing Then Exit Sub
    If rowMap.count = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cRow As Long: cRow = ColumnIndex(lo, "System_Key")
    If cRow = 0 Then Exit Sub
    Dim cCode As Long: cCode = ColumnIndex(lo, "ITEM_CODE")
    If cCode = 0 Then cCode = ColumnIndexLoose(lo, "ITEM_CODE", "ITEMCODE", "ITEM CODE")
    Dim cVend As Long: cVend = ColumnIndex(lo, "VENDORS")
    If cVend = 0 Then cVend = ColumnIndexLoose(lo, "VENDORS", "VENDOR", "VENDOR(S)")
    Dim cVendCode As Long: cVendCode = ColumnIndex(lo, "VENDOR_CODE")
    If cVendCode = 0 Then cVendCode = ColumnIndexLoose(lo, "VENDOR_CODE", "VENDORCODE", "VENDOR CODE")
    Dim cDesc As Long: cDesc = ColumnIndex(lo, "DESCRIPTION")
    If cDesc = 0 Then cDesc = ColumnIndexLoose(lo, "DESCRIPTION", "DESC")
    Dim cItem As Long: cItem = ColumnIndex(lo, "ITEM")
    If cItem = 0 Then cItem = ColumnIndexLoose(lo, "ITEM", "ITEMS", "ITEMNAME", "ITEM NAME")
    Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
    If cUom = 0 Then cUom = ColumnIndexLoose(lo, "UOM", "UNIT", "UNITOFMEASURE", "UNITOFMEASUREMENT")
    Dim cLoc As Long: cLoc = ColumnIndex(lo, "LOCATION")
    If cLoc = 0 Then cLoc = ColumnIndexLoose(lo, "LOCATION", "LOC")

    Dim r As Long
    For r = 1 To lo.DataBodyRange.rows.count
        Dim rowKey As String
        rowKey = NormalizeSystemKey(lo.DataBodyRange.Cells(r, cRow).value)
        If rowKey <> "" Then
            If rowMap.Exists(rowKey) Then
                Dim info As Variant
                info = rowMap(rowKey)
                If cCode > 0 And NzStr(lo.DataBodyRange.Cells(r, cCode).value) = "" Then lo.DataBodyRange.Cells(r, cCode).value = info(1)
                If cVend > 0 And NzStr(lo.DataBodyRange.Cells(r, cVend).value) = "" Then lo.DataBodyRange.Cells(r, cVend).value = info(2)
                If cVendCode > 0 And NzStr(lo.DataBodyRange.Cells(r, cVendCode).value) = "" Then lo.DataBodyRange.Cells(r, cVendCode).value = info(3)
                If cDesc > 0 And NzStr(lo.DataBodyRange.Cells(r, cDesc).value) = "" Then lo.DataBodyRange.Cells(r, cDesc).value = info(4)
                If cItem > 0 And NzStr(lo.DataBodyRange.Cells(r, cItem).value) = "" Then lo.DataBodyRange.Cells(r, cItem).value = info(5)
                If cUom > 0 And NzStr(lo.DataBodyRange.Cells(r, cUom).value) = "" Then lo.DataBodyRange.Cells(r, cUom).value = info(6)
                If cLoc > 0 And NzStr(lo.DataBodyRange.Cells(r, cLoc).value) = "" Then lo.DataBodyRange.Cells(r, cLoc).value = info(7)
            End If
        End If
    Next r
End Sub

Private Function GetIngredientPaletteRows(ByVal recipeId As String, ByVal ingredientId As String) As Collection
    Dim wsPal As Worksheet: Set wsPal = SheetExists("IngredientPalette")
    If wsPal Is Nothing Then Set wsPal = SheetExists("IngredientsPalette")
    If wsPal Is Nothing Then Exit Function

    Dim loPal As ListObject
    Set loPal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "System_Key"))
    If loPal Is Nothing Then
        Set loPal = FindListObjectByNameOrHeaders(wsPal, "Table40", Array("RECIPE_ID", "INGREDIENT_ID", "System_Key"))
    End If
    If loPal Is Nothing Then Exit Function
    If loPal.DataBodyRange Is Nothing Then Exit Function

    Dim cRec As Long: cRec = ColumnIndex(loPal, "RECIPE_ID")
    Dim cIng As Long: cIng = ColumnIndex(loPal, "INGREDIENT_ID")
    Dim cRow As Long: cRow = ColumnIndex(loPal, "System_Key")
    If cRec = 0 Or cIng = 0 Or cRow = 0 Then Exit Function

    Dim normRec As String: normRec = NormalizeIdFirst(recipeId)
    Dim normIng As String: normIng = NormalizeIdLast(ingredientId)

    Dim col As New Collection
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = loPal.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NormalizeIdFirst(NzStr(arr(r, cRec))) = normRec _
            And NormalizeIdLast(NzStr(arr(r, cIng))) = normIng Then
            Dim rowKey As String
            rowKey = NormalizeSystemKey(arr(r, cRow))
            If rowKey <> "" Then
                If Not seen.Exists(rowKey) Then
                    seen.Add rowKey, True
                    col.Add Array(rowKey, PaletteSourceValue(arr, r, loPal, "ITEM"), PaletteSourceValue(arr, r, loPal, "UOM"))
                End If
            End If
        End If
    Next r

    If col.count = 0 Then Exit Function
    Set GetIngredientPaletteRows = col
End Function

Private Function PaletteSourceValue(ByVal arr As Variant, ByVal rowIndex As Long, ByVal lo As ListObject, ByVal headerName As String) As String
    Dim colIndex As Long
    colIndex = ColumnIndex(lo, headerName)
    If colIndex <= 0 Then Exit Function
    PaletteSourceValue = NzStr(arr(rowIndex, colIndex))
End Function

Private Function PaletteEntryField(ByVal entry As Variant, ByVal fieldIndex As Long) As String
    On Error GoTo CleanFail
    If IsArray(entry) Then
        PaletteEntryField = NzStr(entry(LBound(entry) + fieldIndex))
    ElseIf fieldIndex = 0 Then
        PaletteEntryField = NzStr(entry)
    End If
    Exit Function
CleanFail:
    PaletteEntryField = ""
End Function

Public Function GetPaletteTableContext(ByVal lo As ListObject, ByRef recipeId As String, ByRef ingredientId As String, _
    ByRef amount As Variant, ByRef procName As String, ByRef ioVal As String) As Boolean

    GetPaletteTableContext = False
    If lo Is Nothing Then Exit Function
    If mPaletteTableMeta Is Nothing Then Exit Function
    If Not mPaletteTableMeta.Exists(lo.Name) Then Exit Function

    Dim info As Variant
    info = mPaletteTableMeta(lo.Name)
    recipeId = NzStr(info(0))
    ingredientId = NzStr(info(1))
    amount = info(2)
    procName = NzStr(info(3))
    ioVal = NzStr(info(4))
    GetPaletteTableContext = True
End Function

Public Function GetPaletteTableContextInfo(ByVal lo As ListObject) As Variant
    Dim recipeId As String
    Dim ingredientId As String
    Dim amount As Variant
    Dim procName As String
    Dim ioVal As String
    Dim info(1 To 5) As Variant

    If Not GetPaletteTableContext(lo, recipeId, ingredientId, amount, procName, ioVal) Then Exit Function

    info(1) = recipeId
    info(2) = ingredientId
    info(3) = amount
    info(4) = procName
    info(5) = ioVal
    GetPaletteTableContextInfo = info
End Function

Public Function GetAllowedInventorySystemKeysForIngredient(ByVal recipeId As String, ByVal ingredientId As String) As Object
    Set GetAllowedInventorySystemKeysForIngredient = Nothing
    If Trim$(recipeId) = "" Or Trim$(ingredientId) = "" Then Exit Function

    Dim wsPal As Worksheet: Set wsPal = SheetExists("IngredientPalette")
    If wsPal Is Nothing Then Set wsPal = SheetExists("IngredientsPalette")
    If wsPal Is Nothing Then Exit Function

    Dim loPal As ListObject
    Set loPal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "System_Key"))
    If loPal Is Nothing Then
        Set loPal = FindListObjectByNameOrHeaders(wsPal, "Table40", Array("RECIPE_ID", "INGREDIENT_ID", "System_Key"))
    End If
    If loPal Is Nothing Then Exit Function
    If loPal.DataBodyRange Is Nothing Then Exit Function

    Dim cRec As Long: cRec = ColumnIndex(loPal, "RECIPE_ID")
    Dim cIng As Long: cIng = ColumnIndex(loPal, "INGREDIENT_ID")
    Dim cRow As Long: cRow = ColumnIndex(loPal, "System_Key")
    If cRec = 0 Or cIng = 0 Or cRow = 0 Then Exit Function

    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim arr As Variant: arr = loPal.DataBodyRange.value
    Dim r As Long
    Dim normRec As String: normRec = NormalizeIdFirst(recipeId)
    Dim normIng As String: normIng = NormalizeIdLast(ingredientId)
    For r = 1 To UBound(arr, 1)
        If NormalizeIdFirst(NzStr(arr(r, cRec))) = normRec And NormalizeIdLast(NzStr(arr(r, cIng))) = normIng Then
            Dim rowVal As String
            rowVal = NzStr(arr(r, cRow))
            If Trim$(rowVal) <> "" Then
                If Not dict.Exists(rowVal) Then dict.Add rowVal, True
            End If
        End If
    Next r

    If dict.count = 0 Then Exit Function
    Set GetAllowedInventorySystemKeysForIngredient = dict
End Function

Private Sub FindRecipeIngredientInfo(ByVal recipeId As String, ByVal ingredientId As String, _
    ByRef ioVal As String, ByRef pctVal As Variant, ByRef uomVal As String, ByRef amtVal As Variant, _
    Optional ByVal preferredWb As Workbook = Nothing)

    ioVal = ""
    pctVal = ""
    uomVal = ""
    amtVal = ""

    Dim wsRec As Worksheet
    If preferredWb Is Nothing Then
        Set wsRec = SheetExists("Recipes")
    Else
        Set wsRec = WorkbookSheetExists(preferredWb, "Recipes")
    End If
    If wsRec Is Nothing Then Exit Sub
    Dim lo As ListObject: Set lo = GetListObject(wsRec, "Recipes")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cRecId As Long: cRecId = ColumnIndex(lo, "RECIPE_ID")
    Dim cIngId As Long: cIngId = ColumnIndex(lo, "INGREDIENT_ID")
    Dim cIO As Long: cIO = ColumnIndex(lo, "INPUT/OUTPUT")
    Dim cPct As Long: cPct = ColumnIndex(lo, "PERCENT")
    Dim cUom As Long: cUom = ColumnIndex(lo, "UOM")
    Dim cAmt As Long: cAmt = ColumnIndex(lo, "AMOUNT")
    If cRecId = 0 Or cIngId = 0 Then Exit Sub

    Dim arr As Variant: arr = lo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If NzStr(arr(r, cRecId)) = recipeId And NzStr(arr(r, cIngId)) = ingredientId Then
            If cIO > 0 Then ioVal = NzStr(arr(r, cIO))
            If cPct > 0 Then pctVal = arr(r, cPct)
            If cUom > 0 Then uomVal = NzStr(arr(r, cUom))
            If cAmt > 0 Then amtVal = arr(r, cAmt)
            Exit Sub
        End If
    Next r
End Sub

Private Function FirstNonEmptyColumnValue(ByVal lo As ListObject, ByVal colName As String) As String
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Dim c As Long: c = ColumnIndex(lo, colName)
    If c = 0 Then Exit Function
    Dim arr As Variant: arr = lo.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If Trim$(NzStr(arr(r, c))) <> "" Then
            FirstNonEmptyColumnValue = NzStr(arr(r, c))
            Exit Function
        End If
    Next r
End Function

' ===== System 1: Recipe List Builder - Load / Save =====
' System 1: Recipe List Builder - save recipe to Recipes sheet.
Private Sub SaveRecipeToRecipes()
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim wsRec As Worksheet
    Dim stagingWb As Workbook
    Dim stagingReport As String
    Dim designsEnabled As Boolean

    Dim loHeader As ListObject
    Dim loLines As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    Set loLines = GetRecipeBuilderLinesTable(wsProd, loHeader)
    If loLines Is Nothing Then Set loLines = EnsureRecipeBuilderLinesTable(wsProd, loHeader)
    If loHeader Is Nothing Or loLines Is Nothing Then
        MsgBox "Recipe Builder tables not found on Production sheet.", vbExclamation
        Exit Sub
    End If
    Dim nameCell As Range: Set nameCell = GetHeaderDataCell(loHeader, "RECIPE_NAME")
    If nameCell Is Nothing Then
        MsgBox "Recipe Builder header missing RECIPE_NAME column.", vbCritical
        Exit Sub
    End If
    Dim useActiveBuilderLines As Boolean
    Dim processTables As Collection
    Dim sourceTables As Collection
    Set sourceTables = BuildRecipeSaveSourceTables(wsProd, loLines, useActiveBuilderLines, processTables)

    If sourceTables.count = 0 Then
        If loLines.DataBodyRange Is Nothing Then
            MsgBox "Add at least one recipe line before saving.", vbExclamation
            Exit Sub
        End If
        sourceTables.Add loLines
    End If

    Dim cDesc As Long: cDesc = ColumnIndex(loHeader, "DESCRIPTION")
    Dim cBudget As Long: cBudget = ColumnIndex(loHeader, "ROW_BUDGET")
    Dim cGuid As Long: cGuid = ColumnIndex(loHeader, "GUID")
    Dim cRecipeId As Long: cRecipeId = ColumnIndex(loHeader, "RECIPE_ID")
    If cRecipeId = 0 Then
        MsgBox "Recipe Builder header missing RECIPE_NAME or RECIPE_ID.", vbCritical
        Exit Sub
    End If

    Dim recipeName As String: recipeName = NzStr(nameCell.value)
    Dim recipeDesc As String
    If cDesc > 0 Then
        Dim descCell As Range
        Set descCell = GetHeaderDataCell(loHeader, "DESCRIPTION")
        If Not descCell Is Nothing Then recipeDesc = NzStr(descCell.value)
    End If
    If Trim$(recipeName) = "" Then
        MsgBox "Fill RB_AddRecipeName (RECIPE_NAME) or load a recipe before saving.", vbExclamation
        Exit Sub
    End If
    Dim rowBudget As Long
    If cBudget > 0 Then rowBudget = NormalizeProductionRowBudget(loHeader.DataBodyRange.Cells(1, cBudget).Value)
    If rowBudget <= 0 Then rowBudget = PRODUCTION_DEFAULT_ROW_BUDGET
    If cBudget > 0 Then loHeader.DataBodyRange.Cells(1, cBudget).Value = rowBudget

    Dim recipeIdCell As Range: Set recipeIdCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
    Dim recipeId As String: recipeId = NzStr(recipeIdCell.value)
    If recipeId = "" Then
        recipeId = GenerateRecipeId(wsProd.Parent)
        WriteProductionTextCell recipeIdCell, recipeId
    End If
    designsEnabled = ProductionDesignsEnabled()
    If RecipeIdExistsProduction(recipeId, wsProd.Parent) Then
        Dim collisionChoice As VbMsgBoxResult
        Dim collisionMessage As String

        collisionMessage = "Recipe ID '" & CanonicalRecipeIdProduction(recipeId) & "' is already in use." & vbCrLf & vbCrLf
        If designsEnabled Then
            collisionMessage = collisionMessage & _
                "Yes: save a new immutable version under this existing Recipe ID." & vbCrLf
        Else
            collisionMessage = collisionMessage & _
                "Yes: replace the existing legacy recipe under this Recipe ID." & vbCrLf
        End If
        collisionMessage = collisionMessage & _
            "No: assign the next available unused Base-36 ID and save as a new recipe." & vbCrLf & _
            "Cancel: do not save."
        collisionChoice = MsgBox(collisionMessage, vbQuestion + vbYesNoCancel, "Recipe ID Already In Use")
        If collisionChoice = vbCancel Then Exit Sub
        If collisionChoice = vbNo Then
            recipeId = GenerateRecipeId(wsProd.Parent)
            If recipeId = "" Then
                MsgBox "No unused Base-36 Recipe ID is available.", vbExclamation, "Save Recipe"
                Exit Sub
            End If
            WriteProductionTextCell recipeIdCell, recipeId
            If cGuid > 0 Then
                Dim newRecipeGuidCell As Range
                Set newRecipeGuidCell = GetHeaderDataCell(loHeader, "GUID")
                If Not newRecipeGuidCell Is Nothing Then newRecipeGuidCell.ClearContents
            End If
        End If
    End If
    If cGuid > 0 Then
        Dim recipeGuidCell As Range: Set recipeGuidCell = GetHeaderDataCell(loHeader, "GUID")
        Dim recipeGuid As String: recipeGuid = NzStr(recipeGuidCell.value)
        If recipeGuid = "" Then
            recipeGuid = CreateProductionGuid()
            recipeGuidCell.value = recipeGuid
        End If
    End If

    If designsEnabled Then
        Set stagingWb = CreateEmptyDesignRecipeStagingWorkbook(stagingReport)
        If stagingWb Is Nothing Then
            MsgBox stagingReport, vbExclamation, "Production Designs"
            Exit Sub
        End If
        Set wsRec = WorkbookSheetExists(stagingWb, "Recipes")
    Else
        Set wsRec = SheetExists("Recipes")
    End If
    If wsRec Is Nothing Then
        MsgBox "Recipes sheet not found.", vbCritical
        GoTo CleanExit
    End If

    Dim loRecipes As ListObject: Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then
        MsgBox "Recipes table not found on Recipes sheet.", vbCritical
        GoTo CleanExit
    End If
    Dim cRecRecipeId As Long: cRecRecipeId = ColumnIndex(loRecipes, "RECIPE_ID")
    If cRecRecipeId = 0 Then
        MsgBox "Recipes table missing RECIPE_ID column.", vbCritical
        GoTo CleanExit
    End If

    ' Delete existing rows for this recipe ID (overwrite behavior).
    If Not loRecipes.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = loRecipes.DataBodyRange.rows.count To 1 Step -1
            If NzStr(loRecipes.DataBodyRange.Cells(r, cRecRecipeId).value) = recipeId Then
                loRecipes.ListRows(r).Delete
            End If
        Next r
    End If

    ' Column indexes in Recipes table.
    Dim cRecRecipe As Long: cRecRecipe = ColumnIndex(loRecipes, "RECIPE")
    Dim cRecDesc As Long: cRecDesc = ColumnIndex(loRecipes, "DESCRIPTION")
    Dim cRecBudget As Long: cRecBudget = ColumnIndex(loRecipes, "ROW_BUDGET")
    Dim cRecDept As Long: cRecDept = ColumnIndex(loRecipes, "DEPARTMENT")
    Dim cRecProcess As Long: cRecProcess = ColumnIndex(loRecipes, "PROCESS")
    Dim cRecDiagram As Long: cRecDiagram = ColumnIndex(loRecipes, "DIAGRAM_ID")
    Dim cRecIO As Long: cRecIO = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    Dim cRecIngredient As Long: cRecIngredient = ColumnIndex(loRecipes, "INGREDIENT")
    Dim cRecPercent As Long: cRecPercent = ColumnIndex(loRecipes, "PERCENT")
    Dim cRecUom As Long: cRecUom = ColumnIndex(loRecipes, "UOM")
    Dim cRecAmount As Long: cRecAmount = ColumnIndex(loRecipes, "AMOUNT")
    Dim cRecListRow As Long: cRecListRow = ColumnIndex(loRecipes, "RECIPE_LIST_ROW")
    Dim cRecIngId As Long: cRecIngId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    Dim cRecGuid As Long: cRecGuid = ColumnIndex(loRecipes, "GUID")

    Dim savedCount As Long
    Dim seqRow As Long: seqRow = 1
    Dim src As Variant
    For Each src In sourceTables
        AppendRecipeRowsFromTable src, recipeId, recipeName, recipeDesc, loRecipes, _
            cRecRecipeId, cRecRecipe, cRecDesc, cRecBudget, rowBudget, cRecDept, cRecProcess, cRecDiagram, cRecIO, _
            cRecIngredient, cRecPercent, cRecUom, cRecAmount, cRecListRow, cRecIngId, cRecGuid, _
            seqRow, savedCount
    Next src

    Dim templateCount As Long
    If Not processTables Is Nothing Then
        If processTables.count > 0 Then templateCount = RegisterRecipeTemplates(recipeId, processTables)
    End If

    If savedCount = 0 Then
        MsgBox "No recipe lines with data were found to save.", vbExclamation
    Else
        Dim msg As String
        msg = "Saved recipe '" & recipeName & "' (" & savedCount & " lines)."
        If useActiveBuilderLines Then
            templateCount = BuildRecipeProcessTablesFromLines(recipeId, True, Not IsRecipeLinesStaged(loLines))
        End If
        If templateCount > 0 Then msg = msg & vbCrLf & "Templates saved: " & templateCount & "."

        Dim runtimeReport As String
        Dim runtimeSaved As Boolean
        Dim designVersion As String
        If designsEnabled Then
            runtimeSaved = QueueSavedRecipeDesignCreate(loRecipes, recipeId, recipeName, recipeDesc, designVersion, runtimeReport)
            If runtimeSaved Then
                msg = msg & vbCrLf & "Designs Domain version " & designVersion & ": " & runtimeReport
            Else
                msg = msg & vbCrLf & "Designs Domain save did not complete: " & runtimeReport
            End If
        Else
            runtimeSaved = PublishProductionRecipeRowsToRuntime(wsProd.Parent, loRecipes, recipeId, runtimeReport)
            If runtimeSaved Then
                msg = msg & vbCrLf & "Legacy server recipe saved (Designs disabled)."
            ElseIf Trim$(runtimeReport) <> "" Then
                msg = msg & vbCrLf & "Legacy server save did not complete: " & runtimeReport
            End If
        End If
        MsgBox msg, vbInformation
    End If
CleanExit:
    On Error Resume Next
    If Not stagingWb Is Nothing Then stagingWb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub
ErrHandler:
    MsgBox "Save Recipe failed: " & Err.description, vbCritical
    Resume CleanExit
End Sub

Private Function BuildRecipeSaveSourceTables(ByVal wsProd As Worksheet, ByVal loLines As ListObject, _
                                             ByRef useActiveBuilderLines As Boolean, _
                                             ByRef processTables As Collection) As Collection
    Dim sourceTables As New Collection

    useActiveBuilderLines = RecipeBuilderLinesHaveData(loLines)
    Set processTables = GetRecipeBuilderProcessTables(wsProd)

    If useActiveBuilderLines Then
        sourceTables.Add loLines
    ElseIf Not processTables Is Nothing Then
        Dim loProc As ListObject
        For Each loProc In processTables
            If Not loProc.DataBodyRange Is Nothing Then sourceTables.Add loProc
        Next loProc
    End If

    Set BuildRecipeSaveSourceTables = sourceTables
End Function

' System 1: Recipe List Builder - load recipe into builder tables.
Public Sub LoadRecipeFromRecipes(Optional ByVal forceRecipeId As String = "")
    On Error GoTo ErrHandler
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Sub
    Dim syncReport As String
    Dim stagingWb As Workbook
    Dim wsRec As Worksheet

    Dim loHeader As ListObject
    Dim loLines As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    Set loLines = GetRecipeBuilderLinesTable(wsProd, loHeader)
    If loLines Is Nothing Then Set loLines = EnsureRecipeBuilderLinesTable(wsProd, loHeader)
    If loHeader Is Nothing Or loLines Is Nothing Then
        MsgBox "Recipe Builder tables not found on Production sheet.", vbExclamation
        Exit Sub
    End If

    Dim recipeId As String
    Dim recipeName As String
    recipeId = forceRecipeId

    If recipeId = "" Then
        Dim loSel As ListObject
        On Error Resume Next
        Set loSel = Application.ActiveCell.ListObject
        On Error GoTo 0
        If Not loSel Is Nothing Then
            Dim cSelRecipeId As Long: cSelRecipeId = ColumnIndex(loSel, "RECIPE_ID")
            Dim cSelRecipe As Long: cSelRecipe = ColumnIndex(loSel, "RECIPE")
            If cSelRecipeId > 0 Then
                recipeId = NzStr(loSel.DataBodyRange.Cells(Application.ActiveCell.row - loSel.DataBodyRange.row + 1, cSelRecipeId).value)
            End If
            If recipeId = "" And cSelRecipe > 0 Then
                recipeName = NzStr(loSel.DataBodyRange.Cells(Application.ActiveCell.row - loSel.DataBodyRange.row + 1, cSelRecipe).value)
            End If
        End If
    End If

    If recipeId = "" Then
        Dim cHeaderRecipeIdTmp As Long: cHeaderRecipeIdTmp = ColumnIndex(loHeader, "RECIPE_ID")
    If cHeaderRecipeIdTmp > 0 Then
        Dim hdrRecipeIdCell As Range: Set hdrRecipeIdCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
        If Not hdrRecipeIdCell Is Nothing Then recipeId = NzStr(hdrRecipeIdCell.value)
    End If
    End If

    If recipeId = "" And recipeName = "" Then
        recipeId = InputBox("Enter RECIPE_ID to load:", "Load Recipe")
    End If

    If recipeId = "" And recipeName = "" Then Exit Sub

    If ProductionDesignsEnabled() And recipeId <> "" Then
        Set stagingWb = BuildDesignRecipeStagingWorkbook(recipeId, "", syncReport)
        If stagingWb Is Nothing Then
            Set stagingWb = BuildPendingDesignRecipeStagingWorkbook(recipeId, syncReport)
        End If
        If stagingWb Is Nothing Then
            Set stagingWb = BuildLegacyRuntimeRecipeStagingWorkbook(recipeId, syncReport)
        End If
        If stagingWb Is Nothing Then
            MsgBox syncReport, vbExclamation, "Production Designs"
            Exit Sub
        End If
        Set wsRec = WorkbookSheetExists(stagingWb, "Recipes")
    Else
        Set wsRec = SheetExists("Recipes")
        If recipeId <> "" Then
            If Not LocalProductionRecipeRowsExist(wsProd.Parent, recipeId) Then
                RefreshProductionRecipesFromRuntime wsProd.Parent, syncReport
            End If
        ElseIf Not LocalProductionRecipeRowsExist(wsProd.Parent) Then
            RefreshProductionRecipesFromRuntime wsProd.Parent, syncReport
        End If
    End If
    If wsRec Is Nothing Then
        MsgBox "Recipes sheet not found.", vbCritical
        GoTo CleanExit
    End If

    Dim loRecipes As ListObject: Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then
        MsgBox "Recipes table not found on Recipes sheet.", vbCritical
        GoTo CleanExit
    End If

    Dim cRecRecipeId As Long: cRecRecipeId = ColumnIndex(loRecipes, "RECIPE_ID")
    Dim cRecRecipe As Long: cRecRecipe = ColumnIndex(loRecipes, "RECIPE")
    Dim cRecDesc As Long: cRecDesc = ColumnIndex(loRecipes, "DESCRIPTION")
    Dim cRecBudget As Long: cRecBudget = ColumnIndex(loRecipes, "ROW_BUDGET")
    Dim cRecProcess As Long: cRecProcess = ColumnIndex(loRecipes, "PROCESS")
    Dim cRecDiagram As Long: cRecDiagram = ColumnIndex(loRecipes, "DIAGRAM_ID")
    Dim cRecIO As Long: cRecIO = ColumnIndex(loRecipes, "INPUT/OUTPUT")
    Dim cRecIngredient As Long: cRecIngredient = ColumnIndex(loRecipes, "INGREDIENT")
    Dim cRecPercent As Long: cRecPercent = ColumnIndex(loRecipes, "PERCENT")
    Dim cRecUom As Long: cRecUom = ColumnIndex(loRecipes, "UOM")
    Dim cRecAmount As Long: cRecAmount = ColumnIndex(loRecipes, "AMOUNT")
    Dim cRecListRow As Long: cRecListRow = ColumnIndex(loRecipes, "RECIPE_LIST_ROW")
    Dim cRecIngId As Long: cRecIngId = ColumnIndex(loRecipes, "INGREDIENT_ID")
    Dim cRecGuid As Long: cRecGuid = ColumnIndex(loRecipes, "GUID")

    Dim matches As Collection: Set matches = New Collection
    If Not loRecipes.DataBodyRange Is Nothing Then
        Dim r As Long
        For r = 1 To loRecipes.DataBodyRange.rows.count
            Dim rowRecipeId As String
            rowRecipeId = NzStr(loRecipes.DataBodyRange.Cells(r, cRecRecipeId).value)
            Dim rowRecipeName As String
            If cRecRecipe > 0 Then rowRecipeName = NzStr(loRecipes.DataBodyRange.Cells(r, cRecRecipe).value)
            If (recipeId <> "" And rowRecipeId = recipeId) Or (recipeId = "" And rowRecipeName = recipeName And rowRecipeName <> "") Then
                matches.Add r
                If recipeId = "" Then recipeId = rowRecipeId
                If recipeName = "" Then recipeName = rowRecipeName
            End If
        Next r
    End If

    If matches.count = 0 Then
        MsgBox "No recipe rows found for the selected RECIPE_ID.", vbExclamation
        GoTo CleanExit
    End If

    ' Update header table.
    Dim cHeaderName As Long: cHeaderName = ColumnIndex(loHeader, "RECIPE_NAME")
    Dim cHeaderDesc As Long: cHeaderDesc = ColumnIndex(loHeader, "DESCRIPTION")
    Dim cHeaderBudget As Long: cHeaderBudget = ColumnIndex(loHeader, "ROW_BUDGET")
    Dim cHeaderGuid As Long: cHeaderGuid = ColumnIndex(loHeader, "GUID")
    Dim cHeaderRecipeId As Long: cHeaderRecipeId = ColumnIndex(loHeader, "RECIPE_ID")
    Dim hdrNameCell As Range: Set hdrNameCell = GetHeaderDataCell(loHeader, "RECIPE_NAME")
    Dim hdrIdCell As Range: Set hdrIdCell = GetHeaderDataCell(loHeader, "RECIPE_ID")
    Dim hdrDescCell As Range: Set hdrDescCell = GetHeaderDataCell(loHeader, "DESCRIPTION")
    Dim hdrBudgetCell As Range: Set hdrBudgetCell = GetHeaderDataCell(loHeader, "ROW_BUDGET")
    Dim hdrGuidCell As Range: Set hdrGuidCell = GetHeaderDataCell(loHeader, "GUID")
    If Not hdrNameCell Is Nothing Then hdrNameCell.value = recipeName
    If Not hdrIdCell Is Nothing Then WriteProductionTextCell hdrIdCell, recipeId
    If Not hdrDescCell Is Nothing And cRecDesc > 0 Then
        hdrDescCell.value = NzStr(loRecipes.DataBodyRange.Cells(matches(1), cRecDesc).value)
    End If
    If Not hdrBudgetCell Is Nothing Then
        If cRecBudget > 0 Then
            hdrBudgetCell.value = NormalizeProductionRowBudget(loRecipes.DataBodyRange.Cells(matches(1), cRecBudget).value)
        Else
            hdrBudgetCell.value = PRODUCTION_DEFAULT_ROW_BUDGET
        End If
    End If
    If Not hdrGuidCell Is Nothing And cRecGuid > 0 Then
        hdrGuidCell.value = NzStr(loRecipes.DataBodyRange.Cells(matches(1), cRecGuid).value)
    End If

    ' Clear and rebuild RecipeBuilder lines.
    ClearListObjectData loLines
    Dim idx As Long
    For idx = 1 To matches.count
        Dim rr As Long: rr = CLng(matches(idx))
        Dim lr As ListRow: Set lr = loLines.ListRows.Add
        Dim cProc As Long: cProc = ColumnIndex(loLines, "PROCESS")
        Dim cDiag As Long: cDiag = ColumnIndex(loLines, "DIAGRAM_ID")
        Dim cIO As Long: cIO = ColumnIndex(loLines, "INPUT/OUTPUT")
        Dim cIng As Long: cIng = ColumnIndex(loLines, "INGREDIENT")
        Dim cPct As Long: cPct = ColumnIndex(loLines, "PERCENT")
        Dim cUomLine As Long: cUomLine = ColumnIndex(loLines, "UOM")
        Dim cAmt As Long: cAmt = ColumnIndex(loLines, "AMOUNT")
        Dim cListRow As Long: cListRow = ColumnIndex(loLines, "RECIPE_LIST_ROW")
        Dim cIngId As Long: cIngId = ColumnIndex(loLines, "INGREDIENT_ID")
        Dim cGuidLine As Long: cGuidLine = ColumnIndex(loLines, "GUID")

        If cProc > 0 Then lr.Range.Cells(1, cProc).value = loRecipes.DataBodyRange.Cells(rr, cRecProcess).value
        If cDiag > 0 Then lr.Range.Cells(1, cDiag).value = loRecipes.DataBodyRange.Cells(rr, cRecDiagram).value
        If cIO > 0 Then lr.Range.Cells(1, cIO).value = loRecipes.DataBodyRange.Cells(rr, cRecIO).value
        If cIng > 0 Then lr.Range.Cells(1, cIng).value = loRecipes.DataBodyRange.Cells(rr, cRecIngredient).value
        If cPct > 0 Then lr.Range.Cells(1, cPct).value = loRecipes.DataBodyRange.Cells(rr, cRecPercent).value
        If cUomLine > 0 Then lr.Range.Cells(1, cUomLine).value = loRecipes.DataBodyRange.Cells(rr, cRecUom).value
        If cAmt > 0 Then lr.Range.Cells(1, cAmt).value = loRecipes.DataBodyRange.Cells(rr, cRecAmount).value
        If cListRow > 0 Then lr.Range.Cells(1, cListRow).value = loRecipes.DataBodyRange.Cells(rr, cRecListRow).value
        If cIngId > 0 Then lr.Range.Cells(1, cIngId).value = loRecipes.DataBodyRange.Cells(rr, cRecIngId).value
        If cGuidLine > 0 Then lr.Range.Cells(1, cGuidLine).value = loRecipes.DataBodyRange.Cells(rr, cRecGuid).value
    Next idx

    Dim procCount As Long
    Dim hasProc As Boolean
    hasProc = RecipeLinesHasProcess(loLines)
    If hasProc Then
        Dim staged As Boolean
        staged = MoveRecipeBuilderLinesToStaging(loLines)
        procCount = BuildRecipeProcessTablesFromLines(recipeId, True, Not staged)
        ' Keep RecipeBuilder lines table staged until Clear Recipe List Builder.
    End If

    Dim loadMsg As String
    loadMsg = "Loaded recipe '" & recipeName & "' (" & matches.count & " lines)."
    If procCount > 0 Then loadMsg = loadMsg & vbCrLf & "Process tables built: " & procCount & "."
    MsgBox loadMsg, vbInformation
CleanExit:
    On Error Resume Next
    If Not stagingWb Is Nothing Then stagingWb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Sub
ErrHandler:
    MsgBox "Load Recipe failed: " & Err.description, vbCritical
    Resume CleanExit
End Sub

' System 1: Recipe List Builder - write recipe rows to Recipes table.
Private Sub AppendRecipeRowsFromTable(ByVal loSource As ListObject, ByVal recipeId As String, _
    ByVal recipeName As String, ByVal recipeDesc As String, ByVal loRecipes As ListObject, _
    ByVal cRecRecipeId As Long, ByVal cRecRecipe As Long, ByVal cRecDesc As Long, _
    ByVal cRecBudget As Long, ByVal rowBudget As Long, ByVal cRecDept As Long, _
    ByVal cRecProcess As Long, ByVal cRecDiagram As Long, ByVal cRecIO As Long, ByVal cRecIngredient As Long, _
    ByVal cRecPercent As Long, ByVal cRecUom As Long, ByVal cRecAmount As Long, ByVal cRecListRow As Long, _
    ByVal cRecIngId As Long, ByVal cRecGuid As Long, ByRef seqRow As Long, ByRef savedCount As Long)

    If loSource Is Nothing Then Exit Sub
    If loSource.DataBodyRange Is Nothing Then Exit Sub

    Dim cProc As Long: cProc = ColumnIndex(loSource, "PROCESS")
    Dim cDiag As Long: cDiag = ColumnIndex(loSource, "DIAGRAM_ID")
    Dim cIO As Long: cIO = ColumnIndex(loSource, "INPUT/OUTPUT")
    Dim cIng As Long: cIng = ColumnIndex(loSource, "INGREDIENT")
    Dim cPct As Long: cPct = ColumnIndex(loSource, "PERCENT")
    Dim cUomLine As Long: cUomLine = ColumnIndex(loSource, "UOM")
    Dim cAmt As Long: cAmt = ColumnIndex(loSource, "AMOUNT")
    Dim cListRow As Long: cListRow = ColumnIndex(loSource, "RECIPE_LIST_ROW")
    Dim cIngId As Long: cIngId = ColumnIndex(loSource, "INGREDIENT_ID")
    Dim cGuidLine As Long: cGuidLine = ColumnIndex(loSource, "GUID")

    Dim lineArr As Variant: lineArr = loSource.DataBodyRange.value
    Dim rowCount As Long: rowCount = UBound(lineArr, 1)
    Dim processFallback As String: processFallback = ProcessNameFromTable(loSource)

    Dim i As Long
    For i = 1 To rowCount
        Dim hasData As Boolean
        If cIng > 0 Then
            hasData = (Trim$(NzStr(lineArr(i, cIng))) <> "")
        ElseIf cProc > 0 Then
            hasData = (Trim$(NzStr(lineArr(i, cProc))) <> "")
        End If
        If Not hasData Then GoTo NextLine

        Dim processVal As String
        If cProc > 0 Then processVal = NzStr(lineArr(i, cProc))
        If processVal = "" Then processVal = processFallback

        Dim ingId As String
        Dim ioType As String
        If cIO > 0 Then ioType = UCase$(Trim$(NzStr(lineArr(i, cIO))))
        If cIngId > 0 Then ingId = NzStr(lineArr(i, cIngId))
        If ioType = "INSTRUCTION" Then
            ingId = vbNullString
            If cIngId > 0 Then loSource.DataBodyRange.Cells(i, cIngId).Value = vbNullString
        ElseIf ingId = "" Then
            ingId = CreateProductionGuid()
            If cIngId > 0 Then loSource.DataBodyRange.Cells(i, cIngId).value = ingId
        End If

        Dim recListRow As Variant
        If cListRow > 0 Then recListRow = lineArr(i, cListRow)
        If NzStr(recListRow) = "" Then
            recListRow = seqRow
            If cListRow > 0 Then loSource.DataBodyRange.Cells(i, cListRow).value = recListRow
        End If

        Dim rowGuid As String
        If cGuidLine > 0 Then rowGuid = NzStr(lineArr(i, cGuidLine))
        If rowGuid = "" Then
            rowGuid = CreateProductionGuid()
            If cGuidLine > 0 Then loSource.DataBodyRange.Cells(i, cGuidLine).value = rowGuid
        End If

        Dim lr As ListRow: Set lr = loRecipes.ListRows.Add
        If cRecRecipeId > 0 Then WriteProductionTextCell lr.Range.Cells(1, cRecRecipeId), recipeId
        If cRecRecipe > 0 Then lr.Range.Cells(1, cRecRecipe).value = recipeName
        If cRecDesc > 0 Then lr.Range.Cells(1, cRecDesc).value = recipeDesc
        If cRecBudget > 0 Then lr.Range.Cells(1, cRecBudget).value = rowBudget
        If cRecDept > 0 Then lr.Range.Cells(1, cRecDept).value = "" ' optional for now
        If cRecProcess > 0 Then lr.Range.Cells(1, cRecProcess).value = processVal
        If cRecDiagram > 0 And cDiag > 0 Then lr.Range.Cells(1, cRecDiagram).value = lineArr(i, cDiag)
        If cRecIO > 0 And cIO > 0 Then lr.Range.Cells(1, cRecIO).value = lineArr(i, cIO)
        If cRecIngredient > 0 And cIng > 0 Then lr.Range.Cells(1, cRecIngredient).value = lineArr(i, cIng)
        If cRecPercent > 0 And cPct > 0 Then lr.Range.Cells(1, cRecPercent).value = lineArr(i, cPct)
        If cRecUom > 0 And cUomLine > 0 Then lr.Range.Cells(1, cRecUom).value = lineArr(i, cUomLine)
        If cRecAmount > 0 And cAmt > 0 Then lr.Range.Cells(1, cRecAmount).value = lineArr(i, cAmt)
        If cRecListRow > 0 Then lr.Range.Cells(1, cRecListRow).value = recListRow
        If cRecIngId > 0 Then lr.Range.Cells(1, cRecIngId).value = ingId
        If cRecGuid > 0 Then lr.Range.Cells(1, cRecGuid).value = rowGuid

        savedCount = savedCount + 1
        seqRow = seqRow + 1
NextLine:
    Next i
End Sub

Public Function RefreshProductionRecipesFromRuntime(Optional ByVal operatorWb As Workbook = Nothing, _
                                                    Optional ByRef report As String = "") As Boolean
    On Error GoTo FailSoft

    Dim wbOps As Workbook
    Dim surfaceReport As String
    Dim wsRec As Worksheet
    Dim loLocal As ListObject
    Dim wbRuntime As Workbook
    Dim loRuntime As ListObject
    Dim warehouseId As String
    Dim rootPath As String
    Dim openedTransient As Boolean

    Set wbOps = ResolveProductionWorkbook(operatorWb, "Recipes")
    If wbOps Is Nothing Then Set wbOps = ResolveProductionWorkbook(operatorWb, SHEET_PRODUCTION)
    If wbOps Is Nothing Then
        report = "No Production operator workbook was available."
        Exit Function
    End If

    If Not modOperationsPrimitiveBridge.EnsureProductionWorkbookSurface(wbOps.Name, surfaceReport) Then
        report = "Production surface repair failed: " & surfaceReport
        Exit Function
    End If

    Set wsRec = WorkbookSheetExists(wbOps, "Recipes")
    If wsRec Is Nothing Then
        report = "Recipes sheet was not found in the operator workbook."
        Exit Function
    End If
    Set loLocal = GetListObject(wsRec, "Recipes")
    If loLocal Is Nothing Then
        report = "Recipes table was not found in the operator workbook."
        Exit Function
    End If

    If Not ResolveProductionRecipesStorageTarget(warehouseId, rootPath, report) Then Exit Function

    Set wbRuntime = OpenProductionRecipesWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbRuntime Is Nothing Then GoTo CleanExit

    Set loRuntime = EnsureProductionRecipesSchema(wbRuntime, report)
    If loRuntime Is Nothing Then GoTo CleanExit
    If loRuntime.DataBodyRange Is Nothing Then
        report = "Production recipes runtime workbook has no saved recipe rows."
        RefreshProductionRecipesFromRuntime = True
        GoTo CleanExit
    End If

    MergeProductionRecipeRuntimeRowsToLocal loRuntime, loLocal
    RefreshProductionRecipesFromRuntime = True
    report = "Production recipes refreshed from " & wbRuntime.FullName

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveProduction wbRuntime
    Exit Function

FailSoft:
    report = "RefreshProductionRecipesFromRuntime failed: " & Err.Description
    Resume CleanExit
End Function

Public Function RefreshProductionRecipesFromRuntimeForCurrentWorkbook() As String
    Dim report As String
    If RefreshProductionRecipesFromRuntime(ResolveProductionWorkbook(, "Recipes"), report) Then
        RefreshProductionRecipesFromRuntimeForCurrentWorkbook = "OK: " & report
    Else
        RefreshProductionRecipesFromRuntimeForCurrentWorkbook = "FAILED: " & report
    End If
End Function

Private Function LocalProductionRecipeRowsExist(Optional ByVal operatorWb As Workbook = Nothing, _
                                                Optional ByVal recipeId As String = "") As Boolean
    Dim wbOps As Workbook
    Dim wsRec As Worksheet
    Dim loRecipes As ListObject
    Dim cRecipeId As Long
    Dim r As Long

    Set wbOps = ResolveProductionWorkbook(operatorWb, "Recipes")
    If wbOps Is Nothing Then Set wbOps = ResolveProductionWorkbook(operatorWb, SHEET_PRODUCTION)
    If wbOps Is Nothing Then Exit Function

    Set wsRec = WorkbookSheetExists(wbOps, "Recipes")
    If wsRec Is Nothing Then Exit Function
    Set loRecipes = GetListObject(wsRec, "Recipes")
    If loRecipes Is Nothing Then Exit Function
    If loRecipes.DataBodyRange Is Nothing Then Exit Function

    recipeId = Trim$(recipeId)
    If recipeId = "" Then
        LocalProductionRecipeRowsExist = True
        Exit Function
    End If

    cRecipeId = ColumnIndex(loRecipes, "RECIPE_ID")
    If cRecipeId = 0 Then Exit Function
    For r = 1 To loRecipes.DataBodyRange.Rows.Count
        If StrComp(NzStr(loRecipes.DataBodyRange.Cells(r, cRecipeId).Value), recipeId, vbTextCompare) = 0 Then
            LocalProductionRecipeRowsExist = True
            Exit Function
        End If
    Next r
End Function

'@TestOnlyBegin
Public Function TestProductionRecipesRuntimeRoundTrip(ByVal runtimeRoot As String) As String
    On Error GoTo FailSoft

    Dim priorRoot As String
    Dim wbOps As Workbook
    Dim wsRec As Worksheet
    Dim loLocal As ListObject
    Dim lr As ListRow
    Dim report As String
    Dim ok As Boolean
    Dim cRecipeId As Long
    Dim r As Long
    Dim found As Boolean

    priorRoot = modRuntimeWorkbooks.GetCoreDataRootOverride()
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modOperationsPrimitiveBridge.EnsureProductionWorkbookSurface(wbOps.Name, report) Then
        TestProductionRecipesRuntimeRoundTrip = "FAILED: surface: " & report
        GoTo CleanExit
    End If

    Set wsRec = WorkbookSheetExists(wbOps, "Recipes")
    Set loLocal = GetListObject(wsRec, "Recipes")
    ClearListObjectData loLocal

    Set lr = loLocal.ListRows.Add
    SetProductionTableCellByHeader loLocal, lr.Index, "RECIPE", "Round Trip Recipe"
    SetProductionTableCellByHeader loLocal, lr.Index, "RECIPE_ID", "TST"
    SetProductionTableCellByHeader loLocal, lr.Index, "DESCRIPTION", "Runtime persistence test"
    SetProductionTableCellByHeader loLocal, lr.Index, "PROCESS", "MIX"
    SetProductionTableCellByHeader loLocal, lr.Index, "INPUT/OUTPUT", "INPUT"
    SetProductionTableCellByHeader loLocal, lr.Index, "INGREDIENT", "Apple Juice"
    SetProductionTableCellByHeader loLocal, lr.Index, "PERCENT", 100
    SetProductionTableCellByHeader loLocal, lr.Index, "UOM", "GAL"
    SetProductionTableCellByHeader loLocal, lr.Index, "AMOUNT", 1
    SetProductionTableCellByHeader loLocal, lr.Index, "RECIPE_LIST_ROW", 1
    SetProductionTableCellByHeader loLocal, lr.Index, "INGREDIENT_ID", "ING-TST"
    SetProductionTableCellByHeader loLocal, lr.Index, "GUID", "GUID-TST"

    ok = PublishProductionRecipeRowsToRuntime(wbOps, loLocal, "TST", report)
    If Not ok Then
        TestProductionRecipesRuntimeRoundTrip = "FAILED: publish: " & report
        GoTo CleanExit
    End If

    ClearListObjectData loLocal
    ok = RefreshProductionRecipesFromRuntime(wbOps, report)
    If Not ok Then
        TestProductionRecipesRuntimeRoundTrip = "FAILED: refresh: " & report
        GoTo CleanExit
    End If

    cRecipeId = ColumnIndex(loLocal, "RECIPE_ID")
    If Not loLocal.DataBodyRange Is Nothing And cRecipeId > 0 Then
        For r = 1 To loLocal.DataBodyRange.Rows.Count
            If StrComp(NzStr(loLocal.DataBodyRange.Cells(r, cRecipeId).Value), "TST", vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next r
    End If

    If found Then
        TestProductionRecipesRuntimeRoundTrip = "OK"
    Else
        TestProductionRecipesRuntimeRoundTrip = "FAILED: refreshed recipe row was not found."
    End If

CleanExit:
    On Error Resume Next
    If Not wbOps Is Nothing Then wbOps.Close SaveChanges:=False
    modRuntimeWorkbooks.SetCoreDataRootOverride priorRoot
    On Error GoTo 0
    Exit Function

FailSoft:
    TestProductionRecipesRuntimeRoundTrip = "FAILED: " & Err.Description
    Resume CleanExit
End Function

Public Function TestProductionRecipesLocalRowsWinOverStaleRuntime(ByVal runtimeRoot As String) As String
    On Error GoTo FailSoft

    Dim priorRoot As String
    Dim wbOps As Workbook
    Dim wbRuntime As Workbook
    Dim wsRec As Worksheet
    Dim loLocal As ListObject
    Dim loRuntime As ListObject
    Dim lr As ListRow
    Dim report As String
    Dim openedTransient As Boolean
    Dim ingredients As Variant

    priorRoot = modRuntimeWorkbooks.GetCoreDataRootOverride()
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modOperationsPrimitiveBridge.EnsureProductionWorkbookSurface(wbOps.Name, report) Then
        TestProductionRecipesLocalRowsWinOverStaleRuntime = "FAILED: surface: " & report
        GoTo CleanExit
    End If
    wbOps.Activate

    Set wsRec = WorkbookSheetExists(wbOps, "Recipes")
    Set loLocal = GetListObject(wsRec, "Recipes")
    ClearListObjectData loLocal
    Set lr = loLocal.ListRows.Add
    SetProductionTableCellByHeader loLocal, lr.Index, "RECIPE", "Brewed Black Tea"
    SetProductionTableCellByHeader loLocal, lr.Index, "RECIPE_ID", "R-LOCAL"
    SetProductionTableCellByHeader loLocal, lr.Index, "DESCRIPTION", "local edited recipe"
    SetProductionTableCellByHeader loLocal, lr.Index, "PROCESS", "1"
    SetProductionTableCellByHeader loLocal, lr.Index, "INPUT/OUTPUT", "OUTPUT"
    SetProductionTableCellByHeader loLocal, lr.Index, "INGREDIENT", "Brew Black Tea"
    SetProductionTableCellByHeader loLocal, lr.Index, "UOM", "LBS"
    SetProductionTableCellByHeader loLocal, lr.Index, "AMOUNT", 400
    SetProductionTableCellByHeader loLocal, lr.Index, "RECIPE_LIST_ROW", 1
    SetProductionTableCellByHeader loLocal, lr.Index, "INGREDIENT_ID", "OUT-TEA"
    SetProductionTableCellByHeader loLocal, lr.Index, "GUID", "GUID-LOCAL"

    Set wbRuntime = OpenProductionRecipesWorkbook("WH1", runtimeRoot, True, openedTransient, report)
    If wbRuntime Is Nothing Then
        TestProductionRecipesLocalRowsWinOverStaleRuntime = "FAILED: runtime open: " & report
        GoTo CleanExit
    End If
    Set loRuntime = EnsureProductionRecipesSchema(wbRuntime, report)
    If loRuntime Is Nothing Then
        TestProductionRecipesLocalRowsWinOverStaleRuntime = "FAILED: runtime schema: " & report
        GoTo CleanExit
    End If
    ClearListObjectData loRuntime
    Set lr = loRuntime.ListRows.Add
    SetProductionTableCellByHeader loRuntime, lr.Index, "RECIPE", "Brewed Black Tea"
    SetProductionTableCellByHeader loRuntime, lr.Index, "RECIPE_ID", "R-LOCAL"
    SetProductionTableCellByHeader loRuntime, lr.Index, "DESCRIPTION", "stale runtime recipe"
    SetProductionTableCellByHeader loRuntime, lr.Index, "PROCESS", "2"
    SetProductionTableCellByHeader loRuntime, lr.Index, "INPUT/OUTPUT", "OUTPUT"
    SetProductionTableCellByHeader loRuntime, lr.Index, "INGREDIENT", "Brew Black Tea"
    SetProductionTableCellByHeader loRuntime, lr.Index, "UOM", "LBS"
    SetProductionTableCellByHeader loRuntime, lr.Index, "AMOUNT", 400
    SetProductionTableCellByHeader loRuntime, lr.Index, "RECIPE_LIST_ROW", 1
    SetProductionTableCellByHeader loRuntime, lr.Index, "INGREDIENT_ID", "OUT-TEA"
    SetProductionTableCellByHeader loRuntime, lr.Index, "GUID", "GUID-RUNTIME"
    wbRuntime.Save
    CloseWorkbookNoSaveProduction wbRuntime
    Set wbRuntime = Nothing
    wbOps.Activate

    ingredients = LoadIngredientListForRecipe("R-LOCAL")
    If IsEmpty(ingredients) Or Not IsArray(ingredients) Then
        TestProductionRecipesLocalRowsWinOverStaleRuntime = "FAILED: no ingredients returned."
        GoTo CleanExit
    End If
    If NzStr(ingredients(1, 4)) = "1" Then
        TestProductionRecipesLocalRowsWinOverStaleRuntime = "OK"
    Else
        TestProductionRecipesLocalRowsWinOverStaleRuntime = "FAILED: expected local PROCESS 1, got " & NzStr(ingredients(1, 4))
    End If

CleanExit:
    On Error Resume Next
    If Not wbRuntime Is Nothing Then CloseWorkbookNoSaveProduction wbRuntime
    If Not wbOps Is Nothing Then wbOps.Close SaveChanges:=False
    modRuntimeWorkbooks.SetCoreDataRootOverride priorRoot
    On Error GoTo 0
    Exit Function

FailSoft:
    TestProductionRecipesLocalRowsWinOverStaleRuntime = "FAILED: " & Err.Description
    Resume CleanExit
End Function

Public Function TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines(ByVal runtimeRoot As String) As String
    On Error GoTo FailSoft

    Dim priorRoot As String
    Dim wbOps As Workbook
    Dim wbFresh As Workbook
    Dim wsProd As Worksheet
    Dim wsRec As Worksheet
    Dim loHeader As ListObject
    Dim loLines As ListObject
    Dim loRecipes As ListObject
    Dim lr As ListRow
    Dim report As String
    Dim sourceTables As Collection
    Dim processTables As Collection
    Dim useActiveBuilderLines As Boolean
    Dim src As Variant
    Dim savedCount As Long
    Dim seqRow As Long
    Dim rowBudget As Long
    Dim ok As Boolean
    Dim ingredients As Variant

    priorRoot = modRuntimeWorkbooks.GetCoreDataRootOverride()
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    modNasConnection.ClearWarehouseTarget

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modOperationsPrimitiveBridge.EnsureProductionWorkbookSurface(wbOps.Name, report) Then
        TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "FAILED: surface: " & report
        GoTo CleanExit
    End If
    wbOps.Activate

    Set wsProd = WorkbookSheetExists(wbOps, SHEET_PRODUCTION)
    Set wsRec = WorkbookSheetExists(wbOps, "Recipes")
    Set loHeader = FindListObjectByNameOrHeaders(wsProd, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    Set loLines = GetRecipeBuilderLinesTable(wsProd, loHeader)
    If loLines Is Nothing Then Set loLines = EnsureRecipeBuilderLinesTable(wsProd, loHeader)
    Set loRecipes = GetListObject(wsRec, "Recipes")
    If wsProd Is Nothing Or wsRec Is Nothing Or loHeader Is Nothing Or loLines Is Nothing Or loRecipes Is Nothing Then
        TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "FAILED: production recipe test surface incomplete."
        GoTo CleanExit
    End If

    ClearListObjectData loHeader
    EnsureTableHasRow loHeader
    SetProductionTableCellByHeader loHeader, 1, "RECIPE_NAME", "Brewed Black Tea"
    SetProductionTableCellByHeader loHeader, 1, "RECIPE_ID", "R-PERSIST"
    SetProductionTableCellByHeader loHeader, 1, "DESCRIPTION", "loaded then edited"
    SetProductionTableCellByHeader loHeader, 1, "ROW_BUDGET", 50

    ClearListObjectData loLines
    Set lr = loLines.ListRows.Add(AlwaysInsert:=False)
    SetProductionTableCellByHeader loLines, lr.Index, "PROCESS", "2"
    SetProductionTableCellByHeader loLines, lr.Index, "INPUT/OUTPUT", "OUTPUT"
    SetProductionTableCellByHeader loLines, lr.Index, "INGREDIENT", "Brew Black Tea"
    SetProductionTableCellByHeader loLines, lr.Index, "UOM", "LBS"
    SetProductionTableCellByHeader loLines, lr.Index, "AMOUNT", 400
    SetProductionTableCellByHeader loLines, lr.Index, "RECIPE_LIST_ROW", 1
    SetProductionTableCellByHeader loLines, lr.Index, "INGREDIENT_ID", "OUT-TEA"
    SetProductionTableCellByHeader loLines, lr.Index, "GUID", "GUID-OLD"

    SetProductionTableCellByHeader loLines, 1, "PROCESS", "1"
    SetProductionTableCellByHeader loLines, 1, "GUID", "GUID-EDITED"

    Set sourceTables = BuildRecipeSaveSourceTables(wsProd, loLines, useActiveBuilderLines, processTables)
    If Not useActiveBuilderLines Then
        TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "FAILED: staged RecipeBuilder lines were not selected as the save source."
        GoTo CleanExit
    End If

    ClearListObjectData loRecipes
    rowBudget = 50
    seqRow = 1
    For Each src In sourceTables
        AppendRecipeRowsFromTable src, "R-PERSIST", "Brewed Black Tea", "loaded then edited", loRecipes, _
            ColumnIndex(loRecipes, "RECIPE_ID"), ColumnIndex(loRecipes, "RECIPE"), ColumnIndex(loRecipes, "DESCRIPTION"), _
            ColumnIndex(loRecipes, "ROW_BUDGET"), rowBudget, ColumnIndex(loRecipes, "DEPARTMENT"), _
            ColumnIndex(loRecipes, "PROCESS"), ColumnIndex(loRecipes, "DIAGRAM_ID"), ColumnIndex(loRecipes, "INPUT/OUTPUT"), _
            ColumnIndex(loRecipes, "INGREDIENT"), ColumnIndex(loRecipes, "PERCENT"), ColumnIndex(loRecipes, "UOM"), _
            ColumnIndex(loRecipes, "AMOUNT"), ColumnIndex(loRecipes, "RECIPE_LIST_ROW"), _
            ColumnIndex(loRecipes, "INGREDIENT_ID"), ColumnIndex(loRecipes, "GUID"), seqRow, savedCount
    Next src
    If savedCount <> 1 Then
        TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "FAILED: expected 1 saved recipe row, got " & CStr(savedCount)
        GoTo CleanExit
    End If
    If NzStr(loRecipes.DataBodyRange.Cells(1, ColumnIndex(loRecipes, "PROCESS")).Value) <> "1" Then
        TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "FAILED: local saved process was not edited builder value."
        GoTo CleanExit
    End If

    ok = PublishProductionRecipeRowsToRuntime(wbOps, loRecipes, "R-PERSIST", report)
    If Not ok Then
        TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "FAILED: publish: " & report
        GoTo CleanExit
    End If

    wbOps.Close SaveChanges:=False
    Set wbOps = Nothing

    Set wbFresh = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modOperationsPrimitiveBridge.EnsureProductionWorkbookSurface(wbFresh.Name, report) Then
        TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "FAILED: fresh surface: " & report
        GoTo CleanExit
    End If
    wbFresh.Activate

    ingredients = LoadIngredientListForRecipe("R-PERSIST")
    If IsEmpty(ingredients) Or Not IsArray(ingredients) Then
        TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "FAILED: no ingredients after reopen."
        GoTo CleanExit
    End If
    If NzStr(ingredients(1, 4)) = "1" Then
        TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "OK"
    Else
        TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "FAILED: expected reopened PROCESS 1, got " & NzStr(ingredients(1, 4))
    End If

CleanExit:
    On Error Resume Next
    If Not wbOps Is Nothing Then wbOps.Close SaveChanges:=False
    If Not wbFresh Is Nothing Then wbFresh.Close SaveChanges:=False
    modNasConnection.ClearWarehouseTarget
    modRuntimeWorkbooks.SetCoreDataRootOverride priorRoot
    On Error GoTo 0
    Exit Function

FailSoft:
    TestProductionRecipeBuilderSaveAfterLoadPersistsEditedLines = "FAILED: " & Err.Description
    Resume CleanExit
End Function
'@TestOnlyEnd

Private Function PublishProductionRecipeRowsToRuntime(ByVal operatorWb As Workbook, _
                                                      ByVal loLocalRecipes As ListObject, _
                                                      ByVal recipeId As String, _
                                                      ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim warehouseId As String
    Dim rootPath As String
    Dim wbRuntime As Workbook
    Dim loRuntime As ListObject
    Dim openedTransient As Boolean
    Dim copiedCount As Long

    recipeId = Trim$(recipeId)
    If recipeId = "" Then
        report = "RECIPE_ID is required before saving Production recipe to the server."
        Exit Function
    End If
    If loLocalRecipes Is Nothing Then
        report = "Local Recipes table is not available."
        Exit Function
    End If
    If loLocalRecipes.DataBodyRange Is Nothing Then
        report = "Local Recipes table has no rows to publish."
        Exit Function
    End If

    If Not ResolveProductionRecipesStorageTarget(warehouseId, rootPath, report) Then Exit Function

    Set wbRuntime = OpenProductionRecipesWorkbook(warehouseId, rootPath, True, openedTransient, report)
    If wbRuntime Is Nothing Then GoTo CleanExit
    If wbRuntime.ReadOnly Then
        report = "Production recipes runtime workbook is read-only: " & wbRuntime.FullName
        GoTo CleanExit
    End If

    Set loRuntime = EnsureProductionRecipesSchema(wbRuntime, report)
    If loRuntime Is Nothing Then GoTo CleanExit

    DeleteProductionRecipeRowsById loRuntime, recipeId
    copiedCount = CopyProductionRecipeRowsById(loLocalRecipes, loRuntime, recipeId, True)
    If copiedCount = 0 Then
        report = "No local recipe rows matched RECIPE_ID " & recipeId & "."
        GoTo CleanExit
    End If

    wbRuntime.Save
    PublishProductionRecipeRowsToRuntime = True
    report = "Production recipes runtime updated: " & wbRuntime.FullName & " (" & copiedCount & " rows)"

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveProduction wbRuntime
    Exit Function

FailSoft:
    report = "PublishProductionRecipeRowsToRuntime failed: " & Err.Description
    Resume CleanExit
End Function

Private Function ResolveProductionRecipesStorageTarget(ByRef warehouseId As String, _
                                                       ByRef rootPath As String, _
                                                       ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim overrideRoot As String
    Dim existingRoot As String

    warehouseId = CurrentProductionWarehouseId()
    rootPath = NormalizeFolderPathProduction(modNasConnection.GetCurrentTargetRuntimeRoot())

    If warehouseId = "" Then warehouseId = "WH1"
    If rootPath = "" Then
        overrideRoot = NormalizeFolderPathProduction(modRuntimeWorkbooks.GetCoreDataRootOverride())
        If overrideRoot <> "" Then rootPath = overrideRoot
    End If
    If rootPath = "" Then
        existingRoot = NormalizeFolderPathProduction(modRuntimeWorkbooks.TryResolveExistingRuntimeRoot(warehouseId))
        If existingRoot <> "" Then rootPath = existingRoot
    End If

    If rootPath = "" Then
        report = "A connected warehouse target is required before saving or loading Production recipes from the server."
        Exit Function
    End If

    ResolveProductionRecipesStorageTarget = True
    Exit Function

FailSoft:
    report = "ResolveProductionRecipesStorageTarget failed: " & Err.Description
End Function

Private Function OpenProductionRecipesWorkbook(ByVal warehouseId As String, _
                                               ByVal rootPath As String, _
                                               ByVal createIfMissing As Boolean, _
                                               ByRef openedTransient As Boolean, _
                                               ByRef report As String) As Workbook
    On Error GoTo FailSoft

    Dim targetPath As String
    Dim wb As Workbook

    targetPath = ProductionRecipesWorkbookPath(warehouseId, rootPath)
    If targetPath = "" Then
        report = "Production recipes workbook path could not be resolved."
        Exit Function
    End If

    Set wb = FindOpenWorkbookByFullNameProduction(targetPath)
    If Not wb Is Nothing Then
        HideWorkbookWindowsProduction wb
        Set OpenProductionRecipesWorkbook = wb
        Exit Function
    End If

    If Len(Dir$(targetPath, vbNormal)) > 0 Then
        Set wb = OpenWorkbookHiddenProduction(targetPath, False, openedTransient)
    ElseIf createIfMissing Then
        EnsureFolderRecursiveProduction GetParentFolderProduction(targetPath)
        Set wb = Application.Workbooks.Add(xlWBATWorksheet)
        HideWorkbookWindowsProduction wb
        wb.Worksheets(1).Name = SHEET_RUNTIME_RECIPES
        If EnsureProductionRecipesSchema(wb, report) Is Nothing Then
            CloseWorkbookNoSaveProduction wb
            Exit Function
        End If
        wb.SaveAs Filename:=targetPath, FileFormat:=50
        openedTransient = False
    Else
        report = "Production recipes runtime workbook was not found: " & targetPath
        Exit Function
    End If

    Set OpenProductionRecipesWorkbook = wb
    Exit Function

FailSoft:
    report = "OpenProductionRecipesWorkbook failed: " & Err.Description
End Function

Private Function EnsureProductionRecipesSchema(ByVal wb As Workbook, ByRef report As String) As ListObject
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim startCell As Range
    Dim dataRange As Range

    If wb Is Nothing Then Exit Function
    Set ws = WorkbookSheetExists(wb, SHEET_RUNTIME_RECIPES)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_RUNTIME_RECIPES
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_RUNTIME_RECIPES)
    On Error GoTo FailSoft

    headers = ProductionRecipeRuntimeHeaders()
    If lo Is Nothing Then
        Set startCell = ws.Range("A1")
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i
        Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = TABLE_RUNTIME_RECIPES
        If Not lo.DataBodyRange Is Nothing Then lo.ListRows(1).Delete
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureProductionColumnExists lo, CStr(headers(i))
    Next i

    Set EnsureProductionRecipesSchema = lo
    Exit Function

FailSoft:
    report = "EnsureProductionRecipesSchema failed: " & Err.Description
End Function

Private Sub MergeProductionRecipeRuntimeRowsToLocal(ByVal loRuntime As ListObject, ByVal loLocal As ListObject)
    Dim ids As Object
    Dim cRuntimeId As Long
    Dim arr As Variant
    Dim r As Long
    Dim recipeId As String
    Dim key As Variant

    If loRuntime Is Nothing Or loLocal Is Nothing Then Exit Sub
    If loRuntime.DataBodyRange Is Nothing Then Exit Sub

    cRuntimeId = ColumnIndex(loRuntime, "RECIPE_ID")
    If cRuntimeId = 0 Then Exit Sub

    Set ids = CreateObject("Scripting.Dictionary")
    arr = loRuntime.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        recipeId = Trim$(NzStr(arr(r, cRuntimeId)))
        If recipeId <> "" Then ids(recipeId) = True
    Next r

    For Each key In ids.Keys
        DeleteProductionRecipeRowsById loLocal, CStr(key)
    Next key

    CopyProductionRecipeRowsById loRuntime, loLocal, vbNullString, False
End Sub

Private Function CopyProductionRecipeRowsById(ByVal loSource As ListObject, _
                                              ByVal loTarget As ListObject, _
                                              ByVal recipeId As String, _
                                              Optional ByVal stampMetadata As Boolean = False) As Long
    Dim headers As Variant
    Dim cSourceId As Long
    Dim i As Long
    Dim h As Long
    Dim srcCol As Long
    Dim targetCol As Long
    Dim lr As ListRow

    If loSource Is Nothing Or loTarget Is Nothing Then Exit Function
    If loSource.DataBodyRange Is Nothing Then Exit Function

    cSourceId = ColumnIndex(loSource, "RECIPE_ID")
    If cSourceId = 0 Then Exit Function

    headers = ProductionRecipeLocalHeaders()
    recipeId = Trim$(recipeId)
    For i = 1 To loSource.DataBodyRange.Rows.Count
        If recipeId <> "" Then
            If StrComp(NzStr(loSource.DataBodyRange.Cells(i, cSourceId).Value), recipeId, vbTextCompare) <> 0 Then GoTo NextRow
        End If

        Set lr = loTarget.ListRows.Add
        For h = LBound(headers) To UBound(headers)
            srcCol = ColumnIndex(loSource, CStr(headers(h)))
            targetCol = ColumnIndex(loTarget, CStr(headers(h)))
            If srcCol > 0 And targetCol > 0 Then
                lr.Range.Cells(1, targetCol).Value = loSource.DataBodyRange.Cells(i, srcCol).Value
            End If
        Next h

        If stampMetadata Then
            SetProductionTableCellByHeader loTarget, lr.Index, "UPDATED_AT_UTC", Now
            SetProductionTableCellByHeader loTarget, lr.Index, "UPDATED_BY", CurrentProductionUserId()
        End If
        CopyProductionRecipeRowsById = CopyProductionRecipeRowsById + 1
NextRow:
    Next i
End Function

Private Sub DeleteProductionRecipeRowsById(ByVal lo As ListObject, ByVal recipeId As String)
    Dim cRecipeId As Long
    Dim r As Long

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    recipeId = Trim$(recipeId)
    If recipeId = "" Then Exit Sub

    cRecipeId = ColumnIndex(lo, "RECIPE_ID")
    If cRecipeId = 0 Then Exit Sub

    For r = lo.DataBodyRange.Rows.Count To 1 Step -1
        If StrComp(NzStr(lo.DataBodyRange.Cells(r, cRecipeId).Value), recipeId, vbTextCompare) = 0 Then
            lo.ListRows(r).Delete
        End If
    Next r
End Sub

Private Function ProductionRecipeLocalHeaders() As Variant
    ProductionRecipeLocalHeaders = Array("RECIPE", "RECIPE_ID", "DESCRIPTION", "ROW_BUDGET", "DEPARTMENT", _
        "PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT", _
        "RECIPE_LIST_ROW", "INGREDIENT_ID", "GUID")
End Function

Private Function ProductionRecipeRuntimeHeaders() As Variant
    ProductionRecipeRuntimeHeaders = Array("RECIPE", "RECIPE_ID", "DESCRIPTION", "ROW_BUDGET", "DEPARTMENT", _
        "PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT", _
        "RECIPE_LIST_ROW", "INGREDIENT_ID", "GUID", "UPDATED_AT_UTC", "UPDATED_BY")
End Function

Private Function NormalizeProductionRowBudget(ByVal value As Variant) As Long
    NormalizeProductionRowBudget = CLng(Val(NzStr(value)))
    If NormalizeProductionRowBudget <= 0 Then NormalizeProductionRowBudget = PRODUCTION_DEFAULT_ROW_BUDGET
    If NormalizeProductionRowBudget > PRODUCTION_MAX_ROW_BUDGET Then NormalizeProductionRowBudget = PRODUCTION_MAX_ROW_BUDGET
End Function

Private Sub EnsureProductionColumnExists(ByVal lo As ListObject, ByVal columnName As String)
    If lo Is Nothing Then Exit Sub
    If ColumnIndex(lo, columnName) > 0 Then Exit Sub
    lo.ListColumns.Add.Name = columnName
End Sub

Private Sub SetProductionTableCellByHeader(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                           ByVal columnName As String, ByVal value As Variant)
    Dim c As Long
    If lo Is Nothing Then Exit Sub
    c = ColumnIndex(lo, columnName)
    If c = 0 Then Exit Sub
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Sub
    If StrComp(columnName, "RECIPE_ID", vbTextCompare) = 0 Then
        WriteProductionTextCell lo.DataBodyRange.Cells(rowIndex, c), CStr(value)
    Else
        lo.DataBodyRange.Cells(rowIndex, c).Value = value
    End If
End Sub

Private Sub WriteProductionTextCell(ByVal targetCell As Range, ByVal textValue As String)
    If targetCell Is Nothing Then Exit Sub
    targetCell.NumberFormat = "@"
    targetCell.Value2 = CStr(textValue)
End Sub

Private Function CurrentProductionUserId() As String
    On Error Resume Next
    CurrentProductionUserId = Trim$(modRoleEventWriter.ResolveCurrentUserId())
    On Error GoTo 0
    If CurrentProductionUserId = "" Then CurrentProductionUserId = Trim$(Environ$("USERNAME"))
    If CurrentProductionUserId = "" Then CurrentProductionUserId = "unknown"
End Function

Public Function RefreshProductionIngredientPaletteFromRuntime(Optional ByVal operatorWb As Workbook = Nothing, _
                                                              Optional ByRef report As String = "") As Boolean
    On Error GoTo FailSoft

    Dim wbOps As Workbook
    Dim surfaceReport As String
    Dim wsPal As Worksheet
    Dim loLocal As ListObject
    Dim wbRuntime As Workbook
    Dim loRuntime As ListObject
    Dim warehouseId As String
    Dim rootPath As String
    Dim openedTransient As Boolean

    Set wbOps = ResolveProductionWorkbook(operatorWb, "IngredientPalette")
    If wbOps Is Nothing Then Set wbOps = ResolveProductionWorkbook(operatorWb, "IngredientsPalette")
    If wbOps Is Nothing Then Set wbOps = ResolveProductionWorkbook(operatorWb, SHEET_PRODUCTION)
    If wbOps Is Nothing Then
        report = "No Production operator workbook was available."
        Exit Function
    End If

    If Not modOperationsPrimitiveBridge.EnsureProductionWorkbookSurface(wbOps.Name, surfaceReport) Then
        report = "Production surface repair failed: " & surfaceReport
        Exit Function
    End If

    Set wsPal = WorkbookSheetExists(wbOps, "IngredientPalette")
    If wsPal Is Nothing Then Set wsPal = WorkbookSheetExists(wbOps, "IngredientsPalette")
    If wsPal Is Nothing Then
        report = "IngredientPalette sheet was not found in the operator workbook."
        Exit Function
    End If
    Set loLocal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "System_Key"))
    If loLocal Is Nothing Then
        report = "IngredientPalette table was not found in the operator workbook."
        Exit Function
    End If

    If Not ResolveProductionRecipesStorageTarget(warehouseId, rootPath, report) Then Exit Function

    Set wbRuntime = OpenProductionRecipesWorkbook(warehouseId, rootPath, False, openedTransient, report)
    If wbRuntime Is Nothing Then GoTo CleanExit

    Set loRuntime = EnsureProductionIngredientPaletteSchema(wbRuntime, report)
    If loRuntime Is Nothing Then GoTo CleanExit
    If loRuntime.DataBodyRange Is Nothing Then
        report = "Production ingredient palette runtime workbook has no saved assignment rows."
        RefreshProductionIngredientPaletteFromRuntime = True
        GoTo CleanExit
    End If

    MergeIngredientPaletteRuntimeRowsToLocal loRuntime, loLocal
    RefreshProductionIngredientPaletteFromRuntime = True
    report = "Production ingredient palette refreshed from " & wbRuntime.FullName

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveProduction wbRuntime
    Exit Function

FailSoft:
    report = "RefreshProductionIngredientPaletteFromRuntime failed: " & Err.Description
    Resume CleanExit
End Function

Private Function PublishIngredientPaletteRowsToRuntime(ByVal operatorWb As Workbook, _
                                                       ByVal loLocalPalette As ListObject, _
                                                       ByVal recipeId As String, _
                                                       ByVal ingredientId As String, _
                                                       ByRef report As String) As Boolean
    On Error GoTo FailSoft

    Dim warehouseId As String
    Dim rootPath As String
    Dim wbRuntime As Workbook
    Dim loRuntime As ListObject
    Dim openedTransient As Boolean
    Dim copiedCount As Long

    recipeId = Trim$(recipeId)
    ingredientId = Trim$(ingredientId)
    If recipeId = "" Or ingredientId = "" Then
        report = "RECIPE_ID and INGREDIENT_ID are required before saving Production assignments to the server."
        Exit Function
    End If
    If loLocalPalette Is Nothing Then
        report = "Local IngredientPalette table is not available."
        Exit Function
    End If
    If loLocalPalette.DataBodyRange Is Nothing Then
        report = "Local IngredientPalette table has no rows to publish."
        Exit Function
    End If

    If Not ResolveProductionRecipesStorageTarget(warehouseId, rootPath, report) Then Exit Function

    Set wbRuntime = OpenProductionRecipesWorkbook(warehouseId, rootPath, True, openedTransient, report)
    If wbRuntime Is Nothing Then GoTo CleanExit
    If wbRuntime.ReadOnly Then
        report = "Production recipes runtime workbook is read-only: " & wbRuntime.FullName
        GoTo CleanExit
    End If

    Set loRuntime = EnsureProductionIngredientPaletteSchema(wbRuntime, report)
    If loRuntime Is Nothing Then GoTo CleanExit

    DeleteIngredientPaletteRowsByKey loRuntime, recipeId, ingredientId
    copiedCount = CopyIngredientPaletteRowsByKey(loLocalPalette, loRuntime, recipeId, ingredientId, True)
    If copiedCount = 0 Then
        report = "No local IngredientPalette rows matched RECIPE_ID " & recipeId & " / INGREDIENT_ID " & ingredientId & "."
        GoTo CleanExit
    End If

    wbRuntime.Save
    PublishIngredientPaletteRowsToRuntime = True
    report = "Production ingredient palette runtime updated: " & wbRuntime.FullName & " (" & copiedCount & " rows)"

CleanExit:
    If openedTransient Then CloseWorkbookNoSaveProduction wbRuntime
    Exit Function

FailSoft:
    report = "PublishIngredientPaletteRowsToRuntime failed: " & Err.Description
    Resume CleanExit
End Function

Private Function EnsureProductionIngredientPaletteSchema(ByVal wb As Workbook, ByRef report As String) As ListObject
    On Error GoTo FailSoft

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim i As Long
    Dim startCell As Range
    Dim dataRange As Range

    If wb Is Nothing Then Exit Function
    Set ws = WorkbookSheetExists(wb, SHEET_RUNTIME_INGREDIENT_PALETTE)
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = SHEET_RUNTIME_INGREDIENT_PALETTE
    End If

    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_RUNTIME_INGREDIENT_PALETTE)
    On Error GoTo FailSoft

    headers = IngredientPaletteRuntimeHeaders()
    If lo Is Nothing Then
        Set startCell = ws.Range("A1")
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i
        Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = TABLE_RUNTIME_INGREDIENT_PALETTE
        If Not lo.DataBodyRange Is Nothing Then lo.ListRows(1).Delete
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureProductionColumnExists lo, CStr(headers(i))
    Next i

    Set EnsureProductionIngredientPaletteSchema = lo
    Exit Function

FailSoft:
    report = "EnsureProductionIngredientPaletteSchema failed: " & Err.Description
End Function

Private Sub MergeIngredientPaletteRuntimeRowsToLocal(ByVal loRuntime As ListObject, ByVal loLocal As ListObject)
    Dim keys As Object
    Dim cRuntimeRecipe As Long
    Dim cRuntimeIngredient As Long
    Dim arr As Variant
    Dim r As Long
    Dim key As Variant
    Dim recipeId As String
    Dim ingredientId As String

    If loRuntime Is Nothing Or loLocal Is Nothing Then Exit Sub
    If loRuntime.DataBodyRange Is Nothing Then Exit Sub

    cRuntimeRecipe = ColumnIndex(loRuntime, "RECIPE_ID")
    cRuntimeIngredient = ColumnIndex(loRuntime, "INGREDIENT_ID")
    If cRuntimeRecipe = 0 Or cRuntimeIngredient = 0 Then Exit Sub

    Set keys = CreateObject("Scripting.Dictionary")
    arr = loRuntime.DataBodyRange.Value
    For r = 1 To UBound(arr, 1)
        recipeId = Trim$(NzStr(arr(r, cRuntimeRecipe)))
        ingredientId = Trim$(NzStr(arr(r, cRuntimeIngredient)))
        If recipeId <> "" And ingredientId <> "" Then keys(NormalizeIdFirst(recipeId) & "|" & NormalizeIdLast(ingredientId)) = Array(recipeId, ingredientId)
    Next r

    For Each key In keys.Keys
        Dim keyParts As Variant
        keyParts = keys(key)
        DeleteIngredientPaletteRowsByKey loLocal, CStr(keyParts(0)), CStr(keyParts(1))
    Next key

    CopyIngredientPaletteRowsByKey loRuntime, loLocal, vbNullString, vbNullString, False
End Sub

Private Function CopyIngredientPaletteRowsByKey(ByVal loSource As ListObject, _
                                                ByVal loTarget As ListObject, _
                                                ByVal recipeId As String, _
                                                ByVal ingredientId As String, _
                                                Optional ByVal stampMetadata As Boolean = False) As Long
    Dim headers As Variant
    Dim cSourceRecipe As Long
    Dim cSourceIngredient As Long
    Dim i As Long
    Dim h As Long
    Dim srcCol As Long
    Dim targetCol As Long
    Dim lr As ListRow

    If loSource Is Nothing Or loTarget Is Nothing Then Exit Function
    If loSource.DataBodyRange Is Nothing Then Exit Function

    cSourceRecipe = ColumnIndex(loSource, "RECIPE_ID")
    cSourceIngredient = ColumnIndex(loSource, "INGREDIENT_ID")
    If cSourceRecipe = 0 Or cSourceIngredient = 0 Then Exit Function

    headers = IngredientPaletteLocalHeaders()
    recipeId = Trim$(recipeId)
    ingredientId = Trim$(ingredientId)
    For i = 1 To loSource.DataBodyRange.Rows.Count
        If recipeId <> "" Then
            If Not IngredientPaletteRowMatches(loSource, i, cSourceRecipe, cSourceIngredient, recipeId, ingredientId) Then GoTo NextRow
        End If

        Set lr = loTarget.ListRows.Add
        For h = LBound(headers) To UBound(headers)
            srcCol = ColumnIndex(loSource, CStr(headers(h)))
            targetCol = ColumnIndex(loTarget, CStr(headers(h)))
            If srcCol > 0 And targetCol > 0 Then
                lr.Range.Cells(1, targetCol).Value = loSource.DataBodyRange.Cells(i, srcCol).Value
            End If
        Next h

        If stampMetadata Then
            SetProductionTableCellByHeader loTarget, lr.Index, "UPDATED_AT_UTC", Now
            SetProductionTableCellByHeader loTarget, lr.Index, "UPDATED_BY", CurrentProductionUserId()
        End If
        CopyIngredientPaletteRowsByKey = CopyIngredientPaletteRowsByKey + 1
NextRow:
    Next i
End Function

Private Sub DeleteIngredientPaletteRowsByKey(ByVal lo As ListObject, ByVal recipeId As String, ByVal ingredientId As String)
    Dim cRecipeId As Long
    Dim cIngredientId As Long
    Dim r As Long

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    recipeId = Trim$(recipeId)
    ingredientId = Trim$(ingredientId)
    If recipeId = "" Or ingredientId = "" Then Exit Sub

    cRecipeId = ColumnIndex(lo, "RECIPE_ID")
    cIngredientId = ColumnIndex(lo, "INGREDIENT_ID")
    If cRecipeId = 0 Or cIngredientId = 0 Then Exit Sub

    For r = lo.DataBodyRange.Rows.Count To 1 Step -1
        If IngredientPaletteRowMatches(lo, r, cRecipeId, cIngredientId, recipeId, ingredientId) Then
            lo.ListRows(r).Delete
        End If
    Next r
End Sub

Private Function IngredientPaletteRowMatches(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                             ByVal cRecipeId As Long, ByVal cIngredientId As Long, _
                                             ByVal recipeId As String, ByVal ingredientId As String) As Boolean
    IngredientPaletteRowMatches = (NormalizeIdFirst(NzStr(lo.DataBodyRange.Cells(rowIndex, cRecipeId).Value)) = NormalizeIdFirst(recipeId) _
                                   And NormalizeIdLast(NzStr(lo.DataBodyRange.Cells(rowIndex, cIngredientId).Value)) = NormalizeIdLast(ingredientId))
End Function

Private Function IngredientPaletteLocalHeaders() As Variant
    IngredientPaletteLocalHeaders = Array("RECIPE_ID", "INGREDIENT_ID", "INPUT/OUTPUT", "ITEM", "PERCENT", "UOM", "AMOUNT", "System_Key", "ITEM_CODE", "GUID")
End Function

Private Function IngredientPaletteRuntimeHeaders() As Variant
    IngredientPaletteRuntimeHeaders = Array("RECIPE_ID", "INGREDIENT_ID", "INPUT/OUTPUT", "ITEM", "PERCENT", "UOM", "AMOUNT", "System_Key", "ITEM_CODE", "GUID", "UPDATED_AT_UTC", "UPDATED_BY")
End Function

'@TestOnlyBegin
Public Function TestProductionIngredientPaletteRuntimeRoundTrip(ByVal runtimeRoot As String) As String
    On Error GoTo FailSoft

    Dim priorRoot As String
    Dim wbOps As Workbook
    Dim wsPal As Worksheet
    Dim loLocal As ListObject
    Dim lr As ListRow
    Dim report As String
    Dim ok As Boolean
    Dim cRecipeId As Long
    Dim cIngredientId As Long
    Dim cSystemKey As Long
    Dim cItemCode As Long
    Dim r As Long
    Dim found As Boolean

    priorRoot = modRuntimeWorkbooks.GetCoreDataRootOverride()
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot

    Set wbOps = Application.Workbooks.Add(xlWBATWorksheet)
    If Not modOperationsPrimitiveBridge.EnsureProductionWorkbookSurface(wbOps.Name, report) Then
        TestProductionIngredientPaletteRuntimeRoundTrip = "FAILED: surface: " & report
        GoTo CleanExit
    End If

    Set wsPal = WorkbookSheetExists(wbOps, "IngredientPalette")
    If wsPal Is Nothing Then Set wsPal = WorkbookSheetExists(wbOps, "IngredientsPalette")
    Set loLocal = FindListObjectByNameOrHeaders(wsPal, "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "System_Key"))
    ClearListObjectData loLocal

    Set lr = loLocal.ListRows.Add
    SetProductionTableCellByHeader loLocal, lr.Index, "RECIPE_ID", "PAL"
    SetProductionTableCellByHeader loLocal, lr.Index, "INGREDIENT_ID", "ING-PAL"
    SetProductionTableCellByHeader loLocal, lr.Index, "INPUT/OUTPUT", "INPUT"
    SetProductionTableCellByHeader loLocal, lr.Index, "ITEM", "Malawi Black Tea"
    SetProductionTableCellByHeader loLocal, lr.Index, "PERCENT", 100
    SetProductionTableCellByHeader loLocal, lr.Index, "UOM", "LB"
    SetProductionTableCellByHeader loLocal, lr.Index, "AMOUNT", 1
    SetProductionTableCellByHeader loLocal, lr.Index, "System_Key", "SYS-PALETTE-95"
    SetProductionTableCellByHeader loLocal, lr.Index, "ITEM_CODE", "ITM-PAL-095"
    SetProductionTableCellByHeader loLocal, lr.Index, "GUID", "GUID-PAL"

    ok = PublishIngredientPaletteRowsToRuntime(wbOps, loLocal, "PAL", "ING-PAL", report)
    If Not ok Then
        TestProductionIngredientPaletteRuntimeRoundTrip = "FAILED: publish: " & report
        GoTo CleanExit
    End If

    ClearListObjectData loLocal
    ok = RefreshProductionIngredientPaletteFromRuntime(wbOps, report)
    If Not ok Then
        TestProductionIngredientPaletteRuntimeRoundTrip = "FAILED: refresh: " & report
        GoTo CleanExit
    End If

    cRecipeId = ColumnIndex(loLocal, "RECIPE_ID")
    cIngredientId = ColumnIndex(loLocal, "INGREDIENT_ID")
    cSystemKey = ColumnIndex(loLocal, "System_Key")
    cItemCode = ColumnIndex(loLocal, "ITEM_CODE")
    If Not loLocal.DataBodyRange Is Nothing And cRecipeId > 0 And cIngredientId > 0 And cSystemKey > 0 And cItemCode > 0 Then
        For r = 1 To loLocal.DataBodyRange.Rows.Count
            If NormalizeIdFirst(NzStr(loLocal.DataBodyRange.Cells(r, cRecipeId).Value)) = NormalizeIdFirst("PAL") _
               And NormalizeIdLast(NzStr(loLocal.DataBodyRange.Cells(r, cIngredientId).Value)) = NormalizeIdLast("ING-PAL") _
               And StrComp(NzStr(loLocal.DataBodyRange.Cells(r, cSystemKey).Value), "SYS-PALETTE-95", vbTextCompare) = 0 _
               And StrComp(NzStr(loLocal.DataBodyRange.Cells(r, cItemCode).Value), "ITM-PAL-095", vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next r
    End If

    If found Then
        TestProductionIngredientPaletteRuntimeRoundTrip = "OK"
    Else
        TestProductionIngredientPaletteRuntimeRoundTrip = "FAILED: refreshed IngredientPalette row was not found."
    End If

CleanExit:
    On Error Resume Next
    If Not wbOps Is Nothing Then wbOps.Close SaveChanges:=False
    modRuntimeWorkbooks.SetCoreDataRootOverride priorRoot
    On Error GoTo 0
    Exit Function

FailSoft:
    TestProductionIngredientPaletteRuntimeRoundTrip = "FAILED: " & Err.Description
    Resume CleanExit
End Function

Public Function TestProductionInventoryPickerPrefersCanonicalRuntime(ByVal runtimeRoot As String) As String
    On Error GoTo FailSoft

    Dim priorRoot As String
    Dim wbInv As Workbook
    Dim wsLegacy As Worksheet
    Dim wsCatalog As Worksheet
    Dim wsBalance As Worksheet
    Dim inventoryPath As String
    Dim items As Variant
    Dim paletteRows(1 To 1, 1 To 7) As Variant

    ClearProductionOperatorWorkbookBinding
    priorRoot = modRuntimeWorkbooks.GetCoreDataRootOverride()
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    EnsureFolderRecursiveProduction runtimeRoot

    inventoryPath = NormalizeFolderPathProduction(runtimeRoot) & "\WH1.invSys.Data.Inventory.xlsb"
    Set wbInv = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsLegacy = wbInv.Worksheets(1)
    wsLegacy.Name = "InventoryManagement"
    wsLegacy.Range("A1:G1").Value = Array("System_Key", "ITEM", "UOM", "TOTAL INV", "LOCATION", "DESCRIPTION", "ITEM_CODE")
    wsLegacy.Range("A2:G2").Value = Array("SYS-OLD-001", "Old Runtime Item", "EA", 1, "A1", "stale legacy entity", "OLD-001")
    wsLegacy.ListObjects.Add(xlSrcRange, wsLegacy.Range("A1:G2"), , xlYes).Name = "invSys"

    Set wsCatalog = wbInv.Worksheets.Add(After:=wsLegacy)
    wsCatalog.Name = "SkuCatalog"
    wsCatalog.Range("A1:H1").Value = Array("SKU", "System_Key", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION", "ITEM_KIND")
    wsCatalog.Range("A2:H2").Value = Array("ITM-PICK-001", "SYS-PICK-001", "ITM-PICK-001", "Malawi Black Tea", "LB", "A1", "canonical catalog entity", "INVENTORY")
    wsCatalog.ListObjects.Add(xlSrcRange, wsCatalog.Range("A1:H2"), , xlYes).Name = "tblSkuCatalog"

    Set wsBalance = wbInv.Worksheets.Add(After:=wsCatalog)
    wsBalance.Name = "SkuBalance"
    wsBalance.Range("A1:B1").Value = Array("SKU", "QtyOnHand")
    wsBalance.Range("A2:B2").Value = Array("ITM-PICK-001", 1)
    wsBalance.ListObjects.Add(xlSrcRange, wsBalance.Range("A1:B2"), , xlYes).Name = "tblSkuBalance"

    wbInv.SaveAs Filename:=inventoryPath, FileFormat:=50
    wbInv.Close SaveChanges:=False
    Set wbInv = Nothing

    items = LoadProductionInventoryPickerItems("Malawi")
    If IsEmpty(items) Then
        TestProductionInventoryPickerPrefersCanonicalRuntime = "FAILED: picker returned no rows for canonical catalog item."
        GoTo CleanExit
    End If
    If NzStr(items(1, 1)) <> "SYS-PICK-001" Or NzStr(items(1, 2)) <> "Malawi Black Tea" Then
        TestProductionInventoryPickerPrefersCanonicalRuntime = "FAILED: picker did not prefer canonical catalog entity. System_Key=" & _
            NzStr(items(1, 1)) & "; Item=" & NzStr(items(1, 2))
        GoTo CleanExit
    End If
    If NzStr(items(1, 4)) <> "1" Then
        TestProductionInventoryPickerPrefersCanonicalRuntime = "FAILED: picker did not load canonical QtyOnHand. Qty=" & NzStr(items(1, 4))
        GoTo CleanExit
    End If
    If UBound(items, 2) < 7 Or NzStr(items(1, 7)) <> "ITM-PICK-001" Then
        TestProductionInventoryPickerPrefersCanonicalRuntime = _
            "FAILED: picker did not return canonical ITEM_CODE/SKU identity."
        GoTo CleanExit
    End If
    paletteRows(1, 4) = "Malawi Black Tea"
    paletteRows(1, 6) = "LB"
    HydrateIngredientPaletteIdentityRowsProduction paletteRows, 1, items
    If NzStr(paletteRows(1, 5)) <> "SYS-PICK-001" Or NzStr(paletteRows(1, 7)) <> "ITM-PICK-001" Then
        TestProductionInventoryPickerPrefersCanonicalRuntime = _
            "FAILED: palette assignment did not hydrate canonical System_Key and SKU identity."
        GoTo CleanExit
    End If

    TestProductionInventoryPickerPrefersCanonicalRuntime = "OK"

CleanExit:
    On Error Resume Next
    If Not wbInv Is Nothing Then wbInv.Close SaveChanges:=False
    ClearProductionOperatorWorkbookBinding
    modRuntimeWorkbooks.SetCoreDataRootOverride priorRoot
    On Error GoTo 0
    Exit Function

FailSoft:
    TestProductionInventoryPickerPrefersCanonicalRuntime = "FAILED: " & Err.Description
    Resume CleanExit
End Function
'@TestOnlyEnd

Private Function ProductionRecipesWorkbookPath(ByVal warehouseId As String, ByVal rootPath As String) As String
    rootPath = NormalizeFolderPathProduction(rootPath)
    warehouseId = Trim$(warehouseId)
    If rootPath = "" Or warehouseId = "" Then Exit Function
    ProductionRecipesWorkbookPath = rootPath & "\" & warehouseId & ".invSys.Data.ProductionRecipes.xlsb"
End Function

Private Function NormalizeFolderPathProduction(ByVal folderPath As String) As String
    NormalizeFolderPathProduction = Trim$(folderPath)
    Do While Len(NormalizeFolderPathProduction) > 1 And Right$(NormalizeFolderPathProduction, 1) = "\"
        NormalizeFolderPathProduction = Left$(NormalizeFolderPathProduction, Len(NormalizeFolderPathProduction) - 1)
    Loop
End Function

Private Function GetParentFolderProduction(ByVal fullPath As String) As String
    GetParentFolderProduction = modDeploymentPaths.GetParentFolderManaged(fullPath)
End Function

Private Sub EnsureFolderRecursiveProduction(ByVal folderPath As String)
    If Trim$(folderPath) = "" Then Exit Sub
    modDeploymentPaths.EnsureFolderRecursiveManaged folderPath
End Sub

Private Function FindOpenWorkbookByFullNameProduction(ByVal fullNameIn As String) As Workbook
    Dim wb As Workbook

    fullNameIn = Trim$(fullNameIn)
    If fullNameIn = "" Then Exit Function
    For Each wb In Application.Workbooks
        If StrComp(Trim$(wb.FullName), fullNameIn, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByFullNameProduction = wb
            Exit Function
        End If
    Next wb
End Function

Private Function OpenWorkbookHiddenProduction(ByVal workbookPath As String, _
                                              ByVal readOnly As Boolean, _
                                              ByRef openedTransient As Boolean) As Workbook
    On Error GoTo FailSoft

    Dim wb As Workbook
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean

    openedTransient = False
    workbookPath = Trim$(workbookPath)
    If workbookPath = "" Then Exit Function

    Set wb = FindOpenWorkbookByFullNameProduction(workbookPath)
    If Not wb Is Nothing Then
        HideWorkbookWindowsProduction wb
        Set OpenWorkbookHiddenProduction = wb
        Exit Function
    End If

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = Application.Workbooks.Open(Filename:=workbookPath, _
                                        UpdateLinks:=False, _
                                        ReadOnly:=readOnly, _
                                        AddToMru:=False, _
                                        IgnoreReadOnlyRecommended:=True, _
                                        Notify:=False)
    HideWorkbookWindowsProduction wb
    openedTransient = True
    Set OpenWorkbookHiddenProduction = wb

CleanExit:
    On Error Resume Next
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
    Exit Function

FailSoft:
    Resume CleanExit
End Function

Private Sub HideWorkbookWindowsProduction(ByVal wb As Workbook)
    On Error Resume Next
    Dim win As Window
    If wb Is Nothing Then Exit Sub
    For Each win In wb.Windows
        win.Visible = False
    Next win
    On Error GoTo 0
End Sub

Private Sub CloseWorkbookNoSaveProduction(ByVal wb As Workbook)
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function BuildRecipeProcessTablesFromLines(ByVal recipeId As String, Optional ByVal ApplyTemplates As Boolean = False, Optional ByVal anchorBelowLines As Boolean = True) As Long
    ' System 1: Recipe List Builder - build process tables under RB_AddRecipeName.
    Dim wsProd As Worksheet: Set wsProd = SheetExists(SHEET_PRODUCTION)
    If wsProd Is Nothing Then Exit Function

    Dim loLines As ListObject
    Set loLines = GetRecipeBuilderLinesTable(wsProd)
    If loLines Is Nothing Then
        MsgBox "Recipe Builder lines table not found on Production sheet.", vbExclamation
        Exit Function
    End If
    If loLines.DataBodyRange Is Nothing Then Exit Function

    Dim cProc As Long: cProc = ColumnIndex(loLines, "PROCESS")
    If cProc = 0 Then
        MsgBox "Recipe Builder lines missing PROCESS column.", vbCritical
        Exit Function
    End If

    Dim startRow As Long
    Dim startCol As Long
    Dim includeLines As Boolean
    includeLines = anchorBelowLines
    If includeLines Then
        If IsRecipeLinesStaged(loLines) Then includeLines = False
    End If
    If Not GetRecipeBuilderAnchor(wsProd, startRow, startCol, includeLines) Then
        MsgBox "Recipe Builder header table (RB_AddRecipeName) not found on Production sheet.", vbExclamation
        Exit Function
    End If

    Dim headerNames As Variant
    headerNames = RecipeProcessHeaderList()
    Dim colCount As Long: colCount = UBound(headerNames) - LBound(headerNames) + 1
    Dim srcIdx() As Long
    ReDim srcIdx(1 To colCount)
    Dim c As Long
    For c = 1 To colCount
        srcIdx(c) = ColumnIndex(loLines, CStr(headerNames(LBound(headerNames) + c - 1)))
    Next c

    Dim lineArr As Variant: lineArr = loLines.DataBodyRange.value
    Dim procMap As Object: Set procMap = CreateObject("Scripting.Dictionary")
    Dim procOrder As Collection: Set procOrder = New Collection

    Dim r As Long
    For r = 1 To UBound(lineArr, 1)
        Dim procName As String: procName = Trim$(NzStr(lineArr(r, cProc)))
        If procName <> "" Then
            If Not procMap.Exists(procName) Then
                procMap.Add procName, New Collection
                procOrder.Add procName
            End If
            procMap(procName).Add r
        End If
    Next r

    If procOrder.count = 0 Then Exit Function

    DeleteRecipeProcessTables wsProd

    Dim created As New Collection

    Dim procKey As Variant
    Dim nextSeq As Long
    nextSeq = NextRecipeProcessSequence(wsProd)
    For Each procKey In procOrder
        Dim rowsColl As Collection: Set rowsColl = procMap(procKey)
        Dim dataCount As Long: dataCount = rowsColl.count
        If dataCount = 0 Then GoTo NextProc

        Dim tableRange As Range
        Set tableRange = wsProd.Range(wsProd.Cells(startRow, startCol), wsProd.Cells(startRow + dataCount, startCol + colCount - 1))

        If RangeHasListObjectCollision(wsProd, tableRange, loLines) Then
            MsgBox "Not enough space below Recipe Builder to create process tables. Clear space and try again.", vbExclamation
            Exit Function
        End If

        tableRange.Clear
        tableRange.rows(1).value = HeaderRowArray(headerNames)

        Dim dataArr() As Variant
        ReDim dataArr(1 To dataCount, 1 To colCount)
        Dim i As Long
        For i = 1 To dataCount
            Dim srcRow As Long: srcRow = rowsColl(i)
            For c = 1 To colCount
                Dim hdrName As String
                hdrName = CStr(headerNames(LBound(headerNames) + c - 1))
                If StrComp(hdrName, "PROCESS", vbTextCompare) = 0 Then
                    dataArr(i, c) = procKey
                ElseIf srcIdx(c) > 0 Then
                    dataArr(i, c) = lineArr(srcRow, srcIdx(c))
                End If
            Next c
        Next i

        tableRange.Offset(1, 0).Resize(dataCount, colCount).value = dataArr

        Dim newLo As ListObject
        Set newLo = wsProd.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        newLo.Name = UniqueListObjectName(wsProd, BuildRecipeProcessTableName(CStr(nextSeq)))
        On Error Resume Next
        newLo.TableStyle = loLines.TableStyle
        On Error GoTo 0
        created.Add newLo
        nextSeq = nextSeq + 1

        startRow = startRow + dataCount + 3 ' keep 2 blank rows between process tables
NextProc:
    Next procKey

    BuildRecipeProcessTablesFromLines = created.count

    If ApplyTemplates And created.count > 0 And recipeId <> "" Then
        Dim loProc As ListObject
        For Each loProc In created
            Dim procNameTpl As String: procNameTpl = ProcessNameFromTable(loProc)
            ApplyProductionTemplates loProc, TEMPLATE_SCOPE_RECIPE_PROCESS, procNameTpl, "", recipeId
        Next loProc
    End If
End Function

Private Function CreateRecipeProcessTable(ByVal ws As Worksheet, ByVal processName As String, Optional ByVal dataRows As Long = 1) As ListObject
    ' System 1: Recipe List Builder - add a blank process table under RB_AddRecipeName.
    If ws Is Nothing Then Exit Function
    If dataRows < 1 Then dataRows = 1

    Dim loLines As ListObject
    Set loLines = GetRecipeBuilderLinesTable(ws)
    If loLines Is Nothing Then Exit Function

    Dim headers As Variant
    headers = RecipeProcessHeaderList()
    Dim colCount As Long: colCount = UBound(headers) - LBound(headers) + 1
    Dim startRow As Long
    Dim startCol As Long
    Dim includeLines As Boolean
    includeLines = Not IsRecipeLinesStaged(loLines)
    If Not GetRecipeBuilderAnchor(ws, startRow, startCol, includeLines) Then Exit Function

    Dim tableRange As Range
    startRow = NextRecipeBuilderStartRow(ws, startRow)
    Set tableRange = FindAvailableRecipeProcessRange(ws, startRow, startCol, dataRows + 1, colCount, loLines)
    If tableRange Is Nothing Then Exit Function

    Dim seq As Long
    seq = NextRecipeProcessSequence(ws)
    If Trim$(processName) = "" Then processName = CStr(seq)

    tableRange.Clear
    tableRange.rows(1).value = HeaderRowArray(headers)

    Dim cProc As Long
    cProc = HeaderIndex(headers, "PROCESS")
    If cProc > 0 Then
        tableRange.Offset(1, cProc - 1).value = processName
    End If

    Dim newLo As ListObject
    Set newLo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    newLo.Name = UniqueListObjectName(ws, BuildRecipeProcessTableName(CStr(seq)))
    On Error Resume Next
    newLo.TableStyle = loLines.TableStyle
    On Error GoTo 0

    FocusRecipeProcessTable newLo
    Set CreateRecipeProcessTable = newLo
End Function

Private Sub FocusRecipeProcessTable(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    lo.Parent.Activate
    Application.Goto lo.Range, True
    On Error GoTo 0
End Sub

Private Function NextRecipeBuilderStartRow(ByVal ws As Worksheet, ByVal baseRow As Long) As Long
    ' System 1: Recipe List Builder - stack new process tables below the last one.
    Dim startRow As Long
    startRow = baseRow
    If ws Is Nothing Then
        NextRecipeBuilderStartRow = startRow
        Exit Function
    End If

    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsRecipeProcessTable(lo) Then
            Dim bottom As Long
            bottom = lo.Range.row + lo.Range.rows.count - 1
            If bottom + 3 > startRow Then startRow = bottom + 3 ' keep 2 blank rows
        End If
    Next lo

    NextRecipeBuilderStartRow = startRow
End Function

Private Function GetRecipeBuilderAnchor(ByVal ws As Worksheet, ByRef startRow As Long, ByRef startCol As Long, Optional ByVal includeLines As Boolean = True) As Boolean
    ' System 1: Recipe List Builder anchor (under RB_AddRecipeName).
    GetRecipeBuilderAnchor = False
    If ws Is Nothing Then Exit Function
    Dim loHeader As ListObject
    Set loHeader = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    If loHeader Is Nothing Then Exit Function

    startCol = loHeader.Range.Column
    startRow = loHeader.Range.row + loHeader.Range.rows.count + 3 ' keep 2 blank rows before first process table

    If includeLines Then
        Dim loLines As ListObject
        Set loLines = GetRecipeBuilderLinesTable(ws, loHeader)
        If Not loLines Is Nothing Then
            Dim linesBottom As Long
            linesBottom = loLines.Range.row + loLines.Range.rows.count - 1
            If linesBottom + 3 > startRow Then startRow = linesBottom + 3
        End If
    End If
    If startRow > 0 And startCol > 0 Then GetRecipeBuilderAnchor = True
End Function

Private Function EnsureRecipeBuilderLinesTable(ByVal ws As Worksheet, ByVal loHeader As ListObject) As ListObject
    ' System 1: Recipe List Builder - create RecipeBuilder lines table if missing.
    If ws Is Nothing Then Exit Function
    If loHeader Is Nothing Then Exit Function

    Dim existing As ListObject
    Set existing = GetRecipeBuilderLinesTable(ws, loHeader)
    If Not existing Is Nothing Then
        Set EnsureRecipeBuilderLinesTable = existing
        Exit Function
    End If

    Dim headers As Variant
    headers = RecipeProcessHeaderList()
    Dim colCount As Long: colCount = UBound(headers) - LBound(headers) + 1

    Dim startRow As Long
    Dim startCol As Long
    startRow = loHeader.Range.row + loHeader.Range.rows.count + 2 ' one blank row below header
    startCol = loHeader.Range.Column

    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 1, startCol + colCount - 1))
    If RangeHasListObjectCollisionStrict(ws, tableRange, loHeader) Then Exit Function

    tableRange.Clear
    tableRange.rows(1).value = HeaderRowArray(headers)

    Dim newLo As ListObject
    Set newLo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    newLo.Name = UniqueListObjectName(ws, TABLE_RECIPE_BUILDER_LINES)
    On Error Resume Next
    newLo.TableStyle = loHeader.TableStyle
    On Error GoTo 0

    Set EnsureRecipeBuilderLinesTable = newLo
End Function

Private Function GetRecipeBuilderLinesTable(ByVal ws As Worksheet, Optional ByVal loHeader As ListObject) As ListObject
    ' System 1: Recipe List Builder - locate RecipeBuilder lines table under RB_AddRecipeName.
    If ws Is Nothing Then Exit Function

    Dim lo As ListObject
    Set lo = GetListObject(ws, TABLE_RECIPE_BUILDER_LINES)
    If Not lo Is Nothing Then
        Set GetRecipeBuilderLinesTable = lo
        Exit Function
    End If

    Dim headerStartCol As Long
    Dim headerBottom As Long
    If loHeader Is Nothing Then
        Set loHeader = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_BUILDER_HEADER, Array("RECIPE_NAME", "RECIPE_ID"))
    End If
    If Not loHeader Is Nothing Then
        headerStartCol = loHeader.Range.Column
        headerBottom = loHeader.Range.row + loHeader.Range.rows.count - 1
    End If

    Dim candidate As ListObject
    Dim bestRow As Long
    For Each lo In ws.ListObjects
        If ListObjectHasHeaders(lo, Array("PROCESS", "INGREDIENT")) Then
            If IsRecipeProcessTable(lo) Then GoTo NextLo
            If headerStartCol > 0 Then
                If lo.Range.Column <> headerStartCol Then GoTo NextLo
                If lo.Range.row < headerBottom Then GoTo NextLo
            End If
            If bestRow = 0 Or lo.Range.row < bestRow Then
                Set candidate = lo
                bestRow = lo.Range.row
            End If
        End If
NextLo:
    Next lo

    If Not candidate Is Nothing Then
        Set GetRecipeBuilderLinesTable = candidate
        Exit Function
    End If

    If headerStartCol = 0 Then
        Set GetRecipeBuilderLinesTable = FindListObjectByNameOrHeaders(ws, TABLE_RECIPE_BUILDER_LINES, Array("PROCESS", "INGREDIENT"))
    End If
End Function

Private Function RecipeLinesHasProcess(ByVal loLines As ListObject) As Boolean
    ' System 1: Recipe List Builder - detect any PROCESS rows.
    If loLines Is Nothing Then Exit Function
    If loLines.DataBodyRange Is Nothing Then Exit Function
    Dim cProc As Long: cProc = ColumnIndex(loLines, "PROCESS")
    If cProc = 0 Then Exit Function
    Dim arr As Variant: arr = loLines.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If Trim$(NzStr(arr(r, cProc))) <> "" Then
            RecipeLinesHasProcess = True
            Exit Function
        End If
    Next r
End Function

Private Function RecipeBuilderLinesHaveData(ByVal loLines As ListObject) As Boolean
    ' System 1: Recipe List Builder - detect active UserForm/line-table edits.
    If loLines Is Nothing Then Exit Function
    If loLines.DataBodyRange Is Nothing Then Exit Function

    Dim cIng As Long: cIng = ColumnIndex(loLines, "INGREDIENT")
    Dim cProc As Long: cProc = ColumnIndex(loLines, "PROCESS")
    If cIng = 0 And cProc = 0 Then Exit Function

    Dim arr As Variant: arr = loLines.DataBodyRange.value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If cIng > 0 Then
            If Trim$(NzStr(arr(r, cIng))) <> "" Then
                RecipeBuilderLinesHaveData = True
                Exit Function
            End If
        End If
        If cProc > 0 Then
            If Trim$(NzStr(arr(r, cProc))) <> "" Then
                RecipeBuilderLinesHaveData = True
                Exit Function
            End If
        End If
    Next r
End Function

Private Function IsRecipeLinesStaged(ByVal loLines As ListObject) As Boolean
    ' System 1: Recipe List Builder - check if lines table is staged off-screen.
    If loLines Is Nothing Then Exit Function
    IsRecipeLinesStaged = (loLines.Range.row >= RECIPE_LINES_STAGING_ROW)
End Function

Private Function MoveRecipeBuilderLinesToStaging(ByVal loLines As ListObject) As Boolean
    ' System 1: Recipe List Builder - move lines table out of view before building process tables.
    If loLines Is Nothing Then Exit Function
    Dim ws As Worksheet: Set ws = loLines.Parent
    Dim startRow As Long: startRow = RECIPE_LINES_STAGING_ROW
    If loLines.Range.row >= startRow Then
        MoveRecipeBuilderLinesToStaging = True
        Exit Function
    End If

    Dim dest As Range
    Set dest = ws.Cells(startRow, loLines.Range.Column)
    On Error Resume Next
    loLines.Range.Cut Destination:=dest
    ClearExcelClipboardStateProduction
    MoveRecipeBuilderLinesToStaging = (Err.Number = 0)
    If MoveRecipeBuilderLinesToStaging Then
        On Error Resume Next
        loLines.Name = TABLE_RECIPE_BUILDER_LINES
        On Error GoTo 0
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function EnsureInventoryPaletteLinesTable(ByVal ws As Worksheet, Optional ByVal baseStyle As String = "") As ListObject
    ' System 4: Production Input/Output - keep InventoryPalette lines table staged off-screen.
    If ws Is Nothing Then Exit Function

    Dim lo As ListObject
    Set lo = GetListObject(ws, TABLE_INV_PALETTE_GENERATED)

    Dim startRow As Long
    Dim startCol As Long
    startRow = PALETTE_LINES_STAGING_ROW
    startCol = 1

    Dim loProd As ListObject
    Set loProd = FindListObjectByNameOrHeaders(ws, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    If Not loProd Is Nothing Then startCol = loProd.Range.Column

    If Not lo Is Nothing Then
        If lo.Range.row < startRow Then
            Dim dest As Range
            Set dest = ws.Cells(startRow, startCol)
            On Error Resume Next
            lo.Range.Cut Destination:=dest
            ClearExcelClipboardStateProduction
            On Error GoTo 0
        End If
        On Error Resume Next
        If baseStyle <> "" Then lo.TableStyle = baseStyle
        On Error GoTo 0
        Set EnsureInventoryPaletteLinesTable = lo
        Exit Function
    End If

    Dim headers As Variant
    headers = InventoryPaletteHeaderList()
    Dim colCount As Long: colCount = UBound(headers) - LBound(headers) + 1

    Dim tableRange As Range
    Set tableRange = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + 1, startCol + colCount - 1))
    If RangeHasListObjectCollisionStrict(ws, tableRange) Then Exit Function

    tableRange.Clear
    tableRange.rows(1).value = HeaderRowArray(headers)

    Dim newLo As ListObject
    Set newLo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
    newLo.Name = TABLE_INV_PALETTE_GENERATED
    On Error Resume Next
    If baseStyle <> "" Then newLo.TableStyle = baseStyle
    On Error GoTo 0

    Set EnsureInventoryPaletteLinesTable = newLo
End Function

Private Function HeaderIndex(ByVal headers As Variant, ByVal headerName As String) As Long
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        If StrComp(CStr(headers(i)), headerName, vbTextCompare) = 0 Then
            HeaderIndex = i - LBound(headers) + 1
            Exit Function
        End If
    Next i
End Function

Private Function FindAvailableRecipeProcessRange(ByVal ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, _
    ByVal totalRows As Long, ByVal totalCols As Long, ByVal loLines As ListObject) As Range

    If ws Is Nothing Then Exit Function
    If totalRows < 1 Or totalCols < 1 Then Exit Function
    If startRow < 1 Then startRow = 1
    If startCol < 1 Then startCol = 1

    Dim maxRow As Long
    maxRow = ws.rows.count
    Dim tryRow As Long: tryRow = startRow
    Dim candidate As Range

    Do While tryRow + totalRows - 1 <= maxRow
        Set candidate = ws.Range(ws.Cells(tryRow, startCol), ws.Cells(tryRow + totalRows - 1, startCol + totalCols - 1))
        If Not RangeHasListObjectCollisionStrict(ws, candidate, loLines) Then
            Set FindAvailableRecipeProcessRange = candidate
            Exit Function
        End If
        tryRow = tryRow + totalRows + 3 ' keep 2 blank rows between tables
    Loop
End Function

Private Sub DeleteRecipeProcessTables(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    Dim i As Long
    For i = ws.ListObjects.count To 1 Step -1
        Dim lo As ListObject
        Set lo = ws.ListObjects(i)
        If IsRecipeProcessTable(lo) Then
            Dim addr As String
            addr = lo.Range.Address
            On Error Resume Next
            lo.Delete
            ws.Range(addr).Clear
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function GetRecipeBuilderProcessTables(ByVal ws As Worksheet) As Collection
    Dim result As New Collection
    If ws Is Nothing Then
        Set GetRecipeBuilderProcessTables = result
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsRecipeProcessTable(lo) Then result.Add lo
    Next lo
    Set GetRecipeBuilderProcessTables = result
End Function

Private Function IsRecipeProcessTable(ByVal lo As ListObject) As Boolean
    ' System 1: Recipe List Builder - identify process tables.
    If lo Is Nothing Then Exit Function
    Dim nm As String: nm = LCase$(lo.Name)
    If Left$(nm, 5) <> "proc_" Then Exit Function
    If Right$(nm, Len(RECIPE_PROC_TABLE_SUFFIX) + 1) = "_" & LCase$(RECIPE_PROC_TABLE_SUFFIX) Then
        IsRecipeProcessTable = True
    End If
End Function

Private Function GetOrAddTemplateRow(ByVal loTpl As ListObject, ByVal cScope As Long, ByVal cRecipe As Long, _
    ByVal cTargetCol As Long, ByVal cFormula As Long) As ListRow

    If loTpl Is Nothing Then Exit Function
    If loTpl.DataBodyRange Is Nothing Then
        Set GetOrAddTemplateRow = loTpl.ListRows.Add
        Exit Function
    End If

    Dim r As Long
    For r = 1 To loTpl.DataBodyRange.Rows.Count
        If IsTemplateRowEmpty(loTpl, r, cScope, cRecipe, cTargetCol, cFormula) Then
            Set GetOrAddTemplateRow = loTpl.ListRows(r)
            Exit Function
        End If
    Next r

    Set GetOrAddTemplateRow = loTpl.ListRows.Add
End Function

Private Function IsTemplateRowEmpty(ByVal loTpl As ListObject, ByVal rowIdx As Long, ByVal cScope As Long, _
    ByVal cRecipe As Long, ByVal cTargetCol As Long, ByVal cFormula As Long) As Boolean

    If loTpl Is Nothing Then Exit Function
    If loTpl.DataBodyRange Is Nothing Then Exit Function
    If rowIdx < 1 Or rowIdx > loTpl.DataBodyRange.Rows.Count Then Exit Function

    Dim rowRange As Range
    Set rowRange = loTpl.DataBodyRange.Rows(rowIdx)

    Dim scopeVal As String
    Dim recipeVal As String
    Dim targetVal As String
    Dim formulaVal As String

    If cScope > 0 Then scopeVal = NzStr(rowRange.Cells(1, cScope).Value)
    If cRecipe > 0 Then recipeVal = NzStr(rowRange.Cells(1, cRecipe).Value)
    If cTargetCol > 0 Then targetVal = NzStr(rowRange.Cells(1, cTargetCol).Value)

    If cFormula > 0 Then
        Dim fCell As Range
        Set fCell = rowRange.Cells(1, cFormula)
        If Not fCell Is Nothing Then
            If fCell.HasFormula Then
                formulaVal = CStr(fCell.FormulaR1C1)
            Else
                formulaVal = NzStr(fCell.Value)
            End If
        End If
    End If

    If scopeVal <> "" Or recipeVal <> "" Or targetVal <> "" Then Exit Function
    If formulaVal <> "" And formulaVal <> "0" Then Exit Function

    IsTemplateRowEmpty = True
End Function

' System 1: Recipe List Builder - register process formulas as templates.
Private Function RegisterRecipeTemplates(ByVal recipeId As String, ByVal processTables As Collection) As Long
    If processTables Is Nothing Then Exit Function
    If processTables.count = 0 Then Exit Function

    Dim wsTpl As Worksheet: Set wsTpl = SheetExists(SHEET_TEMPLATES)
    If wsTpl Is Nothing Then Exit Function
    Dim loTpl As ListObject: Set loTpl = GetListObject(wsTpl, "TemplatesTable")
    If loTpl Is Nothing Then Exit Function

    Dim cGuid As Long: cGuid = ColumnIndex(loTpl, "GUID")
    Dim cScope As Long: cScope = ColumnIndex(loTpl, "TEMPLATE_SCOPE")
    Dim cRecipe As Long: cRecipe = ColumnIndex(loTpl, "RECIPE_ID")
    Dim cIngredient As Long: cIngredient = ColumnIndex(loTpl, "INGREDIENT_ID")
    Dim cProcess As Long: cProcess = ColumnIndex(loTpl, "PROCESS")
    Dim cTargetTable As Long: cTargetTable = ColumnIndex(loTpl, "TARGET_TABLE")
    Dim cTargetCol As Long: cTargetCol = ColumnIndex(loTpl, "TARGET_COLUMN")
    Dim cFormula As Long: cFormula = ColumnIndex(loTpl, "FORMULA")
    Dim cNotes As Long: cNotes = ColumnIndex(loTpl, "NOTES")
    Dim cActive As Long: cActive = ColumnIndex(loTpl, "ACTIVE")
    Dim cCreated As Long: cCreated = ColumnIndex(loTpl, "CREATED_AT")
    Dim cUpdated As Long: cUpdated = ColumnIndex(loTpl, "UPDATED_AT")

    NormalizeTemplateFormulaColumn loTpl, cFormula

    If Not loTpl.DataBodyRange Is Nothing And cScope > 0 And cRecipe > 0 Then
        Dim r As Long
        For r = loTpl.DataBodyRange.rows.count To 1 Step -1
            If StrComp(NzStr(loTpl.DataBodyRange.Cells(r, cScope).value), TEMPLATE_SCOPE_RECIPE_PROCESS, vbTextCompare) = 0 Then
                If recipeId = "" Or StrComp(NzStr(loTpl.DataBodyRange.Cells(r, cRecipe).value), recipeId, vbTextCompare) = 0 Then
                    loTpl.ListRows(r).Delete
                End If
            End If
        Next r
    End If

    Dim nowVal As Date: nowVal = Now
    Dim added As Long

    Dim loProc As ListObject
    For Each loProc In processTables
        If loProc.DataBodyRange Is Nothing Then GoTo NextProc
        Dim procName As String: procName = ProcessNameFromTable(loProc)
        Dim lc As ListColumn
        For Each lc In loProc.ListColumns
            Dim formulaText As String
            formulaText = GetColumnFormulaText(lc)
            If formulaText = "" Then GoTo NextCol

            Dim lr As ListRow
            Set lr = GetOrAddTemplateRow(loTpl, cScope, cRecipe, cTargetCol, cFormula)
            If cGuid > 0 Then lr.Range.Cells(1, cGuid).Value = CreateProductionGuid()
            If cScope > 0 Then lr.Range.Cells(1, cScope).Value = TEMPLATE_SCOPE_RECIPE_PROCESS
            If cRecipe > 0 Then lr.Range.Cells(1, cRecipe).Value = recipeId
            If cIngredient > 0 Then lr.Range.Cells(1, cIngredient).Value = ""
            If cProcess > 0 Then lr.Range.Cells(1, cProcess).Value = procName
            If cTargetTable > 0 Then lr.Range.Cells(1, cTargetTable).Value = ""
            If cTargetCol > 0 Then lr.Range.Cells(1, cTargetCol).Value = lc.Name
            If cFormula > 0 Then WriteTemplateFormulaCell lr.Range.Cells(1, cFormula), formulaText
            If cNotes > 0 Then lr.Range.Cells(1, cNotes).Value = "Recipe builder"
            If cActive > 0 Then lr.Range.Cells(1, cActive).Value = True
            If cCreated > 0 Then lr.Range.Cells(1, cCreated).Value = nowVal
            If cUpdated > 0 Then lr.Range.Cells(1, cUpdated).Value = nowVal
            added = added + 1
NextCol:
        Next lc
NextProc:
    Next loProc

    RegisterRecipeTemplates = added
End Function

Private Function ProcessNameFromTable(ByVal lo As ListObject) As String
    If lo Is Nothing Then Exit Function
    Dim cProc As Long: cProc = ColumnIndex(lo, "PROCESS")
    If cProc > 0 And Not lo.DataBodyRange Is Nothing Then
        ProcessNameFromTable = NzStr(lo.DataBodyRange.Cells(1, cProc).value)
    End If
    If ProcessNameFromTable = "" Then ProcessNameFromTable = ExtractProcessKeyFromTableName(lo.Name)
End Function

Private Function GetColumnFormulaText(ByVal lc As ListColumn) As String
    If lc Is Nothing Then Exit Function
    If lc.DataBodyRange Is Nothing Then Exit Function
    Dim cell As Range
    Set cell = lc.DataBodyRange.Cells(1, 1)
    On Error Resume Next
    If cell.HasFormula Then GetColumnFormulaText = CStr(cell.FormulaR1C1)
    On Error GoTo 0
    If Left$(GetColumnFormulaText, 1) <> "=" Then GetColumnFormulaText = ""
End Function

Private Function SaveFormulaTemplatesForRecipe(ByVal recipeId As String, ByVal wsProd As Worksheet) As Long
    If Trim$(recipeId) = "" Then Exit Function
    If wsProd Is Nothing Then Exit Function

    Dim wsTpl As Worksheet: Set wsTpl = SheetExists(SHEET_TEMPLATES)
    If wsTpl Is Nothing Then
        MsgBox "TemplatesTable sheet not found.", vbCritical
        Exit Function
    End If
    Dim loTpl As ListObject: Set loTpl = GetListObject(wsTpl, "TemplatesTable")
    If loTpl Is Nothing Then
        MsgBox "TemplatesTable not found.", vbCritical
        Exit Function
    End If

    Dim cGuid As Long, cScope As Long, cRecipe As Long, cIngredient As Long, cProcess As Long
    Dim cTargetTable As Long, cTargetCol As Long, cFormula As Long, cNotes As Long
    Dim cActive As Long, cCreated As Long, cUpdated As Long
    If Not GetTemplateColumnIndexes(loTpl, cGuid, cScope, cRecipe, cIngredient, cProcess, cTargetTable, _
        cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated) Then Exit Function
    NormalizeTemplateFormulaColumn loTpl, cFormula

    Dim totalAdded As Long
    Dim nowVal As Date: nowVal = Now

    ' Scope: Recipe process tables (builder/chooser share formulas).
    ClearTemplatesForScope loTpl, recipeId, TEMPLATE_SCOPE_RECIPE_PROCESS, cScope, cRecipe
    Dim procTables As Collection
    Set procTables = GetRecipeBuilderProcessTables(wsProd)
    If Not procTables Is Nothing Then
        Dim loProc As ListObject
        For Each loProc In procTables
            totalAdded = totalAdded + AddTemplateRowsFromTable(loTpl, loProc, recipeId, TEMPLATE_SCOPE_RECIPE_PROCESS, _
                ProcessNameFromTable(loProc), "", nowVal, cGuid, cScope, cRecipe, cIngredient, cProcess, _
                cTargetTable, cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated, "Recipe process")
        Next loProc
    End If

    ' Scope: Inventory Palette Builder tables.
    ClearTemplatesForScope loTpl, recipeId, TEMPLATE_SCOPE_PALETTE_BUILDER, cScope, cRecipe
    Dim loIng As ListObject
    Dim loItems As ListObject
    Set loIng = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseIngredient", Array("INGREDIENT", "INGREDIENT_ID"))
    Set loItems = FindListObjectByNameOrHeaders(wsProd, "IP_ChooseItem", Array("ITEMS", "RECIPE_ID", "INGREDIENT_ID"))
    totalAdded = totalAdded + AddTemplateRowsFromTable(loTpl, loIng, recipeId, TEMPLATE_SCOPE_PALETTE_BUILDER, _
        "", "IP_ChooseIngredient", nowVal, cGuid, cScope, cRecipe, cIngredient, cProcess, _
        cTargetTable, cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated, "Palette builder")
    totalAdded = totalAdded + AddTemplateRowsFromTable(loTpl, loItems, recipeId, TEMPLATE_SCOPE_PALETTE_BUILDER, _
        "", "IP_ChooseItem", nowVal, cGuid, cScope, cRecipe, cIngredient, cProcess, _
        cTargetTable, cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated, "Palette builder")

    ' Scope: Production run tables.
    ClearTemplatesForScope loTpl, recipeId, TEMPLATE_SCOPE_PROD_RUN, cScope, cRecipe
    Dim lo As ListObject
    For Each lo In wsProd.ListObjects
        If LCase$(lo.Name) Like "proc_*_palette" Then
            totalAdded = totalAdded + AddTemplateRowsFromTable(loTpl, lo, recipeId, TEMPLATE_SCOPE_PROD_RUN, _
                ProcessNameFromTable(lo), TEMPLATE_TABLEKEY_PALETTE, nowVal, cGuid, cScope, cRecipe, cIngredient, cProcess, _
                cTargetTable, cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated, "Production run")
        End If
    Next lo
    Dim loOut As ListObject
    Set loOut = FindListObjectByNameOrHeaders(wsProd, "ProductionOutput", Array("PROCESS", "OUTPUT"))
    totalAdded = totalAdded + AddTemplateRowsFromTable(loTpl, loOut, recipeId, TEMPLATE_SCOPE_PROD_RUN, _
        "", "ProductionOutput", nowVal, cGuid, cScope, cRecipe, cIngredient, cProcess, _
        cTargetTable, cTargetCol, cFormula, cNotes, cActive, cCreated, cUpdated, "Production run")

    SaveFormulaTemplatesForRecipe = totalAdded
End Function

Private Function GetTemplateColumnIndexes(ByVal loTpl As ListObject, ByRef cGuid As Long, ByRef cScope As Long, _
    ByRef cRecipe As Long, ByRef cIngredient As Long, ByRef cProcess As Long, ByRef cTargetTable As Long, _
    ByRef cTargetCol As Long, ByRef cFormula As Long, ByRef cNotes As Long, ByRef cActive As Long, _
    ByRef cCreated As Long, ByRef cUpdated As Long) As Boolean

    If loTpl Is Nothing Then Exit Function
    cGuid = ColumnIndex(loTpl, "GUID")
    cScope = ColumnIndex(loTpl, "TEMPLATE_SCOPE")
    cRecipe = ColumnIndex(loTpl, "RECIPE_ID")
    cIngredient = ColumnIndex(loTpl, "INGREDIENT_ID")
    cProcess = ColumnIndex(loTpl, "PROCESS")
    cTargetTable = ColumnIndex(loTpl, "TARGET_TABLE")
    cTargetCol = ColumnIndex(loTpl, "TARGET_COLUMN")
    cFormula = ColumnIndex(loTpl, "FORMULA")
    cNotes = ColumnIndex(loTpl, "NOTES")
    cActive = ColumnIndex(loTpl, "ACTIVE")
    cCreated = ColumnIndex(loTpl, "CREATED_AT")
    cUpdated = ColumnIndex(loTpl, "UPDATED_AT")

    If cScope = 0 Or cRecipe = 0 Or cTargetCol = 0 Or cFormula = 0 Then
        MsgBox "TemplatesTable is missing required columns (TEMPLATE_SCOPE, RECIPE_ID, TARGET_COLUMN, FORMULA).", vbCritical
        Exit Function
    End If
    GetTemplateColumnIndexes = True
End Function

Private Sub ClearTemplatesForScope(ByVal loTpl As ListObject, ByVal recipeId As String, ByVal scopeName As String, _
    ByVal cScope As Long, ByVal cRecipe As Long)

    If loTpl Is Nothing Then Exit Sub
    If cScope = 0 Or cRecipe = 0 Then Exit Sub
    If loTpl.DataBodyRange Is Nothing Then Exit Sub

    Dim r As Long
    For r = loTpl.DataBodyRange.Rows.Count To 1 Step -1
        If StrComp(NzStr(loTpl.DataBodyRange.Cells(r, cScope).Value), scopeName, vbTextCompare) = 0 Then
            If StrComp(NzStr(loTpl.DataBodyRange.Cells(r, cRecipe).Value), recipeId, vbTextCompare) = 0 Then
                loTpl.ListRows(r).Delete
            End If
        End If
    Next r
End Sub

Private Function AddTemplateRowsFromTable(ByVal loTpl As ListObject, ByVal loSource As ListObject, ByVal recipeId As String, _
    ByVal scopeName As String, ByVal processName As String, ByVal targetTableName As String, ByVal nowVal As Date, _
    ByVal cGuid As Long, ByVal cScope As Long, ByVal cRecipe As Long, ByVal cIngredient As Long, ByVal cProcess As Long, _
    ByVal cTargetTable As Long, ByVal cTargetCol As Long, ByVal cFormula As Long, ByVal cNotes As Long, _
    ByVal cActive As Long, ByVal cCreated As Long, ByVal cUpdated As Long, Optional ByVal noteText As String = "") As Long

    If loTpl Is Nothing Or loSource Is Nothing Then Exit Function
    If loSource.DataBodyRange Is Nothing Then Exit Function

    Dim added As Long
    Dim lc As ListColumn
    For Each lc In loSource.ListColumns
        Dim formulaText As String
        formulaText = GetColumnFormulaText(lc)
        If formulaText = "" Then GoTo NextCol

        Dim lr As ListRow
        Set lr = GetOrAddTemplateRow(loTpl, cScope, cRecipe, cTargetCol, cFormula)
        If cGuid > 0 Then lr.Range.Cells(1, cGuid).Value = CreateProductionGuid()
        If cScope > 0 Then lr.Range.Cells(1, cScope).Value = scopeName
        If cRecipe > 0 Then lr.Range.Cells(1, cRecipe).Value = recipeId
        If cIngredient > 0 Then lr.Range.Cells(1, cIngredient).Value = ""
        If cProcess > 0 Then lr.Range.Cells(1, cProcess).Value = processName
        If cTargetTable > 0 Then lr.Range.Cells(1, cTargetTable).Value = targetTableName
        If cTargetCol > 0 Then lr.Range.Cells(1, cTargetCol).Value = lc.Name
        If cFormula > 0 Then WriteTemplateFormulaCell lr.Range.Cells(1, cFormula), formulaText
        If cNotes > 0 Then lr.Range.Cells(1, cNotes).Value = noteText
        If cActive > 0 Then lr.Range.Cells(1, cActive).Value = True
        If cCreated > 0 Then lr.Range.Cells(1, cCreated).Value = nowVal
        If cUpdated > 0 Then lr.Range.Cells(1, cUpdated).Value = nowVal
        added = added + 1
NextCol:
    Next lc

    AddTemplateRowsFromTable = added
End Function

Private Sub NormalizeTemplateFormulaColumn(ByVal loTpl As ListObject, ByVal cFormula As Long)
    If loTpl Is Nothing Then Exit Sub
    If cFormula = 0 Then Exit Sub
    Dim lc As ListColumn
    Set lc = loTpl.ListColumns(cFormula)
    If lc Is Nothing Then Exit Sub
    On Error Resume Next
    lc.Range.NumberFormat = "@"
    On Error GoTo 0
    If lc.DataBodyRange Is Nothing Then Exit Sub

    Dim cell As Range
    For Each cell In lc.DataBodyRange.Cells
        Dim formulaText As String
        If cell.HasFormula Then
            formulaText = CStr(cell.FormulaR1C1)
        Else
            formulaText = NzStr(cell.Value)
        End If
        If Left$(formulaText, 1) = "=" Then
            WriteTemplateFormulaCell cell, formulaText
        End If
    Next cell
End Sub

Private Function SafeProcessKey(ByVal rawKey As String) As String
    Dim cleaned As String
    cleaned = Trim$(rawKey)
    If cleaned = "" Then cleaned = "process"

    Dim i As Long, ch As String, key As String
    For i = 1 To Len(cleaned)
        ch = Mid$(cleaned, i, 1)
        If ch Like "[A-Za-z0-9_]" Then
            key = key & LCase$(ch)
        Else
            key = key & "_"
        End If
    Next i

    Do While InStr(key, "__") > 0
        key = Replace(key, "__", "_")
    Loop
    key = Trim$(key)
    If key = "" Then key = "process"
    If Not key Like "[A-Za-z_]*" Then key = "p_" & key
    SafeProcessKey = key
End Function

Private Function BuildRecipeProcessTableName(ByVal processKey As String) As String
    ' System 1: Recipe List Builder - process table naming.
    Dim key As String: key = Trim$(processKey)
    If key <> "" And IsNumeric(key) Then
        BuildRecipeProcessTableName = "proc_" & CLng(key) & "_" & RECIPE_PROC_TABLE_SUFFIX
    Else
        key = SafeProcessKey(processKey)
        BuildRecipeProcessTableName = "proc_" & key & "_" & RECIPE_PROC_TABLE_SUFFIX
    End If
End Function

Private Function NextRecipeProcessSequence(ByVal ws As Worksheet) As Long
    ' System 1: Recipe List Builder - next numeric process table sequence.
    Dim maxSeq As Long
    If ws Is Nothing Then
        NextRecipeProcessSequence = 1
        Exit Function
    End If
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If IsRecipeProcessTable(lo) Then
            Dim seq As Long
            seq = RecipeProcessSequenceFromName(lo.Name)
            If seq > maxSeq Then maxSeq = seq
        End If
    Next lo
    NextRecipeProcessSequence = maxSeq + 1
End Function

Private Function RecipeProcessSequenceFromName(ByVal tableName As String) As Long
    ' System 1: Recipe List Builder - parse numeric process table sequence.
    Dim nm As String: nm = LCase$(tableName)
    If Left$(nm, 5) <> "proc_" Then Exit Function
    If Right$(nm, Len(RECIPE_PROC_TABLE_SUFFIX) + 1) <> "_" & LCase$(RECIPE_PROC_TABLE_SUFFIX) Then Exit Function
    Dim core As String
    core = Mid$(nm, 6, Len(nm) - 5 - (Len(RECIPE_PROC_TABLE_SUFFIX) + 1))
    If core = "" Then Exit Function
    If Left$(core, 2) = "p_" Then core = Mid$(core, 3)
    RecipeProcessSequenceFromName = CLng(val(core))
End Function

Private Function RecipeProcessHeaderList() As Variant
    ' System 1: Recipe List Builder - process table headers.
    RecipeProcessHeaderList = Array( _
        "PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT", _
        "OOO", "INSTRUCTION", "RECIPE_LIST_ROW", "INGREDIENT_ID", "GUID")
End Function

Private Function HeaderRowArray(ByVal headers As Variant) As Variant
    Dim cols As Long: cols = UBound(headers) - LBound(headers) + 1
    Dim arr() As Variant
    ReDim arr(1 To 1, 1 To cols)
    Dim i As Long
    For i = 1 To cols
        arr(1, i) = headers(LBound(headers) + i - 1)
    Next i
    HeaderRowArray = arr
End Function

Private Function UniqueListObjectName(ByVal ws As Worksheet, ByVal baseName As String) As String
    Dim nameTry As String: nameTry = baseName
    Dim idx As Long: idx = 1
    Do While Not GetListObject(ws, nameTry) Is Nothing
        nameTry = baseName & "_" & CStr(idx)
        idx = idx + 1
    Loop
    UniqueListObjectName = nameTry
End Function

Private Function ExtractProcessKeyFromTableName(ByVal tableName As String) As String
    Dim nm As String: nm = LCase$(tableName)
    If Left$(nm, 5) <> "proc_" Then Exit Function
    Dim parts As Variant: parts = Split(nm, "_")
    If UBound(parts) < 2 Then Exit Function
    Dim i As Long
    For i = 1 To UBound(parts) - 1
        If ExtractProcessKeyFromTableName <> "" Then ExtractProcessKeyFromTableName = ExtractProcessKeyFromTableName & "_"
        ExtractProcessKeyFromTableName = ExtractProcessKeyFromTableName & parts(i)
    Next i
End Function

Private Function RangeHasListObjectCollision(ByVal ws As Worksheet, ByVal targetRange As Range, ParamArray allowedTables() As Variant) As Boolean
    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo Is Nothing Then GoTo NextLo
        If IsListObjectAllowed(lo, allowedTables) Then GoTo NextLo
        If Not Intersect(lo.Range, targetRange) Is Nothing Then
            RangeHasListObjectCollision = True
            Exit Function
        End If
NextLo:
    Next lo
End Function

Private Function IsListObjectAllowed(ByVal lo As ListObject, ByVal allowedTables As Variant) As Boolean
    Dim v As Variant
    For Each v In allowedTables
        If TypeName(v) = "ListObject" Then
            If lo Is v Then
                IsListObjectAllowed = True
                Exit Function
            End If
        End If
    Next v
    If IsRecipeProcessTable(lo) Then IsListObjectAllowed = True
End Function

Private Function RangeHasListObjectCollisionStrict(ByVal ws As Worksheet, ByVal targetRange As Range, ParamArray allowedTables() As Variant) As Boolean
    If ws Is Nothing Then Exit Function
    If targetRange Is Nothing Then Exit Function
    Dim lo As ListObject
    For Each lo In ws.ListObjects
        If lo Is Nothing Then GoTo NextLo
        If IsListObjectAllowedStrict(lo, False, allowedTables) Then GoTo NextLo
        If Not Intersect(lo.Range, targetRange) Is Nothing Then
            RangeHasListObjectCollisionStrict = True
            Exit Function
        End If
NextLo:
    Next lo
End Function

Private Function IsListObjectAllowedStrict(ByVal lo As ListObject, ByVal allowRecipeTables As Boolean, ByVal allowedTables As Variant) As Boolean
    Dim v As Variant
    For Each v In allowedTables
        If TypeName(v) = "ListObject" Then
            If lo Is v Then
                IsListObjectAllowedStrict = True
                Exit Function
            End If
        End If
    Next v
    If allowRecipeTables Then
        If IsRecipeProcessTable(lo) Then IsListObjectAllowedStrict = True
    End If
End Function

Private Sub EnsureTableHasRow(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If Not lo.DataBodyRange Is Nothing Then Exit Sub
    On Error Resume Next
    lo.ListRows.Add AlwaysInsert:=True
    On Error GoTo 0
End Sub

' System 2: Inventory Palette Builder - clear values but keep a single data row.
Private Sub ResetPaletteTable(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    EnsureTableHasRow lo
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cell As Range
    For Each cell In lo.DataBodyRange.Cells
        If Not cell.HasFormula Then cell.ClearContents
    Next cell
End Sub

Private Function GetHeaderDataCell(lo As ListObject, colName As String) As Range
    If lo Is Nothing Then Exit Function
    Dim idx As Long: idx = ColumnIndex(lo, colName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then
        Set GetHeaderDataCell = lo.HeaderRowRange.Offset(1, 0).Cells(1, idx)
    Else
        Set GetHeaderDataCell = lo.DataBodyRange.Cells(1, idx)
    End If
End Function

Private Sub ClearListObjectData(lo As ListObject)
    If lo Is Nothing Then Exit Sub
    On Error Resume Next
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    On Error GoTo 0
End Sub

Private Sub ClearListObjectFormulas(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    Dim cell As Range
    For Each cell In lo.DataBodyRange.Cells
        If cell.HasFormula Then cell.ClearContents
    Next cell
End Sub

Private Sub ApplyProductionTemplates(ByVal targetLo As ListObject, ByVal templateScope As String, _
    Optional ByVal processKey As String = "", Optional ByVal tableNameOverride As String = "", _
    Optional ByVal recipeId As String = "", Optional ByVal ingredientId As String = "")

    If targetLo Is Nothing Then Exit Sub
    If targetLo.Parent Is Nothing Then Exit Sub

    Dim wb As Workbook
    Set wb = targetLo.Parent.Parent
    If wb Is Nothing Then Exit Sub

    Dim wsTpl As Worksheet
    Set wsTpl = WorkbookSheetExists(wb, SHEET_TEMPLATES)
    If wsTpl Is Nothing Then Exit Sub

    Dim loTpl As ListObject
    Set loTpl = GetListObject(wsTpl, TABLE_TEMPLATES)
    If loTpl Is Nothing Then Exit Sub
    If loTpl.DataBodyRange Is Nothing Then Exit Sub

    Dim cScope As Long: cScope = ColumnIndex(loTpl, "TEMPLATE_SCOPE")
    Dim cRecipe As Long: cRecipe = ColumnIndex(loTpl, "RECIPE_ID")
    Dim cIngredient As Long: cIngredient = ColumnIndex(loTpl, "INGREDIENT_ID")
    Dim cProcess As Long: cProcess = ColumnIndex(loTpl, "PROCESS")
    Dim cTargetTable As Long: cTargetTable = ColumnIndex(loTpl, "TARGET_TABLE")
    Dim cTargetCol As Long: cTargetCol = ColumnIndex(loTpl, "TARGET_COLUMN")
    Dim cFormula As Long: cFormula = ColumnIndex(loTpl, "FORMULA")
    Dim cActive As Long: cActive = ColumnIndex(loTpl, "ACTIVE")
    If cTargetCol = 0 Or cFormula = 0 Then Exit Sub

    Dim arr As Variant
    arr = loTpl.DataBodyRange.Value
    Dim r As Long
    For r = 1 To UBound(arr, 1)
        If cActive > 0 Then
            If LCase$(Trim$(NzStr(arr(r, cActive)))) = "false" Then GoTo NextTemplate
        End If
        If cScope > 0 Then
            If StrComp(NzStr(arr(r, cScope)), templateScope, vbTextCompare) <> 0 Then GoTo NextTemplate
        End If
        If cRecipe > 0 And recipeId <> "" Then
            If StrComp(NzStr(arr(r, cRecipe)), recipeId, vbTextCompare) <> 0 Then GoTo NextTemplate
        End If
        If cIngredient > 0 And ingredientId <> "" Then
            If StrComp(NzStr(arr(r, cIngredient)), ingredientId, vbTextCompare) <> 0 Then GoTo NextTemplate
        End If
        If cProcess > 0 And processKey <> "" Then
            If StrComp(NzStr(arr(r, cProcess)), processKey, vbTextCompare) <> 0 Then GoTo NextTemplate
        End If
        If cTargetTable > 0 Then
            Dim templateTable As String
            templateTable = NzStr(arr(r, cTargetTable))
            If templateTable <> "" Then
                Dim targetName As String
                If tableNameOverride <> "" Then
                    targetName = tableNameOverride
                Else
                    targetName = targetLo.Name
                End If
                If StrComp(templateTable, targetName, vbTextCompare) <> 0 Then GoTo NextTemplate
            End If
        End If

        Dim colName As String
        colName = NzStr(arr(r, cTargetCol))
        If colName = "" Then GoTo NextTemplate

        Dim formulaText As String
        formulaText = NzStr(arr(r, cFormula))
        If formulaText = "" Or Left$(formulaText, 1) <> "=" Then
            Dim formulaCell As Range
            Set formulaCell = loTpl.DataBodyRange.Cells(r, cFormula)
            If Not formulaCell Is Nothing Then
                If formulaCell.HasFormula Then formulaText = CStr(formulaCell.FormulaR1C1)
            End If
        End If
        If formulaText = "" Or Left$(formulaText, 1) <> "=" Then GoTo NextTemplate

        ApplyProductionTemplateFormulaToColumn targetLo, colName, formulaText
NextTemplate:
    Next r
End Sub

Private Sub ApplyProductionTemplateFormulaToColumn(ByVal targetLo As ListObject, ByVal colName As String, ByVal formulaText As String)
    If targetLo Is Nothing Then Exit Sub
    If targetLo.DataBodyRange Is Nothing Then Exit Sub

    Dim colIdx As Long
    colIdx = ColumnIndex(targetLo, colName)
    If colIdx = 0 Then Exit Sub

    Dim rng As Range
    Set rng = targetLo.ListColumns(colIdx).DataBodyRange
    On Error Resume Next
    Err.Clear
    rng.FormulaR1C1 = formulaText
    If Err.Number <> 0 Then
        Err.Clear
        rng.Formula = formulaText
    End If
    On Error GoTo 0
End Sub

Private Sub WriteTemplateFormulaCell(ByVal targetCell As Range, ByVal formulaText As String)
    If targetCell Is Nothing Then Exit Sub
    If formulaText = "" Then Exit Sub
    On Error Resume Next
    targetCell.NumberFormat = "@"
    targetCell.Value = "'" & formulaText
    On Error GoTo 0
End Sub

Private Sub RemoveRecipeBuilderLinesTable(ByVal lo As ListObject)
    ' System 1: Recipe List Builder - remove RecipeBuilder lines table after load.
    If lo Is Nothing Then Exit Sub
    Dim ws As Worksheet: Set ws = lo.Parent
    Dim addr As String: addr = lo.Range.Address
    On Error Resume Next
    lo.Delete
    ws.Range(addr).Clear
    On Error GoTo 0
End Sub

Private Function NzStr(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NzStr = ""
    Else
        NzStr = CStr(v)
    End If
End Function

Private Function NzDbl(v As Variant) As Double
    If IsError(v) Then
        NzDbl = 0#
    ElseIf IsNull(v) Or IsEmpty(v) Then
        NzDbl = 0#
    ElseIf Trim$(CStr(v)) = "" Then
        NzDbl = 0#
    ElseIf IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0#
    End If
End Function

Private Function NzLng(v As Variant) As Long
    If IsError(v) Then
        NzLng = 0
    ElseIf IsNull(v) Or IsEmpty(v) Then
        NzLng = 0
    ElseIf Trim$(CStr(v)) = "" Then
        NzLng = 0
    ElseIf IsNumeric(v) Then
        NzLng = CLng(v)
    Else
        NzLng = 0
    End If
End Function

Private Function NormalizeIdFirst(ByVal v As String) As String
    Dim tokens As Variant
    tokens = SplitTokens(v)
    If IsEmpty(tokens) Then Exit Function
    NormalizeIdFirst = CStr(tokens(LBound(tokens)))
End Function

Private Function NormalizeIdLast(ByVal v As String) As String
    Dim tokens As Variant
    tokens = SplitTokens(v)
    If IsEmpty(tokens) Then Exit Function
    NormalizeIdLast = CStr(tokens(UBound(tokens)))
End Function

Private Function SplitTokens(ByVal v As String) As Variant
    Dim s As String
    s = Trim$(v)
    If s = "" Then Exit Function
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    s = Application.WorksheetFunction.Trim(s)
    Dim parts As Variant
    parts = Split(s, " ")
    Dim cleaned() As String
    Dim i As Long, n As Long
    For i = LBound(parts) To UBound(parts)
        If Trim$(parts(i)) <> "" Then
            n = n + 1
            ReDim Preserve cleaned(0 To n - 1)
            cleaned(n - 1) = Trim$(parts(i))
        End If
    Next i
    If n = 0 Then Exit Function
    SplitTokens = cleaned
End Function
