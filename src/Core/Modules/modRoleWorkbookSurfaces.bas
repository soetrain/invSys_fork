Attribute VB_Name = "modRoleWorkbookSurfaces"
Option Explicit

Private Const SHIPPING_BACKEND_SHEET As String = "ShippingBackend"

Public Function EnsureReceivingWorkbookSurface(Optional ByVal targetWb As Workbook = Nothing, _
                                               Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Set wb = ResolveTargetWorkbookSurface(targetWb)

    EnsureTableSurface wb, "ReceivedTally", "ReceivedTally", Array("REF_NUMBER", "ITEMS", "QUANTITY", "ROW"), True, "C3"
    EnsureTableSurface wb, "ReceivedTally", "AggregateReceived", Array("REF_NUMBER", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW"), False, "J3"
    EnsureTableSurface wb, "ReceivedTally", "invSysData_Receiving", InventoryManagementHeadersSurface(), False, "V3"
    EnsureInventoryManagementSurface wb
    EnsureTableSurface wb, "ReceivedLog", "ReceivedLog", Array("SNAPSHOT_ID", "ENTRY_DATE", "USER", "REF_NUMBER", "ITEMS", "QUANTITY", "UOM", "VENDOR", "LOCATION", "ITEM_CODE", "ROW"), False
    ArrangeReceivingTablesSurface wb
    EnsureReceivingButtonsSurface wb
    FormatWorkbookSurface wb

    EnsureReceivingWorkbookSurface = True
    Exit Function

FailEnsure:
    report = "EnsureReceivingWorkbookSurface failed: " & Err.Description
End Function

Public Function EnsureShippingWorkbookSurface(Optional ByVal targetWb As Workbook = Nothing, _
                                              Optional ByRef report As String = "") As Boolean
    Static inProgress As Boolean

    If inProgress Then
        EnsureShippingWorkbookSurface = True
        Exit Function
    End If
    inProgress = True
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Set wb = ResolveTargetWorkbookSurface(targetWb)

    EnsureWorksheetSurface wb, "ShipmentsTally"
    MoveTableToSheetSurface wb, "ShipmentsTally", SHIPPING_BACKEND_SHEET
    MoveTableToSheetSurface wb, "NotShipped", SHIPPING_BACKEND_SHEET
    MoveTableToSheetSurface wb, "AggregateBoxBOM", SHIPPING_BACKEND_SHEET
    MoveTableToSheetSurface wb, "AggregatePackages", SHIPPING_BACKEND_SHEET
    MoveTableToSheetSurface wb, "Check_invSys", SHIPPING_BACKEND_SHEET
    MoveTableToSheetSurface wb, "invSysData_Shipping", SHIPPING_BACKEND_SHEET
    MoveTableToSheetSurface wb, "ShippingBOMView", SHIPPING_BACKEND_SHEET
    MoveTableToSheetSurface wb, "AggregateBoxBOM_Log", SHIPPING_BACKEND_SHEET
    MoveTableToSheetSurface wb, "AggregatePackages_Log", SHIPPING_BACKEND_SHEET

    EnsureTableSurface wb, SHIPPING_BACKEND_SHEET, "ShipmentsTally", Array("LINE_ID", "SERVER_RESERVE_EVENT_ID", "REF_NUMBER", "ITEMS", "QUANTITY", "ROW", "UOM", "LOCATION", "DESCRIPTION", "AREA", "CARRIER"), True
    EnsureTableSurface wb, SHIPPING_BACKEND_SHEET, "NotShipped", Array("LINE_ID", "SERVER_RESERVE_EVENT_ID", "REF_NUMBER", "ITEMS", "QUANTITY", "ROW", "UOM", "LOCATION", "DESCRIPTION", "AREA", "CARRIER"), False
    EnsureTableSurface wb, SHIPPING_BACKEND_SHEET, "AggregateBoxBOM", Array("ROW", "ITEM_CODE", "ITEM", "QUANTITY", "UOM", "LOCATION"), False
    EnsureTableSurface wb, SHIPPING_BACKEND_SHEET, "AggregatePackages", Array("ROW", "ITEM_CODE", "ITEM", "QUANTITY", "UOM", "LOCATION"), False
    EnsureTableSurface wb, "ShipmentsTally", "BoxBuilder", Array("Box Name", "UOM", "LOCATION", "DESCRIPTION"), True
    EnsureTableSurface wb, "ShipmentsTally", "BoxBOM", Array("ITEM", "ROW", "QUANTITY", "UOM", "LOCATION", "DESCRIPTION"), True
    EnsureTableSurface wb, SHIPPING_BACKEND_SHEET, "Check_invSys", Array("ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "USED", "MADE", "SHIPMENTS", "TOTAL INV"), False
    EnsureTableSurface wb, SHIPPING_BACKEND_SHEET, "invSysData_Shipping", InventoryManagementHeadersSurface(), False
    EnsureTableSurface wb, SHIPPING_BACKEND_SHEET, "ShippingBOMView", ShippingBomViewHeadersSurface(), False
    EnsureTableSurface wb, SHIPPING_BACKEND_SHEET, "AggregateBoxBOM_Log", Array("GUID", "USER", "ACTION", "ROW", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP"), False
    EnsureTableSurface wb, SHIPPING_BACKEND_SHEET, "AggregatePackages_Log", Array("GUID", "USER", "ACTION", "ROW", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP"), False
    ArrangeShippingBackendTablesSurface wb
    On Error Resume Next
    Application.CutCopyMode = False
    On Error GoTo FailEnsure
    EnsureInventoryManagementSurface wb, report, False
    DeleteWorksheetSurface wb, "ShippingBOM"
    DeleteWorksheetSurface wb, "AggregateBoxBOM_Log"
    DeleteWorksheetSurface wb, "AggregatePackages_Log"
    HideWorksheetSurface wb, "ShipmentsTally"
    HideWorksheetSurface wb, "InventoryManagement"
    HideWorksheetSurface wb, SHIPPING_BACKEND_SHEET

    EnsureShippingWorkbookSurface = True
    inProgress = False
    Exit Function

FailEnsure:
    report = "EnsureShippingWorkbookSurface failed: " & Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    inProgress = False
    On Error GoTo 0
End Function

Public Function EnsureProductionWorkbookSurface(Optional ByVal targetWb As Workbook = Nothing, _
                                                Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Set wb = ResolveTargetWorkbookSurface(targetWb)

    EnsureTableSurface wb, "Production", "RB_AddRecipeName", Array("RECIPE_NAME", "RECIPE_ID", "DESCRIPTION", "GUID"), True
    EnsureTableSurface wb, "Production", "RecipeBuilder", Array("PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT", "OOO", "INSTRUCTION", "RECIPE_LIST_ROW", "INGREDIENT_ID", "GUID"), True
    EnsureTableSurface wb, "Production", "IP_ChooseRecipe", Array("RECIPE_NAME", "DESCRIPTION", "GUID", "RECIPE_ID"), True
    EnsureTableSurface wb, "Production", "IP_ChooseIngredient", Array("INGREDIENT", "UOM", "QUANTITY", "DESCRIPTION", "GUID", "RECIPE_ID", "INGREDIENT_ID", "PROCESS"), True
    EnsureTableSurface wb, "Production", "IP_ChooseItem", Array("ITEMS", "UOM", "DESCRIPTION", "ROW", "RECIPE_ID", "INGREDIENT_ID"), True
    EnsureTableSurface wb, "Production", "RC_RecipeChoose", Array("RECIPE", "RECIPE_ID", "DESCRIPTION", "DEPARTMENT", "PROCESS"), True
    EnsureTableSurface wb, "Production", "RecipeChooser_generated", Array("PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT NEEDED", "INGREDIENT_ID", "RECIPE_LIST_ROW"), False
    EnsureTableSurface wb, "Production", "InventoryPalette_generated", Array("ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", "QUANTITY", "PROCESS", "LOCATION", "ROW", "INPUT/OUTPUT"), False
    EnsureTableSurface wb, "Production", "ProductionOutput", Array("PROCESS", "OUTPUT", "UOM", "REAL OUTPUT", "BATCH", "RECALL CODE", "ROW"), False
    EnsureTableSurface wb, "Production", "Prod_invSys_Check", Array("ROW", "ITEM_CODE", "ITEM", "UOM", "USED", "TOTAL INV"), False
    EnsureTableSurface wb, "Recipes", "Recipes", Array("RECIPE", "RECIPE_ID", "DESCRIPTION", "DEPARTMENT", "PROCESS", "DIAGRAM_ID", "INPUT/OUTPUT", "INGREDIENT", "PERCENT", "UOM", "AMOUNT", "RECIPE_LIST_ROW", "INGREDIENT_ID", "GUID"), False
    EnsureTableSurface wb, ResolveIngredientPaletteSheetSurface(wb), "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "INPUT/OUTPUT", "ITEM", "PERCENT", "UOM", "AMOUNT", "ROW", "GUID"), False
    EnsureTableSurface wb, "TemplatesTable", "TemplatesTable", Array("TEMPLATE_SCOPE", "RECIPE_ID", "INGREDIENT_ID", "PROCESS", "TARGET_TABLE", "TARGET_COLUMN", "FORMULA", "GUID", "NOTES", "ACTIVE", "CREATED_AT", "UPDATED_AT"), False
    EnsureTableSurface wb, "ProductionLog", "ProductionLog", Array("TIMESTAMP", "USER", "RECIPE", "RECIPE_ID", "DEPARTMENT", "DESCRIPTION", "PROCESS", "OUTPUT", "PREDICTED OUTPUT", "REAL OUTPUT", "BATCH", "BATCH_ID", "RECALL CODE", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW", "INPUT/OUTPUT", "INGREDIENT_ID", "GUID"), False
    EnsureTableSurface wb, "BatchCodesLog", "BatchCodesLog", Array("RECIPE", "RECIPE_ID", "PROCESS", "OUTPUT", "UOM", "REAL OUTPUT", "BATCH", "RECALL CODE", "TIMESTAMP", "LOCATION", "USER", "GUID"), False
    EnsureInventoryManagementSurface wb
    FormatWorkbookSurface wb

    EnsureProductionWorkbookSurface = True
    Exit Function

FailEnsure:
    report = "EnsureProductionWorkbookSurface failed: " & Err.Description
End Function

Public Function EnsureAdminLegacyWorkbookSurface(Optional ByVal targetWb As Workbook = Nothing, _
                                                 Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Set wb = ResolveTargetWorkbookSurface(targetWb)

    EnsureTableSurface wb, "UserCredentials", "UserCredentials", Array("USER_ID", "USERNAME", "PIN", "ROLE", "STATUS", "LAST LOGIN"), False
    EnsureTableSurface wb, "Emails", "Emails", Array("EMAIL_ID", "EMAIL_ADDRESS", "DISPLAY_NAME", "STATUS"), False
    FormatWorkbookSurface wb

    EnsureAdminLegacyWorkbookSurface = True
    Exit Function

FailEnsure:
    report = "EnsureAdminLegacyWorkbookSurface failed: " & Err.Description
End Function

Public Function EnsureInventoryManagementSurface(Optional ByVal targetWb As Workbook = Nothing, _
                                                 Optional ByRef report As String = "", _
                                                 Optional ByVal applyPresentation As Boolean = True) As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Set wb = ResolveTargetWorkbookSurface(targetWb)

    EnsureTableSurface wb, "InventoryManagement", "invSys", InventoryManagementHeadersSurface(), False
    RemoveInventoryDomainSupportSurface wb
    If applyPresentation Then ApplyInventoryManagementPresentationSurface wb

    EnsureInventoryManagementSurface = True
    Exit Function

FailEnsure:
    report = "EnsureInventoryManagementSurface failed: " & Err.Description
End Function

Private Function InventoryManagementHeadersSurface() As Variant
    InventoryManagementHeadersSurface = Array( _
        "ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION", "VENDOR(s)", "VENDOR_CODE", "CATEGORY", _
        "RECEIVED", "USED", "MADE", "SHIPMENTS", "TOTAL INV", "LAST EDITED", "TOTAL INV LAST EDIT", _
        "QtyAvailable", "LocationSummary", "LastRefreshUTC", _
        "SnapshotId", "SourceType", "IsStale")
End Function

Private Function ShippingBomViewHeadersSurface() As Variant
    ShippingBomViewHeadersSurface = Array( _
        "PackageRow", "PackageItem", "PackageUOM", "PackageLocation", "PackageDescription", _
        "BomVersion", "BomVersionLabel", "IsActive", "EffectiveFromUTC", "EffectiveToUTC", "RetiredAtUTC", _
        "ComponentRow", "ComponentItem", "ComponentQty", "ComponentUOM", "ComponentLocation", "ComponentDescription", _
        "UpdatedAtUTC", "UpdatedBy")
End Function

Private Sub EnsureInventoryDomainSupportSurface(ByVal wb As Workbook)
    EnsureTableSurface wb, "InventoryLog", "tblInventoryLog", _
        Array("EventID", "UndoOfEventId", "AppliedSeq", "EventType", "OccurredAtUTC", "AppliedAtUTC", _
              "WarehouseId", "StationId", "UserId", "SKU", "QtyDelta", "Location", "Note"), False

    EnsureTableSurface wb, "AppliedEvents", "tblAppliedEvents", _
        Array("EventID", "UndoOfEventId", "AppliedSeq", "AppliedAtUTC", "RunId", "SourceInbox", "Status"), False

    EnsureTableSurface wb, "Locks", "tblLocks", _
        Array("LockName", "OwnerStationId", "OwnerUserId", "RunId", "AcquiredAtUTC", "ExpiresAtUTC", "HeartbeatAtUTC", "Status"), False
End Sub

Private Sub RemoveInventoryDomainSupportSurface(ByVal wb As Workbook)
    DeleteWorksheetSurface wb, "InventoryLog"
    DeleteWorksheetSurface wb, "AppliedEvents"
    DeleteWorksheetSurface wb, "Locks"
End Sub

Private Sub ApplyInventoryManagementPresentationSurface(ByVal wb As Workbook)
    Dim lo As ListObject
    Dim visibleCols As Variant
    Dim hiddenCols As Variant
    Dim key As Variant

    Set lo = FindTableByNameSurface(wb, "invSys")
    If lo Is Nothing Then Exit Sub

    visibleCols = Array("ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION", "VENDOR(s)", "CATEGORY", _
                        "RECEIVED", "USED", "MADE", "SHIPMENTS", "TOTAL INV", "QtyAvailable", "LocationSummary", "LAST EDITED", _
                        "LastRefreshUTC", "SnapshotId", "SourceType", "IsStale")
    hiddenCols = Array("ROW", "VENDOR_CODE", "TOTAL INV LAST EDIT")

    For Each key In visibleCols
        SetTableColumnHiddenSurface lo, CStr(key), False
    Next key
    For Each key In hiddenCols
        SetTableColumnHiddenSurface lo, CStr(key), True
    Next key

    ApplyInventoryManagementFormatsSurface lo
End Sub

Private Sub SetTableColumnHiddenSurface(ByVal lo As ListObject, ByVal columnName As String, ByVal isHidden As Boolean)
    Dim idx As Long

    idx = GetColumnIndexSurface(lo, columnName)
    If idx = 0 Then Exit Sub

    On Error Resume Next
    lo.ListColumns(idx).Range.EntireColumn.Hidden = isHidden
    On Error GoTo 0
End Sub

Private Sub ApplyInventoryManagementFormatsSurface(ByVal lo As ListObject)
    Dim qtyCols As Variant
    Dim dateCols As Variant
    Dim key As Variant
    Dim idx As Long

    If lo Is Nothing Then Exit Sub

    qtyCols = Array("RECEIVED", "USED", "MADE", "SHIPMENTS", "TOTAL INV", "QtyAvailable")
    For Each key In qtyCols
        idx = GetColumnIndexSurface(lo, CStr(key))
        If idx > 0 Then lo.ListColumns(idx).Range.NumberFormat = "0.########"
    Next key

    dateCols = Array("LAST EDITED", "TOTAL INV LAST EDIT", "LastRefreshUTC")
    For Each key In dateCols
        idx = GetColumnIndexSurface(lo, CStr(key))
        If idx > 0 Then lo.ListColumns(idx).Range.NumberFormat = "yyyy-mm-dd hh:mm:ss"
    Next key
End Sub

Private Function ResolveTargetWorkbookSurface(ByVal targetWb As Workbook) As Workbook
    If targetWb Is Nothing Then
        Set ResolveTargetWorkbookSurface = ThisWorkbook
    Else
        Set ResolveTargetWorkbookSurface = targetWb
    End If
End Function

Private Function ResolveIngredientPaletteSheetSurface(ByVal wb As Workbook) As String
    If Not EnsureWorksheetSurface(wb, "IngredientsPalette") Is Nothing Then
        ResolveIngredientPaletteSheetSurface = "IngredientsPalette"
        Exit Function
    End If
    ResolveIngredientPaletteSheetSurface = "IngredientPalette"
End Function

Public Function ShouldBootstrapRoleWorkbookSurface(Optional ByVal targetWb As Workbook = Nothing) As Boolean
    Dim wb As Workbook
    Dim wbName As String

    Set wb = targetWb
    If wb Is Nothing Then Exit Function
    If wb.IsAddin Then Exit Function

    wbName = LCase$(Trim$(wb.Name))
    If wbName = "" Then Exit Function
    If Left$(wbName, 2) = "~$" Then Exit Function
    If wbName = "personal.xlsb" Then Exit Function
    If wbName Like "*inventory_management*.xls*" Then Exit Function
    If wbName Like "*.xla" Or wbName Like "*.xlam" Then Exit Function
    If IsRuntimeWorkbookNameSurface(wbName) Then Exit Function

    ShouldBootstrapRoleWorkbookSurface = True
End Function

Private Function IsRuntimeWorkbookNameSurface(ByVal wbName As String) As Boolean
    If wbName Like "*.invsys.*.xlsb" _
       Or wbName Like "*.invsys.*.xlsx" _
       Or wbName Like "*.invsys.*.xlsm" Then
        IsRuntimeWorkbookNameSurface = True
        Exit Function
    End If

    If wbName Like "invsys.inbox.*.xlsb" _
       Or wbName Like "invsys.inbox.*.xlsx" _
       Or wbName Like "invsys.inbox.*.xlsm" Then
        IsRuntimeWorkbookNameSurface = True
        Exit Function
    End If

    If wbName Like "*.outbox.events.xlsb" _
       Or wbName Like "*.outbox.events.xlsx" _
       Or wbName Like "*.outbox.events.xlsm" _
       Or wbName Like "*.snapshot.inventory.xlsb" _
       Or wbName Like "*.snapshot.inventory.xlsx" _
       Or wbName Like "*.snapshot.inventory.xlsm" Then
        IsRuntimeWorkbookNameSurface = True
    End If
End Function

Private Sub EnsureTableSurface(ByVal wb As Workbook, _
                               ByVal sheetName As String, _
                               ByVal tableName As String, _
                               ByVal headers As Variant, _
                               ByVal seedEntryRow As Boolean, _
                               Optional ByVal startAddress As String = "")
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim startCell As Range
    Dim dataRange As Range

    Set ws = EnsureWorksheetSurface(wb, sheetName)
    EnsureWorksheetEditableSurface ws

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If lo Is Nothing Then
        If Trim$(startAddress) <> "" Then
            Set startCell = ws.Range(startAddress)
        Else
            Set startCell = GetNextTableStartCellSurface(ws)
        End If
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i

        Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = tableName
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumnSurface lo, CStr(headers(i))
    Next i
    PruneInventoryAliasColumnsSurface lo
    RemoveAutogeneratedColumnsSurface lo

    If seedEntryRow Then
        EnsureTableHasDataRowSurface lo
    ElseIf Not lo.DataBodyRange Is Nothing Then
        If lo.ListRows.Count = 1 And TableRowIsBlankSurface(lo, 1) Then lo.ListRows(1).Delete
    End If
End Sub

Private Sub ArrangeReceivingTablesSurface(ByVal wb As Workbook)
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets("ReceivedTally")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    MoveTableTopLeftSurface ws, "ReceivedTally", "C3"
    MoveTableTopLeftSurface ws, "AggregateReceived", "J3"
    MoveTableTopLeftSurface ws, "invSysData_Receiving", "V3"
End Sub

Private Sub MoveTableTopLeftSurface(ByVal ws As Worksheet, ByVal tableName As String, ByVal targetAddress As String)
    Dim lo As ListObject
    Dim targetCell As Range

    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    Set targetCell = ws.Range(targetAddress)
    On Error GoTo 0
    If lo Is Nothing Or targetCell Is Nothing Then Exit Sub
    If lo.Range.Cells(1, 1).Address(False, False) = targetCell.Address(False, False) Then Exit Sub

    RebuildTableAtSurface lo, targetCell
End Sub

Private Sub MoveTableToSheetSurface(ByVal wb As Workbook, ByVal tableName As String, ByVal targetSheetName As String)
    Dim lo As ListObject
    Dim targetWs As Worksheet
    Dim targetCell As Range

    If wb Is Nothing Then Exit Sub
    Set lo = FindTableByNameSurface(wb, tableName)
    If lo Is Nothing Then Exit Sub
    Set targetWs = EnsureWorksheetSurface(wb, targetSheetName)
    If lo.Parent Is targetWs Then Exit Sub

    EnsureWorksheetEditableSurface lo.Parent
    EnsureWorksheetEditableSurface targetWs
    targetWs.Visible = xlSheetVisible
    Set targetCell = GetNextTableStartCellSurface(targetWs)

    RebuildTableAtSurface lo, targetCell
End Sub

Private Sub RebuildTableAtSurface(ByVal lo As ListObject, ByVal targetCell As Range)
    On Error GoTo CleanExit

    Dim sourceRange As Range
    Dim targetRange As Range
    Dim data As Variant
    Dim tableName As String
    Dim tableStyle As String
    Dim showTotals As Boolean
    Dim rowCount As Long
    Dim colCount As Long
    Dim newLo As ListObject

    If lo Is Nothing Or targetCell Is Nothing Then Exit Sub
    Set sourceRange = lo.Range
    If sourceRange Is Nothing Then Exit Sub
    rowCount = sourceRange.Rows.Count
    colCount = sourceRange.Columns.Count
    If rowCount <= 0 Or colCount <= 0 Then Exit Sub

    data = sourceRange.Value
    tableName = lo.Name
    tableStyle = lo.TableStyle
    showTotals = lo.ShowTotals

    lo.Unlist
    sourceRange.Clear
    Set targetRange = targetCell.Resize(rowCount, colCount)
    targetRange.Clear
    targetRange.Value = data
    Set newLo = targetCell.Worksheet.ListObjects.Add(xlSrcRange, targetRange, , xlYes)
    newLo.Name = tableName
    If Trim$(tableStyle) <> "" Then newLo.TableStyle = tableStyle
    newLo.ShowTotals = showTotals

CleanExit:
    On Error Resume Next
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

Private Sub ArrangeShippingBackendTablesSurface(ByVal wb As Workbook)
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    Set ws = EnsureWorksheetSurface(wb, SHIPPING_BACKEND_SHEET)
    ws.Visible = xlSheetVisible

    MoveTableTopLeftSurface ws, "ShipmentsTally", "A1"
    MoveTableTopLeftSurface ws, "NotShipped", "M1"
    MoveTableTopLeftSurface ws, "AggregateBoxBOM", "Y1"
    MoveTableTopLeftSurface ws, "AggregatePackages", "AF1"
    MoveTableTopLeftSurface ws, "Check_invSys", "AM1"
    MoveTableTopLeftSurface ws, "invSysData_Shipping", "AW1"
    MoveTableTopLeftSurface ws, "ShippingBOMView", "BT1"
    MoveTableTopLeftSurface ws, "AggregateBoxBOM_Log", "CO1"
    MoveTableTopLeftSurface ws, "AggregatePackages_Log", "CY1"
    On Error Resume Next
    Application.CutCopyMode = False
    On Error GoTo 0
End Sub

Private Sub HideWorksheetSurface(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    If Not ws Is Nothing Then ws.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

Private Sub EnsureReceivingButtonsSurface(ByVal wb As Workbook)
    Const BTN_TOP As Double = 6
    Const BTN_HEIGHT As Double = 20
    Const BTN_WIDTH As Double = 118
    Const BTN_SPACING As Double = 8

    Dim ws As Worksheet
    Dim leftPos As Double

    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    Set ws = wb.Worksheets("ReceivedTally")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    DeleteReceivingButtonsSurface ws

    leftPos = ws.Range("C1").Left
    EnsureReceivingButtonSurface ws, "btnConfirmWrites", "Confirm Writes", "'invSys.Receiving.xlam'!modTS_Received.ConfirmWrites", leftPos, BTN_TOP, BTN_WIDTH, BTN_HEIGHT
    leftPos = leftPos + BTN_WIDTH + BTN_SPACING
    EnsureReceivingButtonSurface ws, "btnUndoMacro", "Undo", "'invSys.Receiving.xlam'!modTS_Received.MacroUndo", leftPos, BTN_TOP, 82, BTN_HEIGHT
    leftPos = leftPos + 82 + BTN_SPACING
    EnsureReceivingButtonSurface ws, "btnRedoMacro", "Redo", "'invSys.Receiving.xlam'!modTS_Received.MacroRedo", leftPos, BTN_TOP, 82, BTN_HEIGHT
End Sub

Private Sub EnsureReceivingButtonSurface(ByVal ws As Worksheet, _
                                         ByVal shapeName As String, _
                                         ByVal caption As String, _
                                         ByVal onActionMacro As String, _
                                         ByVal leftPos As Double, _
                                         ByVal topPos As Double, _
                                         ByVal widthPts As Double, _
                                         ByVal heightPts As Double)
    Dim shp As Shape

    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    Set shp = ws.Shapes(shapeName)
    On Error GoTo 0

    If shp Is Nothing Then
        Set shp = ws.Shapes.AddFormControl(xlButtonControl, leftPos, topPos, widthPts, heightPts)
        shp.Name = shapeName
    Else
        shp.Left = leftPos
        shp.Top = topPos
        shp.Width = widthPts
        shp.Height = heightPts
    End If

    On Error Resume Next
    shp.TextFrame.Characters.Text = caption
    shp.OnAction = onActionMacro
    On Error GoTo 0
End Sub

Private Sub DeleteReceivingButtonsSurface(ByVal ws As Worksheet)
    Dim shp As Shape
    Dim names As Collection
    Dim item As Variant
    Dim shpName As String

    If ws Is Nothing Then Exit Sub
    Set names = New Collection
    For Each shp In ws.Shapes
        shpName = LCase$(Trim$(shp.Name))
        If shpName = "btnconfirmwrites" Or shpName = "btnundomacro" Or shpName = "btnredomacro" Then names.Add shp.Name
    Next shp

    For Each item In names
        On Error Resume Next
        ws.Shapes(CStr(item)).Delete
        On Error GoTo 0
    Next item
End Sub

Private Sub PruneInventoryAliasColumnsSurface(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub

    Select Case LCase$(Trim$(lo.Name))
        Case "invsys", "invsysdata_receiving", "invsysdata_shipping"
            DeleteListColumnIfPresentSurface lo, "SKU"
            DeleteListColumnIfPresentSurface lo, "ItemName"
            DeleteListColumnIfPresentSurface lo, "QtyOnHand"
            DeleteListColumnIfPresentSurface lo, "LastAppliedUTC"
            DeleteListColumnIfPresentSurface lo, "TIMESTAMP"
    End Select
End Sub

Private Sub DeleteListColumnIfPresentSurface(ByVal lo As ListObject, ByVal columnName As String)
    Dim idx As Long

    idx = GetColumnIndexSurface(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.ListColumns(idx).Delete
End Sub

Private Sub RemoveAutogeneratedColumnsSurface(ByVal lo As ListObject)
    Dim i As Long

    If lo Is Nothing Then Exit Sub
    For i = lo.ListColumns.Count To 1 Step -1
        If LCase$(Left$(Trim$(lo.ListColumns(i).Name), 6)) = "column" Then
            lo.ListColumns(i).Delete
        End If
    Next i
End Sub

Private Function FindTableByNameSurface(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTableByNameSurface = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTableByNameSurface Is Nothing Then Exit Function
    Next ws
End Function

Private Function EnsureWorksheetSurface(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheetSurface = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheetSurface Is Nothing Then
        Set EnsureWorksheetSurface = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheetSurface.Name = sheetName
    End If
End Function

Private Function GetNextTableStartCellSurface(ByVal ws As Worksheet) As Range
    Dim lo As ListObject
    Dim maxRow As Long

    For Each lo In ws.ListObjects
        If lo.Range.Row + lo.Range.Rows.Count > maxRow Then
            maxRow = lo.Range.Row + lo.Range.Rows.Count
        End If
    Next lo

    If maxRow = 0 Then
        Set GetNextTableStartCellSurface = ws.Range("A1")
    Else
        Set GetNextTableStartCellSurface = ws.Cells(maxRow + 2, 1)
    End If
End Function

Private Sub EnsureListColumnSurface(ByVal lo As ListObject, ByVal columnName As String)
    If GetColumnIndexSurface(lo, columnName) > 0 Then Exit Sub
    lo.ListColumns.Add lo.ListColumns.Count + 1
    lo.ListColumns(lo.ListColumns.Count).Name = columnName
End Sub

Private Function GetColumnIndexSurface(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexSurface = i
            Exit Function
        End If
    Next i
End Function

Private Sub EnsureTableHasDataRowSurface(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

Private Function TableRowIsBlankSurface(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    Dim cell As Range

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then
        TableRowIsBlankSurface = True
        Exit Function
    End If

    For Each cell In lo.ListRows(rowIndex).Range.Cells
        If Trim$(CStr(cell.Value)) <> "" Then Exit Function
    Next cell
    TableRowIsBlankSurface = True
End Function

Private Sub EnsureWorksheetEditableSurface(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 2751, "modRoleWorkbookSurfaces.EnsureWorksheetEditableSurface", _
                  "Worksheet '" & ws.Name & "' is protected and could not be unprotected before updating workbook surfaces."
    End If
End Sub

Private Sub DeleteWorksheetSurface(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    Dim prevAlerts As Boolean

    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    EnsureWorksheetEditableSurface ws
    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = prevAlerts
End Sub

Private Sub FormatWorkbookSurface(ByVal wb As Workbook)
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        ws.Cells.EntireColumn.AutoFit
        ws.Rows(1).Font.Bold = True
        ApplyWorksheetTabColorSurface ws
    Next ws
End Sub

Private Sub ApplyWorksheetTabColorSurface(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Tab.Color = ResolveWorksheetTabColorSurface(ws.Name)
    On Error GoTo 0
End Sub

Private Function ResolveWorksheetTabColorSurface(ByVal sheetName As String) As Long
    Select Case LCase$(Trim$(sheetName))
        Case "receivedtally", "receivedlog"
            ResolveWorksheetTabColorSurface = RGB(217, 119, 6)
        Case "shipmentstally", "aggregateboxbom_log", "aggregatepackages_log"
            ResolveWorksheetTabColorSurface = RGB(3, 105, 161)
        Case "recipes", "ingredientpalette", "ingredientspalette", "shippingbom", "templatestable"
            ResolveWorksheetTabColorSurface = RGB(190, 24, 93)
        Case "production", "productionlog", "batchcodeslog"
            ResolveWorksheetTabColorSurface = RGB(22, 101, 52)
        Case "inventorymanagement"
            ResolveWorksheetTabColorSurface = RGB(109, 40, 217)
        Case "usercredentials", "emails"
            ResolveWorksheetTabColorSurface = RGB(161, 98, 7)
        Case Else
            ResolveWorksheetTabColorSurface = RGB(107, 114, 128)
    End Select
End Function
