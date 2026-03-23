Attribute VB_Name = "modRoleWorkbookSurfaces"
Option Explicit

Public Function EnsureReceivingWorkbookSurface(Optional ByVal targetWb As Workbook = Nothing, _
                                               Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Set wb = ResolveTargetWorkbookSurface(targetWb)

    EnsureTableSurface wb, "ReceivedTally", "ReceivedTally", Array("REF_NUMBER", "ITEMS", "QUANTITY", "ROW"), True
    EnsureTableSurface wb, "ReceivedTally", "AggregateReceived", Array("REF_NUMBER", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "DESCRIPTION", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW"), False
    EnsureTableSurface wb, "ReceivedTally", "invSysData_Receiving", InventoryManagementHeadersSurface(), False
    EnsureInventoryManagementSurface wb
    EnsureTableSurface wb, "ReceivedLog", "ReceivedLog", Array("SNAPSHOT_ID", "ENTRY_DATE", "REF_NUMBER", "ITEMS", "QUANTITY", "UOM", "VENDOR", "LOCATION", "ITEM_CODE", "ROW"), False
    FormatWorkbookSurface wb

    EnsureReceivingWorkbookSurface = True
    Exit Function

FailEnsure:
    report = "EnsureReceivingWorkbookSurface failed: " & Err.Description
End Function

Public Function EnsureShippingWorkbookSurface(Optional ByVal targetWb As Workbook = Nothing, _
                                              Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Set wb = ResolveTargetWorkbookSurface(targetWb)

    EnsureTableSurface wb, "ShipmentsTally", "ShipmentsTally", Array("REF_NUMBER", "ITEMS", "QUANTITY", "ROW", "UOM", "LOCATION", "DESCRIPTION"), True
    EnsureTableSurface wb, "ShipmentsTally", "NotShipped", Array("REF_NUMBER", "ITEMS", "QUANTITY", "ROW", "UOM", "LOCATION", "DESCRIPTION"), False
    EnsureTableSurface wb, "ShipmentsTally", "AggregateBoxBOM", Array("ROW", "ITEM_CODE", "ITEM", "QUANTITY", "UOM", "LOCATION"), False
    EnsureTableSurface wb, "ShipmentsTally", "AggregatePackages", Array("ROW", "ITEM_CODE", "ITEM", "QUANTITY", "UOM", "LOCATION"), False
    EnsureTableSurface wb, "ShipmentsTally", "BoxBuilder", Array("Box Name", "UOM", "LOCATION", "DESCRIPTION", "ROW"), True
    EnsureTableSurface wb, "ShipmentsTally", "BoxBOM", Array("ITEM", "ROW", "QUANTITY", "UOM", "LOCATION", "DESCRIPTION"), True
    EnsureTableSurface wb, "ShipmentsTally", "Check_invSys", Array("ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "USED", "MADE", "SHIPMENTS", "TOTAL INV"), False
    EnsureTableSurface wb, "ShipmentsTally", "invSysData_Shipping", InventoryManagementHeadersSurface(), False
    EnsureTableSurface wb, "AggregateBoxBOM_Log", "AggregateBoxBOM_Log", Array("GUID", "USER", "ACTION", "ROW", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP"), False
    EnsureTableSurface wb, "AggregatePackages_Log", "AggregatePackages_Log", Array("GUID", "USER", "ACTION", "ROW", "ITEM_CODE", "ITEM", "QTY_DELTA", "NEW_VALUE", "TIMESTAMP"), False
    EnsureInventoryManagementSurface wb
    EnsureWorksheetSurface wb, "ShippingBOM"
    FormatWorkbookSurface wb

    EnsureShippingWorkbookSurface = True
    Exit Function

FailEnsure:
    report = "EnsureShippingWorkbookSurface failed: " & Err.Description
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
    EnsureTableSurface wb, "IngredientPalette", "IngredientPalette", Array("RECIPE_ID", "INGREDIENT_ID", "INPUT/OUTPUT", "ITEM", "PERCENT", "UOM", "AMOUNT", "ROW", "GUID"), False
    EnsureTableSurface wb, "TemplatesTable", "TemplatesTable", Array("TEMPLATE_SCOPE", "RECIPE_ID", "INGREDIENT_ID", "PROCESS", "TARGET_TABLE", "TARGET_COLUMN", "FORMULA", "GUID", "NOTES", "ACTIVE", "CREATED_AT", "UPDATED_AT"), False
    EnsureTableSurface wb, "ProductionLog", "ProductionLog", Array("TIMESTAMP", "RECIPE", "RECIPE_ID", "DEPARTMENT", "DESCRIPTION", "PROCESS", "OUTPUT", "PREDICTED OUTPUT", "REAL OUTPUT", "BATCH", "BATCH_ID", "RECALL CODE", "ITEM_CODE", "VENDORS", "VENDOR_CODE", "ITEM", "UOM", "QUANTITY", "LOCATION", "ROW", "INPUT/OUTPUT", "INGREDIENT_ID", "GUID"), False
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
                                                 Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Set wb = ResolveTargetWorkbookSurface(targetWb)

    EnsureTableSurface wb, "InventoryManagement", "invSys", InventoryManagementHeadersSurface(), False

    EnsureInventoryManagementSurface = True
    Exit Function

FailEnsure:
    report = "EnsureInventoryManagementSurface failed: " & Err.Description
End Function

Private Function InventoryManagementHeadersSurface() As Variant
    InventoryManagementHeadersSurface = Array( _
        "ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION", "VENDOR(s)", "VENDOR_CODE", "CATEGORY", _
        "RECEIVED", "USED", "MADE", "SHIPMENTS", "TOTAL INV", "LAST EDITED", "TOTAL INV LAST EDIT", "TIMESTAMP")
End Function

Private Function ResolveTargetWorkbookSurface(ByVal targetWb As Workbook) As Workbook
    If targetWb Is Nothing Then
        Set ResolveTargetWorkbookSurface = ThisWorkbook
    Else
        Set ResolveTargetWorkbookSurface = targetWb
    End If
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
    If wbName Like "*.xla" Or wbName Like "*.xlam" Then Exit Function
    If InStr(1, wbName, ".invsys.", vbTextCompare) > 0 Then Exit Function

    ShouldBootstrapRoleWorkbookSurface = True
End Function

Private Sub EnsureTableSurface(ByVal wb As Workbook, _
                               ByVal sheetName As String, _
                               ByVal tableName As String, _
                               ByVal headers As Variant, _
                               ByVal seedEntryRow As Boolean)
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
        Set startCell = GetNextTableStartCellSurface(ws)
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

    If seedEntryRow Then
        EnsureTableHasDataRowSurface lo
    ElseIf Not lo.DataBodyRange Is Nothing Then
        If lo.ListRows.Count = 1 And TableRowIsBlankSurface(lo, 1) Then lo.ListRows(1).Delete
    End If
End Sub

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

Private Sub FormatWorkbookSurface(ByVal wb As Workbook)
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        ws.Cells.EntireColumn.AutoFit
        ws.Rows(1).Font.Bold = True
    Next ws
End Sub
