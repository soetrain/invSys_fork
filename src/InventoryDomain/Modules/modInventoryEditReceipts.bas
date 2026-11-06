Attribute VB_Name = "modInventoryEditReceipts"
Option Explicit

' Declared Core/Domain boundary: context/command strings in, Boolean/report
' scalar array out. The canonical receipt is owned here, never by a form.
Public Function SaveCatalog(ByVal context As String, ByVal command As String) As Variant
    Dim plan As Variant, book As Workbook, candidate As Workbook, catalog As ListObject
    Dim receipts As ListObject, sheet As Worksheet, names As Variant, values As Variant
    Dim receiptNames As Variant, receiptValues As Variant, index As Long, row As Long, selected As Long
    Dim report As String, saved As Boolean, events As Boolean, security As Long, owned As Boolean
    On Error GoTo Failed
    events = Application.EnableEvents: security = Application.AutomationSecurity
    plan = modInventoryEditCommands.ClaimWritePlan(context, command)
    If Not IsArray(plan) Then GoTo Failed
    ' An open source belongs to its existing owner. Never save or close it here.
    For Each candidate In Application.Workbooks
        If StrComp(candidate.FullName, CStr(plan(0)), vbTextCompare) = 0 Then GoTo Failed
    Next candidate
    If Not modInventoryCommands.CheckForm(context, report) Then GoTo Failed
    Application.EnableEvents = False: Application.AutomationSecurity = 3
    Set book = Application.Workbooks.Open(CStr(plan(0)), UpdateLinks:=0, ReadOnly:=False, _
        IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    owned = True
    If book.ReadOnly Or Not book.Saved Then GoTo Failed
    Set catalog = FindTable(book, "tblSkuCatalog")
    If catalog Is Nothing Then GoTo Failed
    receiptNames = plan(3): receiptValues = plan(4)
    index = ColumnIndex(catalog, "SKU")
    If index = 0 Then GoTo Failed
    For row = 1 To catalog.ListRows.Count
        If CStr(catalog.DataBodyRange.Cells(row, index).Value2) = CStr(receiptValues(4)) Then
            If selected <> 0 Then GoTo Failed
            selected = row
        End If
    Next row
    If selected = 0 Then GoTo Failed
    If Not modInventoryCommands.BindCatalogWriter(context, book.Name, report) Then GoTo Failed
    names = plan(1): values = plan(2)
    If Not modInventoryCommands.CheckForm(context, report) Then GoTo Failed
    If catalog.Parent.ProtectContents Then catalog.Parent.Unprotect
    For index = LBound(names) To UBound(names)
        If Not modInventoryCommands.CheckForm(context, report) Then GoTo Failed
        WriteCell catalog, selected, CStr(names(index)), values(index), True
    Next index
    If Not modInventoryCommands.CheckForm(context, report) Then GoTo Failed
    catalog.Parent.Protect UserInterfaceOnly:=True
    Set receipts = FindTable(book, "tblInventoryEditReceipts")
    If receipts Is Nothing Then
        ' Only this explicit owning EDIT initializes this single private table.
        If Not modInventoryCommands.CheckForm(context, report) Then GoTo Failed
        Set sheet = book.Worksheets.Add(After:=book.Worksheets(book.Worksheets.Count))
        For index = 0 To UBound(receiptNames)
            sheet.Cells(1, index + 1).Value2 = CStr(receiptNames(index))
        Next index
        Set receipts = sheet.ListObjects.Add(xlSrcRange, _
            sheet.Range(sheet.Cells(1, 1), sheet.Cells(2, UBound(receiptNames) + 1)), , xlYes)
        receipts.Name = "tblInventoryEditReceipts"
        sheet.Visible = xlSheetVeryHidden
    End If
    If Not modInventoryCommands.CheckForm(context, report) Then GoTo Failed
    row = receipts.ListRows.Add.Index
    For index = 0 To UBound(receiptNames)
        If Not modInventoryCommands.CheckForm(context, report) Then GoTo Failed
        WriteCell receipts, row, CStr(receiptNames(index)), receiptValues(index), False
    Next index
    If Not modInventoryCommands.BeginCatalogSave(context, report) Then GoTo Failed
    book.Save
    If Not book.Saved Then GoTo Failed
    If Not modInventoryCommands.ConfirmCatalog(context, report) Then GoTo Failed
    saved = True
    report = "Catalog changes and the EDIT receipt were saved."
    GoTo Finished
Failed:
    report = "Partial result: catalog/receipt durability is unconfirmed. Reopen Add/Edit Inventory before retrying."
Finished:
    On Error Resume Next
    If owned Then book.Close SaveChanges:=False
    Application.AutomationSecurity = security
    Application.EnableEvents = events
    On Error GoTo 0
    SaveCatalog = Array(saved, report)
End Function

' Domain-internal objects stay in this project. Only an exact existing receipt
' distinguishes the worksheet quantity phase from retained ordinary CREATE calls.
Public Function ReadCreationPhase(ByVal book As Workbook, ByVal evt As Object, _
        ByVal eventType As String, ByVal items As Collection, ByRef worksheetCreate As Boolean, _
        ByRef errorCode As String, ByRef errorMessage As String) As Boolean
    Dim table As ListObject, column As Long, row As Long, selected As Long
    Dim headers As Variant, values As Variant, index As Long, item As Object
    On Error GoTo Invalid
    worksheetCreate = False
    ReadCreationPhase = True
    If eventType <> modInventoryApply.EVENT_TYPE_INVENTORY_CREATE Then Exit Function
    Set table = FindTable(book, "tblInventoryEditReceipts")
    If table Is Nothing Then Exit Function
    column = ColumnIndex(table, "PlannedEventId")
    If column = 0 Then Exit Function
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, column).Value2) = CStr(evt("EventID")) Then
            If selected <> 0 Then GoTo Invalid
            selected = row
        End If
    Next row
    If selected = 0 Then Exit Function
    ReadCreationPhase = False
    If items.Count <> 1 Then GoTo Invalid
    headers = modInventoryEditWire.ReceiptHeaders()
    ReDim values(0 To UBound(headers))
    For index = 0 To UBound(headers)
        values(index) = CStr(ReadCell(table, selected, CStr(headers(index))))
    Next index
    Set item = items(1)
    If Not CreationReceiptMatches(headers, values, evt, item) Then GoTo Invalid
    If Not FirstEntityAvailable(book, CStr(values(4)), CStr(values(8))) Then GoTo Invalid
    worksheetCreate = True
    ReadCreationPhase = True
    Exit Function
Invalid:
    ReadCreationPhase = False
    errorCode = "EDIT_CREATION_UNCONFIRMED"
    errorMessage = "The saved worksheet EDIT creation plan could not be verified."
End Function

Private Function CreationReceiptMatches(ByVal headers As Variant, ByVal values As Variant, _
        ByVal evt As Object, ByVal item As Object) As Boolean
    Dim names(0 To 13) As String, bound(0 To 13) As String, index As Long, field As Variant
    On Error GoTo Invalid
    If values(12) <> "CREATE" Or values(13) <> "SAVED" Then Exit Function
    For index = 0 To 13
        names(index) = CStr(headers(index)): bound(index) = CStr(values(index))
    Next index
    If values(14) = "" Or values(14) <> modInventoryEditWire.HashFields(names, bound) Then Exit Function
    For Each field In Array(0, 7, 8)
        If Not modTrainingWire.ValidId(LCase$(CStr(values(field)))) Then Exit Function
    Next field
    If values(0) = values(7) Or values(0) = values(8) Or values(7) = values(8) Then Exit Function
    For index = 1 To 6
        If values(index) = "" Then Exit Function
    Next index
    For index = 1 To 3
        If values(index) <> CStr(evt(headers(index))) Then Exit Function
    Next index
    If values(7) <> CStr(evt("EventID")) Or values(11) <> CStr(evt("PayloadJson")) Then Exit Function
    If Not IsNumeric(values(9)) Or Not IsNumeric(values(10)) Then Exit Function
    If CDbl(values(9)) <> 0 Or CDbl(values(10)) <= 0 Then Exit Function
    If item.Count <> 8 Then Exit Function
    For Each field In Array("System_Key", "SKU", "Qty", "Location", "Condition", "AttributesJson", "Note", "IoType")
        If Not item.Exists(CStr(field)) Then Exit Function
    Next field
    If CStr(item("System_Key")) <> values(8) Or CStr(item("SKU")) <> values(4) Then Exit Function
    If CDbl(item("Qty")) <> CDbl(values(10)) Or CStr(item("IoType")) <> "CREATE" Then Exit Function
    If CStr(item("Condition")) <> "GOOD" Or CStr(item("AttributesJson")) <> "" Then Exit Function
    If Trim$(CStr(item("Note"))) = "" Then Exit Function
    CreationReceiptMatches = True
Invalid:
End Function

Private Function FirstEntityAvailable(ByVal book As Workbook, ByVal sku As String, ByVal key As String) As Boolean
    Dim table As ListObject, row As Long, matches As Long, category As String
    On Error GoTo Invalid
    Set table = FindTable(book, "tblSkuCatalog")
    If table Is Nothing Then Exit Function
    For row = 1 To table.ListRows.Count
        If StrComp(CStr(ReadCell(table, row, "SKU")), sku, vbTextCompare) = 0 Then
            matches = matches + 1
            If CStr(ReadCell(table, row, "SKU")) <> sku Or CStr(ReadCell(table, row, "ITEM_CODE")) <> sku Then Exit Function
            If UCase$(Trim$(CStr(ReadCell(table, row, "CATALOG_STATE")))) <> "ACTIVE" Then Exit Function
            If UCase$(Trim$(CStr(ReadCell(table, row, "TRACK_QTY")))) <> "TRUE" Then Exit Function
            If UCase$(Trim$(CStr(ReadCell(table, row, "ITEM_KIND")))) <> "INVENTORY" Then Exit Function
            category = UCase$(Trim$(CStr(ReadCell(table, row, "CATEGORY"))))
            If category = "UTILITY" Or category = "SERVICE" Then Exit Function
        End If
    Next row
    If matches <> 1 Then Exit Function
    Set table = FindTable(book, "tblInventoryEntities")
    If table Is Nothing Then Exit Function
    For row = 1 To table.ListRows.Count
        If CStr(ReadCell(table, row, "System_Key")) = key Then Exit Function
        If StrComp(CStr(ReadCell(table, row, "SKU")), sku, vbTextCompare) = 0 Then Exit Function
    Next row
    FirstEntityAvailable = True
Invalid:
End Function

Private Function ReadCell(ByVal table As ListObject, ByVal row As Long, ByVal name As String) As Variant
    Dim index As Long
    index = ColumnIndex(table, name)
    If index = 0 Then Err.Raise vbObjectError + 2224, , "Required EDIT receipt or source header is unavailable."
    ReadCell = table.DataBodyRange.Cells(row, index).Value2
End Function

Private Function FindTable(ByVal book As Workbook, ByVal name As String) As ListObject
    Dim sheet As Worksheet, table As ListObject
    For Each sheet In book.Worksheets
        For Each table In sheet.ListObjects
            If StrComp(table.Name, name, vbTextCompare) = 0 Then Set FindTable = table: Exit Function
        Next table
    Next sheet
End Function

Private Function ColumnIndex(ByVal table As ListObject, ByVal name As String) As Long
    Dim column As ListColumn
    For Each column In table.ListColumns
        If StrComp(Trim$(column.Name), name, vbTextCompare) = 0 Then
            If ColumnIndex <> 0 Then Err.Raise vbObjectError + 2222, , "Ambiguous catalog or receipt header."
            ColumnIndex = column.Index
        End If
    Next column
End Function

Private Sub WriteCell(ByVal table As ListObject, ByVal row As Long, ByVal name As String, _
        ByVal value As Variant, ByVal allowCatalogColumn As Boolean)
    Dim index As Long
    index = ColumnIndex(table, name)
    If index = 0 Then
        If Not allowCatalogColumn Then Err.Raise vbObjectError + 2223, , "Required receipt header is unavailable."
        index = table.ListColumns.Add.Index
        table.ListColumns(index).Name = name
    End If
    If allowCatalogColumn Then
        table.DataBodyRange.Cells(row, index).Value = value
    Else
        table.DataBodyRange.Cells(row, index).NumberFormat = "@"
        table.DataBodyRange.Cells(row, index).Value2 = CStr(value)
    End If
End Sub
