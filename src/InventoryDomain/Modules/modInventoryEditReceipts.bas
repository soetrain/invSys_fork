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
