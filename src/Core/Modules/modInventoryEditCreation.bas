Attribute VB_Name = "modInventoryEditCreation"
Option Explicit
Option Private Module

' CREATE is allowed only for the first entity of an exact active catalog item.
' An unresolved/retired entity is never treated as an empty inventory history.
Public Function Eligible(ByVal book As Workbook, ByVal sku As String, _
        Optional ByVal plannedKey As String = "") As Boolean
    Dim table As ListObject, fields As Object, keys As Object, row As Long
    Dim key As String, matches As Long
    On Error GoTo Invalid
    Set table = modExecutionTarget.Table(book, "tblSkuCatalog")
    Set fields = modAggregationSources.RequiredColumns(table, Array("SKU", "ITEM_CODE", "CATALOG_STATE"))
    If fields Is Nothing Then Exit Function
    For row = 1 To table.ListRows.Count
        If StrComp(CStr(table.DataBodyRange.Cells(row, fields("SKU")).Value2), sku, vbTextCompare) = 0 Then
            matches = matches + 1
            If CStr(table.DataBodyRange.Cells(row, fields("SKU")).Value2) <> sku Then Exit Function
            If CStr(table.DataBodyRange.Cells(row, fields("ITEM_CODE")).Value2) <> sku Then Exit Function
            If UCase$(Trim$(CStr(table.DataBodyRange.Cells(row, fields("CATALOG_STATE")).Value2))) <> "ACTIVE" Then Exit Function
        End If
    Next row
    If matches <> 1 Then Exit Function
    Set table = modExecutionTarget.Table(book, "tblInventoryEntities")
    Set fields = modAggregationSources.RequiredColumns(table, _
        Array("System_Key", "SKU", "QtyOnHand", "Location", "Condition", "AttributesJson", "InventoryState"))
    If fields Is Nothing Then Exit Function
    Set keys = CreateObject("Scripting.Dictionary"): keys.CompareMode = vbBinaryCompare
    For row = 1 To table.ListRows.Count
        key = CStr(table.DataBodyRange.Cells(row, fields("System_Key")).Value2)
        If key = "" Or keys.Exists(key) Then Exit Function
        keys.Add key, True
        If plannedKey <> "" And key = plannedKey Then Exit Function
        If StrComp(CStr(table.DataBodyRange.Cells(row, fields("SKU")).Value2), sku, vbTextCompare) = 0 Then Exit Function
    Next row
    Eligible = True
Invalid:
End Function

Public Function PendingUnchanged(ByVal context As String, ByVal receipt As Variant) As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, report As String
    On Error GoTo Invalid
    If Not modInventoryCommands.CheckForm(context, report) Then Exit Function
    Set owner = New cSnapshotInventorySource
    Set book = owner.Read(CStr(receipt(1)), Nothing)
    If book Is Nothing Then Exit Function
    If Not book.Saved Then Exit Function
    If Not Eligible(book, CStr(receipt(4)), CStr(receipt(8))) Then Exit Function
    If Not CatalogQuantityUnchanged(book, receipt) Then Exit Function
    Set book = Nothing: Set owner = Nothing
    PendingUnchanged = modInventoryCommands.CheckForm(context, report)
Invalid:
End Function

Private Function CatalogQuantityUnchanged(ByVal book As Workbook, ByVal receipt As Variant) As Boolean
    Dim table As ListObject, fields As Object, item As Object, row As Long
    Dim payload As String, report As String, category As String
    On Error GoTo Invalid
    payload = CStr(receipt(11))
    Set item = modRoleEventJson.ParseJsonObjectRole(Mid$(payload, 2, Len(payload) - 2), report)
    If item Is Nothing Then Exit Function
    Set table = modExecutionTarget.Table(book, "tblSkuCatalog")
    Set fields = modAggregationSources.RequiredColumns(table, Array("SKU", "LOCATION", "TRACK_QTY", "ITEM_KIND", "CATEGORY"))
    If fields Is Nothing Then Exit Function
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, fields("SKU")).Value2) = CStr(receipt(4)) Then
            If CStr(table.DataBodyRange.Cells(row, fields("LOCATION")).Value2) <> CStr(item("Location")) Then Exit Function
            If UCase$(Trim$(CStr(table.DataBodyRange.Cells(row, fields("TRACK_QTY")).Value2))) <> "TRUE" Then Exit Function
            If UCase$(Trim$(CStr(table.DataBodyRange.Cells(row, fields("ITEM_KIND")).Value2))) <> "INVENTORY" Then Exit Function
            category = UCase$(Trim$(CStr(table.DataBodyRange.Cells(row, fields("CATEGORY")).Value2)))
            If category = "UTILITY" Or category = "SERVICE" Then Exit Function
            CatalogQuantityUnchanged = True
            Exit Function
        End If
    Next row
Invalid:
End Function

Public Function Payload(ByVal submitted As cInventoryEditRow, ByVal systemKey As String) As String
    Dim item As Object, items As New Collection, note As String
    note = "Admin completed catalog-only inventory item. Reason: " & submitted.SubmittedText("EDITREASON")
    ' The catalog phase was saved already. Quantity retry must not carry its
    ' old names, custom fields or other catalog values back into the Domain.
    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem(systemKey, submitted.SubmittedText("ITEMCODE"), _
        submitted.TargetQuantity, submitted.SubmittedText("DEFAULTLOCATION"), "GOOD", "", note)
    items.Add item
    Payload = modRoleEventWriter.BuildPayloadJsonFromCollection(items)
End Function

Public Function ValidPayload(ByVal receipt As Variant) As Boolean
    Dim item As Object, report As String, payload As String, name As Variant
    On Error GoTo Invalid
    payload = CStr(receipt(11))
    If Left$(payload, 1) <> "[" Or Right$(payload, 1) <> "]" Then Exit Function
    Set item = modRoleEventJson.ParseJsonObjectRole(Mid$(payload, 2, Len(payload) - 2), report)
    If item Is Nothing Then Exit Function
    If item.Count <> 8 Then Exit Function
    For Each name In Array("System_Key", "SKU", "Qty", "Location", "Condition", "AttributesJson", "Note", "IoType")
        If Not item.Exists(CStr(name)) Then Exit Function
    Next name
    If "[" & modRoleEventJson.DictionaryToJson(item) & "]" <> payload Then Exit Function
    If CStr(item("System_Key")) <> receipt(8) Or CStr(item("SKU")) <> receipt(4) Then Exit Function
    If CDbl(item("Qty")) <> CDbl(receipt(10)) Or CStr(item("IoType")) <> "CREATE" Then Exit Function
    If CStr(item("Condition")) <> "GOOD" Or CStr(item("AttributesJson")) <> "" Then Exit Function
    If Trim$(CStr(item("Note"))) = "" Then Exit Function
    ValidPayload = True
Invalid:
End Function
