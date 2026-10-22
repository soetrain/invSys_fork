Attribute VB_Name = "modInventoryEditPlan"
Option Explicit
Option Private Module

Public Function Build(ByVal context As String, ByVal submitted As cInventoryEditRow, _
        ByVal command As String, ByVal eventId As String, ByRef receipt As Variant, _
        ByRef catalogPlan As Variant, ByRef report As String) As Boolean
    Dim owner As cSnapshotInventorySource, book As Workbook, table As ListObject
    Dim fields As Object, state As Object, column As ListColumn, row As Long, selected As Long
    Dim headers() As String, names As Variant, values As Variant, custom As Variant, history As String
    Dim entity As Variant, delta As Double, payload As String, intent As String
    Dim items As New Collection, item As Object, source As Variant, i As Long
    On Error GoTo Invalid
    If Not modInventoryCommands.CheckForm(context, report) Then Exit Function
    source = submitted.Source
    Set owner = New cSnapshotInventorySource
    Set book = owner.Read(CStr(source(0)), Nothing)
    If book Is Nothing Then GoTo Invalid
    If Not book.Saved Then GoTo Invalid
    Set table = modExecutionTarget.Table(book, "tblSkuCatalog")
    Set fields = modAggregationSources.RequiredColumns(table, Array("SKU", "ITEM_CODE"))
    If fields Is Nothing Then GoTo Invalid
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, fields("SKU")).Value2) = submitted.SubmittedText("ITEMCODE") Then
            If selected <> 0 Then GoTo Invalid
            selected = row
            If CStr(table.DataBodyRange.Cells(row, fields("ITEM_CODE")).Value2) <> submitted.SubmittedText("ITEMCODE") Then GoTo Invalid
        End If
    Next row
    If selected = 0 Then GoTo Invalid
    Set state = CreateObject("Scripting.Dictionary"): state.CompareMode = vbBinaryCompare
    ReDim headers(0 To table.ListColumns.Count - 1)
    For Each column In table.ListColumns
        headers(column.Index - 1) = column.Name
        state.Add "CATALOG:" & UCase$(Trim$(column.Name)), _
            CStr(table.DataBodyRange.Cells(selected, column.Index).Value2) & vbNullChar & _
            CStr(table.DataBodyRange.Cells(selected, column.Index).Formula)
        If StrComp(column.Name, "EDIT_HISTORY", vbTextCompare) = 0 Then _
            history = Trim$(CStr(table.DataBodyRange.Cells(selected, column.Index).Value))
    Next column
    custom = submitted.CatalogCustomFields()
    catalogPlan = modInventoryCatalogWire.PlanFields(submitted.CatalogValues, custom(0), custom(1), history, headers)
    names = catalogPlan(0): values = catalogPlan(1)
    For i = LBound(names) To UBound(names)
        If StrComp(CStr(names(i)), "SKU", vbTextCompare) = 0 Then
            If CStr(values(i)) <> submitted.SubmittedText("ITEMCODE") Then GoTo Invalid
        End If
        If Len(CStr(values(i))) > 32767 Then GoTo Invalid
    Next i
    intent = "NONE": entity = Array("", "", "", "", "")
    If submitted.HasQuantity Then
        entity = modInventoryQuantityCommands.ResolveSingleEntity(CStr(source(0)), submitted.SubmittedText("ITEMCODE"), report)
        If Not IsArray(entity) Then GoTo Invalid
        state.Add "ENTITY", Join(entity, vbNullChar)
        delta = submitted.TargetQuantity - CDbl(entity(1))
        intent = "NOOP"
        If Abs(delta) >= 0.0000001 Then
            intent = "ADJUST"
            If eventId = "" Then eventId = UCase$(modSystemIdentity.NewId())
            Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem(CStr(entity(0)), submitted.SubmittedText("ITEMCODE"), _
                delta, CStr(entity(2)), CStr(entity(3)), CStr(entity(4)), _
                "Admin set inventory quantity. Reason: " & submitted.SubmittedText("EDITREASON"))
            item("IoType") = "ADJUST": item("PreviousQty") = CDbl(entity(1))
            item("TargetQty") = submitted.TargetQuantity: item("Reason") = submitted.SubmittedText("EDITREASON")
            items.Add item
            payload = modRoleEventWriter.BuildPayloadJsonFromCollection(items)
        End If
    End If
    If intent <> "ADJUST" And eventId <> "" Then GoTo Invalid
    receipt = Array(command, CStr(source(0)), CStr(source(1)), CStr(source(2)), submitted.SubmittedText("ITEMCODE"), _
        submitted.InputHash, modInventoryEditWire.HashFields(state.Keys, state.Items), _
        eventId, CStr(entity(0)), CStr(entity(1)), "", payload, intent, "SAVED", "")
    If submitted.HasQuantity Then receipt(10) = CStr(submitted.TargetQuantity)
    receipt(14) = BindingHash(receipt)
    If receipt(6) = "" Or receipt(14) = "" Then GoTo Invalid
    For i = 0 To UBound(receipt)
        If Len(CStr(receipt(i))) > 32767 Then GoTo Invalid
    Next i
    If Not ValidReceipt(receipt) Then GoTo Invalid
    Set book = Nothing: Set owner = Nothing
    Build = modInventoryCommands.CheckForm(context, report)
    Exit Function
Invalid:
    report = "The exact EDIT catalog or quantity plan could not be verified. No new event was queued."
End Function

Public Function BindingHash(ByVal receipt As Variant) As String
    Dim headers As Variant, names(0 To 13) As String, values(0 To 13) As String, i As Long
    On Error GoTo Invalid
    If LBound(receipt) <> 0 Or UBound(receipt) <> 14 Then Exit Function
    headers = modInventoryEditWire.ReceiptHeaders()
    For i = 0 To 13: names(i) = headers(i): values(i) = CStr(receipt(i)): Next i
    BindingHash = modInventoryEditWire.HashFields(names, values)
Invalid:
End Function

Public Function ValidReceipt(ByVal receipt As Variant) As Boolean
    Dim i As Long
    On Error GoTo Invalid
    If Not modTrainingWire.ValidId(LCase$(CStr(receipt(0)))) Then Exit Function
    If CStr(receipt(0)) = CStr(receipt(7)) Or CStr(receipt(0)) = CStr(receipt(8)) Then Exit Function
    If CStr(receipt(7)) <> "" And CStr(receipt(7)) = CStr(receipt(8)) Then Exit Function
    For i = 1 To 6
        If CStr(receipt(i)) = "" Then Exit Function
    Next i
    If CStr(receipt(13)) <> "SAVED" Or CStr(receipt(14)) = "" Then Exit Function
    If CStr(receipt(14)) <> BindingHash(receipt) Then Exit Function
    Select Case CStr(receipt(12))
        Case "NONE"
            For i = 7 To 11
                If CStr(receipt(i)) <> "" Then Exit Function
            Next i
        Case "NOOP", "ADJUST"
            If CStr(receipt(8)) = "" Or Not IsNumeric(receipt(9)) Or Not IsNumeric(receipt(10)) Then Exit Function
            If CDbl(receipt(10)) < 0 Then Exit Function
            If CStr(receipt(12)) = "NOOP" Then
                If CStr(receipt(7)) <> "" Or CStr(receipt(11)) <> "" Then Exit Function
                If Abs(CDbl(receipt(9)) - CDbl(receipt(10))) >= 0.0000001 Then Exit Function
            Else
                If Not modTrainingWire.ValidId(LCase$(CStr(receipt(7)))) Or CStr(receipt(11)) = "" Then Exit Function
                If Abs(CDbl(receipt(9)) - CDbl(receipt(10))) < 0.0000001 Then Exit Function
            End If
        Case Else: Exit Function
    End Select
    ValidReceipt = True
Invalid:
End Function

Public Function Matches(ByVal submitted As cInventoryEditRow, ByVal reference As String, _
        ByVal receipt As Variant) As Boolean
    Dim source As Variant
    On Error GoTo Invalid
    source = submitted.Source
    If CStr(receipt(0)) <> modInventoryEditWire.ReferenceField(reference, "Command:") Then Exit Function
    If CStr(receipt(1)) <> source(0) Or CStr(receipt(2)) <> source(1) Or CStr(receipt(3)) <> source(2) Then Exit Function
    If CStr(receipt(4)) <> submitted.SubmittedText("ITEMCODE") Or CStr(receipt(5)) <> submitted.InputHash Then Exit Function
    If CStr(receipt(13)) <> "SAVED" Then Exit Function
    If CStr(receipt(7)) <> modInventoryEditWire.ReferenceField(reference, "Event:") Then Exit Function
    If CStr(receipt(14)) = "" Or CStr(receipt(14)) <> BindingHash(receipt) Then Exit Function
    If CStr(receipt(14)) <> modInventoryEditWire.ReferenceField(reference, "Edit binding:") Then Exit Function
    Matches = True
Invalid:
End Function

Public Function PendingEntityUnchanged(ByVal context As String, ByVal receipt As Variant) As Boolean
    Dim entity As Variant, item As Object, report As String
    On Error GoTo Invalid
    If Not modInventoryCommands.CheckForm(context, report) Then Exit Function
    entity = modInventoryQuantityCommands.ResolveSingleEntity(CStr(receipt(1)), CStr(receipt(4)), report)
    If Not IsArray(entity) Then Exit Function
    If CStr(entity(0)) <> receipt(8) Or CDbl(entity(1)) <> CDbl(receipt(9)) Then Exit Function
    ' Compare the full originally planned payload, including location, condition
    ' and attributes. A pending retry must not be rebased onto changed inventory.
    Set item = modRoleEventJson.ParseJsonObjectRole(Mid$(CStr(receipt(11)), 2, Len(CStr(receipt(11))) - 2), report)
    If item Is Nothing Then Exit Function
    If CStr(item("Location")) <> entity(2) Or CStr(item("Condition")) <> entity(3) Or _
        CStr(item("AttributesJson")) <> entity(4) Then Exit Function
    PendingEntityUnchanged = modInventoryCommands.CheckForm(context, report)
Invalid:
End Function
