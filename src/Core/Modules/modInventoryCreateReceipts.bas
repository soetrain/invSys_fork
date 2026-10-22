Attribute VB_Name = "modInventoryCreateReceipts"
Option Explicit

' Exact event lookup in existing queues and canonical receipts; no new store.
Public Function NewEventId(ByVal context As String, ByRef report As String) As String
    If modInventoryCommands.CheckForm(context, report) Then NewEventId = UCase$(modSystemIdentity.NewId())
End Function

Public Function EventIdFromResult(ByVal text As String) As String
    EventIdFromResult = ResultField(text, "Event:")
End Function

Private Function ResultField(ByVal text As String, ByVal prefix As String) As String
    Dim line As Variant, value As String, found As String
    text = Replace(text, vbCr, "")
    For Each line In Split(text, vbLf)
        If StrComp(Left$(Trim$(CStr(line)), Len(prefix)), prefix, vbTextCompare) = 0 Then
            value = Trim$(Mid$(Trim$(CStr(line)), Len(prefix) + 1))
            If value = "" Or InStr(value, " ") > 0 Or Len(value) > 200 Then Exit Function
            If found <> "" And found <> value Then Exit Function
            found = value
        End If
    Next line
    ResultField = found
End Function

' A worksheet success label is not application proof. Confirm the exact saved
' event and immutable entity identity through the captured live source.
Public Function ConfirmAppliedResult(ByVal context As String, ByVal result As String, _
        ByRef report As String) As Boolean
    Dim values As Variant, eventId As String, expectedKey As String, systemKey As String
    Dim found As Boolean, applied As Boolean, payload As String
    On Error GoTo Unconfirmed
    eventId = EventIdFromResult(result): expectedKey = ResultField(result, "System_Key:")
    If eventId = "" Or expectedKey = "" Then GoTo Unconfirmed
    values = modInventoryCommands.SourceValues(context)
    If Not IsArray(values) Then GoTo Unconfirmed
    If Not Resolve(context, CStr(values(0)), CStr(values(1)), CStr(values(2)), _
            eventId, found, systemKey, payload, applied, report) Then GoTo Unconfirmed
    If Not found Or Not applied Then GoTo Unconfirmed
    If StrComp(systemKey, expectedKey, vbBinaryCompare) <> 0 Then GoTo Unconfirmed
    ConfirmAppliedResult = True: report = ""
    Exit Function
Unconfirmed:
    report = "The imported ADD receipt is unconfirmed. No new event was queued."
End Function

Public Function Resolve(ByVal context As String, ByVal warehouse As String, ByVal station As String, _
        ByVal user As String, ByVal eventId As String, ByRef found As Boolean, ByRef systemKey As String, _
        ByRef payload As String, ByRef applied As Boolean, ByRef report As String) As Boolean
    Dim values As Variant, chosen As Object, item As Object
    On Error GoTo Failed
    found = False: applied = False: systemKey = "": payload = ""
    If eventId = "" Or context = "" Then GoTo Failed
    values = modInventoryCommands.SourceValues(context)
    If Not IsArray(values) Then Exit Function
    If values(0) <> warehouse Or values(1) <> station Or values(2) <> user Then GoTo Failed
    If Not ReadStaging(context, warehouse, station, user, eventId, chosen, report) Then Exit Function
    If Not ReadInbox(context, warehouse, station, user, eventId, chosen, report) Then Exit Function
    found = Not chosen Is Nothing
    If found Then
        payload = CStr(chosen("PayloadJson"))
        Set item = SingleItem(payload)
        If item Is Nothing Then GoTo Failed
        systemKey = CStr(item("System_Key"))
        If systemKey = "" Or CStr(item("SKU")) = "" Then GoTo Failed
    End If
    If Not ReadSavedState(context, eventId, payload, found, applied, report) Then Exit Function
    Resolve = modInventoryCommands.CheckForm(context, report)
    Exit Function
Failed:
    report = "The exact existing create receipt could not be verified. No new event was queued."
End Function

Public Function PayloadMatches(ByVal expected As String, ByVal queued As String) As Boolean
    Dim expectedItem As Object, queuedItem As Object, key As Variant, a As Variant, b As Variant
    On Error GoTo Invalid
    Set expectedItem = SingleItem(expected): Set queuedItem = SingleItem(queued)
    If expectedItem Is Nothing Or queuedItem Is Nothing Then Exit Function
    If expectedItem.Count <> queuedItem.Count Then Exit Function
    For Each key In expectedItem.Keys
        If Not queuedItem.Exists(key) Then Exit Function
        a = expectedItem(key): b = queuedItem(key)
        If IsNumeric(a) And IsNumeric(b) Then
            If CDbl(a) <> CDbl(b) Then Exit Function
        ElseIf CStr(a) <> CStr(b) Then
            Exit Function
        End If
    Next key
    PayloadMatches = True
Invalid:
End Function

Private Function SingleItem(ByVal payload As String) As Object
    Dim item As Object, report As String
    On Error GoTo Invalid
    payload = Trim$(payload)
    If Left$(payload, 1) <> "[" Or Right$(payload, 1) <> "]" Then Exit Function
    Set item = modRoleEventJson.ParseJsonObjectRole(Mid$(payload, 2, Len(payload) - 2), report)
    If item Is Nothing Then Exit Function
    ' Queue payloads are emitted by this same serializer. Reject ignored tails,
    ' duplicate keys or multiple entities rather than applying an unproved batch.
    If "[" & modRoleEventJson.DictionaryToJson(item) & "]" <> payload Then Exit Function
    Set SingleItem = item
Invalid:
End Function

Private Sub Consider(ByVal record As Object, ByVal warehouse As String, ByVal station As String, _
        ByVal user As String, ByVal eventId As String, ByRef chosen As Object, _
        Optional ByVal eventType As String = "INVENTORY_CREATE")
    If CStr(record("EventID")) <> eventId Then Exit Sub
    If CStr(record("EventType")) <> eventType Or CStr(record("WarehouseId")) <> warehouse _
            Or CStr(record("StationId")) <> station Or CStr(record("UserId")) <> user Then
        Err.Raise vbObjectError + 2142, , "The existing event belongs to another command scope."
    End If
    If chosen Is Nothing Then
        Set chosen = record
    ElseIf CStr(chosen("PayloadJson")) <> CStr(record("PayloadJson")) Then
        Err.Raise vbObjectError + 2142, , "Conflicting exact event payloads require reconciliation."
    End If
End Sub

Private Function ReadStaging(ByVal context As String, ByVal warehouse As String, ByVal station As String, _
        ByVal user As String, ByVal eventId As String, ByRef chosen As Object, ByRef report As String, _
        Optional ByVal eventType As String = "INVENTORY_CREATE") As Boolean
    Dim fso As Object, folder As String, file As Object, rows As Collection, row As Variant, hash As String
    On Error GoTo Failed
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = modRoleEventWriter.LocalStagingSearchRootRole(warehouse, station)
    If folder = "" Then GoTo Failed
    If fso.FolderExists(folder) Then
        For Each file In fso.GetFolder(folder).Files
            If LCase$(file.Name) Like "*.staging.jsonl" Or LCase$(file.Name) Like "*.syncing" Then
                If Not modInventoryCommands.CheckForm(context, report) Then Exit Function
                hash = modTrainingWire.FileSha256(CStr(file.Path))
                If hash = "" Then GoTo Failed
                Set rows = modRoleEventJson.ReadStagedInboxRowsRole(CStr(file.Path), report)
                If rows Is Nothing Then GoTo Failed
                For Each row In rows: Consider row, warehouse, station, user, eventId, chosen, eventType: Next row
                If hash <> modTrainingWire.FileSha256(CStr(file.Path)) Then GoTo Failed
                If Not modInventoryCommands.CheckForm(context, report) Then Exit Function
            End If
        Next file
    End If
    ReadStaging = True
    Exit Function
Failed:
    report = "Existing staged create receipts could not be verified. No new event was queued."
End Function

Private Function ReadInbox(ByVal context As String, ByVal warehouse As String, ByVal station As String, _
        ByVal user As String, ByVal eventId As String, ByRef chosen As Object, ByRef report As String, _
        Optional ByVal eventType As String = "INVENTORY_CREATE") As Boolean
    Dim path As String, hash As String, book As Workbook, copy As String, table As ListObject
    Dim fields As Object, row As Long, record As Object, name As Variant
    On Error GoTo Failed
    path = modRoleEventWriter.ResolveInboxWorkbookPath(eventType, warehouse, station, report)
    If path = "" Or modAggregationFiles.Dirty(path) Then GoTo Failed
    If Len(Dir$(path)) = 0 Then ReadInbox = True: Exit Function
    hash = modTrainingWire.FileSha256(path)
    If hash = "" Then GoTo Failed
    Set book = modAggregationFiles.OpenCopy(path, copy, True)
    If book Is Nothing Then GoTo Failed
    Set table = modExecutionTarget.Table(book, "tblInboxProd")
    Set fields = modAggregationSources.RequiredColumns(table, _
        Array("EventID", "EventType", "WarehouseId", "StationId", "UserId", "PayloadJson"))
    If fields Is Nothing Then GoTo Failed
    If eventType = "ADMIN_INVENTORY_ADJUST" Then
        Set fields = modAggregationSources.RequiredColumns(table, _
            Array("EventID", "EventType", "WarehouseId", "StationId", "UserId", "PayloadJson", "Status"))
        If fields Is Nothing Then GoTo Failed
    End If
    For row = 1 To table.ListRows.Count
        Set record = CreateObject("Scripting.Dictionary")
        For Each name In fields.Keys: record.Add CStr(name), CStr(table.DataBodyRange.Cells(row, fields(name)).Value2): Next name
        If CStr(record("EventID")) = eventId And eventType = "ADMIN_INVENTORY_ADJUST" Then
            If CStr(record("Status")) <> "NEW" Then GoTo Failed
        End If
        Consider record, warehouse, station, user, eventId, chosen, eventType
    Next row
    If hash <> modTrainingWire.FileSha256(path) Then GoTo Failed
    modAggregationFiles.CloseCopy book, copy
    ReadInbox = modInventoryCommands.CheckForm(context, report)
    Exit Function
Failed:
    modAggregationFiles.CloseCopy book, copy
    report = "Existing inbox create receipts could not be verified. No new event was queued."
End Function

' Used only after independent saved-state inspection found this exact planned
' adjustment not applied. Queue presence establishes pending work, not success.
Public Function ReadPendingAdjustment(ByVal context As String, ByVal eventId As String, _
        ByRef found As Boolean, ByRef payload As String, ByRef report As String) As Boolean
    Dim source As Variant, chosen As Object
    On Error GoTo Failed
    found = False: payload = ""
    source = modInventoryCommands.SourceValues(context)
    If Not IsArray(source) Or eventId = "" Then GoTo Failed
    If Not ReadStaging(context, CStr(source(0)), CStr(source(1)), CStr(source(2)), _
        eventId, chosen, report, "ADMIN_INVENTORY_ADJUST") Then GoTo Failed
    If Not ReadInbox(context, CStr(source(0)), CStr(source(1)), CStr(source(2)), _
        eventId, chosen, report, "ADMIN_INVENTORY_ADJUST") Then GoTo Failed
    found = Not chosen Is Nothing
    If found Then payload = CStr(chosen("PayloadJson"))
    ReadPendingAdjustment = modInventoryCommands.CheckForm(context, report)
    Exit Function
Failed:
    report = "The exact pending EDIT event could not be verified. No new event was queued."
End Function

Private Function ReadSavedState(ByVal context As String, ByVal eventId As String, ByVal payload As String, _
        ByVal found As Boolean, ByRef applied As Boolean, ByRef report As String) As Boolean
    Dim path As String, hash As String, book As Workbook, copy As String, table As ListObject
    Dim fields As Object, row As Long, receipts As Long, lines As Long, plan As cInventoryCreatePlan
    On Error GoTo Failed
    path = modInventoryCommands.SourceInventoryPath(context)
    If path = "" Then GoTo Failed
    hash = modTrainingWire.FileSha256(path)
    Set book = modAggregationFiles.OpenCopy(path, copy, True)
    If book Is Nothing Then GoTo Failed
    If hash = "" Or modTrainingWire.FileSha256(copy) <> hash Then GoTo Failed
    Set table = modExecutionTarget.Table(book, "tblAppliedEvents")
    Set fields = modAggregationSources.RequiredColumns(table, Array("EventID", "Status"))
    If fields Is Nothing Then GoTo Failed
    For row = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(row, fields("EventID")).Value2) = eventId Then
            receipts = receipts + 1
            If CStr(table.DataBodyRange.Cells(row, fields("Status")).Value2) <> "APPLIED" Then GoTo Failed
        End If
    Next row
    If receipts > 1 Then GoTo Failed
    If receipts = 1 Then
        If Not found Then GoTo Failed
        Set plan = New cInventoryCreatePlan
        If Not plan.VerifySaved(book, eventId, payload) Then GoTo Failed
        applied = True
    Else
        Set table = modExecutionTarget.Table(book, "tblInventoryLog")
        Set fields = modAggregationSources.RequiredColumns(table, Array("EventID"))
        If fields Is Nothing Then GoTo Failed
        For row = 1 To table.ListRows.Count
            If CStr(table.DataBodyRange.Cells(row, fields("EventID")).Value2) = eventId Then GoTo Failed
        Next row
    End If
    If hash <> modTrainingWire.FileSha256(path) Then GoTo Failed
    modAggregationFiles.CloseCopy book, copy
    ReadSavedState = modInventoryCommands.CheckForm(context, report)
    Exit Function
Failed:
    modAggregationFiles.CloseCopy book, copy
    report = "Exact saved create application is unconfirmed. No new event was queued."
End Function
