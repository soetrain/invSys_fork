Attribute VB_Name = "modInventoryCommandReceipts"
Option Explicit
Option Private Module

' Match one exact existing command payload in the durable queue. Other pending
' events are inspected read-only and never selected for application by this API.
Public Function FindExact(ByVal warehouse As String, ByVal station As String, ByVal user As String, _
        ByVal payload As String, ByRef eventId As String, ByRef report As String) As Boolean
    Dim ids As Object, fso As Object, file As Object, folder As String, rows As Collection
    Dim row As Variant, hash As String, path As String, book As Workbook, copy As String
    Dim table As ListObject, fields As Object, index As Long, record As Object, name As Variant
    On Error GoTo Failed
    eventId = ""
    If warehouse = "" Or station = "" Or user = "" Or payload = "" Then GoTo Failed
    Set ids = CreateObject("Scripting.Dictionary"): ids.CompareMode = vbBinaryCompare
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = modRoleEventWriter.LocalStagingSearchRootRole(warehouse, station)
    If folder = "" Then GoTo Failed
    If fso.FolderExists(folder) Then
        For Each file In fso.GetFolder(folder).Files
            If LCase$(file.Name) Like "*.staging.jsonl" Or LCase$(file.Name) Like "*.syncing" Then
                hash = modTrainingWire.FileSha256(CStr(file.Path))
                If hash = "" Then GoTo Failed
                report = "": Set rows = modRoleEventJson.ReadStagedInboxRowsRole(CStr(file.Path), report)
                If rows Is Nothing Then GoTo Failed
                For Each row In rows
                    AddExact ids, row, warehouse, station, user, payload
                Next row
                If hash <> modTrainingWire.FileSha256(CStr(file.Path)) Then GoTo Failed
            End If
        Next file
    End If
    path = modRoleEventWriter.ResolveInboxWorkbookPath("ADMIN_INVENTORY_ADJUST", warehouse, station, report)
    If path = "" Or modAggregationFiles.Dirty(path) Then GoTo Failed
    If fso.FileExists(path) Then
        hash = modTrainingWire.FileSha256(path)
        If hash = "" Then GoTo Failed
        Set book = modAggregationFiles.OpenCopy(path, copy)
        If book Is Nothing Then GoTo Failed
        Set table = modExecutionTarget.Table(book, "tblInboxProd")
        Set fields = modAggregationSources.RequiredColumns(table, _
            Array("EventID", "EventType", "WarehouseId", "StationId", "UserId", "Status", "PayloadJson"))
        If fields Is Nothing Then GoTo Failed
        For index = 1 To table.ListRows.Count
            Set record = CreateObject("Scripting.Dictionary")
            For Each name In fields.Keys
                record.Add CStr(name), CStr(table.DataBodyRange.Cells(index, fields(name)).Value2)
            Next name
            AddExact ids, record, warehouse, station, user, payload
        Next index
        If hash <> modTrainingWire.FileSha256(path) Then GoTo Failed
    End If
    modAggregationFiles.CloseCopy book, copy
    If ids.Count > 1 Then
        report = "Multiple saved events match this exact adjustment. Reconcile them before retry; no new event was queued."
        Exit Function
    End If
    For Each name In ids.Keys: eventId = CStr(name): Next name
    report = "": FindExact = True
    Exit Function
Failed:
    modAggregationFiles.CloseCopy book, copy
    report = "Existing adjustment receipts could not be verified. No new event was queued."
End Function

Private Sub AddExact(ByVal ids As Object, ByVal row As Object, ByVal warehouse As String, _
        ByVal station As String, ByVal user As String, ByVal payload As String)
    Dim status As String, key As String
    If row.Exists("Status") Then status = UCase$(Trim$(CStr(row("Status"))))
    If status = "APPLIED" Or status = "SKIP_DUP" Then Exit Sub
    If CStr(row("WarehouseId")) <> warehouse Or CStr(row("StationId")) <> station _
            Or CStr(row("UserId")) <> user Or CStr(row("EventType")) <> "ADMIN_INVENTORY_ADJUST" Then Exit Sub
    If StrComp(CStr(row("PayloadJson")), payload, vbBinaryCompare) <> 0 Then Exit Sub
    key = CStr(row("EventID"))
    If key = "" Then Err.Raise 5, , "Matched adjustment has no event identity."
    ids(key) = True
End Sub
