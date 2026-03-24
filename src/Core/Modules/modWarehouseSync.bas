Attribute VB_Name = "modWarehouseSync"
Option Explicit

Private Const SHEET_OUTBOX As String = "OutboxEvents"
Private Const TABLE_OUTBOX As String = "tblOutboxEvents"

Private Const SHEET_SNAPSHOT As String = "InventorySnapshot"
Private Const TABLE_SNAPSHOT As String = "tblInventorySnapshot"

Public Function AppendEventToOutbox(ByVal evt As Object, _
                                    Optional ByVal inventoryWb As Workbook = Nothing, _
                                    Optional ByVal outboxWb As Workbook = Nothing, _
                                    Optional ByVal runId As String = "", _
                                    Optional ByRef report As String = "") As Boolean
    On Error GoTo FailAppend

    Dim warehouseId As String
    Dim wbOutbox As Workbook
    Dim loOutbox As ListObject
    Dim appliedMeta As Object
    Dim eventId As String
    Dim rowIndex As Long
    Dim r As ListRow

    warehouseId = GetEventStringSync(evt, "WarehouseId")
    eventId = GetEventStringSync(evt, "EventID")
    If eventId = "" Then
        report = "Outbox write requires EventID."
        Exit Function
    End If

    Set wbOutbox = ResolveOutboxWorkbook(warehouseId, outboxWb, True)
    If wbOutbox Is Nothing Then
        report = "Outbox workbook not found."
        Exit Function
    End If
    If Not EnsureOutboxSchema(wbOutbox, report) Then Exit Function

    Set loOutbox = wbOutbox.Worksheets(SHEET_OUTBOX).ListObjects(TABLE_OUTBOX)
    Set appliedMeta = ResolveAppliedMeta(eventId, inventoryWb)
    If appliedMeta Is Nothing Then
        report = "Applied metadata not found for EventID " & eventId
        Exit Function
    End If

    rowIndex = FindRowByValueSync(loOutbox, "EventID", eventId)
    If rowIndex = 0 Then
        EnsureTableSheetEditableSync loOutbox, TABLE_OUTBOX
        Set r = loOutbox.ListRows.Add
        rowIndex = r.Index
    End If

    SetTableRowValueSync loOutbox, rowIndex, "EventID", eventId
    SetTableRowValueSync loOutbox, rowIndex, "UndoOfEventId", GetEventStringSync(evt, "UndoOfEventId")
    SetTableRowValueSync loOutbox, rowIndex, "EventType", GetEventStringSync(evt, "EventType")
    SetTableRowValueSync loOutbox, rowIndex, "WarehouseId", warehouseId
    SetTableRowValueSync loOutbox, rowIndex, "StationId", GetEventStringSync(evt, "StationId")
    SetTableRowValueSync loOutbox, rowIndex, "OccurredAtUTC", GetEventValueSync(evt, "CreatedAtUTC")
    SetTableRowValueSync loOutbox, rowIndex, "AppliedAtUTC", appliedMeta("AppliedAtUTC")
    SetTableRowValueSync loOutbox, rowIndex, "AppliedByUserId", GetEventStringSync(evt, "UserId")
    SetTableRowValueSync loOutbox, rowIndex, "RunId", ResolveStringSync(appliedMeta, "RunId", runId)
    SetTableRowValueSync loOutbox, rowIndex, "DeltaJson", BuildDeltaJsonForOutbox(evt)
    SaveWorkbookSync wbOutbox

    report = "OK"
    AppendEventToOutbox = True
    Exit Function

FailAppend:
    report = "AppendEventToOutbox failed: " & Err.Description
End Function

Public Function EnsureOutboxSchema(Optional ByVal targetWb As Workbook = Nothing, _
                                   Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim startCell As Range
    Dim i As Long

    If targetWb Is Nothing Then
        Set wb = ResolveOutboxWorkbook(modConfig.GetWarehouseId(), Nothing, True)
    Else
        Set wb = targetWb
    End If
    If wb Is Nothing Then
        report = "Outbox workbook not resolved."
        Exit Function
    End If

    headers = Array("EventID", "UndoOfEventId", "EventType", "WarehouseId", "StationId", "OccurredAtUTC", _
                    "AppliedAtUTC", "AppliedByUserId", "RunId", "DeltaJson")

    Set ws = EnsureWorksheetSync(wb, SHEET_OUTBOX)
    EnsureWorksheetEditableSync ws
    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_OUTBOX)
    On Error GoTo 0

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCellSync(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers))), , xlYes)
        lo.Name = TABLE_OUTBOX
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumnSync lo, CStr(headers(i))
    Next i
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add

    report = "OK"
    EnsureOutboxSchema = True
    Exit Function

FailEnsure:
    report = "EnsureOutboxSchema failed: " & Err.Description
End Function

Public Function GenerateWarehouseSnapshot(Optional ByVal warehouseId As String = "", _
                                          Optional ByVal inventoryWb As Workbook = Nothing, _
                                          Optional ByVal outputPath As String = "", _
                                          Optional ByVal snapshotWb As Workbook = Nothing, _
                                          Optional ByRef report As String = "") As Boolean
    On Error GoTo FailSnapshot

    Dim wbInv As Workbook
    Dim wbSnap As Workbook
    Dim loLog As ListObject
    Dim summary As Object
    Dim lastApplied As Object
    Dim sku As String
    Dim qty As Double
    Dim rowDate As Variant
    Dim i As Long
    Dim savePath As String

    If warehouseId = "" Then warehouseId = modConfig.GetWarehouseId()
    Set wbInv = ResolveInventoryWorkbookBridge(warehouseId, inventoryWb)
    If wbInv Is Nothing Then
        report = "Inventory workbook not found."
        Exit Function
    End If

    Set loLog = FindListObjectByNameSync(wbInv, "tblInventoryLog")
    If loLog Is Nothing Then
        report = "Inventory log table not found."
        Exit Function
    End If

    Set summary = CreateObject("Scripting.Dictionary")
    summary.CompareMode = vbTextCompare
    Set lastApplied = CreateObject("Scripting.Dictionary")
    lastApplied.CompareMode = vbTextCompare

    If Not loLog.DataBodyRange Is Nothing Then
        For i = 1 To loLog.ListRows.Count
            sku = SafeTrimSync(GetCellByColumnSync(loLog, i, "SKU"))
            If sku <> "" Then
                qty = 0
                If IsNumeric(GetCellByColumnSync(loLog, i, "QtyDelta")) Then qty = CDbl(GetCellByColumnSync(loLog, i, "QtyDelta"))
                If summary.Exists(sku) Then
                    summary(sku) = CDbl(summary(sku)) + qty
                Else
                    summary.Add sku, qty
                End If

                rowDate = GetCellByColumnSync(loLog, i, "AppliedAtUTC")
                If IsDate(rowDate) Then
                    If (Not lastApplied.Exists(sku)) Or CDate(rowDate) > CDate(lastApplied(sku)) Then lastApplied(sku) = CDate(rowDate)
                End If
            End If
        Next i
    End If

    Set wbSnap = ResolveSnapshotWorkbook(warehouseId, outputPath, snapshotWb, True)
    If wbSnap Is Nothing Then
        report = "Snapshot workbook not resolved."
        Exit Function
    End If
    savePath = wbSnap.FullName
    If Not EnsureSnapshotSchema(wbSnap, report) Then Exit Function
    WriteSnapshotRows wbSnap, warehouseId, summary, lastApplied
    wbSnap.Save

    report = savePath
    GenerateWarehouseSnapshot = True
    Exit Function

FailSnapshot:
    report = "GenerateWarehouseSnapshot failed: " & Err.Description
End Function

Public Function ResolveOutboxWorkbook(Optional ByVal warehouseId As String = "", _
                                      Optional ByVal targetWb As Workbook = Nothing, _
                                      Optional ByVal createIfMissing As Boolean = False) As Workbook
    Dim targetPath As String

    If Not targetWb Is Nothing Then
        Set ResolveOutboxWorkbook = targetWb
        Exit Function
    End If

    targetPath = ResolveOutboxPath(warehouseId)
    Set ResolveOutboxWorkbook = ResolveWorkbookByPathSync(targetPath, createIfMissing)
End Function

Public Function ResolveSnapshotWorkbook(Optional ByVal warehouseId As String = "", _
                                        Optional ByVal outputPath As String = "", _
                                        Optional ByVal targetWb As Workbook = Nothing, _
                                        Optional ByVal createIfMissing As Boolean = False) As Workbook
    Dim targetPath As String

    If Not targetWb Is Nothing Then
        Set ResolveSnapshotWorkbook = targetWb
        Exit Function
    End If

    targetPath = ResolveSnapshotPath(warehouseId, outputPath)
    Set ResolveSnapshotWorkbook = ResolveWorkbookByPathSync(targetPath, createIfMissing)
End Function

Private Function EnsureSnapshotSchema(ByVal wb As Workbook, ByRef report As String) As Boolean
    On Error GoTo FailEnsure

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim startCell As Range
    Dim i As Long

    headers = Array("WarehouseId", "SKU", "QtyOnHand", "LastAppliedAtUTC")
    Set ws = EnsureWorksheetSync(wb, SHEET_SNAPSHOT)
    EnsureWorksheetEditableSync ws

    On Error Resume Next
    Set lo = ws.ListObjects(TABLE_SNAPSHOT)
    On Error GoTo 0

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCellSync(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers))), , xlYes)
        lo.Name = TABLE_SNAPSHOT
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumnSync lo, CStr(headers(i))
    Next i
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add

    report = "OK"
    EnsureSnapshotSchema = True
    Exit Function

FailEnsure:
    report = "EnsureSnapshotSchema failed: " & Err.Description
End Function

Private Sub WriteSnapshotRows(ByVal wb As Workbook, _
                              ByVal warehouseId As String, _
                              ByVal summary As Object, _
                              ByVal lastApplied As Object)
    Dim lo As ListObject
    Dim key As Variant
    Dim rowIndex As Long

    Set lo = wb.Worksheets(SHEET_SNAPSHOT).ListObjects(TABLE_SNAPSHOT)
    DeleteAllRowsSync lo

    If summary Is Nothing Or summary.Count = 0 Then
        EnsureTableSheetEditableSync lo, TABLE_SNAPSHOT
        lo.ListRows.Add
        SetTableRowValueSync lo, 1, "WarehouseId", warehouseId
        SetTableRowValueSync lo, 1, "SKU", ""
        SetTableRowValueSync lo, 1, "QtyOnHand", 0
        SetTableRowValueSync lo, 1, "LastAppliedAtUTC", vbNullString
        Exit Sub
    End If

    For Each key In summary.Keys
        EnsureTableSheetEditableSync lo, TABLE_SNAPSHOT
        lo.ListRows.Add
        rowIndex = lo.ListRows.Count
        SetTableRowValueSync lo, rowIndex, "WarehouseId", warehouseId
        SetTableRowValueSync lo, rowIndex, "SKU", CStr(key)
        SetTableRowValueSync lo, rowIndex, "QtyOnHand", CDbl(summary(key))
        If lastApplied.Exists(key) Then SetTableRowValueSync lo, rowIndex, "LastAppliedAtUTC", lastApplied(key)
    Next key
End Sub

Private Function ResolveAppliedMeta(ByVal eventId As String, ByVal inventoryWb As Workbook) As Object
    Dim wb As Workbook
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim meta As Object

    Set wb = ResolveInventoryWorkbookBridge("", inventoryWb)
    If wb Is Nothing Then Exit Function

    Set lo = FindListObjectByNameSync(wb, "tblAppliedEvents")
    If lo Is Nothing Then Exit Function
    rowIndex = FindRowByValueSync(lo, "EventID", eventId)
    If rowIndex = 0 Then Exit Function

    Set meta = CreateObject("Scripting.Dictionary")
    meta.CompareMode = vbTextCompare
    meta("AppliedAtUTC") = GetCellByColumnSync(lo, rowIndex, "AppliedAtUTC")
    meta("RunId") = GetCellByColumnSync(lo, rowIndex, "RunId")
    meta("Status") = GetCellByColumnSync(lo, rowIndex, "Status")
    meta("SourceInbox") = GetCellByColumnSync(lo, rowIndex, "SourceInbox")
    Set ResolveAppliedMeta = meta
End Function

Private Function BuildDeltaJsonForOutbox(ByVal evt As Object) As String
    Dim payloadJson As String
    Dim items As Collection
    Dim item As Object

    payloadJson = GetEventStringSync(evt, "PayloadJson")
    If payloadJson <> "" Then
        BuildDeltaJsonForOutbox = payloadJson
        Exit Function
    End If

    Set items = New Collection
    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("SKU") = GetEventStringSync(evt, "SKU")
    item("QtyDelta") = GetEventValueSync(evt, "Qty")
    item("Location") = GetEventStringSync(evt, "Location")
    item("Note") = GetEventStringSync(evt, "Note")
    items.Add item
    BuildDeltaJsonForOutbox = modRoleEventWriter.BuildPayloadJsonFromCollection(items)
End Function

Private Function ResolveOutboxPath(ByVal warehouseId As String) As String
    Dim rootPath As String
    If warehouseId = "" Then warehouseId = modConfig.GetWarehouseId()
    rootPath = modConfig.GetString("PathDataRoot", Environ$("TEMP"))
    ResolveOutboxPath = NormalizeFolderPathSync(rootPath) & warehouseId & ".Outbox.Events.xlsb"
End Function

Private Function ResolveSnapshotPath(ByVal warehouseId As String, ByVal outputPath As String) As String
    Dim rootPath As String
    If Trim$(outputPath) <> "" Then
        ResolveSnapshotPath = outputPath
        Exit Function
    End If
    If warehouseId = "" Then warehouseId = modConfig.GetWarehouseId()
    rootPath = modConfig.GetString("PathDataRoot", Environ$("TEMP"))
    ResolveSnapshotPath = NormalizeFolderPathSync(rootPath) & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
End Function

Private Function ResolveWorkbookByPathSync(ByVal targetPath As String, ByVal createIfMissing As Boolean) As Workbook
    On Error GoTo FailOpen

    Dim wb As Workbook
    Dim fileExists As Boolean
    Dim prevEvents As Boolean
    Dim eventsSuppressed As Boolean

    If Trim$(targetPath) = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, targetPath, vbTextCompare) = 0 Then
            Set ResolveWorkbookByPathSync = wb
            Exit Function
        End If
    Next wb

    fileExists = (Len(Dir$(targetPath)) > 0)
    If fileExists Then
        Set ResolveWorkbookByPathSync = Application.Workbooks.Open(targetPath)
        Exit Function
    End If

    If Not createIfMissing Then Exit Function

    EnsureFolderForFileSync targetPath
    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    eventsSuppressed = True
    Set wb = Application.Workbooks.Add(xlWBATWorksheet)
    wb.SaveAs Filename:=targetPath, FileFormat:=50
    Application.EnableEvents = prevEvents
    eventsSuppressed = False
    Set ResolveWorkbookByPathSync = wb
    Exit Function

FailOpen:
    On Error Resume Next
    If eventsSuppressed Then Application.EnableEvents = prevEvents
    On Error GoTo 0
End Function

Private Function EnsureWorksheetSync(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheetSync = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheetSync Is Nothing Then
        Set EnsureWorksheetSync = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheetSync.Name = sheetName
    End If
End Function

Private Sub EnsureWorksheetEditableSync(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 3001, "modWarehouseSync.EnsureWorksheetEditableSync", _
                  "Worksheet '" & ws.Name & "' is protected and could not be unprotected."
    End If
End Sub

Private Sub EnsureTableSheetEditableSync(ByVal lo As ListObject, ByVal tableName As String)
    If lo Is Nothing Then Exit Sub
    EnsureWorksheetEditableSync lo.Parent
    If lo.Parent.ProtectContents Then
        Err.Raise vbObjectError + 3002, "modWarehouseSync.EnsureTableSheetEditableSync", _
                  "Worksheet '" & lo.Parent.Name & "' is protected and could not be unprotected before updating " & tableName & "."
    End If
End Sub

Private Function GetNextTableStartCellSync(ByVal ws As Worksheet) As Range
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        Set GetNextTableStartCellSync = ws.Range("A1")
    Else
        Set GetNextTableStartCellSync = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End If
End Function

Private Sub EnsureListColumnSync(ByVal lo As ListObject, ByVal columnName As String)
    If GetColumnIndexSync(lo, columnName) > 0 Then Exit Sub
    lo.ListColumns.Add lo.ListColumns.Count + 1
    lo.ListColumns(lo.ListColumns.Count).Name = columnName
End Sub

Private Sub DeleteAllRowsSync(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    EnsureTableSheetEditableSync lo, lo.Name
    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
End Sub

Private Function FindListObjectByNameSync(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameSync = ws.ListObjects(tableName)
        If Not FindListObjectByNameSync Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function FindRowByValueSync(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrimSync(GetCellByColumnSync(lo, i, columnName)), expectedValue, vbTextCompare) = 0 Then
            FindRowByValueSync = i
            Exit Function
        End If
    Next i
End Function

Private Sub SaveWorkbookSync(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If wb.Path = "" Then Exit Sub
    wb.Save
End Sub

Private Function GetCellByColumnSync(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndexSync(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnSync = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Sub SetTableRowValueSync(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = GetColumnIndexSync(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetColumnIndexSync(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexSync = i
            Exit Function
        End If
    Next i
End Function

Private Function GetEventStringSync(ByVal evt As Object, ByVal key As String) As String
    Dim v As Variant
    v = GetEventValueSync(evt, key)
    GetEventStringSync = SafeTrimSync(v)
End Function

Private Function GetEventValueSync(ByVal evt As Object, ByVal key As String) As Variant
    On Error Resume Next
    If evt Is Nothing Then Exit Function
    GetEventValueSync = evt(key)
    On Error GoTo 0
End Function

Private Function ResolveStringSync(ByVal d As Object, ByVal key As String, ByVal fallbackValue As String) As String
    On Error Resume Next
    ResolveStringSync = SafeTrimSync(d(key))
    On Error GoTo 0
    If ResolveStringSync = "" Then ResolveStringSync = fallbackValue
End Function

Private Function SafeTrimSync(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimSync = Trim$(CStr(valueIn))
End Function

Private Function NormalizeFolderPathSync(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then
        NormalizeFolderPathSync = Environ$("TEMP") & "\"
        Exit Function
    End If
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathSync = folderPath
End Function

Private Sub EnsureFolderForFileSync(ByVal filePath As String)
    Dim folderPath As String
    Dim sepPos As Long

    sepPos = InStrRev(filePath, "\")
    If sepPos <= 0 Then Exit Sub
    folderPath = Left$(filePath, sepPos - 1)
    CreateFolderRecursiveSync folderPath
End Sub

Private Sub CreateFolderRecursiveSync(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long

    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    sepPos = InStrRev(folderPath, "\")
    If sepPos > 0 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then CreateFolderRecursiveSync parentPath
    End If

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub
