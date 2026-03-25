Attribute VB_Name = "modInventorySchema"
Option Explicit

Private Const SHEET_INVENTORY_LOG As String = "InventoryLog"
Private Const TABLE_INVENTORY_LOG As String = "tblInventoryLog"
Private Const SHEET_APPLIED_EVENTS As String = "AppliedEvents"
Private Const TABLE_APPLIED_EVENTS As String = "tblAppliedEvents"
Private Const SHEET_LOCKS As String = "Locks"
Private Const TABLE_LOCKS As String = "tblLocks"
Private Const SHEET_SKU_BALANCE As String = "SkuBalance"
Private Const TABLE_SKU_BALANCE As String = "tblSkuBalance"
Private Const SHEET_LOCATION_BALANCE As String = "LocationBalance"
Private Const TABLE_LOCATION_BALANCE As String = "tblLocationBalance"
Private Const SHEET_LEDGER_STATUS As String = "LedgerStatus"
Private Const TABLE_LEDGER_STATUS As String = "tblInventoryLedgerStatus"

Public Function EnsureInventorySchema(Optional ByVal targetWb As Workbook = Nothing, _
                                      Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Dim issues As Collection

    If targetWb Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetWb
    End If

    Set issues = New Collection

    EnsureTableWithHeaders wb, SHEET_INVENTORY_LOG, TABLE_INVENTORY_LOG, _
        Array("EventID", "UndoOfEventId", "AppliedSeq", "EventType", "OccurredAtUTC", "AppliedAtUTC", _
              "WarehouseId", "StationId", "UserId", "SKU", "QtyDelta", "Location", "Note"), issues

    EnsureTableWithHeaders wb, SHEET_APPLIED_EVENTS, TABLE_APPLIED_EVENTS, _
        Array("EventID", "UndoOfEventId", "AppliedSeq", "AppliedAtUTC", "RunId", "SourceInbox", "Status"), issues

    EnsureTableWithHeaders wb, SHEET_LOCKS, TABLE_LOCKS, _
        Array("LockName", "OwnerStationId", "OwnerUserId", "RunId", "AcquiredAtUTC", "ExpiresAtUTC", "HeartbeatAtUTC", "Status"), issues

    EnsureTableWithHeaders wb, SHEET_SKU_BALANCE, TABLE_SKU_BALANCE, _
        Array("SKU", "QtyOnHand", "LastAppliedUTC"), issues

    EnsureTableWithHeaders wb, SHEET_LOCATION_BALANCE, TABLE_LOCATION_BALANCE, _
        Array("SKU", "Location", "QtyOnHand", "LastAppliedUTC"), issues

    EnsureTableWithHeaders wb, SHEET_LEDGER_STATUS, TABLE_LEDGER_STATUS, _
        Array("WarehouseId", "LastAppliedSeq", "LastEventId", "LastAppliedAtUTC", "TotalEventRows", _
              "TotalAppliedEvents", "DistinctSkuCount", "DistinctLocationCount", "ProjectionRebuiltAtUTC", "Notes"), issues

    report = JoinCollection(issues, "; ")
    EnsureInventorySchema = True
    Exit Function

FailEnsure:
    report = "EnsureInventorySchema failed: " & Err.Description
    EnsureInventorySchema = False
End Function

Private Sub EnsureTableWithHeaders(ByVal wb As Workbook, _
                                   ByVal sheetName As String, _
                                   ByVal tableName As String, _
                                   ByVal headers As Variant, _
                                   ByRef issues As Collection)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim startCell As Range
    Dim tableRange As Range

    Set ws = EnsureWorksheet(wb, sheetName)
    EnsureWorksheetEditableSchema ws
    Set lo = FindListObjectByName(wb, tableName)

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCell(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i

        Set tableRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, tableRange, , xlYes)
        lo.Name = tableName
        issues.Add tableName & " created"
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumn lo, CStr(headers(i)), issues
    Next i

    RemoveBlankSeedRow lo
    StyleProtectedHeaders lo, headers
End Sub

Private Sub EnsureWorksheetEditableSchema(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 2501, "modInventorySchema.EnsureWorksheetEditableSchema", _
                  "Worksheet '" & ws.Name & "' is protected and could not be unprotected. " & _
                  "Excel automation cannot create or extend tables while the sheet remains protected."
    End If
End Sub

Private Function EnsureWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheet Is Nothing Then
        Set EnsureWorksheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheet.Name = sheetName
    End If
End Function

Private Function FindListObjectByName(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByName = ws.ListObjects(tableName)
        If Not FindListObjectByName Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function GetNextTableStartCell(ByVal ws As Worksheet) As Range
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        Set GetNextTableStartCell = ws.Range("A1")
    Else
        Set GetNextTableStartCell = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End If
End Function

Private Sub EnsureListColumn(ByVal lo As ListObject, ByVal columnName As String, ByRef issues As Collection)
    If GetColumnIndex(lo, columnName) > 0 Then Exit Sub

    lo.ListColumns.Add lo.ListColumns.Count + 1
    lo.ListColumns(lo.ListColumns.Count).Name = columnName
    issues.Add lo.Name & "." & columnName & " created"
End Sub

Private Function GetColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub RemoveBlankSeedRow(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If lo.ListRows.Count <> 1 Then Exit Sub
    If Not TableRowIsBlank(lo, 1) Then Exit Sub
    lo.ListRows(1).Delete
End Sub

Private Function TableRowIsBlank(ByVal lo As ListObject, ByVal rowIndex As Long) As Boolean
    Dim c As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex <= 0 Or rowIndex > lo.ListRows.Count Then Exit Function

    TableRowIsBlank = True
    For c = 1 To lo.ListColumns.Count
        If Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, c).Value)) <> "" Then
            TableRowIsBlank = False
            Exit Function
        End If
    Next c
End Function

Private Sub StyleProtectedHeaders(ByVal lo As ListObject, ByVal protectedHeaders As Variant)
    Dim key As Variant
    Dim idx As Long
    Dim hdr As Range
    Dim ws As Worksheet

    If lo Is Nothing Then Exit Sub
    Set ws = lo.Parent

    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Locked = False

    For Each key In protectedHeaders
        idx = GetColumnIndex(lo, CStr(key))
        If idx > 0 Then
            Set hdr = lo.HeaderRowRange.Cells(1, idx)
            hdr.Locked = True
            hdr.Interior.Color = RGB(60, 60, 60)
            hdr.Font.Color = RGB(255, 255, 255)
            hdr.Font.Bold = True
        End If
    Next key

    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Function JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String
    Dim itm As Variant
    For Each itm In items
        If Len(JoinCollection) > 0 Then JoinCollection = JoinCollection & delimiter
        JoinCollection = JoinCollection & CStr(itm)
    Next itm
End Function
