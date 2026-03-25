Attribute VB_Name = "modAdminConsole"
Option Explicit

Private Const SHEET_ADMIN_CONSOLE As String = "AdminConsole"
Private Const SHEET_ADMIN_AUDIT As String = "AdminAudit"
Private Const SHEET_ADMIN_POISON As String = "PoisonQueue"

Private Const TABLE_ADMIN_AUDIT As String = "tblAdminAudit"
Private Const TABLE_ADMIN_POISON As String = "tblAdminPoisonQueue"
Private Const TABLE_LOCKS As String = "tblLocks"
Private Const TABLE_APPLIED As String = "tblAppliedEvents"
Private Const TABLE_LOG As String = "tblInventoryLog"

Private Const TABLE_INBOX_RECEIVE As String = "tblInboxReceive"
Private Const TABLE_INBOX_SHIP As String = "tblInboxShip"
Private Const TABLE_INBOX_PROD As String = "tblInboxProd"

Private Const SNAPSHOT_SHEET As String = "InventorySnapshot"
Private Const SNAPSHOT_TABLE As String = "tblInventorySnapshot"

Public Function OpenAdminConsole(Optional ByVal adminWb As Workbook = Nothing, _
                                 Optional ByRef report As String = "") As Boolean
    Dim wb As Workbook

    Set wb = ResolveAdminWorkbook(adminWb)
    If wb Is Nothing Then
        report = "Admin workbook not resolved."
        Exit Function
    End If

    If Not EnsureAdminSchema(wb, report) Then Exit Function
    If Not RefreshAdminConsole(wb, report) Then Exit Function

    wb.Worksheets(SHEET_ADMIN_CONSOLE).Activate
    OpenAdminConsole = True
End Function

Public Function OpenUserManagement(Optional ByVal adminWb As Workbook = Nothing, _
                                   Optional ByRef report As String = "") As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet

    Set wb = ResolveAdminWorkbook(adminWb)
    If wb Is Nothing Then
        report = "Admin workbook not resolved."
        Exit Function
    End If

    If Not EnsureAdminSchema(wb, report) Then Exit Function

    On Error Resume Next
    Set ws = wb.Worksheets("UserCredentials")
    On Error GoTo 0

    If ws Is Nothing Then
        wb.Worksheets(SHEET_ADMIN_CONSOLE).Activate
        report = "UserCredentials sheet not found; opened AdminConsole instead."
    Else
        ws.Activate
        report = "OK"
    End If

    OpenUserManagement = True
End Function

Public Function EnsureAdminSchema(Optional ByVal adminWb As Workbook = Nothing, _
                                  Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Dim wsConsole As Worksheet
    Dim issues As Collection

    Set wb = ResolveAdminWorkbook(adminWb)
    If wb Is Nothing Then
        report = "Admin workbook not resolved."
        Exit Function
    End If

    Set issues = New Collection
    Set wsConsole = EnsureWorksheetAdmin(wb, SHEET_ADMIN_CONSOLE)
    InitializeAdminConsoleLayout wsConsole

    EnsureAdminTableWithHeaders wb, SHEET_ADMIN_AUDIT, TABLE_ADMIN_AUDIT, _
        Array("LoggedAtUTC", "Action", "UserId", "WarehouseId", "StationId", "TargetType", "TargetId", "Reason", "Detail", "Result"), issues
    EnsureAdminTableWithHeaders wb, SHEET_ADMIN_POISON, TABLE_ADMIN_POISON, _
        Array("SourceWorkbook", "SourceTable", "RowIndex", "EventID", "ParentEventId", "UndoOfEventId", "EventType", "CreatedAtUTC", _
              "WarehouseId", "StationId", "UserId", "SKU", "Qty", "Location", "Note", "PayloadJson", "Status", "RetryCount", _
              "ErrorCode", "ErrorMessage", "FailedAtUTC"), issues

    report = JoinIssuesAdmin(issues)
    EnsureAdminSchema = True
    Exit Function

FailEnsure:
    report = "EnsureAdminSchema failed: " & Err.Description
End Function

Public Function RefreshAdminConsole(Optional ByVal adminWb As Workbook = Nothing, _
                                    Optional ByRef report As String = "") As Boolean
    On Error GoTo FailRefresh

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim inventoryWb As Workbook
    Dim poisonReport As String
    Dim poisonCount As Long
    Dim newCount As Long
    Dim processedCount As Long
    Dim skipDupCount As Long
    Dim lockRow As Long
    Dim loLocks As ListObject
    Dim loApplied As ListObject

    Set wb = ResolveAdminWorkbook(adminWb)
    If wb Is Nothing Then
        report = "Admin workbook not resolved."
        Exit Function
    End If
    If Not EnsureAdminSchema(wb, report) Then Exit Function

    If Not modConfig.LoadConfig("", "") Then
        report = "Config load failed: " & modConfig.Validate()
        Exit Function
    End If
    If Not modAuth.LoadAuth(modConfig.GetWarehouseId()) Then
        report = "Auth load failed: " & modAuth.ValidateAuth()
        Exit Function
    End If

    poisonCount = RefreshPoisonQueue(wb, poisonReport)
    If poisonReport <> "" And poisonReport <> "OK" Then
        report = poisonReport
        Exit Function
    End If
    CountInboxStatuses newCount, processedCount, skipDupCount

    Set ws = wb.Worksheets(SHEET_ADMIN_CONSOLE)
    Set inventoryWb = modInventoryApply.ResolveInventoryWorkbook(modConfig.GetWarehouseId())
    If Not inventoryWb Is Nothing Then
        Set loLocks = FindListObjectByNameAdmin(inventoryWb, TABLE_LOCKS)
        Set loApplied = FindListObjectByNameAdmin(inventoryWb, TABLE_APPLIED)
    End If

    lockRow = FindLockRowAdmin(loLocks, "INVENTORY")

    ws.Range("B3").Value = modConfig.GetWarehouseId()
    ws.Range("B4").Value = modConfig.GetStationId()
    ws.Range("B5").Value = modConfig.GetString("ProcessorServiceUserId", "svc_processor")
    ws.Range("B6").Value = newCount
    ws.Range("B7").Value = poisonCount
    ws.Range("B8").Value = processedCount
    ws.Range("B9").Value = skipDupCount
    ws.Range("B10").Value = GetLockFieldAdmin(loLocks, lockRow, "Status")
    ws.Range("B11").Value = GetLockFieldAdmin(loLocks, lockRow, "RunId")
    ws.Range("B12").Value = GetLockFieldAdmin(loLocks, lockRow, "OwnerUserId")
    ws.Range("B13").Value = GetLatestAppliedValue(loApplied, "AppliedAtUTC")
    ws.Range("B14").Value = GetLatestAppliedValue(loApplied, "EventID")
    ws.Range("B15").Value = Now

    report = "OK"
    RefreshAdminConsole = True
    Exit Function

FailRefresh:
    report = "RefreshAdminConsole failed: " & Err.Description
End Function

Public Function RunProcessorFromConsole(Optional ByVal adminUserId As String = "", _
                                        Optional ByVal warehouseId As String = "", _
                                        Optional ByVal adminWb As Workbook = Nothing, _
                                        Optional ByRef report As String = "") As Long
    Dim resolvedWh As String
    Dim resolvedSt As String
    Dim resolvedUser As String
    Dim auditWb As Workbook
    Dim refreshReport As String
    Dim resultCode As String

    Set auditWb = ResolveAdminWorkbook(adminWb)

    If Not EnsureAdminContext(adminUserId, warehouseId, resolvedUser, resolvedWh, resolvedSt, report) Then Exit Function
    If Not RequireAdminMaintenance(resolvedUser, resolvedWh, resolvedSt, report) Then Exit Function

    RunProcessorFromConsole = modProcessor.RunBatch(resolvedWh, 0, report)
    resultCode = IIf(Left$(report, 15) = "RunBatch failed", "FAIL", "OK")
    AppendAuditEntry auditWb, "RUN_PROCESSOR", resolvedUser, resolvedWh, resolvedSt, "WAREHOUSE", resolvedWh, "", report, resultCode
    Call RefreshAdminConsole(auditWb, refreshReport)
End Function

Public Function BreakInventoryLock(Optional ByVal reason As String = "", _
                                   Optional ByVal adminUserId As String = "", _
                                   Optional ByVal warehouseId As String = "", _
                                   Optional ByVal inventoryWb As Workbook = Nothing, _
                                   Optional ByVal adminWb As Workbook = Nothing, _
                                   Optional ByRef report As String = "") As Boolean
    Dim resolvedWh As String
    Dim resolvedSt As String
    Dim resolvedUser As String
    Dim auditWb As Workbook
    Dim refreshReport As String

    If Trim$(reason) = "" Then
        report = "Reason is required to break a lock."
        Exit Function
    End If

    Set auditWb = ResolveAdminWorkbook(adminWb)
    If Not EnsureAdminContext(adminUserId, warehouseId, resolvedUser, resolvedWh, resolvedSt, report) Then Exit Function
    If Not RequireAdminMaintenance(resolvedUser, resolvedWh, resolvedSt, report) Then Exit Function

    BreakInventoryLock = modLockManager.BreakLock("INVENTORY", resolvedWh, resolvedUser, reason, inventoryWb, report)
    AppendAuditEntry auditWb, "BREAK_LOCK", resolvedUser, resolvedWh, resolvedSt, "LOCK", "INVENTORY", reason, report, IIf(BreakInventoryLock, "OK", "FAIL")
    Call RefreshAdminConsole(auditWb, refreshReport)
End Function

Public Function RefreshPoisonQueue(Optional ByVal adminWb As Workbook = Nothing, _
                                   Optional ByRef report As String = "") As Long
    On Error GoTo FailRefresh

    Dim wb As Workbook
    Dim loPoison As ListObject
    Dim inboxWb As Workbook
    Dim loInbox As ListObject

    Set wb = ResolveAdminWorkbook(adminWb)
    If wb Is Nothing Then
        report = "Admin workbook not resolved."
        Exit Function
    End If
    If Not EnsureAdminSchema(wb, report) Then Exit Function

    Set loPoison = wb.Worksheets(SHEET_ADMIN_POISON).ListObjects(TABLE_ADMIN_POISON)
    DeleteAllAdminTableRows loPoison

    For Each inboxWb In Application.Workbooks
        Set loInbox = FindListObjectByNameAdmin(inboxWb, TABLE_INBOX_RECEIVE)
        If Not loInbox Is Nothing Then RefreshPoisonQueue = RefreshPoisonQueue + AppendPoisonRowsFromInbox(loPoison, inboxWb, loInbox)

        Set loInbox = FindListObjectByNameAdmin(inboxWb, TABLE_INBOX_SHIP)
        If Not loInbox Is Nothing Then RefreshPoisonQueue = RefreshPoisonQueue + AppendPoisonRowsFromInbox(loPoison, inboxWb, loInbox)

        Set loInbox = FindListObjectByNameAdmin(inboxWb, TABLE_INBOX_PROD)
        If Not loInbox Is Nothing Then RefreshPoisonQueue = RefreshPoisonQueue + AppendPoisonRowsFromInbox(loPoison, inboxWb, loInbox)
    Next inboxWb

    If loPoison.DataBodyRange Is Nothing Then loPoison.ListRows.Add
    report = "OK"
    Exit Function

FailRefresh:
    report = "RefreshPoisonQueue failed: " & Err.Description
End Function

Public Function ReissuePoisonEvent(ByVal sourceWorkbookName As String, _
                                   ByVal tableName As String, _
                                   ByVal sourceEventId As String, _
                                   Optional ByVal adminUserId As String = "", _
                                   Optional ByVal corrections As Object = Nothing, _
                                   Optional ByVal reason As String = "", _
                                   Optional ByVal adminWb As Workbook = Nothing, _
                                   Optional ByRef newEventIdOut As String = "", _
                                   Optional ByRef report As String = "") As Boolean
    On Error GoTo FailReissue

    Dim resolvedWh As String
    Dim resolvedSt As String
    Dim resolvedUser As String
    Dim sourceWb As Workbook
    Dim loInbox As ListObject
    Dim sourceRow As Long
    Dim targetRow As ListRow
    Dim colIdx As Long
    Dim key As Variant
    Dim refreshReport As String

    If Not EnsureAdminContext(adminUserId, "", resolvedUser, resolvedWh, resolvedSt, report) Then Exit Function
    If Not RequireAdminMaintenance(resolvedUser, resolvedWh, resolvedSt, report) Then Exit Function

    Set sourceWb = ResolveOpenWorkbookByNameAdmin(sourceWorkbookName)
    If sourceWb Is Nothing Then
        report = "Source workbook not open: " & sourceWorkbookName
        Exit Function
    End If

    Set loInbox = FindListObjectByNameAdmin(sourceWb, tableName)
    If loInbox Is Nothing Then
        report = "Source table not found: " & tableName
        Exit Function
    End If

    sourceRow = FindRowByColumnValueAdmin(loInbox, "EventID", sourceEventId)
    If sourceRow = 0 Then
        report = "Source EventID not found: " & sourceEventId
        Exit Function
    End If
    If UCase$(SafeTrimAdmin(GetCellByColumnAdmin(loInbox, sourceRow, "Status"))) <> INBOX_STATUS_POISON Then
        report = "Source event is not in POISON status."
        Exit Function
    End If

    EnsureTableSheetEditableAdmin loInbox, loInbox.Name
    Set targetRow = loInbox.ListRows.Add
    For colIdx = 1 To loInbox.ListColumns.Count
        targetRow.Range.Cells(1, colIdx).Value = loInbox.DataBodyRange.Cells(sourceRow, colIdx).Value
    Next colIdx

    newEventIdOut = "EVT-REISSUE-" & GenerateGuidAdmin()
    SetTableRowValueAdmin loInbox, targetRow.Index, "EventID", newEventIdOut
    SetTableRowValueAdmin loInbox, targetRow.Index, "ParentEventId", sourceEventId
    SetTableRowValueAdmin loInbox, targetRow.Index, "CreatedAtUTC", Now
    SetTableRowValueAdmin loInbox, targetRow.Index, "Status", INBOX_STATUS_NEW
    SetTableRowValueAdmin loInbox, targetRow.Index, "RetryCount", 0
    SetTableRowValueAdmin loInbox, targetRow.Index, "ErrorCode", vbNullString
    SetTableRowValueAdmin loInbox, targetRow.Index, "ErrorMessage", vbNullString
    SetTableRowValueAdmin loInbox, targetRow.Index, "FailedAtUTC", vbNullString

    If Not corrections Is Nothing Then
        For Each key In corrections.Keys
            If HasListColumnAdmin(loInbox, CStr(key)) Then
                SetTableRowValueAdmin loInbox, targetRow.Index, CStr(key), corrections(key)
            End If
        Next key
    End If

    If Len(Trim$(reason)) = 0 Then reason = "Reissued from poison queue."
    AppendAuditEntry ResolveAdminWorkbook(adminWb), "REISSUE_POISON", resolvedUser, resolvedWh, resolvedSt, _
                     tableName, sourceEventId, reason, "NewEventID=" & newEventIdOut, "OK"
    Call RefreshAdminConsole(adminWb, refreshReport)
    report = "OK"
    ReissuePoisonEvent = True
    Exit Function

FailReissue:
    report = "ReissuePoisonEvent failed: " & Err.Description
End Function

Public Function GenerateInventorySnapshot(Optional ByVal adminUserId As String = "", _
                                          Optional ByVal warehouseId As String = "", _
                                          Optional ByVal inventoryWb As Workbook = Nothing, _
                                          Optional ByVal outputPath As String = "", _
                                          Optional ByVal adminWb As Workbook = Nothing, _
                                          Optional ByRef report As String = "") As Boolean
    On Error GoTo FailSnapshot

    Dim resolvedWh As String
    Dim resolvedSt As String
    Dim resolvedUser As String
    Dim sourceInvWb As Workbook
    Dim snapPath As String

    If Not EnsureAdminContext(adminUserId, warehouseId, resolvedUser, resolvedWh, resolvedSt, report) Then Exit Function
    If Not RequireAdminMaintenance(resolvedUser, resolvedWh, resolvedSt, report) Then Exit Function

    Set sourceInvWb = modInventoryApply.ResolveInventoryWorkbook(resolvedWh, inventoryWb)
    If sourceInvWb Is Nothing Then
        report = "Inventory workbook not found."
        Exit Function
    End If

    snapPath = vbNullString
    If Not modWarehouseSync.GenerateWarehouseSnapshot(resolvedWh, sourceInvWb, outputPath, Nothing, snapPath) Then
        report = snapPath
        Exit Function
    End If

    AppendAuditEntry ResolveAdminWorkbook(adminWb), "GENERATE_SNAPSHOT", resolvedUser, resolvedWh, resolvedSt, _
                     "SNAPSHOT", snapPath, "", snapPath, "OK"
    report = snapPath
    GenerateInventorySnapshot = True
    Exit Function

FailSnapshot:
    report = "GenerateInventorySnapshot failed: " & Err.Description
End Function

Private Function ResolveAdminWorkbook(ByVal adminWb As Workbook) As Workbook
    If Not adminWb Is Nothing Then
        Set ResolveAdminWorkbook = adminWb
    Else
        Set ResolveAdminWorkbook = ThisWorkbook
    End If
End Function

Private Function EnsureWorksheetAdmin(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheetAdmin = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheetAdmin Is Nothing Then
        Set EnsureWorksheetAdmin = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheetAdmin.Name = sheetName
    End If
End Function

Private Sub InitializeAdminConsoleLayout(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    ws.Range("A1").Value = "Admin Console"
    ws.Range("A2").Value = "Metric"
    ws.Range("B2").Value = "Value"
    ws.Range("A3").Value = "WarehouseId"
    ws.Range("A4").Value = "StationId"
    ws.Range("A5").Value = "ProcessorServiceUserId"
    ws.Range("A6").Value = "InboxNewCount"
    ws.Range("A7").Value = "InboxPoisonCount"
    ws.Range("A8").Value = "InboxProcessedCount"
    ws.Range("A9").Value = "InboxSkipDupCount"
    ws.Range("A10").Value = "InventoryLockStatus"
    ws.Range("A11").Value = "InventoryLockRunId"
    ws.Range("A12").Value = "InventoryLockOwner"
    ws.Range("A13").Value = "LastAppliedAtUTC"
    ws.Range("A14").Value = "LastAppliedEventId"
    ws.Range("A15").Value = "LastRefreshUTC"
    ws.Columns("A:B").AutoFit
End Sub

Private Sub EnsureAdminTableWithHeaders(ByVal wb As Workbook, _
                                        ByVal sheetName As String, _
                                        ByVal tableName As String, _
                                        ByVal headers As Variant, _
                                        ByVal issues As Collection)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim startCell As Range
    Dim i As Long

    Set ws = EnsureWorksheetAdmin(wb, sheetName)
    EnsureTableSheetEditableAdminSheet ws

    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCellAdmin(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers))), , xlYes)
        lo.Name = tableName
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumnAdmin lo, CStr(headers(i))
    Next i

    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

Private Function GetNextTableStartCellAdmin(ByVal ws As Worksheet) As Range
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        Set GetNextTableStartCellAdmin = ws.Range("A1")
    Else
        Set GetNextTableStartCellAdmin = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End If
End Function

Private Sub EnsureListColumnAdmin(ByVal lo As ListObject, ByVal columnName As String)
    If HasListColumnAdmin(lo, columnName) Then Exit Sub
    lo.ListColumns.Add lo.ListColumns.Count + 1
    lo.ListColumns(lo.ListColumns.Count).Name = columnName
End Sub

Private Function HasListColumnAdmin(ByVal lo As ListObject, ByVal columnName As String) As Boolean
    HasListColumnAdmin = (GetColumnIndexAdmin(lo, columnName) > 0)
End Function

Private Function JoinIssuesAdmin(ByVal issues As Collection) As String
    Dim itm As Variant
    If issues Is Nothing Then Exit Function
    For Each itm In issues
        If Len(JoinIssuesAdmin) > 0 Then JoinIssuesAdmin = JoinIssuesAdmin & "; "
        JoinIssuesAdmin = JoinIssuesAdmin & CStr(itm)
    Next itm
End Function

Private Function EnsureAdminContext(ByVal adminUserId As String, _
                                    ByVal warehouseId As String, _
                                    ByRef resolvedUser As String, _
                                    ByRef resolvedWh As String, _
                                    ByRef resolvedSt As String, _
                                    ByRef report As String) As Boolean
    If Not modConfig.LoadConfig(warehouseId, "") Then
        report = "Config load failed: " & modConfig.Validate()
        Exit Function
    End If
    resolvedWh = modConfig.GetString("WarehouseId", warehouseId)
    resolvedSt = modConfig.GetString("StationId", "")

    If Not modAuth.LoadAuth(resolvedWh) Then
        report = "Auth load failed: " & modAuth.ValidateAuth()
        Exit Function
    End If

    resolvedUser = Trim$(adminUserId)
    If resolvedUser = "" Then resolvedUser = modRoleEventWriter.ResolveCurrentUserId()
    If resolvedUser = "" Then
        report = "Admin user could not be resolved."
        Exit Function
    End If

    EnsureAdminContext = True
End Function

Private Function RequireAdminMaintenance(ByVal adminUserId As String, _
                                         ByVal warehouseId As String, _
                                         ByVal stationId As String, _
                                         ByRef report As String) As Boolean
    If modAuth.CanPerform("ADMIN_MAINT", adminUserId, warehouseId, stationId, "ADMIN", "ADMIN-MAINT") Then
        RequireAdminMaintenance = True
    Else
        report = "User lacks ADMIN_MAINT capability."
    End If
End Function

Private Sub AppendAuditEntry(ByVal adminWb As Workbook, _
                             ByVal actionName As String, _
                             ByVal userId As String, _
                             ByVal warehouseId As String, _
                             ByVal stationId As String, _
                             ByVal targetType As String, _
                             ByVal targetId As String, _
                             ByVal reason As String, _
                             ByVal detail As String, _
                             ByVal resultCode As String)
    Dim report As String
    Dim lo As ListObject
    Dim r As ListRow
    Dim wb As Workbook

    Set wb = ResolveAdminWorkbook(adminWb)
    If wb Is Nothing Then Exit Sub
    If Not EnsureAdminSchema(wb, report) Then Exit Sub

    Set lo = wb.Worksheets(SHEET_ADMIN_AUDIT).ListObjects(TABLE_ADMIN_AUDIT)
    EnsureTableSheetEditableAdmin lo, TABLE_ADMIN_AUDIT
    Set r = lo.ListRows.Add
    SetTableRowValueAdmin lo, r.Index, "LoggedAtUTC", Now
    SetTableRowValueAdmin lo, r.Index, "Action", actionName
    SetTableRowValueAdmin lo, r.Index, "UserId", userId
    SetTableRowValueAdmin lo, r.Index, "WarehouseId", warehouseId
    SetTableRowValueAdmin lo, r.Index, "StationId", stationId
    SetTableRowValueAdmin lo, r.Index, "TargetType", targetType
    SetTableRowValueAdmin lo, r.Index, "TargetId", targetId
    SetTableRowValueAdmin lo, r.Index, "Reason", reason
    SetTableRowValueAdmin lo, r.Index, "Detail", detail
    SetTableRowValueAdmin lo, r.Index, "Result", resultCode
End Sub

Private Sub CountInboxStatuses(ByRef newCount As Long, ByRef processedCount As Long, ByRef skipDupCount As Long)
    Dim wb As Workbook
    Dim lo As ListObject

    For Each wb In Application.Workbooks
        Set lo = FindListObjectByNameAdmin(wb, TABLE_INBOX_RECEIVE)
        If Not lo Is Nothing Then CountStatusesInTable lo, newCount, processedCount, skipDupCount

        Set lo = FindListObjectByNameAdmin(wb, TABLE_INBOX_SHIP)
        If Not lo Is Nothing Then CountStatusesInTable lo, newCount, processedCount, skipDupCount

        Set lo = FindListObjectByNameAdmin(wb, TABLE_INBOX_PROD)
        If Not lo Is Nothing Then CountStatusesInTable lo, newCount, processedCount, skipDupCount
    Next wb
End Sub

Private Sub CountStatusesInTable(ByVal lo As ListObject, _
                                 ByRef newCount As Long, _
                                 ByRef processedCount As Long, _
                                 ByRef skipDupCount As Long)
    Dim i As Long
    Dim statusVal As String

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    For i = 1 To lo.ListRows.Count
        If SafeTrimAdmin(GetCellByColumnAdmin(lo, i, "EventID")) <> "" Then
            statusVal = UCase$(SafeTrimAdmin(GetCellByColumnAdmin(lo, i, "Status")))
            Select Case statusVal
                Case "", INBOX_STATUS_NEW
                    newCount = newCount + 1
                Case INBOX_STATUS_PROCESSED
                    processedCount = processedCount + 1
                Case INBOX_STATUS_SKIP_DUP
                    skipDupCount = skipDupCount + 1
            End Select
        End If
    Next i
End Sub

Private Function AppendPoisonRowsFromInbox(ByVal loPoison As ListObject, ByVal sourceWb As Workbook, ByVal loInbox As ListObject) As Long
    Dim i As Long
    Dim statusVal As String
    Dim r As ListRow
    Dim columnName As Variant

    If loInbox Is Nothing Or loInbox.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To loInbox.ListRows.Count
        statusVal = UCase$(SafeTrimAdmin(GetCellByColumnAdmin(loInbox, i, "Status")))
        If statusVal = INBOX_STATUS_POISON And SafeTrimAdmin(GetCellByColumnAdmin(loInbox, i, "EventID")) <> "" Then
            EnsureTableSheetEditableAdmin loPoison, TABLE_ADMIN_POISON
            Set r = loPoison.ListRows.Add
            SetTableRowValueAdmin loPoison, r.Index, "SourceWorkbook", sourceWb.Name
            SetTableRowValueAdmin loPoison, r.Index, "SourceTable", loInbox.Name
            SetTableRowValueAdmin loPoison, r.Index, "RowIndex", i
            For Each columnName In Array("EventID", "ParentEventId", "UndoOfEventId", "EventType", "CreatedAtUTC", "WarehouseId", "StationId", _
                                         "UserId", "SKU", "Qty", "Location", "Note", "PayloadJson", "Status", "RetryCount", _
                                         "ErrorCode", "ErrorMessage", "FailedAtUTC")
                If HasListColumnAdmin(loInbox, CStr(columnName)) Then
                    SetTableRowValueAdmin loPoison, r.Index, CStr(columnName), GetCellByColumnAdmin(loInbox, i, CStr(columnName))
                End If
            Next columnName
            AppendPoisonRowsFromInbox = AppendPoisonRowsFromInbox + 1
        End If
    Next i
End Function

Private Sub DeleteAllAdminTableRows(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    EnsureTableSheetEditableAdmin lo, lo.Name
    Do While lo.ListRows.Count > 0
        lo.ListRows(lo.ListRows.Count).Delete
    Loop
End Sub

Private Function FindListObjectByNameAdmin(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameAdmin = ws.ListObjects(tableName)
        If Not FindListObjectByNameAdmin Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function FindLockRowAdmin(ByVal lo As ListObject, ByVal lockName As String) As Long
    Dim i As Long
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(UCase$(SafeTrimAdmin(GetCellByColumnAdmin(lo, i, "LockName"))), UCase$(lockName), vbTextCompare) = 0 Then
            FindLockRowAdmin = i
            Exit Function
        End If
    Next i
End Function

Private Function GetLockFieldAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    If lo Is Nothing Then Exit Function
    If rowIndex <= 0 Then Exit Function
    GetLockFieldAdmin = GetCellByColumnAdmin(lo, rowIndex, columnName)
End Function

Private Function GetLatestAppliedValue(ByVal lo As ListObject, ByVal columnName As String) As Variant
    Dim i As Long
    Dim latestRow As Long
    Dim latestDate As Date
    Dim currentDate As Variant

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        currentDate = GetCellByColumnAdmin(lo, i, "AppliedAtUTC")
        If IsDate(currentDate) Then
            If latestRow = 0 Or CDate(currentDate) > latestDate Then
                latestRow = i
                latestDate = CDate(currentDate)
            End If
        End If
    Next i
    If latestRow > 0 Then GetLatestAppliedValue = GetCellByColumnAdmin(lo, latestRow, columnName)
End Function

Private Function FindRowByColumnValueAdmin(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrimAdmin(GetCellByColumnAdmin(lo, i, columnName)), expectedValue, vbTextCompare) = 0 Then
            FindRowByColumnValueAdmin = i
            Exit Function
        End If
    Next i
End Function

Private Function GetCellByColumnAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndexAdmin(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnAdmin = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Sub SetTableRowValueAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = GetColumnIndexAdmin(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetColumnIndexAdmin(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexAdmin = i
            Exit Function
        End If
    Next i
End Function

Private Function SafeTrimAdmin(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimAdmin = Trim$(CStr(valueIn))
End Function

Private Function ResolveOpenWorkbookByNameAdmin(ByVal workbookName As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Or _
           StrComp(wb.FullName, workbookName, vbTextCompare) = 0 Then
            Set ResolveOpenWorkbookByNameAdmin = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub EnsureTableSheetEditableAdmin(ByVal lo As ListObject, ByVal tableName As String)
    If lo Is Nothing Then Exit Sub
    EnsureTableSheetEditableAdminSheet lo.Parent
    If lo.Parent.ProtectContents Then
        Err.Raise vbObjectError + 2901, "modAdminConsole.EnsureTableSheetEditableAdmin", _
                  "Worksheet '" & lo.Parent.Name & "' is protected and could not be unprotected before updating " & tableName & "."
    End If
End Sub

Private Sub EnsureTableSheetEditableAdminSheet(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
End Sub

Private Function ResolveSnapshotPathAdmin(ByVal warehouseId As String, ByVal outputPath As String) As String
    Dim rootPath As String
    If Trim$(outputPath) <> "" Then
        ResolveSnapshotPathAdmin = outputPath
        Exit Function
    End If

    rootPath = modConfig.GetString("PathDataRoot", Environ$("TEMP"))
    If Right$(rootPath, 1) = "\" Then rootPath = Left$(rootPath, Len(rootPath) - 1)
    ResolveSnapshotPathAdmin = rootPath & "\Snapshots\" & warehouseId & ".invSys.InventorySnapshot." & _
                               Format$(Now, "yyyymmdd_hhnnss") & ".xlsb"
End Function

Private Function GenerateGuidAdmin() As String
    Dim i As Long
    Dim token As String
    Const chars As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

    Randomize
    For i = 1 To 32
        token = token & Mid$(chars, Int((Len(chars) * Rnd) + 1), 1)
    Next i
    GenerateGuidAdmin = Left$(token, 8) & "-" & Mid$(token, 9, 4) & "-" & Mid$(token, 13, 4) & "-" & Mid$(token, 17, 4) & "-" & Right$(token, 12)
End Function

Private Sub EnsureFolderForFileAdmin(ByVal filePath As String)
    Dim folderPath As String
    Dim sepPos As Long

    sepPos = InStrRev(filePath, "\")
    If sepPos <= 0 Then Exit Sub
    folderPath = Left$(filePath, sepPos - 1)
    CreateFolderRecursiveAdmin folderPath
End Sub

Private Sub CreateFolderRecursiveAdmin(ByVal folderPath As String)
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
        If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then CreateFolderRecursiveAdmin parentPath
    End If

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub

Private Sub CloseWorkbookByFullNameAdmin(ByVal fullNameIn As String)
    Dim wb As Workbook
    If Trim$(fullNameIn) = "" Then Exit Sub
    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullNameIn, vbTextCompare) = 0 Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            On Error GoTo 0
            Exit For
        End If
    Next wb
End Sub
