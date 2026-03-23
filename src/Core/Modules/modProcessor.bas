Attribute VB_Name = "modProcessor"
Option Explicit

Public Const INBOX_STATUS_NEW As String = "NEW"
Public Const INBOX_STATUS_PROCESSED As String = "PROCESSED"
Public Const INBOX_STATUS_SKIP_DUP As String = "SKIP_DUP"
Public Const INBOX_STATUS_POISON As String = "POISON"

Private Const SHEET_INBOX_RECEIVE As String = "InboxReceive"
Private Const SHEET_INBOX_SHIP As String = "InboxShip"
Private Const SHEET_INBOX_PROD As String = "InboxProd"

Private Const TABLE_INBOX_RECEIVE As String = "tblInboxReceive"
Private Const TABLE_INBOX_SHIP As String = "tblInboxShip"
Private Const TABLE_INBOX_PROD As String = "tblInboxProd"

Public Function RunBatch(Optional ByVal warehouseId As String = "", _
                         Optional ByVal batchSize As Long = 0, _
                         Optional ByRef report As String = "") As Long
    On Error GoTo FailRun

    Dim inventoryWb As Workbook
    Dim inboxTargets As Collection
    Dim target As Variant
    Dim loInbox As ListObject
    Dim rowIndex As Long
    Dim runId As String
    Dim message As String
    Dim serviceUserId As String
    Dim skipDupCount As Long
    Dim poisonCount As Long
    Dim heartbeatSeconds As Long
    Dim lastHeartbeat As Date
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim evt As Object
    Dim lockHeld As Boolean
    Dim capability As String
    Dim artifactWarnings As Long
    Dim artifactReport As String

    If Not EnsurePhase2Context(warehouseId, report) Then Exit Function

    warehouseId = modConfig.GetString("WarehouseId", warehouseId)
    If warehouseId = "" Then
        report = "WarehouseId not resolved."
        Exit Function
    End If

    serviceUserId = modConfig.GetString("ProcessorServiceUserId", "svc_processor")
    If serviceUserId = "" Then serviceUserId = "svc_processor"

    If batchSize <= 0 Then batchSize = modConfig.GetLong("BatchSize", 500)
    If batchSize <= 0 Then batchSize = 500

    heartbeatSeconds = modConfig.GetLong("HeartbeatIntervalSeconds", 30)
    If heartbeatSeconds <= 0 Then heartbeatSeconds = 30

    If Not modAuth.CanPerform("INBOX_PROCESS", serviceUserId, warehouseId, modConfig.GetString("StationId", ""), "PROCESSOR", "PROCESSOR-RUN") Then
        report = "Processor service identity lacks INBOX_PROCESS."
        Exit Function
    End If

    Set inventoryWb = ResolveInventoryWorkbookBridge(warehouseId)
    If inventoryWb Is Nothing Then
        report = "Inventory workbook not found."
        Exit Function
    End If

    If Not modLockManager.AcquireLock("INVENTORY", warehouseId, serviceUserId, modConfig.GetString("StationId", ""), inventoryWb, runId, message) Then
        report = message
        Exit Function
    End If
    lockHeld = True
    lastHeartbeat = Now

    Set inboxTargets = ResolveInboxTargets()
    For Each target In inboxTargets
        If Not EnsureInboxTargetSchema(target("Workbook"), CStr(target("TableName")), report) Then GoTo ContinueInbox

        Set loInbox = FindListObjectByNameProcessor(target("Workbook"), CStr(target("TableName")))
        If loInbox Is Nothing Then GoTo ContinueInbox
        If loInbox.DataBodyRange Is Nothing Then GoTo ContinueInbox

        For rowIndex = 1 To loInbox.ListRows.Count
            If RunBatch >= batchSize Then Exit For
            If Not IsProcessableInboxRow(loInbox, rowIndex, warehouseId) Then GoTo ContinueRow

            Set evt = BuildInboxEvent(loInbox, rowIndex, target("Workbook").Name, CStr(target("TableName")), CStr(target("DefaultEventType")))
            If evt Is Nothing Then
                UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_POISON, "INVALID_EVENT", "Unable to read inbox row."
                poisonCount = poisonCount + 1
                GoTo MaybeHeartbeat
            End If

            capability = CapabilityForEventType(GetDictionaryString(evt, "EventType"))
            If capability = "" Then
                UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_POISON, "INVALID_EVENT_TYPE", "Unsupported EventType."
                poisonCount = poisonCount + 1
                GoTo MaybeHeartbeat
            End If

            If Not modAuth.CanPerform(capability, GetDictionaryString(evt, "UserId"), GetDictionaryString(evt, "WarehouseId"), GetDictionaryString(evt, "StationId"), "PROCESSOR_VALIDATE", GetDictionaryString(evt, "EventID")) Then
                UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_POISON, "AUTH_DENIED", "Event creator lacks " & capability & " capability."
                poisonCount = poisonCount + 1
                GoTo MaybeHeartbeat
            End If

            statusOut = vbNullString
            errorCode = vbNullString
            errorMessage = vbNullString

            If ApplyInventoryEventBridge(evt, inventoryWb, runId, statusOut, errorCode, errorMessage) Then
                Select Case UCase$(statusOut)
                    Case APPLY_STATUS_APPLIED
                        artifactReport = vbNullString
                        If Not modWarehouseSync.AppendEventToOutbox(evt, inventoryWb, Nothing, runId, artifactReport) Then artifactWarnings = artifactWarnings + 1
                        UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_PROCESSED
                        RunBatch = RunBatch + 1
                    Case APPLY_STATUS_SKIP_DUP
                        artifactReport = vbNullString
                        If Not modWarehouseSync.AppendEventToOutbox(evt, inventoryWb, Nothing, runId, artifactReport) Then artifactWarnings = artifactWarnings + 1
                        UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_SKIP_DUP
                        skipDupCount = skipDupCount + 1
                    Case Else
                        UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_POISON, "UNKNOWN_APPLY_STATUS", "Unknown apply status."
                        poisonCount = poisonCount + 1
                End Select
            Else
                UpdateInboxRowStatus loInbox, rowIndex, INBOX_STATUS_POISON, errorCode, errorMessage
                poisonCount = poisonCount + 1
            End If

MaybeHeartbeat:
            If DateDiff("s", lastHeartbeat, Now) >= heartbeatSeconds Then
                Call modLockManager.UpdateHeartbeat("INVENTORY", runId, inventoryWb)
                lastHeartbeat = Now
            End If

ContinueRow:
        Next rowIndex

        If RunBatch >= batchSize Then Exit For
ContinueInbox:
    Next target

    report = "Applied=" & CStr(RunBatch) & "; SkipDup=" & CStr(skipDupCount) & "; Poison=" & CStr(poisonCount) & "; RunId=" & runId
    If artifactWarnings > 0 Then report = report & "; ArtifactWarnings=" & CStr(artifactWarnings)

    artifactReport = vbNullString
    If Not modWarehouseSync.GenerateWarehouseSnapshot(warehouseId, inventoryWb, "", Nothing, artifactReport) Then
        If report <> "" Then report = report & "; "
        report = report & "SnapshotError=" & artifactReport
    End If

CleanExit:
    If lockHeld Then Call modLockManager.ReleaseLock("INVENTORY", runId, inventoryWb)
    Exit Function

FailRun:
    report = "RunBatch failed: " & Err.Description
    Resume CleanExit
End Function

Public Function RunBatchForAutomation(Optional ByVal warehouseId As String = "", _
                                      Optional ByVal batchSize As Long = 0) As Long
    Dim report As String

    RunBatchForAutomation = RunBatch(warehouseId, batchSize, report)
End Function

Public Function RunBatchReportForAutomation(Optional ByVal warehouseId As String = "", _
                                            Optional ByVal batchSize As Long = 0) As String
    Dim report As String
    Dim processedCount As Long

    processedCount = RunBatch(warehouseId, batchSize, report)
    RunBatchReportForAutomation = "Processed=" & CStr(processedCount) & "; Report=" & report
End Function

Public Function EnsureReceiveInboxSchema(Optional ByVal targetWb As Workbook = Nothing, _
                                         Optional ByRef report As String = "") As Boolean
    EnsureReceiveInboxSchema = EnsureInboxSchemaCore(targetWb, report, SHEET_INBOX_RECEIVE, TABLE_INBOX_RECEIVE, EVENT_TYPE_RECEIVE)
End Function

Public Function EnsureShipInboxSchema(Optional ByVal targetWb As Workbook = Nothing, _
                                      Optional ByRef report As String = "") As Boolean
    EnsureShipInboxSchema = EnsureInboxSchemaCore(targetWb, report, SHEET_INBOX_SHIP, TABLE_INBOX_SHIP, EVENT_TYPE_SHIP)
End Function

Public Function EnsureProductionInboxSchema(Optional ByVal targetWb As Workbook = Nothing, _
                                            Optional ByRef report As String = "") As Boolean
    EnsureProductionInboxSchema = EnsureInboxSchemaCore(targetWb, report, SHEET_INBOX_PROD, TABLE_INBOX_PROD, EVENT_TYPE_PROD_CONSUME)
End Function

Private Function EnsureInboxTargetSchema(ByVal targetWb As Workbook, ByVal tableName As String, ByRef report As String) As Boolean
    Select Case UCase$(tableName)
        Case UCase$(TABLE_INBOX_RECEIVE)
            EnsureInboxTargetSchema = EnsureReceiveInboxSchema(targetWb, report)
        Case UCase$(TABLE_INBOX_SHIP)
            EnsureInboxTargetSchema = EnsureShipInboxSchema(targetWb, report)
        Case UCase$(TABLE_INBOX_PROD)
            EnsureInboxTargetSchema = EnsureProductionInboxSchema(targetWb, report)
        Case Else
            report = "Unknown inbox table: " & tableName
    End Select
End Function

Private Function EnsureInboxSchemaCore(ByVal targetWb As Workbook, _
                                       ByRef report As String, _
                                       ByVal sheetName As String, _
                                       ByVal tableName As String, _
                                       ByVal defaultEventType As String) As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim headers As Variant
    Dim startCell As Range
    Dim dataRange As Range
    Dim i As Long

    If targetWb Is Nothing Then
        Set wb = ResolveSingleInboxWorkbook(tableName)
    Else
        Set wb = targetWb
    End If
    If wb Is Nothing Then
        report = "Inbox workbook not found."
        Exit Function
    End If

    headers = Array("EventID", "ParentEventId", "UndoOfEventId", "EventType", "CreatedAtUTC", "WarehouseId", "StationId", _
                    "UserId", "SKU", "Qty", "Location", "Note", "PayloadJson", "Status", "RetryCount", "ErrorCode", _
                    "ErrorMessage", "FailedAtUTC")

    Set ws = EnsureWorksheetProcessor(wb, sheetName)
    SetSheetProtectionProcessor ws, False
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCellProcessor(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i

        Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = tableName
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumnProcessor lo, CStr(headers(i))
    Next i

    EnsureTableHasRowProcessor lo
    EnsureInboxDefaultEventType lo, defaultEventType
    report = "OK"
    EnsureInboxSchemaCore = True
    SetSheetProtectionProcessor ws, True
    Exit Function

FailEnsure:
    On Error Resume Next
    If Not ws Is Nothing Then SetSheetProtectionProcessor ws, True
    On Error GoTo 0
    report = "EnsureInboxSchema failed: " & Err.Description
End Function

Private Sub EnsureInboxDefaultEventType(ByVal lo As ListObject, ByVal defaultEventType As String)
    Dim i As Long
    Dim currentValue As String
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To lo.ListRows.Count
        currentValue = SafeTrimProcessor(GetCellByColumnProcessor(lo, i, "EventType"))
        If currentValue = "" Then SetCellByColumnProcessor lo, i, "EventType", defaultEventType
    Next i
End Sub

Private Function EnsurePhase2Context(ByVal warehouseId As String, ByRef report As String) As Boolean
    If Not modConfig.LoadConfig(warehouseId, "") Then
        report = "Config load failed: " & modConfig.Validate()
        Exit Function
    End If

    If Not modAuth.LoadAuth(modConfig.GetString("WarehouseId", warehouseId)) Then
        report = "Auth load failed: " & modAuth.ValidateAuth()
        Exit Function
    End If

    EnsurePhase2Context = True
End Function

Private Function ResolveInboxTargets() As Collection
    Dim wb As Workbook
    Dim seen As Object

    Set ResolveInboxTargets = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    For Each wb In Application.Workbooks
        AddInboxTarget ResolveInboxTargets, seen, wb, TABLE_INBOX_RECEIVE, SHEET_INBOX_RECEIVE, EVENT_TYPE_RECEIVE, _
                       IsReceiveInboxWorkbookName(wb.Name) Or WorkbookHasListObjectProcessor(wb, TABLE_INBOX_RECEIVE)
        AddInboxTarget ResolveInboxTargets, seen, wb, TABLE_INBOX_SHIP, SHEET_INBOX_SHIP, EVENT_TYPE_SHIP, _
                       IsShipInboxWorkbookName(wb.Name) Or WorkbookHasListObjectProcessor(wb, TABLE_INBOX_SHIP)
        AddInboxTarget ResolveInboxTargets, seen, wb, TABLE_INBOX_PROD, SHEET_INBOX_PROD, EVENT_TYPE_PROD_CONSUME, _
                       IsProductionInboxWorkbookName(wb.Name) Or WorkbookHasListObjectProcessor(wb, TABLE_INBOX_PROD)
    Next wb
End Function

Private Sub AddInboxTarget(ByVal targets As Collection, _
                           ByVal seen As Object, _
                           ByVal wb As Workbook, _
                           ByVal tableName As String, _
                           ByVal sheetName As String, _
                           ByVal defaultEventType As String, _
                           ByVal shouldAdd As Boolean)
    Dim target As Object
    Dim key As String

    If Not shouldAdd Then Exit Sub
    key = wb.Name & "|" & tableName
    If seen.Exists(key) Then Exit Sub

    Set target = CreateObject("Scripting.Dictionary")
    target.CompareMode = vbTextCompare
    target.Add "Workbook", wb
    target.Add "TableName", tableName
    target.Add "SheetName", sheetName
    target.Add "DefaultEventType", defaultEventType
    targets.Add target
    seen.Add key, True
End Sub

Private Function ResolveSingleInboxWorkbook(ByVal tableName As String) As Workbook
    Dim targets As Collection
    Dim target As Variant

    Set targets = ResolveInboxTargets()
    For Each target In targets
        If StrComp(CStr(target("TableName")), tableName, vbTextCompare) = 0 Then
            Set ResolveSingleInboxWorkbook = target("Workbook")
            Exit Function
        End If
    Next target
End Function

Private Function BuildInboxEvent(ByVal lo As ListObject, _
                                 ByVal rowIndex As Long, _
                                 ByVal workbookName As String, _
                                 ByVal tableName As String, _
                                 ByVal defaultEventType As String) As Object
    Dim evt As Object
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Set evt = CreateObject("Scripting.Dictionary")
    evt.CompareMode = vbTextCompare
    evt("EventID") = GetCellByColumnProcessor(lo, rowIndex, "EventID")
    evt("ParentEventId") = GetCellByColumnProcessor(lo, rowIndex, "ParentEventId")
    evt("UndoOfEventId") = GetCellByColumnProcessor(lo, rowIndex, "UndoOfEventId")
    evt("EventType") = GetInboxEventType(lo, rowIndex, defaultEventType)
    evt("CreatedAtUTC") = GetCellByColumnProcessor(lo, rowIndex, "CreatedAtUTC")
    evt("WarehouseId") = GetCellByColumnProcessor(lo, rowIndex, "WarehouseId")
    evt("StationId") = GetCellByColumnProcessor(lo, rowIndex, "StationId")
    evt("UserId") = GetCellByColumnProcessor(lo, rowIndex, "UserId")
    evt("SKU") = GetCellByColumnProcessor(lo, rowIndex, "SKU")
    evt("Qty") = GetCellByColumnProcessor(lo, rowIndex, "Qty")
    evt("Location") = GetCellByColumnProcessor(lo, rowIndex, "Location")
    evt("Note") = GetCellByColumnProcessor(lo, rowIndex, "Note")
    evt("PayloadJson") = GetCellByColumnProcessor(lo, rowIndex, "PayloadJson")
    evt("SourceInbox") = workbookName & ":" & tableName
    Set BuildInboxEvent = evt
End Function

Private Function GetInboxEventType(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal defaultEventType As String) As String
    GetInboxEventType = SafeTrimProcessor(GetCellByColumnProcessor(lo, rowIndex, "EventType"))
    If GetInboxEventType = "" Then GetInboxEventType = defaultEventType
End Function

Private Function IsProcessableInboxRow(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal warehouseId As String) As Boolean
    Dim statusVal As String
    Dim eventId As String
    Dim rowWarehouse As String

    eventId = SafeTrimProcessor(GetCellByColumnProcessor(lo, rowIndex, "EventID"))
    If eventId = "" Then Exit Function

    statusVal = UCase$(SafeTrimProcessor(GetCellByColumnProcessor(lo, rowIndex, "Status")))
    If statusVal <> "" And statusVal <> INBOX_STATUS_NEW Then Exit Function

    rowWarehouse = SafeTrimProcessor(GetCellByColumnProcessor(lo, rowIndex, "WarehouseId"))
    If warehouseId <> "" And rowWarehouse <> "" Then
        If StrComp(warehouseId, rowWarehouse, vbTextCompare) <> 0 Then Exit Function
    End If

    IsProcessableInboxRow = True
End Function

Private Function CapabilityForEventType(ByVal eventType As String) As String
    Select Case UCase$(SafeTrimProcessor(eventType))
        Case EVENT_TYPE_RECEIVE
            CapabilityForEventType = "RECEIVE_POST"
        Case EVENT_TYPE_SHIP
            CapabilityForEventType = "SHIP_POST"
        Case EVENT_TYPE_PROD_CONSUME, EVENT_TYPE_PROD_COMPLETE
            CapabilityForEventType = "PROD_POST"
    End Select
End Function

Private Sub UpdateInboxRowStatus(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal newStatus As String, _
                                 Optional ByVal errorCode As String = "", Optional ByVal errorMessage As String = "")
    Dim retryCount As Long
    If lo Is Nothing Then Exit Sub

    SetSheetProtectionProcessor lo.Parent, False

    SetCellByColumnProcessor lo, rowIndex, "Status", newStatus

    Select Case UCase$(newStatus)
        Case INBOX_STATUS_POISON
            retryCount = 0
            If IsNumeric(GetCellByColumnProcessor(lo, rowIndex, "RetryCount")) Then
                retryCount = CLng(GetCellByColumnProcessor(lo, rowIndex, "RetryCount"))
            End If
            SetCellByColumnProcessor lo, rowIndex, "RetryCount", retryCount + 1
            SetCellByColumnProcessor lo, rowIndex, "ErrorCode", errorCode
            SetCellByColumnProcessor lo, rowIndex, "ErrorMessage", errorMessage
            SetCellByColumnProcessor lo, rowIndex, "FailedAtUTC", Now
        Case Else
            SetCellByColumnProcessor lo, rowIndex, "ErrorCode", vbNullString
            SetCellByColumnProcessor lo, rowIndex, "ErrorMessage", vbNullString
            SetCellByColumnProcessor lo, rowIndex, "FailedAtUTC", vbNullString
    End Select

    SetSheetProtectionProcessor lo.Parent, True
End Sub

Private Function GetDictionaryString(ByVal d As Object, ByVal key As String) As String
    On Error Resume Next
    GetDictionaryString = SafeTrimProcessor(d(key))
    On Error GoTo 0
End Function

Private Function IsReceiveInboxWorkbookName(ByVal wbName As String) As Boolean
    Dim n As String
    n = LCase$(wbName)
    IsReceiveInboxWorkbookName = (n Like "invsys.inbox.receiving.*.xlsb") Or _
                                 (n Like "invsys.inbox.receiving.*.xlsx") Or _
                                 (n Like "invsys.inbox.receiving.*.xlsm")
End Function

Private Function IsShipInboxWorkbookName(ByVal wbName As String) As Boolean
    Dim n As String
    n = LCase$(wbName)
    IsShipInboxWorkbookName = (n Like "invsys.inbox.shipping.*.xlsb") Or _
                              (n Like "invsys.inbox.shipping.*.xlsx") Or _
                              (n Like "invsys.inbox.shipping.*.xlsm")
End Function

Private Function IsProductionInboxWorkbookName(ByVal wbName As String) As Boolean
    Dim n As String
    n = LCase$(wbName)
    IsProductionInboxWorkbookName = (n Like "invsys.inbox.production.*.xlsb") Or _
                                    (n Like "invsys.inbox.production.*.xlsx") Or _
                                    (n Like "invsys.inbox.production.*.xlsm")
End Function

Private Sub EnsureListColumnProcessor(ByVal lo As ListObject, ByVal columnName As String)
    If GetColumnIndexProcessor(lo, columnName) > 0 Then Exit Sub
    lo.ListColumns.Add lo.ListColumns.Count + 1
    lo.ListColumns(lo.ListColumns.Count).Name = columnName
End Sub

Private Sub EnsureTableHasRowProcessor(ByVal lo As ListObject)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then lo.ListRows.Add
End Sub

Private Function EnsureWorksheetProcessor(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureWorksheetProcessor = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureWorksheetProcessor Is Nothing Then
        Set EnsureWorksheetProcessor = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureWorksheetProcessor.Name = sheetName
    End If
End Function

Private Function GetNextTableStartCellProcessor(ByVal ws As Worksheet) As Range
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        Set GetNextTableStartCellProcessor = ws.Range("A1")
    Else
        Set GetNextTableStartCellProcessor = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End If
End Function

Private Function WorkbookHasListObjectProcessor(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    WorkbookHasListObjectProcessor = Not (FindListObjectByNameProcessor(wb, tableName) Is Nothing)
End Function

Private Function FindListObjectByNameProcessor(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameProcessor = ws.ListObjects(tableName)
        If Not FindListObjectByNameProcessor Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function GetCellByColumnProcessor(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndexProcessor(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumnProcessor = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

Private Sub SetCellByColumnProcessor(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim idx As Long
    idx = GetColumnIndexProcessor(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function GetColumnIndexProcessor(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long
    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexProcessor = i
            Exit Function
        End If
    Next i
End Function

Private Function SafeTrimProcessor(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimProcessor = Trim$(CStr(valueIn))
End Function

Private Sub SetSheetProtectionProcessor(ByVal ws As Worksheet, ByVal protectAfter As Boolean)
    If ws Is Nothing Then Exit Sub
    If protectAfter Then
        On Error Resume Next
        ws.Protect UserInterfaceOnly:=True
        On Error GoTo 0
    Else
        If Not ws.ProtectContents Then Exit Sub
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
        If ws.ProtectContents Then
            Err.Raise vbObjectError + 2401, "modProcessor.SetSheetProtectionProcessor", _
                      "Worksheet '" & ws.Name & "' is protected and could not be unprotected. " & _
                      "Excel automation cannot update inbox tables while the sheet remains protected."
        End If
    End If
End Sub
