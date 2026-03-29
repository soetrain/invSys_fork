Attribute VB_Name = "modRoleEventWriter"
Option Explicit

Private Const SHEET_INBOX_RECEIVE As String = "InboxReceive"
Private Const SHEET_INBOX_SHIP As String = "InboxShip"
Private Const SHEET_INBOX_PROD As String = "InboxProd"

Private Const TABLE_INBOX_RECEIVE As String = "tblInboxReceive"
Private Const TABLE_INBOX_SHIP As String = "tblInboxShip"
Private Const TABLE_INBOX_PROD As String = "tblInboxProd"
Private Const ROLE_EVENT_TYPE_RECEIVE As String = "RECEIVE"
Private Const ROLE_EVENT_TYPE_SHIP As String = "SHIP"
Private Const ROLE_EVENT_TYPE_PROD_CONSUME As String = "PROD_CONSUME"
Private Const ROLE_EVENT_TYPE_PROD_COMPLETE As String = "PROD_COMPLETE"

Public Function ResolveCurrentUserId() As String
    ResolveCurrentUserId = Trim$(Environ$("USERNAME"))
    If ResolveCurrentUserId = "" Then ResolveCurrentUserId = Trim$(Application.UserName)
End Function

Public Function OpenInboxWorkbook(ByVal eventType As String, _
                                  Optional ByVal warehouseId As String = "", _
                                  Optional ByVal stationId As String = "", _
                                  Optional ByRef errorMessage As String = "") As Workbook
    Dim resolvedWh As String
    Dim resolvedSt As String

    If Not EnsureContextResolved(resolvedWh, resolvedSt, warehouseId, stationId, errorMessage) Then Exit Function
    Set OpenInboxWorkbook = ResolveInboxWorkbookForEventType(eventType, resolvedWh, resolvedSt, errorMessage)
End Function

Public Function ResolveInboxWorkbookPath(ByVal eventType As String, _
                                         Optional ByVal warehouseId As String = "", _
                                         Optional ByVal stationId As String = "", _
                                         Optional ByRef errorMessage As String = "") As String
    Dim resolvedWh As String
    Dim resolvedSt As String

    If Not EnsureContextResolved(resolvedWh, resolvedSt, warehouseId, stationId, errorMessage) Then Exit Function
    ResolveInboxWorkbookPath = ResolveInboxWorkbookPathResolvedRole(eventType, resolvedWh, resolvedSt, errorMessage)
End Function

Public Function QueueReceiveEvent(Optional ByVal warehouseId As String = "", _
                                  Optional ByVal stationId As String = "", _
                                  Optional ByVal userId As String = "", _
                                  Optional ByVal sku As String = "", _
                                  Optional ByVal qty As Double = 0, _
                                  Optional ByVal location As String = "", _
                                  Optional ByVal noteVal As String = "", _
                                  Optional ByVal parentEventId As String = "", _
                                  Optional ByVal undoOfEventId As String = "", _
                                  Optional ByVal createdAtUtc As Date = 0, _
                                  Optional ByVal targetInboxWb As Workbook = Nothing, _
                                  Optional ByRef eventIdOut As String = "", _
                                  Optional ByRef errorMessage As String = "") As Boolean
    QueueReceiveEvent = QueueEventCore(ROLE_EVENT_TYPE_RECEIVE, warehouseId, stationId, userId, sku, qty, location, noteVal, "", parentEventId, undoOfEventId, createdAtUtc, targetInboxWb, eventIdOut, errorMessage)
End Function

Public Function QueueReceiveEventCurrent(Optional ByVal userId As String = "", _
                                         Optional ByVal sku As String = "", _
                                         Optional ByVal qty As Double = 0, _
                                         Optional ByVal location As String = "", _
                                         Optional ByVal noteVal As String = "", _
                                         Optional ByRef eventIdOut As String = "", _
                                         Optional ByRef errorMessage As String = "") As Boolean
    Dim targetInboxWb As Workbook

    QueueReceiveEventCurrent = QueueReceiveEvent("", "", userId, sku, qty, location, noteVal, "", "", 0, targetInboxWb, eventIdOut, errorMessage)
End Function

Public Function QueuePayloadEvent(ByVal eventType As String, _
                                  Optional ByVal warehouseId As String = "", _
                                  Optional ByVal stationId As String = "", _
                                  Optional ByVal userId As String = "", _
                                  Optional ByVal payloadJson As String = "", _
                                  Optional ByVal noteVal As String = "", _
                                  Optional ByVal parentEventId As String = "", _
                                  Optional ByVal undoOfEventId As String = "", _
                                  Optional ByVal createdAtUtc As Date = 0, _
                                  Optional ByVal targetInboxWb As Workbook = Nothing, _
                                  Optional ByRef eventIdOut As String = "", _
                                  Optional ByRef errorMessage As String = "") As Boolean
    QueuePayloadEvent = QueueEventCore(eventType, warehouseId, stationId, userId, "", 0, "", noteVal, payloadJson, parentEventId, undoOfEventId, createdAtUtc, targetInboxWb, eventIdOut, errorMessage)
End Function

Public Function QueuePayloadEventCurrent(ByVal eventType As String, _
                                         Optional ByVal userId As String = "", _
                                         Optional ByVal payloadJson As String = "", _
                                         Optional ByVal noteVal As String = "", _
                                         Optional ByRef eventIdOut As String = "", _
                                         Optional ByRef errorMessage As String = "") As Boolean
    Dim targetInboxWb As Workbook

    QueuePayloadEventCurrent = QueuePayloadEvent(eventType, "", "", userId, payloadJson, noteVal, "", "", 0, targetInboxWb, eventIdOut, errorMessage)
End Function

Public Function BuildPayloadJson(ParamArray items() As Variant) As String
    Dim i As Long
    Dim item As Object

    BuildPayloadJson = "["
    For i = LBound(items) To UBound(items)
        Set item = items(i)
        If i > LBound(items) Then BuildPayloadJson = BuildPayloadJson & ","
        BuildPayloadJson = BuildPayloadJson & DictionaryToJson(item)
    Next i
    BuildPayloadJson = BuildPayloadJson & "]"
End Function

Public Function BuildPayloadJsonFromCollection(ByVal items As Collection) As String
    Dim i As Long

    BuildPayloadJsonFromCollection = "["
    If items Is Nothing Then
        BuildPayloadJsonFromCollection = BuildPayloadJsonFromCollection & "]"
        Exit Function
    End If

    For i = 1 To items.Count
        If i > 1 Then BuildPayloadJsonFromCollection = BuildPayloadJsonFromCollection & ","
        BuildPayloadJsonFromCollection = BuildPayloadJsonFromCollection & DictionaryToJson(items(i))
    Next i
    BuildPayloadJsonFromCollection = BuildPayloadJsonFromCollection & "]"
End Function

Public Function CreatePayloadItem(ByVal rowVal As Long, _
                                  ByVal sku As String, _
                                  ByVal qty As Double, _
                                  Optional ByVal location As String = "", _
                                  Optional ByVal noteVal As String = "", _
                                  Optional ByVal ioType As String = "") As Object
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("Row") = rowVal
    item("SKU") = sku
    item("Qty") = qty
    item("Location") = location
    item("Note") = noteVal
    If ioType <> "" Then item("IoType") = ioType
    Set CreatePayloadItem = item
End Function

Private Function QueueEventCore(ByVal eventType As String, _
                                ByVal warehouseId As String, _
                                ByVal stationId As String, _
                                ByVal userId As String, _
                                ByVal sku As String, _
                                ByVal qty As Double, _
                                ByVal location As String, _
                                ByVal noteVal As String, _
                                ByVal payloadJson As String, _
                                ByVal parentEventId As String, _
                                ByVal undoOfEventId As String, _
                                ByVal createdAtUtc As Date, _
                                ByVal targetInboxWb As Workbook, _
                                ByRef eventIdOut As String, _
                                ByRef errorMessage As String) As Boolean
    On Error GoTo FailQueue

    Dim resolvedWh As String
    Dim resolvedSt As String
    Dim resolvedUser As String
    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim sheetWasProtected As Boolean
    Dim ws As Worksheet
    Dim report As String
    Dim capability As String
    Dim openPaths As Object
    Dim openedTransient As Boolean

    If Not EnsureContextResolved(resolvedWh, resolvedSt, warehouseId, stationId, errorMessage) Then Exit Function

    resolvedUser = Trim$(userId)
    If resolvedUser = "" Then resolvedUser = ResolveCurrentUserId()
    If resolvedUser = "" Then
        errorMessage = "Unable to resolve current user identity."
        Exit Function
    End If

    capability = CapabilityForEventTypeRole(eventType)
    If capability = "" Then
        errorMessage = "Unsupported event type '" & eventType & "'."
        Exit Function
    End If

    If Not modAuth.LoadAuth(resolvedWh) Then
        errorMessage = "Auth load failed: " & modAuth.ValidateAuth()
        Exit Function
    End If
    If Not modAuth.CanPerform(capability, resolvedUser, resolvedWh, resolvedSt, "ROLE_UI", eventIdOut) Then
        errorMessage = "Current user lacks " & capability & " capability."
        Exit Function
    End If

    If eventIdOut = "" Then eventIdOut = CreateEventIdRole()
    If createdAtUtc = 0 Then createdAtUtc = Now

    If targetInboxWb Is Nothing Then
        Set openPaths = CaptureOpenWorkbookPathsRole()
        Set wbInbox = ResolveInboxWorkbookForEventType(eventType, resolvedWh, resolvedSt, errorMessage)
    Else
        Set wbInbox = targetInboxWb
    End If
    If wbInbox Is Nothing Then Exit Function
    openedTransient = (targetInboxWb Is Nothing) And (Not WorkbookWasAlreadyOpenRole(openPaths, wbInbox))
    If openedTransient Then HideWorkbookWindowsRole wbInbox

    Select Case UCase$(Trim$(eventType))
        Case ROLE_EVENT_TYPE_RECEIVE
            If Not modProcessor.EnsureReceiveInboxSchema(wbInbox, report) Then
                errorMessage = report
                GoTo CleanExit
            End If
            Set lo = FindListObjectByNameRole(wbInbox, TABLE_INBOX_RECEIVE)
        Case ROLE_EVENT_TYPE_SHIP
            If Not modProcessor.EnsureShipInboxSchema(wbInbox, report) Then
                errorMessage = report
                GoTo CleanExit
            End If
            Set lo = FindListObjectByNameRole(wbInbox, TABLE_INBOX_SHIP)
        Case ROLE_EVENT_TYPE_PROD_CONSUME, ROLE_EVENT_TYPE_PROD_COMPLETE
            If Not modProcessor.EnsureProductionInboxSchema(wbInbox, report) Then
                errorMessage = report
                GoTo CleanExit
            End If
            Set lo = FindListObjectByNameRole(wbInbox, TABLE_INBOX_PROD)
    End Select
    If lo Is Nothing Then
        errorMessage = "Inbox table not found for event type '" & eventType & "'."
        GoTo CleanExit
    End If

    Set ws = lo.Parent
    sheetWasProtected = ws.ProtectContents
    EnsureWorksheetEditableRole ws, lo.Name

    rowIndex = lo.ListRows.Add.Index
    SetTableRowValueRole lo, rowIndex, "EventID", eventIdOut
    SetTableRowValueRole lo, rowIndex, "ParentEventId", parentEventId
    SetTableRowValueRole lo, rowIndex, "UndoOfEventId", undoOfEventId
    SetTableRowValueRole lo, rowIndex, "EventType", UCase$(Trim$(eventType))
    SetTableRowValueRole lo, rowIndex, "CreatedAtUTC", createdAtUtc
    SetTableRowValueRole lo, rowIndex, "WarehouseId", resolvedWh
    SetTableRowValueRole lo, rowIndex, "StationId", resolvedSt
    SetTableRowValueRole lo, rowIndex, "UserId", resolvedUser
    SetTableRowValueRole lo, rowIndex, "SKU", sku
    If qty <> 0 Then SetTableRowValueRole lo, rowIndex, "Qty", qty
    SetTableRowValueRole lo, rowIndex, "Location", location
    SetTableRowValueRole lo, rowIndex, "Note", noteVal
    SetTableRowValueRole lo, rowIndex, "PayloadJson", payloadJson
    SetTableRowValueRole lo, rowIndex, "Status", "NEW"
    SetTableRowValueRole lo, rowIndex, "RetryCount", 0
    SetTableRowValueRole lo, rowIndex, "ErrorCode", ""
    SetTableRowValueRole lo, rowIndex, "ErrorMessage", ""
    SetTableRowValueRole lo, rowIndex, "FailedAtUTC", ""

    SaveWorkbookRole wbInbox

    QueueEventCore = True
CleanExit:
    On Error Resume Next
    If Not ws Is Nothing Then
        If sheetWasProtected Then RestoreWorksheetProtectionRole ws
    End If
    If openedTransient Then CloseTransientRoleWorkbook wbInbox
    On Error GoTo 0
    Exit Function

FailQueue:
    errorMessage = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then
        If sheetWasProtected Then RestoreWorksheetProtectionRole ws
    End If
    If openedTransient Then CloseTransientRoleWorkbook wbInbox
End Function

Private Function EnsureContextResolved(ByRef resolvedWh As String, _
                                       ByRef resolvedSt As String, _
                                       ByVal warehouseId As String, _
                                       ByVal stationId As String, _
                                       ByRef errorMessage As String) As Boolean
    resolvedWh = Trim$(warehouseId)
    resolvedSt = Trim$(stationId)

    If Not modConfig.LoadConfig(resolvedWh, resolvedSt) Then
        errorMessage = "Config load failed: " & modConfig.Validate()
        Exit Function
    End If

    If resolvedWh = "" Then resolvedWh = modConfig.GetWarehouseId()
    If resolvedSt = "" Then resolvedSt = modConfig.GetStationId()
    If resolvedWh = "" Or resolvedSt = "" Then
        errorMessage = "WarehouseId and StationId are required."
        Exit Function
    End If

    EnsureContextResolved = True
End Function

Private Function ResolveInboxWorkbookForEventType(ByVal eventType As String, _
                                                  ByVal warehouseId As String, _
                                                  ByVal stationId As String, _
                                                  ByRef errorMessage As String) As Workbook
    On Error GoTo FailOpen

    Dim wb As Workbook
    Dim expectedName As String
    Dim fullPath As String
    Dim prevEvents As Boolean
    Dim eventsSuppressed As Boolean
    Dim prevAlerts As Boolean
    Dim alertsSuppressed As Boolean
    Dim prevScreenUpdating As Boolean

    prevScreenUpdating = Application.ScreenUpdating
    expectedName = InboxWorkbookNameRole(eventType, stationId)
    fullPath = ResolveInboxWorkbookPathResolvedRole(eventType, warehouseId, stationId, errorMessage)
    If fullPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 _
           Or StrComp(wb.Name, expectedName, vbTextCompare) = 0 Then
            Set ResolveInboxWorkbookForEventType = wb
            Exit Function
        End If
    Next wb

    If Len(Dir$(fullPath, vbNormal)) > 0 Then
        prevAlerts = Application.DisplayAlerts
        Application.DisplayAlerts = False
        alertsSuppressed = True
        prevScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False
        Set ResolveInboxWorkbookForEventType = Application.Workbooks.Open(fullPath)
        If Not ResolveInboxWorkbookForEventType Is Nothing Then HideWorkbookWindowsRole ResolveInboxWorkbookForEventType
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevAlerts
        alertsSuppressed = False
    Else
        prevEvents = Application.EnableEvents
        Application.EnableEvents = False
        eventsSuppressed = True
        Set wb = Application.Workbooks.Add(xlWBATWorksheet)
        SaveWorkbookAsXlsbRole wb, fullPath
        Application.EnableEvents = prevEvents
        eventsSuppressed = False
        HideWorkbookWindowsRole wb
        Set ResolveInboxWorkbookForEventType = wb
    End If
    Exit Function

FailOpen:
    On Error Resume Next
    If eventsSuppressed Then Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreenUpdating
    If alertsSuppressed Then Application.DisplayAlerts = prevAlerts
    On Error GoTo 0
    errorMessage = "Inbox workbook open/create failed: " & Err.Description
End Function

Private Function ResolveInboxWorkbookPathResolvedRole(ByVal eventType As String, _
                                                      ByVal warehouseId As String, _
                                                      ByVal stationId As String, _
                                                      ByRef errorMessage As String) As String
    Dim expectedName As String
    Dim targetDir As String

    expectedName = InboxWorkbookNameRole(eventType, stationId)
    If expectedName = "" Then
        errorMessage = "Unsupported event type '" & eventType & "'."
        Exit Function
    End If

    targetDir = ResolveInboxDirectoryRole(warehouseId, stationId)
    If targetDir = "" Then
        errorMessage = "Unable to resolve inbox directory."
        Exit Function
    End If

    EnsureFolderExistsRole targetDir
    ResolveInboxWorkbookPathResolvedRole = CombinePathRole(targetDir, expectedName)
End Function

Private Function ResolveInboxDirectoryRole(ByVal warehouseId As String, ByVal stationId As String) As String
    Dim rawPath As String

    rawPath = Trim$(modConfig.GetString("PathInboxRoot", ""))
    rawPath = ExpandConfigPathRole(rawPath, warehouseId, stationId)
    If rawPath = "" Then
        rawPath = Trim$(modConfig.GetString("PathDataRoot", ""))
        rawPath = ExpandConfigPathRole(rawPath, warehouseId, stationId)
    End If
    If rawPath = "" Then rawPath = ThisWorkbook.Path
    If rawPath = "" Then rawPath = Environ$("TEMP")
    ResolveInboxDirectoryRole = rawPath
End Function

Private Function ExpandConfigPathRole(ByVal rawPath As String, ByVal warehouseId As String, ByVal stationId As String) As String
    ExpandConfigPathRole = Trim$(rawPath)
    If ExpandConfigPathRole = "" Then Exit Function

    ExpandConfigPathRole = Replace$(ExpandConfigPathRole, "{WarehouseId}", warehouseId)
    ExpandConfigPathRole = Replace$(ExpandConfigPathRole, "{StationId}", stationId)
    ExpandConfigPathRole = Replace$(ExpandConfigPathRole, "/", "\")
    Do While Right$(ExpandConfigPathRole, 1) = "\"
        ExpandConfigPathRole = Left$(ExpandConfigPathRole, Len(ExpandConfigPathRole) - 1)
    Loop
End Function

Private Function InboxWorkbookNameRole(ByVal eventType As String, ByVal stationId As String) As String
    Select Case UCase$(Trim$(eventType))
        Case ROLE_EVENT_TYPE_RECEIVE
            InboxWorkbookNameRole = "invSys.Inbox.Receiving." & stationId & ".xlsb"
        Case ROLE_EVENT_TYPE_SHIP
            InboxWorkbookNameRole = "invSys.Inbox.Shipping." & stationId & ".xlsb"
        Case ROLE_EVENT_TYPE_PROD_CONSUME, ROLE_EVENT_TYPE_PROD_COMPLETE
            InboxWorkbookNameRole = "invSys.Inbox.Production." & stationId & ".xlsb"
    End Select
End Function

Private Function CapabilityForEventTypeRole(ByVal eventType As String) As String
    Select Case UCase$(Trim$(eventType))
        Case ROLE_EVENT_TYPE_RECEIVE
            CapabilityForEventTypeRole = "RECEIVE_POST"
        Case ROLE_EVENT_TYPE_SHIP
            CapabilityForEventTypeRole = "SHIP_POST"
        Case ROLE_EVENT_TYPE_PROD_CONSUME, ROLE_EVENT_TYPE_PROD_COMPLETE
            CapabilityForEventTypeRole = "PROD_POST"
    End Select
End Function

Private Sub EnsureFolderExistsRole(ByVal folderPath As String)
    Dim parts As Variant
    Dim currentPath As String
    Dim i As Long

    If folderPath = "" Then Exit Sub
    parts = Split(folderPath, "\")
    If UBound(parts) < 0 Then Exit Sub

    currentPath = parts(0)
    If Right$(currentPath, 1) = ":" Then currentPath = currentPath & "\"

    For i = 1 To UBound(parts)
        If parts(i) <> "" Then
            currentPath = CombinePathRole(currentPath, parts(i))
            If Len(Dir$(currentPath, vbDirectory)) = 0 Then MkDir currentPath
        End If
    Next i
End Sub

Private Function CombinePathRole(ByVal basePath As String, ByVal childName As String) As String
    If basePath = "" Then
        CombinePathRole = childName
    ElseIf Right$(basePath, 1) = "\" Then
        CombinePathRole = basePath & childName
    Else
        CombinePathRole = basePath & "\" & childName
    End If
End Function

Private Sub SaveWorkbookAsXlsbRole(ByVal wb As Workbook, ByVal fullPath As String)
    If wb Is Nothing Then Exit Sub
    wb.SaveAs fullPath, 50
End Sub

Private Sub SaveWorkbookRole(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If wb.Path = "" Then Exit Sub
    wb.Save
End Sub

Private Function CaptureOpenWorkbookPathsRole() As Object
    Dim wb As Workbook
    Dim paths As Object

    Set paths = CreateObject("Scripting.Dictionary")
    paths.CompareMode = vbTextCompare

    For Each wb In Application.Workbooks
        If Trim$(wb.FullName) <> "" Then paths(Trim$(wb.FullName)) = True
    Next wb

    Set CaptureOpenWorkbookPathsRole = paths
End Function

Private Function WorkbookWasAlreadyOpenRole(ByVal openPaths As Object, ByVal wb As Workbook) As Boolean
    If openPaths Is Nothing Then Exit Function
    If wb Is Nothing Then Exit Function
    If Trim$(wb.FullName) = "" Then Exit Function
    WorkbookWasAlreadyOpenRole = openPaths.Exists(Trim$(wb.FullName))
End Function

Private Sub HideWorkbookWindowsRole(ByVal wb As Workbook)
    Dim i As Long

    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To wb.Windows.Count
        wb.Windows(i).Visible = False
    Next i
    modUiQuiet.ReactivateQuietOwner
    On Error GoTo 0
End Sub

Private Sub CloseTransientRoleWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    HideWorkbookWindowsRole wb
    If Not wb.ReadOnly Then
        If wb.Saved = False Then wb.Save
    End If
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub EnsureWorksheetEditableRole(ByVal ws As Worksheet, ByVal context As String)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    If ws.ProtectContents Then
        Err.Raise vbObjectError + 7801, "modRoleEventWriter.EnsureWorksheetEditableRole", _
                  "Worksheet '" & ws.Name & "' is protected and could not be unprotected before writing to " & context & "."
    End If
End Sub

Private Sub RestoreWorksheetProtectionRole(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Function FindListObjectByNameRole(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindListObjectByNameRole = ws.ListObjects(tableName)
        If Not FindListObjectByNameRole Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Sub SetTableRowValueRole(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueIn As Variant)
    Dim idx As Long

    idx = GetColumnIndexRole(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueIn
End Sub

Private Function GetColumnIndexRole(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexRole = i
            Exit Function
        End If
    Next i
End Function

Private Function CreateEventIdRole() As String
    On Error Resume Next
    CreateEventIdRole = CreateObject("Scriptlet.TypeLib").GUID
    On Error GoTo 0
    If CreateEventIdRole = "" Then
        CreateEventIdRole = CreateGuidFallbackRole()
    End If
    CreateEventIdRole = Replace$(CreateEventIdRole, Chr$(0), "")
    CreateEventIdRole = Replace$(CreateEventIdRole, "{", "")
    CreateEventIdRole = Replace$(CreateEventIdRole, "}", "")
End Function

Private Function CreateGuidFallbackRole() As String
    Dim i As Long
    Dim token As String
    Dim chars As String

    chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
    Randomize
    For i = 1 To 32
        token = token & Mid$(chars, Int((Len(chars) * Rnd) + 1), 1)
    Next i
    CreateGuidFallbackRole = Left$(token, 8) & "-" & Mid$(token, 9, 4) & "-" & Mid$(token, 13, 4) & "-" & Mid$(token, 17, 4) & "-" & Right$(token, 12)
End Function

Private Function DictionaryToJson(ByVal d As Object) As String
    Dim keys As Variant
    Dim i As Long
    Dim key As String

    DictionaryToJson = "{"
    keys = d.Keys
    For i = LBound(keys) To UBound(keys)
        key = CStr(keys(i))
        If i > LBound(keys) Then DictionaryToJson = DictionaryToJson & ","
        DictionaryToJson = DictionaryToJson & """" & EscapeJsonRole(key) & """:" & JsonValueRole(d(key))
    Next i
    DictionaryToJson = DictionaryToJson & "}"
End Function

Private Function JsonValueRole(ByVal valueIn As Variant) As String
    Select Case True
        Case IsObject(valueIn)
            JsonValueRole = "null"
        Case IsNull(valueIn), IsEmpty(valueIn)
            JsonValueRole = "null"
        Case VarType(valueIn) = vbBoolean
            JsonValueRole = IIf(CBool(valueIn), "true", "false")
        Case IsNumeric(valueIn)
            JsonValueRole = Replace$(CStr(valueIn), ",", "")
        Case Else
            JsonValueRole = """" & EscapeJsonRole(CStr(valueIn)) & """"
    End Select
End Function

Private Function EscapeJsonRole(ByVal textIn As String) As String
    EscapeJsonRole = textIn
    EscapeJsonRole = Replace$(EscapeJsonRole, "\", "\\")
    EscapeJsonRole = Replace$(EscapeJsonRole, Chr$(34), "\" & Chr$(34))
    EscapeJsonRole = Replace$(EscapeJsonRole, vbCrLf, "\n")
    EscapeJsonRole = Replace$(EscapeJsonRole, vbCr, "\n")
    EscapeJsonRole = Replace$(EscapeJsonRole, vbLf, "\n")
    EscapeJsonRole = Replace$(EscapeJsonRole, vbTab, "\t")
End Function
