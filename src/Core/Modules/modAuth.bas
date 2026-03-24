Attribute VB_Name = "modAuth"
Option Explicit

Public Const ERR_AUTH_DENIED As Long = vbObjectError + 7200

Private mUsers As Object
Private mAllowCaps As Collection
Private mDenyCaps As Collection
Private mValidationIssues As Collection
Private mAuthWorkbook As String
Private mIsLoaded As Boolean
Private mLoadedAt As Date
Private mCacheTtlSeconds As Long

Public Function LoadAuth(Optional ByVal whId As String = "") As Boolean
    On Error GoTo FailLoad

    Dim wb As Workbook
    Dim loUsers As ListObject
    Dim loCaps As ListObject

    InitializeState

    Set wb = ResolveAuthWorkbook(whId)
    If wb Is Nothing Then
        AddValidationIssue "ERROR", "AUTH_MISSING", "No open auth workbook found."
        GoTo FailSoft
    End If
    mAuthWorkbook = wb.Name

    If Not EnsureAuthSchema(wb, whId, modConfig.GetString("ProcessorServiceUserId", "svc_processor")) Then
        AddValidationIssue "ERROR", "AUTH_SELF_HEAL_FAILED", "Failed to create/repair auth tables."
        GoTo FailSoft
    End If

    Set loUsers = FindListObjectByName(wb, "tblUsers")
    Set loCaps = FindListObjectByName(wb, "tblCapabilities")
    If loUsers Is Nothing Then
        AddValidationIssue "ERROR", "AUTH_TABLE_MISSING", "tblUsers not found."
        GoTo FailSoft
    End If
    If loCaps Is Nothing Then
        AddValidationIssue "ERROR", "AUTH_TABLE_MISSING", "tblCapabilities not found."
        GoTo FailSoft
    End If

    LoadUsers loUsers
    LoadCapabilities loCaps

    mCacheTtlSeconds = modConfig.GetLong("AuthCacheTTLSeconds", 300)
    If mCacheTtlSeconds <= 0 Then mCacheTtlSeconds = 300
    mLoadedAt = Now

    If CountFatalIssues() > 0 Then GoTo FailSoft
    mIsLoaded = True
    LoadAuth = True
    Exit Function

FailSoft:
    mIsLoaded = False
    LoadAuth = False
    Exit Function

FailLoad:
    AddValidationIssue "ERROR", "AUTH_LOAD_EXCEPTION", Err.Description
    Resume FailSoft
End Function

Public Function EnsureAuthSchema(Optional ByVal targetWb As Workbook = Nothing, _
                                 Optional ByVal warehouseId As String = "", _
                                 Optional ByVal processorServiceUserId As String = "", _
                                 Optional ByRef report As String = "") As Boolean
    On Error GoTo FailEnsure

    Dim wb As Workbook

    If targetWb Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetWb
    End If

    If Not EnsureAuthTables(wb) Then GoTo FailSoft
    SeedAuthDefaults wb, warehouseId, processorServiceUserId
    FormatAuthSurface wb

    EnsureAuthSchema = True
    Exit Function

FailSoft:
    report = "EnsureAuthSchema failed."
    Exit Function

FailEnsure:
    report = "EnsureAuthSchema failed: " & Err.Description
End Function

Public Function ReloadAuth() As Boolean
    ReloadAuth = LoadAuth(modConfig.GetString("WarehouseId", ""))
End Function

Public Function IsAuthLoaded() As Boolean
    IsAuthLoaded = mIsLoaded
End Function

Public Function CanPerform(ByVal capability As String, _
                           ByVal userId As String, _
                           Optional ByVal warehouseId As String = "", _
                           Optional ByVal stationId As String = "", _
                           Optional ByVal source As String = "UI", _
                           Optional ByVal requestId As String = "") As Boolean
    Dim nowTs As Date
    Dim resolvedWh As String
    Dim resolvedSt As String
    Dim allowed As Boolean
    Dim denied As Boolean

    nowTs = Now

    If Not EnsureFreshCache() Then
        LogDecision requestId, userId, capability, warehouseId, stationId, "DENY", source, "auth-cache-unavailable"
        Exit Function
    End If

    resolvedWh = warehouseId
    If resolvedWh = "" Then resolvedWh = modConfig.GetString("WarehouseId", "")
    resolvedSt = stationId
    If resolvedSt = "" Then resolvedSt = modConfig.GetString("StationId", "")

    If Not IsUserActive(userId, nowTs) Then
        LogDecision requestId, userId, capability, resolvedWh, resolvedSt, "DENY", source, "user-inactive-or-missing"
        Exit Function
    End If

    allowed = HasCapabilityMatch(mAllowCaps, userId, capability, resolvedWh, resolvedSt, nowTs)
    denied = HasCapabilityMatch(mDenyCaps, userId, capability, resolvedWh, resolvedSt, nowTs)

    CanPerform = (allowed And Not denied)
    If CanPerform Then
        LogDecision requestId, userId, capability, resolvedWh, resolvedSt, "ALLOW", source, ""
    Else
        LogDecision requestId, userId, capability, resolvedWh, resolvedSt, "DENY", source, "capability-not-granted"
    End If
End Function

Public Function Require(ByVal capability As String, _
                        ByVal userId As String, _
                        Optional ByVal warehouseId As String = "", _
                        Optional ByVal stationId As String = "", _
                        Optional ByVal source As String = "UI", _
                        Optional ByVal requestId As String = "") As Boolean
    If Not CanPerform(capability, userId, warehouseId, stationId, source, requestId) Then
        Err.Raise ERR_AUTH_DENIED, "modAuth.Require", "Capability denied: " & capability
    End If
    Require = True
End Function

Public Function ValidateAuth() As String
    Dim itm As Variant
    Dim parts() As String
    Dim lineOut As String

    If mValidationIssues Is Nothing Or mValidationIssues.Count = 0 Then Exit Function

    For Each itm In mValidationIssues
        parts = Split(CStr(itm), "|")
        If UBound(parts) >= 2 Then
            lineOut = parts(0) & " " & parts(1) & ": " & parts(2)
        Else
            lineOut = CStr(itm)
        End If

        If Len(ValidateAuth) > 0 Then ValidateAuth = ValidateAuth & "; "
        ValidateAuth = ValidateAuth & lineOut
    Next itm
End Function

Private Sub InitializeState()
    Set mUsers = CreateObject("Scripting.Dictionary")
    mUsers.CompareMode = vbTextCompare
    Set mAllowCaps = New Collection
    Set mDenyCaps = New Collection
    Set mValidationIssues = New Collection
    mAuthWorkbook = vbNullString
    mIsLoaded = False
    mLoadedAt = 0
    mCacheTtlSeconds = 300
End Sub

Private Function EnsureFreshCache() As Boolean
    If mIsLoaded Then
        If DateDiff("s", mLoadedAt, Now) <= mCacheTtlSeconds Then
            EnsureFreshCache = True
            Exit Function
        End If
    End If

    EnsureFreshCache = ReloadAuth()
End Function

Private Sub LoadUsers(ByVal loUsers As ListObject)
    Dim i As Long
    Dim userId As String
    Dim userInfo As Object

    If loUsers.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To loUsers.ListRows.Count
        userId = SafeTrim(loUsers.DataBodyRange.Cells(i, GetColumnIndex(loUsers, "UserId")).Value)
        If userId <> "" Then
            Set userInfo = CreateObject("Scripting.Dictionary")
            userInfo.CompareMode = vbTextCompare
            userInfo("UserId") = userId
            userInfo("Status") = UCase$(SafeTrim(GetCellByColumn(loUsers, i, "Status")))
            userInfo("ValidFrom") = GetCellByColumn(loUsers, i, "ValidFrom")
            userInfo("ValidTo") = GetCellByColumn(loUsers, i, "ValidTo")
            Set mUsers(userId) = userInfo
        End If
    Next i
End Sub

Private Sub LoadCapabilities(ByVal loCaps As ListObject)
    Dim i As Long
    Dim cap As Object
    Dim status As String

    If loCaps.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To loCaps.ListRows.Count
        Set cap = CreateObject("Scripting.Dictionary")
        cap.CompareMode = vbTextCompare

        cap("UserId") = SafeTrim(GetCellByColumn(loCaps, i, "UserId"))
        cap("Capability") = UCase$(SafeTrim(GetCellByColumn(loCaps, i, "Capability")))
        cap("WarehouseId") = SafeTrim(GetCellByColumn(loCaps, i, "WarehouseId"))
        cap("StationId") = SafeTrim(GetCellByColumn(loCaps, i, "StationId"))
        cap("Status") = UCase$(SafeTrim(GetCellByColumn(loCaps, i, "Status")))
        cap("ValidFrom") = GetCellByColumn(loCaps, i, "ValidFrom")
        cap("ValidTo") = GetCellByColumn(loCaps, i, "ValidTo")

        If cap("UserId") = "" Or cap("Capability") = "" Then
            AddValidationIssue "WARN", "AUTH_CAP_ROW_SKIPPED", "Capability row missing UserId or Capability."
            GoTo ContinueLoop
        End If

        status = cap("Status")
        Select Case status
            Case "DENY"
                mDenyCaps.Add cap
            Case "ACTIVE", "ALLOW", ""
                mAllowCaps.Add cap
            Case Else
                ' Disabled/unknown rows are ignored by default.
        End Select

ContinueLoop:
    Next i
End Sub

Private Function IsUserActive(ByVal userId As String, ByVal nowTs As Date) As Boolean
    Dim d As Object
    If Not mUsers.Exists(userId) Then Exit Function

    Set d = mUsers(userId)
    If d.Exists("Status") Then
        If d("Status") <> "" And d("Status") <> "ACTIVE" Then Exit Function
    End If

    IsUserActive = IsWithinDateRange(d("ValidFrom"), d("ValidTo"), nowTs)
End Function

Private Function HasCapabilityMatch(ByVal caps As Collection, _
                                    ByVal userId As String, _
                                    ByVal capability As String, _
                                    ByVal warehouseId As String, _
                                    ByVal stationId As String, _
                                    ByVal nowTs As Date) As Boolean
    Dim entry As Object
    Dim wantedCap As String
    wantedCap = UCase$(SafeTrim(capability))

    For Each entry In caps
        If StrComp(SafeTrim(entry("UserId")), SafeTrim(userId), vbTextCompare) <> 0 Then GoTo NextEntry
        If Not CapabilityMatches(SafeTrim(entry("Capability")), wantedCap) Then GoTo NextEntry
        If Not ScopeMatches(SafeTrim(entry("WarehouseId")), warehouseId) Then GoTo NextEntry
        If Not ScopeMatches(SafeTrim(entry("StationId")), stationId) Then GoTo NextEntry
        If Not IsWithinDateRange(entry("ValidFrom"), entry("ValidTo"), nowTs) Then GoTo NextEntry

        HasCapabilityMatch = True
        Exit Function
NextEntry:
    Next entry
End Function

Private Function CapabilityMatches(ByVal entryCapability As String, ByVal wantedCapability As String) As Boolean
    entryCapability = UCase$(SafeTrim(entryCapability))
    If entryCapability = "*" Then
        CapabilityMatches = True
    Else
        CapabilityMatches = (entryCapability = wantedCapability)
    End If
End Function

Private Function ScopeMatches(ByVal scopeValue As String, ByVal currentValue As String) As Boolean
    scopeValue = SafeTrim(scopeValue)
    currentValue = SafeTrim(currentValue)

    If scopeValue = "" Or scopeValue = "*" Then
        ScopeMatches = True
    ElseIf currentValue = "" Then
        ScopeMatches = False
    Else
        ScopeMatches = (StrComp(scopeValue, currentValue, vbTextCompare) = 0)
    End If
End Function

Private Function IsWithinDateRange(ByVal validFrom As Variant, ByVal validTo As Variant, ByVal nowTs As Date) As Boolean
    IsWithinDateRange = True
    If IsDate(validFrom) Then
        If nowTs < CDate(validFrom) Then IsWithinDateRange = False
    End If
    If IsDate(validTo) Then
        If nowTs > CDate(validTo) Then IsWithinDateRange = False
    End If
End Function

Private Sub LogDecision(ByVal requestId As String, _
                        ByVal userId As String, _
                        ByVal capability As String, _
                        ByVal warehouseId As String, _
                        ByVal stationId As String, _
                        ByVal result As String, _
                        ByVal source As String, _
                        ByVal detail As String)
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss"), _
                "AUTH", _
                IIf(requestId = "", "-", requestId), _
                userId, _
                capability, _
                IIf(warehouseId = "", "-", warehouseId), _
                IIf(stationId = "", "-", stationId), _
                result, _
                source, _
                detail
End Sub

Private Function ResolveAuthWorkbook(ByVal whId As String) As Workbook
    Dim wb As Workbook
    Dim bootstrapReport As String
    Dim bootstrapWh As String

    For Each wb In Application.Workbooks
        If IsAuthWorkbookName(wb.Name) Then
            If whId = "" Or InStr(1, wb.Name, whId, vbTextCompare) > 0 Then
                Set ResolveAuthWorkbook = wb
                Exit Function
            End If
        End If
    Next wb

    If whId <> "" Then
        For Each wb In Application.Workbooks
            If WorkbookHasListObject(wb, "tblUsers") And WorkbookHasListObject(wb, "tblCapabilities") Then
                If WorkbookHasAuthScope(wb, whId) Then
                    Set ResolveAuthWorkbook = wb
                    Exit Function
                End If
            End If
        Next wb
    End If

    For Each wb In Application.Workbooks
        If WorkbookHasListObject(wb, "tblUsers") And WorkbookHasListObject(wb, "tblCapabilities") Then
            Set ResolveAuthWorkbook = wb
            Exit Function
        End If
    Next wb

    bootstrapWh = whId
    If bootstrapWh = "" Then bootstrapWh = modConfig.GetString("WarehouseId", "")

    If bootstrapWh <> "" Then
        Set ResolveAuthWorkbook = modRuntimeWorkbooks.OpenOrCreateAuthWorkbookRuntime(bootstrapWh, modConfig.GetString("ProcessorServiceUserId", "svc_processor"), "", bootstrapReport)
        If Not ResolveAuthWorkbook Is Nothing Then Exit Function
    End If

    Set ResolveAuthWorkbook = modRuntimeWorkbooks.OpenFirstRuntimeAuthWorkbook(bootstrapReport)
End Function

Private Function IsAuthWorkbookName(ByVal wbName As String) As Boolean
    Dim n As String
    n = LCase$(wbName)
    IsAuthWorkbookName = (n Like "wh*.invsys.auth.xlsb") Or _
                         (n Like "wh*.invsys.auth.xlsx") Or _
                         (n Like "wh*.invsys.auth.xlsm")
End Function

Private Function EnsureAuthTables(ByVal wb As Workbook) As Boolean
    On Error GoTo FailEnsure

    EnsureListObjectWithHeaders wb, "Users", "tblUsers", Array("UserId", "DisplayName", "PinHash", "Status", "ValidFrom", "ValidTo")
    EnsureListObjectWithHeaders wb, "Capabilities", "tblCapabilities", Array("UserId", "Capability", "WarehouseId", "StationId", "Status", "ValidFrom", "ValidTo")

    EnsureAuthTables = True
    Exit Function

FailEnsure:
    EnsureAuthTables = False
End Function

Private Sub EnsureListObjectWithHeaders(ByVal wb As Workbook, _
                                        ByVal sheetName As String, _
                                        ByVal tableName As String, _
                                        ByVal headers As Variant)
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long
    Dim dataRange As Range
    Dim startCell As Range

    Set ws = EnsureWorksheet(wb, sheetName)
    EnsureWorksheetEditableAuth ws
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    If lo Is Nothing Then
        Set startCell = GetNextTableStartCell(ws)
        For i = LBound(headers) To UBound(headers)
            startCell.Offset(0, i - LBound(headers)).Value = headers(i)
        Next i

        Set dataRange = ws.Range(startCell, startCell.Offset(1, UBound(headers) - LBound(headers)))
        Set lo = ws.ListObjects.Add(xlSrcRange, dataRange, , xlYes)
        lo.Name = tableName
        AddValidationIssue "WARN", "AUTH_TABLE_CREATED", tableName & " created."
    End If

    For i = LBound(headers) To UBound(headers)
        EnsureListColumn lo, CStr(headers(i))
    Next i
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

Private Sub SeedAuthDefaults(ByVal wb As Workbook, ByVal warehouseId As String, ByVal processorServiceUserId As String)
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim serviceUser As String
    Dim currentUser As String
    Dim resolvedWh As String

    Set loUsers = FindListObjectByName(wb, "tblUsers")
    Set loCaps = FindListObjectByName(wb, "tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then Exit Sub

    serviceUser = SafeTrim(processorServiceUserId)
    If serviceUser = "" Then serviceUser = "svc_processor"
    resolvedWh = SafeTrim(warehouseId)
    If resolvedWh = "" Then resolvedWh = modConfig.GetString("WarehouseId", "")
    currentUser = ResolveCurrentUserIdAuth()

    EnsureUserRow loUsers, serviceUser, "Processor Service"
    If currentUser <> "" Then EnsureUserRow loUsers, currentUser, currentUser

    If AuthTableHasCapabilityRows(loCaps) Then Exit Sub

    If resolvedWh <> "" And currentUser <> "" Then
        EnsureCapabilityRow loCaps, currentUser, "RECEIVE_POST", resolvedWh, "*", "ACTIVE"
        EnsureCapabilityRow loCaps, currentUser, "SHIP_POST", resolvedWh, "*", "ACTIVE"
        EnsureCapabilityRow loCaps, currentUser, "PROD_POST", resolvedWh, "*", "ACTIVE"
    End If
    If resolvedWh <> "" Then
        EnsureCapabilityRow loCaps, serviceUser, "INBOX_PROCESS", resolvedWh, "*", "ACTIVE"
    End If
End Sub

Private Sub EnsureUserRow(ByVal lo As ListObject, ByVal userId As String, ByVal displayName As String)
    Dim rowIndex As Long

    rowIndex = FindAuthUserRow(lo, userId)
    If rowIndex = 0 Then
        rowIndex = 1
        If Not lo.DataBodyRange Is Nothing Then
            If SafeTrim(lo.DataBodyRange.Cells(1, lo.ListColumns("UserId").Index).Value) <> "" Then
                lo.ListRows.Add
                rowIndex = lo.ListRows.Count
            End If
        End If
    End If

    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("UserId").Index).Value = userId
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("DisplayName").Index).Value = displayName
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("PinHash").Index).Value = ""
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("Status").Index).Value = "Active"
End Sub

Private Function FindAuthUserRow(ByVal lo As ListObject, ByVal userId As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrim(lo.DataBodyRange.Cells(i, lo.ListColumns("UserId").Index).Value), userId, vbTextCompare) = 0 Then
            FindAuthUserRow = i
            Exit Function
        End If
    Next i
End Function

Private Function AuthTableHasCapabilityRows(ByVal lo As ListObject) As Boolean
    Dim i As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If SafeTrim(lo.DataBodyRange.Cells(i, lo.ListColumns("UserId").Index).Value) <> "" _
           And SafeTrim(lo.DataBodyRange.Cells(i, lo.ListColumns("Capability").Index).Value) <> "" Then
            AuthTableHasCapabilityRows = True
            Exit Function
        End If
    Next i
End Function

Private Function ResolveCurrentUserIdAuth() As String
    ResolveCurrentUserIdAuth = Trim$(Environ$("USERNAME"))
    If ResolveCurrentUserIdAuth = "" Then ResolveCurrentUserIdAuth = Trim$(Application.UserName)
End Function

Private Sub EnsureCapabilityRow(ByVal lo As ListObject, _
                                ByVal userId As String, _
                                ByVal capability As String, _
                                ByVal warehouseId As String, _
                                ByVal stationId As String, _
                                ByVal statusVal As String)
    Dim rowIndex As Long

    rowIndex = FindCapabilityRow(lo, userId, capability, warehouseId, stationId)
    If rowIndex = 0 Then
        rowIndex = NextWritableAuthRow(lo, "UserId")
        If rowIndex > lo.ListRows.Count Then lo.ListRows.Add
        AddValidationIssue "WARN", "AUTH_CAPABILITY_CREATED", userId & "." & capability & " created."
    End If

    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("UserId").Index).Value = userId
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("Capability").Index).Value = capability
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("WarehouseId").Index).Value = warehouseId
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("StationId").Index).Value = stationId
    lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("Status").Index).Value = statusVal
End Sub

Private Function FindCapabilityRow(ByVal lo As ListObject, _
                                   ByVal userId As String, _
                                   ByVal capability As String, _
                                   ByVal warehouseId As String, _
                                   ByVal stationId As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrim(lo.DataBodyRange.Cells(i, lo.ListColumns("UserId").Index).Value), userId, vbTextCompare) = 0 _
           And StrComp(UCase$(SafeTrim(lo.DataBodyRange.Cells(i, lo.ListColumns("Capability").Index).Value)), UCase$(capability), vbTextCompare) = 0 _
           And StrComp(SafeTrim(lo.DataBodyRange.Cells(i, lo.ListColumns("WarehouseId").Index).Value), warehouseId, vbTextCompare) = 0 _
           And StrComp(SafeTrim(lo.DataBodyRange.Cells(i, lo.ListColumns("StationId").Index).Value), stationId, vbTextCompare) = 0 Then
            FindCapabilityRow = i
            Exit Function
        End If
    Next i
End Function

Private Function NextWritableAuthRow(ByVal lo As ListObject, ByVal keyColumnName As String) As Long
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        NextWritableAuthRow = 1
        Exit Function
    End If

    NextWritableAuthRow = 1
    If SafeTrim(lo.DataBodyRange.Cells(1, lo.ListColumns(keyColumnName).Index).Value) <> "" Then
        NextWritableAuthRow = lo.ListRows.Count + 1
    End If
End Function

Private Sub FormatAuthSurface(ByVal wb As Workbook)
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        ws.Cells.EntireColumn.AutoFit
        ws.Rows(1).Font.Bold = True
    Next ws
End Sub

Private Function GetNextTableStartCell(ByVal ws As Worksheet) As Range
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        Set GetNextTableStartCell = ws.Range("A1")
    Else
        Set GetNextTableStartCell = ws.Cells(ws.Rows.Count, 1).End(xlUp).Offset(2, 0)
    End If
End Function

Private Sub EnsureWorksheetEditableAuth(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub
    If Not ws.ProtectContents Then Exit Sub

    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 2702, "modAuth.EnsureWorksheetEditableAuth", _
                  "Worksheet '" & ws.Name & "' is protected and could not be unprotected before updating auth tables."
    End If
End Sub

Private Sub EnsureListColumn(ByVal lo As ListObject, ByVal columnName As String)
    If GetColumnIndex(lo, columnName) > 0 Then Exit Sub
    lo.ListColumns.Add lo.ListColumns.Count + 1
    lo.ListColumns(lo.ListColumns.Count).Name = columnName
    AddValidationIssue "WARN", "AUTH_COLUMN_CREATED", lo.Name & "." & columnName & " created."
End Sub

Private Function GetCellByColumn(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long
    idx = GetColumnIndex(lo, columnName)
    If idx = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    GetCellByColumn = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

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

Private Function WorkbookHasListObject(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    WorkbookHasListObject = Not (FindListObjectByName(wb, tableName) Is Nothing)
End Function

Private Function WorkbookHasAuthScope(ByVal wb As Workbook, ByVal whId As String) As Boolean
    Dim lo As ListObject
    Dim i As Long
    Dim scopeVal As String

    Set lo = FindListObjectByName(wb, "tblCapabilities")
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        scopeVal = SafeTrim(GetCellByColumn(lo, i, "WarehouseId"))
        If StrComp(scopeVal, whId, vbTextCompare) = 0 Or scopeVal = "*" Then
            WorkbookHasAuthScope = True
            Exit Function
        End If
    Next i
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

Private Function SafeTrim(ByVal v As Variant) As String
    On Error Resume Next
    SafeTrim = Trim$(CStr(v))
End Function

Private Sub AddValidationIssue(ByVal severity As String, ByVal code As String, ByVal message As String)
    If mValidationIssues Is Nothing Then Set mValidationIssues = New Collection
    mValidationIssues.Add severity & "|" & code & "|" & message
End Sub

Private Function CountFatalIssues() As Long
    Dim itm As Variant
    Dim parts() As String
    For Each itm In mValidationIssues
        parts = Split(CStr(itm), "|")
        If UBound(parts) >= 0 Then
            If UCase$(parts(0)) = "ERROR" Then CountFatalIssues = CountFatalIssues + 1
        End If
    Next itm
End Function
