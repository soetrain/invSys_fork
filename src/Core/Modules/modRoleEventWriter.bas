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
Private Const ROLE_EVENT_TYPE_BOX_BUILD As String = "BOX_BUILD"
Private Const ROLE_EVENT_TYPE_BOX_UNBOX As String = "BOX_UNBOX"
Private Const ROLE_EVENT_TYPE_PROD_CONSUME As String = "PROD_CONSUME"
Private Const ROLE_EVENT_TYPE_PROD_COMPLETE As String = "PROD_COMPLETE"
Private Const ROLE_EVENT_TYPE_MIGRATION_SEED As String = "MIGRATION_SEED"
Private Const SETTINGS_APP As String = "invSys"
Private Const SETTINGS_SECTION_RUNTIME As String = "Runtime"
Private Const SETTINGS_CURRENT_USER_ID As String = "CurrentUserId"
Private Const LOCAL_STAGING_ROOT_FOLDER As String = "invSys\staging"
Private Const LOCAL_STAGING_ARCHIVE_FOLDER As String = "Archive"

Private mStagingSyncInProgress As Boolean

Public Function ResolveCurrentUserId() As String
    ResolveCurrentUserId = Trim$(GetCurrentUserOverride())
    If ResolveCurrentUserId <> "" Then Exit Function

    ResolveCurrentUserId = Trim$(Environ$("USERNAME"))
    If ResolveCurrentUserId = "" Then ResolveCurrentUserId = Trim$(Application.UserName)
End Function

Public Function GetCurrentUserOverride() As String
    On Error Resume Next
    GetCurrentUserOverride = Trim$(GetSetting(SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_CURRENT_USER_ID, ""))
    On Error GoTo 0
End Function

Public Sub SetCurrentUserId(ByVal userId As String)
    userId = Trim$(userId)
    If userId = "" Then
        On Error Resume Next
        DeleteSetting SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_CURRENT_USER_ID
        On Error GoTo 0
    Else
        SaveSetting SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_CURRENT_USER_ID, userId
    End If
    modRibbonRuntimeStatus.InvalidateCurrentUserRibbons
End Sub

Public Sub PromptSetCurrentUser()
    PromptSetCurrentUserForCapability ""
End Sub

Public Sub PromptSetCurrentUserForCapability(Optional ByVal requiredCapability As String = "")
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim authStatus As AuthStatusCode
    Dim requireNasTarget As Boolean

    requireNasTarget = CapabilityRequiresNasTargetRole(requiredCapability)
    Set target = ResolveRoleWarehouseTarget(requireNasTarget, statusCode)
    If target Is Nothing Then
        MsgBox "Warehouse storage is not connected. Use Connect Server or Runtime Context before signing in.", vbExclamation, "invSys Current User"
        modRibbonRuntimeStatus.InvalidateCurrentUserRibbons
        Exit Sub
    End If

    authStatus = modAuth.ShowSignInPrompt(target, requiredCapability)
    If authStatus = AUTH_CANCELLED Then Exit Sub

    If authStatus <> AUTH_OK Then
        MsgBox AuthStatusMessageRole(authStatus, target, requiredCapability), vbExclamation, "invSys Current User"
        Exit Sub
    End If

    MsgBox "Current invSys user: " & CurrentInvSysUserDisplayRole(), vbInformation, "invSys Current User"
End Sub

Public Sub ShowCurrentUser()
    MsgBox "Current invSys user: " & CurrentInvSysUserDisplayRole(), vbInformation, "invSys Current User"
End Sub

Public Sub ConnectWarehouseStorageForCapability(Optional ByVal requiredCapability As String = "")
    Dim connectedRoot As String
    Dim statusText As String
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode
    Dim requireNasTarget As Boolean

    requireNasTarget = CapabilityRequiresNasTargetRole(requiredCapability)
    If modNasConnection.ConnectKnownWarehouseServer(connectedRoot, statusText, True) Then
        Set target = ResolveConnectedRoleWarehouseTarget(connectedRoot, requireNasTarget, statusCode)
        modRibbonRuntimeStatus.InvalidateWarehouseTargets
        modRibbonRuntimeStatus.InvalidateCurrentUserRibbons
        If target Is Nothing Then
            MsgBox "Connected to the warehouse server, but no warehouse target was selected." & vbCrLf & vbCrLf & _
                   "Connected root: " & ValueOrPlaceholderRole(connectedRoot) & vbCrLf & _
                   "Status: " & ValueOrPlaceholderRole(modNasConnection.GetConnectionStatus()) & vbCrLf & vbCrLf & _
                   "Use Send To to choose the NAS warehouse. If it is not listed, use Admin > View Warehouses to inspect the server root.", _
                   vbExclamation, "invSys Warehouse Storage"
        End If
        Exit Sub
    End If

    If ShouldPromptForServerCredentialsRole(statusText) Then
        If modNasConnection.ShowWarehouseConnectionPromptForTarget(ServerCredentialPromptRole(statusText)) Then
            Set target = modNasConnection.GetCurrentTarget()
            modRibbonRuntimeStatus.InvalidateWarehouseTargets
            modRibbonRuntimeStatus.InvalidateCurrentUserRibbons
            If modNasConnection.IsWarehouseTargetAllowed(target, requireNasTarget) Then Exit Sub
        End If
    End If

    modRibbonRuntimeStatus.InvalidateWarehouseTargets
    modRibbonRuntimeStatus.InvalidateCurrentUserRibbons
    MsgBox "Could not connect to the warehouse server." & vbCrLf & vbCrLf & _
           "Saved root: " & ValueOrPlaceholderRole(modNasConnection.GetPromptDefaultRoot()) & vbCrLf & _
           "Status: " & ValueOrPlaceholderRole(statusText) & vbCrLf & vbCrLf & _
           "Use Admin/setup to add or repair the warehouse server root, then try Connect Server again.", _
           vbExclamation, "invSys Warehouse Storage"
End Sub

Private Function ShouldPromptForServerCredentialsRole(ByVal statusText As String) As Boolean
    statusText = LCase$(Trim$(statusText))
    ShouldPromptForServerCredentialsRole = _
        (InStr(1, statusText, "credential rejected", vbTextCompare) > 0) Or _
        (InStr(1, statusText, "credential", vbTextCompare) > 0 And InStr(1, statusText, "expired", vbTextCompare) > 0) Or _
        (InStr(1, statusText, "conflicting nas session", vbTextCompare) > 0)
End Function

Private Function ServerCredentialPromptRole(ByVal statusText As String) As String
    ServerCredentialPromptRole = _
        "The warehouse server path is saved, but Windows does not have a usable NAS/server credential for this session." & vbCrLf & vbCrLf & _
        "Enter the NAS/server account for storage access, connect, select the Zenbook warehouse runtime, then sign in with your invSys user account." & vbCrLf & vbCrLf & _
        "Status: " & ValueOrPlaceholderRole(statusText)
End Function

Private Function ResolveConnectedRoleWarehouseTarget(ByVal connectedRoot As String, _
                                                     ByVal requireNasTarget As Boolean, _
                                                     ByRef statusCode As NasStatusCode) As WarehouseTarget
    Dim target As WarehouseTarget
    Dim runtimeRoots As Collection
    Dim runtimeRoot As Variant

    If modRibbonRuntimeStatus.TryApplyRememberedWarehouseTarget() Then
        Set target = modNasConnection.GetCurrentTarget()
        If modNasConnection.IsWarehouseTargetAllowed(target, requireNasTarget) _
           And WarehouseTargetMatchesConnectedRootRole(target, connectedRoot) Then
            statusCode = NAS_OK
            Set ResolveConnectedRoleWarehouseTarget = target
            Exit Function
        End If
    End If

    Set target = modNasConnection.GetCurrentTarget()
    If modNasConnection.IsWarehouseTargetAllowed(target, requireNasTarget) _
       And WarehouseTargetMatchesConnectedRootRole(target, connectedRoot) Then
        statusCode = NAS_OK
        Set ResolveConnectedRoleWarehouseTarget = target
        Exit Function
    End If

    Set runtimeRoots = modNasConnection.ScanNasRoot(connectedRoot)
    For Each runtimeRoot In runtimeRoots
        statusCode = modNasConnection.SelectWarehouseTarget(connectedRoot, CStr(runtimeRoot), target)
        If statusCode = NAS_OK Then
            If modNasConnection.IsWarehouseTargetAllowed(target, requireNasTarget) Then
                Set ResolveConnectedRoleWarehouseTarget = target
                Exit Function
            End If
        End If
    Next runtimeRoot

    statusCode = WH_NO_TARGET
End Function

Private Function ResolveRoleWarehouseTarget(ByVal requireNasTarget As Boolean, _
                                            ByRef statusCode As NasStatusCode) As WarehouseTarget
    Dim target As WarehouseTarget

    If modRibbonRuntimeStatus.TryApplyRememberedWarehouseTarget() Then
        Set target = modNasConnection.GetCurrentTarget()
        If modNasConnection.IsWarehouseTargetAllowed(target, requireNasTarget) Then
            Set ResolveRoleWarehouseTarget = target
            Exit Function
        End If
    End If

    Set target = modNasConnection.GetCurrentTarget()
    If modNasConnection.IsWarehouseTargetAllowed(target, requireNasTarget) Then
        Set ResolveRoleWarehouseTarget = target
        Exit Function
    End If
End Function

Public Sub SignOutCurrentUser()
    modAuth.SignOut
    SetCurrentUserId vbNullString
    modRibbonRuntimeStatus.InvalidateCurrentUserRibbons
    If modAuth.IsSignedIn() Or Trim$(modAuth.GetCurrentUserId()) <> "" Or Trim$(GetCurrentUserOverride()) <> "" Then
        MsgBox "Sign out did not complete. Close and reopen Excel, then try again.", vbExclamation, "invSys Current User"
    Else
        MsgBox "Signed out of invSys. Warehouse storage remains selected.", vbInformation, "invSys Current User"
    End If
End Sub

Private Function ValidateCurrentUserCredential(ByVal userId As String, _
                                               ByVal pinText As String, _
                                               ByVal requiredCapability As String, _
                                               ByRef report As String) As Boolean
    Dim whId As String
    Dim stId As String
    Dim credentialStatus As String

    If Not modConfig.LoadConfig("", "") Then
        report = "Runtime config could not be loaded: " & modConfig.Validate()
        Exit Function
    End If

    whId = modConfig.GetWarehouseId()
    stId = modConfig.GetStationId()
    If Not modAuth.LoadAuth(whId) Then
        report = "Auth workbook could not be loaded: " & modAuth.ValidateAuth()
        Exit Function
    End If

    If Not modAuth.ValidateUserCredentialForCapability(userId, pinText, "") Then
        credentialStatus = modAuth.DiagnoseUserCredential(userId, pinText)
        report = "Invalid credentials for '" & userId & "'." & vbCrLf & _
                 "Auth workbook: " & modAuth.GetResolvedAuthWorkbookName() & vbCrLf & _
                 "Detail: " & credentialStatus
        Exit Function
    End If

    If Trim$(requiredCapability) <> "" _
       And Not modAuth.HasProvisionedCapabilityForSystem(requiredCapability, userId, whId, stId) Then
        report = "'" & userId & "' lacks " & requiredCapability & " for " & whId & " / " & stId & "." & vbCrLf & _
                 "Auth workbook: " & modAuth.GetResolvedAuthWorkbookName()
        Exit Function
    End If

    ValidateCurrentUserCredential = True
End Function

Private Function CapabilityRequiresNasTargetRole(ByVal requiredCapability As String) As Boolean
    Select Case UCase$(Trim$(requiredCapability))
        Case "RECEIVE_POST", "SHIP_POST", "PROD_POST"
            CapabilityRequiresNasTargetRole = True
    End Select
End Function

Private Function CurrentInvSysUserDisplayRole() As String
    If modAuth.IsSignedIn() Then
        CurrentInvSysUserDisplayRole = Trim$(modAuth.GetCurrentUserDisplayName())
    End If
    If CurrentInvSysUserDisplayRole = "" Then CurrentInvSysUserDisplayRole = "<not signed in>"
End Function

Private Function WarehouseTargetMatchesConnectedRootRole(ByVal target As WarehouseTarget, _
                                                        ByVal connectedRoot As String) As Boolean
    If target Is Nothing Then Exit Function
    connectedRoot = NormalizeFolderRole(connectedRoot)
    If connectedRoot = "" Then
        WarehouseTargetMatchesConnectedRootRole = True
        Exit Function
    End If

    WarehouseTargetMatchesConnectedRootRole = _
        PathIsUnderRootRole(target.HubRoot, connectedRoot) Or _
        PathIsUnderRootRole(target.RuntimeRoot, connectedRoot) Or _
        PathIsUnderRootRole(connectedRoot, target.HubRoot) Or _
        PathIsUnderRootRole(connectedRoot, target.RuntimeRoot)
End Function

Private Function PathIsUnderRootRole(ByVal pathText As String, ByVal rootText As String) As Boolean
    pathText = LCase$(NormalizeFolderRole(pathText))
    rootText = LCase$(NormalizeFolderRole(rootText))
    If pathText = "" Or rootText = "" Then Exit Function
    If pathText = rootText Then
        PathIsUnderRootRole = True
    ElseIf Len(pathText) > Len(rootText) Then
        PathIsUnderRootRole = (Left$(pathText, Len(rootText) + 1) = rootText & "\")
    End If
End Function

Private Function NormalizeFolderRole(ByVal folderPath As String) As String
    NormalizeFolderRole = Trim$(Replace$(folderPath, "/", "\"))
    Do While Len(NormalizeFolderRole) > 3 And Right$(NormalizeFolderRole, 1) = "\"
        NormalizeFolderRole = Left$(NormalizeFolderRole, Len(NormalizeFolderRole) - 1)
    Loop
End Function

Private Function AuthStatusMessageRole(ByVal authStatus As AuthStatusCode, _
                                       ByVal target As WarehouseTarget, _
                                       ByVal requiredCapability As String) As String
    Dim targetLabel As String

    If target Is Nothing Then
        targetLabel = "<no warehouse target>"
    Else
        targetLabel = target.WarehouseId & " / " & IIf(Trim$(target.StationId) = "", "<roaming>", target.StationId)
    End If

    Select Case authStatus
        Case AUTH_WAREHOUSE_MISMATCH
            AuthStatusMessageRole = "Warehouse target mismatch."
        Case AUTH_USER_NOT_FOUND
            AuthStatusMessageRole = "User was not found for " & targetLabel & "."
        Case AUTH_CREDENTIAL_REJECTED
            AuthStatusMessageRole = "PIN/password was rejected."
        Case AUTH_WORKBOOK_UNREADABLE
            AuthStatusMessageRole = "Auth workbook could not be read for " & targetLabel & "."
        Case AUTH_NO_CAPABILITIES
            AuthStatusMessageRole = "User lacks " & UCase$(Trim$(requiredCapability)) & " for " & targetLabel & "."
        Case AUTH_REAUTH_REQUIRED, AUTH_CACHE_EXPIRED
            AuthStatusMessageRole = "Sign-in expired. Sign in again."
        Case Else
            AuthStatusMessageRole = "Sign-in failed. Status: " & CStr(authStatus)
    End Select
End Function

Private Function ValueOrPlaceholderRole(ByVal valueText As String) As String
    valueText = Trim$(valueText)
    If valueText = "" Then
        ValueOrPlaceholderRole = "<not saved>"
    Else
        ValueOrPlaceholderRole = valueText
    End If
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

Public Function DescribeInboxPendingRows(ByVal eventType As String, _
                                         Optional ByVal warehouseId As String = "", _
                                         Optional ByVal stationId As String = "", _
                                         Optional ByVal eventIdsCsv As String = "", _
                                         Optional ByRef pendingCount As Long = 0, _
                                         Optional ByRef matchingPendingCount As Long = 0, _
                                         Optional ByRef errorMessage As String = "") As String
    On Error GoTo FailDescribe

    Dim resolvedWh As String
    Dim resolvedSt As String
    Dim fullPath As String
    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim openPaths As Object
    Dim openedTransient As Boolean
    Dim rowIndex As Long
    Dim statusVal As String
    Dim eventId As String
    Dim createdVal As Variant
    Dim oldestCreated As String
    Dim newestCreated As String
    Dim sampleIds As String
    Dim report As String

    pendingCount = 0
    matchingPendingCount = 0
    eventIdsCsv = Trim$(eventIdsCsv)

    If Not EnsureContextResolved(resolvedWh, resolvedSt, warehouseId, stationId, errorMessage) Then Exit Function

    fullPath = ResolveInboxWorkbookPathResolvedRole(eventType, resolvedWh, resolvedSt, errorMessage)
    If fullPath = "" Then Exit Function

    Set openPaths = CaptureOpenWorkbookPathsRole()
    Set wbInbox = ResolveInboxWorkbookForEventType(eventType, resolvedWh, resolvedSt, errorMessage)
    If wbInbox Is Nothing Then Exit Function
    openedTransient = Not WorkbookWasAlreadyOpenRole(openPaths, wbInbox)

    Select Case UCase$(Trim$(eventType))
        Case ROLE_EVENT_TYPE_RECEIVE
            If Not modProcessor.EnsureReceiveInboxSchema(wbInbox, report) Then
                errorMessage = report
                GoTo CleanExit
            End If
        Case ROLE_EVENT_TYPE_SHIP, ROLE_EVENT_TYPE_BOX_BUILD, ROLE_EVENT_TYPE_BOX_UNBOX
            If Not modProcessor.EnsureShipInboxSchema(wbInbox, report) Then
                errorMessage = report
                GoTo CleanExit
            End If
        Case ROLE_EVENT_TYPE_PROD_CONSUME, ROLE_EVENT_TYPE_PROD_COMPLETE, ROLE_EVENT_TYPE_MIGRATION_SEED
            If Not modProcessor.EnsureProductionInboxSchema(wbInbox, report) Then
                errorMessage = report
                GoTo CleanExit
            End If
    End Select

    Set lo = FindListObjectByNameRole(wbInbox, InboxTableNameRole(eventType))
    If lo Is Nothing Then
        errorMessage = "Inbox table not found for event type '" & eventType & "'."
        GoTo CleanExit
    End If

    If Not lo.DataBodyRange Is Nothing Then
        For rowIndex = 1 To lo.ListRows.Count
            statusVal = UCase$(Trim$(CStr(GetTableRowValueRole(lo, rowIndex, "Status"))))
            If statusVal = "" Or statusVal = "NEW" Then
                pendingCount = pendingCount + 1
                eventId = Trim$(CStr(GetTableRowValueRole(lo, rowIndex, "EventID")))
                If EventIdListedRole(eventId, eventIdsCsv) Then matchingPendingCount = matchingPendingCount + 1
                createdVal = GetTableRowValueRole(lo, rowIndex, "CreatedAtUTC")
                If IsDate(createdVal) Then
                    If oldestCreated = "" Then oldestCreated = Format$(CDate(createdVal), "yyyy-mm-dd hh:nn:ss")
                    newestCreated = Format$(CDate(createdVal), "yyyy-mm-dd hh:nn:ss")
                End If
                If sampleIds = "" Then
                    sampleIds = eventId
                ElseIf pendingCount <= 3 Then
                    sampleIds = sampleIds & "," & eventId
                End If
            End If
        Next rowIndex
    End If

    DescribeInboxPendingRows = "Path=" & fullPath & _
        "; PendingRows=" & CStr(pendingCount) & _
        "; MatchingPendingRows=" & CStr(matchingPendingCount) & _
        "; OldestCreatedAt=" & oldestCreated & _
        "; NewestCreatedAt=" & newestCreated & _
        "; SampleEventIds=" & sampleIds

CleanExit:
    On Error Resume Next
    If openedTransient Then CloseTransientRoleWorkbook wbInbox
    On Error GoTo 0
    Exit Function

FailDescribe:
    errorMessage = "Inbox pending inspection failed: " & Err.Description
    Resume CleanExit
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
                                  Optional ByRef errorMessage As String = "", _
                                  Optional ByVal perfRunId As String = "") As Boolean
    QueueReceiveEvent = QueueEventCore(ROLE_EVENT_TYPE_RECEIVE, warehouseId, stationId, userId, sku, qty, location, noteVal, "", "", parentEventId, undoOfEventId, createdAtUtc, targetInboxWb, eventIdOut, errorMessage, perfRunId)
End Function

Public Function QueueReceiveEventCurrent(Optional ByVal userId As String = "", _
                                         Optional ByVal sku As String = "", _
                                         Optional ByVal qty As Double = 0, _
                                         Optional ByVal location As String = "", _
                                         Optional ByVal noteVal As String = "", _
                                         Optional ByRef eventIdOut As String = "", _
                                         Optional ByRef errorMessage As String = "", _
                                         Optional ByVal perfRunId As String = "") As Boolean
    Dim targetInboxWb As Workbook
    Dim resolvedUser As String

    If Not EnsureCurrentRoleWriteAllowed("RECEIVE_POST", userId, resolvedUser, errorMessage) Then Exit Function
    QueueReceiveEventCurrent = QueueReceiveEvent("", "", resolvedUser, sku, qty, location, noteVal, "", "", 0, targetInboxWb, eventIdOut, errorMessage, perfRunId)
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
                                  Optional ByRef errorMessage As String = "", _
                                  Optional ByVal perfRunId As String = "") As Boolean
    QueuePayloadEvent = QueueEventCore(eventType, warehouseId, stationId, userId, "", 0, "", noteVal, payloadJson, "", parentEventId, undoOfEventId, createdAtUtc, targetInboxWb, eventIdOut, errorMessage, perfRunId)
End Function

Public Function QueueMigrationSeedEvent(Optional ByVal warehouseId As String = "", _
                                        Optional ByVal stationId As String = "", _
                                        Optional ByVal userId As String = "", _
                                        Optional ByVal payloadJson As String = "", _
                                        Optional ByVal migrationSourceId As String = "", _
                                        Optional ByVal noteVal As String = "", _
                                        Optional ByVal createdAtUtc As Date = 0, _
                                        Optional ByVal targetInboxWb As Workbook = Nothing, _
                                        Optional ByRef eventIdOut As String = "", _
                                        Optional ByRef errorMessage As String = "", _
                                        Optional ByVal perfRunId As String = "") As Boolean
    QueueMigrationSeedEvent = QueueEventCore(ROLE_EVENT_TYPE_MIGRATION_SEED, warehouseId, stationId, userId, "", 0, "", noteVal, payloadJson, migrationSourceId, "", "", createdAtUtc, targetInboxWb, eventIdOut, errorMessage, perfRunId)
End Function

Public Function QueuePayloadEventCurrent(ByVal eventType As String, _
                                         Optional ByVal userId As String = "", _
                                         Optional ByVal payloadJson As String = "", _
                                         Optional ByVal noteVal As String = "", _
                                         Optional ByRef eventIdOut As String = "", _
                                         Optional ByRef errorMessage As String = "", _
                                         Optional ByVal perfRunId As String = "") As Boolean
    Dim targetInboxWb As Workbook
    Dim resolvedUser As String
    Dim capability As String
    Dim target As WarehouseTarget

    capability = CapabilityForEventTypeRole(eventType)
    If Not EnsureCurrentRoleWriteAllowed(capability, userId, resolvedUser, errorMessage) Then Exit Function
    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        errorMessage = "A connected NAS warehouse target is required before posting role events."
        Exit Function
    End If
    QueuePayloadEventCurrent = QueueEventCore(eventType, _
                                             target.WarehouseId, _
                                             target.StationId, _
                                             resolvedUser, _
                                             "", _
                                             0, _
                                             "", _
                                             noteVal, _
                                             payloadJson, _
                                             "", _
                                             "", _
                                             "", _
                                             0, _
                                             targetInboxWb, _
                                             eventIdOut, _
                                             errorMessage, _
                                             perfRunId, _
                                             True)
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
                                ByVal migrationSourceId As String, _
                                ByVal parentEventId As String, _
                                ByVal undoOfEventId As String, _
                                ByVal createdAtUtc As Date, _
                                ByVal targetInboxWb As Workbook, _
                                ByRef eventIdOut As String, _
                                ByRef errorMessage As String, _
                                Optional ByVal perfRunId As String = "", _
                                Optional ByVal currentAuthAlreadyChecked As Boolean = False) As Boolean
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
    Dim queueStart As Single
    Dim queueRunId As String
    Dim perfStarted As Boolean
    Dim rowValues As Object
    Dim stagingPath As String
    Dim localStageOnly As Boolean

    queueStart = Timer

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

    localStageOnly = (targetInboxWb Is Nothing And ShouldStageEventLocallyRole(eventType))
    If Not (localStageOnly And currentAuthAlreadyChecked) Then
        If Not modAuth.LoadAuth(resolvedWh) Then
            errorMessage = "Auth load failed: " & modAuth.ValidateAuth()
            Exit Function
        End If
        If Not modAuth.HasProvisionedCapabilityForSystem(capability, resolvedUser, resolvedWh, resolvedSt) Then
            errorMessage = "Current user lacks " & capability & " capability." & vbCrLf & _
                           "User=" & ValueOrPlaceholderRole(resolvedUser) & _
                           "; Warehouse=" & ValueOrPlaceholderRole(resolvedWh) & _
                           "; Station=" & ValueOrPlaceholderRole(resolvedSt) & _
                           "; Auth=" & ValueOrPlaceholderRole(modAuth.GetResolvedAuthWorkbookName())
            Exit Function
        End If
    End If

    If eventIdOut = "" Then eventIdOut = CreateEventIdRole()
    If createdAtUtc = 0 Then createdAtUtc = Now
    queueRunId = Trim$(perfRunId)
    If queueRunId = "" Then queueRunId = eventIdOut
    If queueRunId <> "" Then
        PerfBeginSafeRole queueRunId, "RoleEventWriter.QueueEvent"
        perfStarted = True
    End If

    Set rowValues = BuildInboxRowValuesRole(eventIdOut, parentEventId, undoOfEventId, _
        eventType, createdAtUtc, resolvedWh, resolvedSt, resolvedUser, migrationSourceId, _
        sku, qty, location, noteVal, payloadJson)

    If localStageOnly Then
        stagingPath = LocalStagingPathRole(eventType, resolvedWh, resolvedSt)
        If Not AppendInboxRowToLocalStagingRole(rowValues, stagingPath, errorMessage) Then GoTo CleanExit
        If queueRunId <> "" Then PerfMarkSafeRole queueRunId, "LocalStagingWrite", CLng((Timer - queueStart) * 1000)
        QueueEventCore = True
        GoTo CleanExit
    End If

    If targetInboxWb Is Nothing Then
        Set openPaths = CaptureOpenWorkbookPathsRole()
        Set wbInbox = ResolveInboxWorkbookForEventType(eventType, resolvedWh, resolvedSt, errorMessage)
    Else
        Set wbInbox = targetInboxWb
    End If
    If wbInbox Is Nothing Then Exit Function
    If wbInbox.ReadOnly Then
        errorMessage = "Inbox workbook is read-only or locked by another Excel session."
        GoTo CleanExit
    End If
    openedTransient = (targetInboxWb Is Nothing) And (Not WorkbookWasAlreadyOpenRole(openPaths, wbInbox))
    If openedTransient Then HideWorkbookWindowsRole wbInbox

    Select Case UCase$(Trim$(eventType))
        Case ROLE_EVENT_TYPE_RECEIVE
            If Not modProcessor.EnsureReceiveInboxSchema(wbInbox, report) Then
                errorMessage = report
                GoTo CleanExit
            End If
            Set lo = FindListObjectByNameRole(wbInbox, TABLE_INBOX_RECEIVE)
        Case ROLE_EVENT_TYPE_SHIP, ROLE_EVENT_TYPE_BOX_BUILD, ROLE_EVENT_TYPE_BOX_UNBOX
            If Not modProcessor.EnsureShipInboxSchema(wbInbox, report) Then
                errorMessage = report
                GoTo CleanExit
            End If
            Set lo = FindListObjectByNameRole(wbInbox, TABLE_INBOX_SHIP)
        Case ROLE_EVENT_TYPE_PROD_CONSUME, ROLE_EVENT_TYPE_PROD_COMPLETE, ROLE_EVENT_TYPE_MIGRATION_SEED
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

    WriteInboxRowValuesRole lo, rowValues

    SaveWorkbookRole wbInbox
    If queueRunId <> "" Then PerfMarkSafeRole queueRunId, "InboxWrite", CLng((Timer - queueStart) * 1000)

    QueueEventCore = True
CleanExit:
    If perfStarted Then PerfEndSafeRole queueRunId, CLng((Timer - queueStart) * 1000), IIf(QueueEventCore, "OK", "FAIL") & " EventType=" & UCase$(Trim$(eventType))
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

Public Function SyncLocalStagedInboxRows(Optional ByRef report As String = "", _
                                         Optional ByVal warehouseId As String = "", _
                                         Optional ByVal stationId As String = "") As Boolean
    On Error GoTo FailSync

    Dim files As Collection
    Dim filePath As Variant
    Dim mergedTotal As Long
    Dim failedCount As Long
    Dim fileMerged As Long
    Dim fileReport As String
    Dim searchRoot As String

    If mStagingSyncInProgress Then
        report = "Local staging sync already in progress."
        SyncLocalStagedInboxRows = True
        Exit Function
    End If

    mStagingSyncInProgress = True
    Set files = New Collection
    searchRoot = LocalStagingSearchRootRole(warehouseId, stationId)
    CollectLocalStagingFilesRole searchRoot, files
    If files.Count = 0 Then
        report = "No local staged inbox rows."
        SyncLocalStagedInboxRows = True
        GoTo CleanExit
    End If

    For Each filePath In files
        fileMerged = 0
        fileReport = vbNullString
        If MergeLocalStagingFileRole(CStr(filePath), fileMerged, fileReport) Then
            mergedTotal = mergedTotal + fileMerged
        Else
            failedCount = failedCount + 1
            If report <> "" Then report = report & "; "
            report = report & fileReport
        End If
    Next filePath

    If report <> "" Then report = report & "; "
    report = report & "LocalStagingMerged=" & CStr(mergedTotal) & "; LocalStagingFailed=" & CStr(failedCount)
    SyncLocalStagedInboxRows = (failedCount = 0)

CleanExit:
    mStagingSyncInProgress = False
    Exit Function

FailSync:
    report = "Local staging sync failed: " & Err.Description
    Resume CleanExit
End Function

Public Function GetLocalStagedBoxInventoryDeltas(Optional ByVal warehouseId As String = "", _
                                                 Optional ByVal stationId As String = "") As Object
    On Error GoTo FailSoft

    Dim result As Object
    Dim files As Collection
    Dim filePath As Variant
    Dim rows As Collection
    Dim rowValues As Variant
    Dim rowDict As Object
    Dim eventType As String
    Dim payloadJson As String
    Dim report As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set GetLocalStagedBoxInventoryDeltas = result

    Set files = New Collection
    CollectLocalStagingFilesRole LocalStagingSearchRootRole(warehouseId, stationId), files
    For Each filePath In files
        report = vbNullString
        Set rows = ReadStagedInboxRowsRole(CStr(filePath), report)
        If rows Is Nothing Then GoTo NextFile
        For Each rowValues In rows
            Set rowDict = rowValues
            eventType = UCase$(Trim$(CStr(rowDict("EventType"))))
            If eventType = ROLE_EVENT_TYPE_BOX_BUILD Or eventType = ROLE_EVENT_TYPE_BOX_UNBOX Then
                payloadJson = CStr(rowDict("PayloadJson"))
                AccumulateBoxPayloadInventoryDeltasRole payloadJson, eventType, result
            End If
        Next rowValues
NextFile:
    Next filePath
    Exit Function

FailSoft:
    If GetLocalStagedBoxInventoryDeltas Is Nothing Then
        Set GetLocalStagedBoxInventoryDeltas = CreateObject("Scripting.Dictionary")
        GetLocalStagedBoxInventoryDeltas.CompareMode = vbTextCompare
    End If
End Function

Public Function GetLocalStagedBoxVersionInventoryDeltas(ByVal packageRow As Long, _
                                                        Optional ByVal warehouseId As String = "", _
                                                        Optional ByVal stationId As String = "") As Object
    On Error GoTo FailSoft

    Dim result As Object
    Dim files As Collection
    Dim filePath As Variant
    Dim rows As Collection
    Dim rowValues As Variant
    Dim rowDict As Object
    Dim eventType As String
    Dim payloadJson As String
    Dim report As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set GetLocalStagedBoxVersionInventoryDeltas = result
    If packageRow <= 0 Then Exit Function

    Set files = New Collection
    CollectLocalStagingFilesRole LocalStagingSearchRootRole(warehouseId, stationId), files
    For Each filePath In files
        report = vbNullString
        Set rows = ReadStagedInboxRowsRole(CStr(filePath), report)
        If rows Is Nothing Then GoTo NextFile
        For Each rowValues In rows
            Set rowDict = rowValues
            eventType = UCase$(Trim$(CStr(rowDict("EventType"))))
            If eventType = ROLE_EVENT_TYPE_SHIP Or eventType = ROLE_EVENT_TYPE_BOX_BUILD Or eventType = ROLE_EVENT_TYPE_BOX_UNBOX Then
                payloadJson = CStr(rowDict("PayloadJson"))
                AccumulateBoxPayloadVersionInventoryDeltasRole payloadJson, eventType, packageRow, result
            End If
        Next rowValues
NextFile:
    Next filePath
    Exit Function

FailSoft:
    If GetLocalStagedBoxVersionInventoryDeltas Is Nothing Then
        Set GetLocalStagedBoxVersionInventoryDeltas = CreateObject("Scripting.Dictionary")
        GetLocalStagedBoxVersionInventoryDeltas.CompareMode = vbTextCompare
    End If
End Function

Private Function BuildInboxRowValuesRole(ByVal eventId As String, _
                                         ByVal parentEventId As String, _
                                         ByVal undoOfEventId As String, _
                                         ByVal eventType As String, _
                                         ByVal createdAtUtc As Date, _
                                         ByVal warehouseId As String, _
                                         ByVal stationId As String, _
                                         ByVal userId As String, _
                                         ByVal migrationSourceId As String, _
                                         ByVal sku As String, _
                                         ByVal qty As Double, _
                                         ByVal location As String, _
                                         ByVal noteVal As String, _
                                         ByVal payloadJson As String) As Object
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    d("EventID") = eventId
    d("ParentEventId") = parentEventId
    d("UndoOfEventId") = undoOfEventId
    d("EventType") = UCase$(Trim$(eventType))
    d("CreatedAtUTC") = Format$(createdAtUtc, "yyyy-mm-dd hh:nn:ss")
    d("WarehouseId") = warehouseId
    d("StationId") = stationId
    d("UserId") = userId
    d("MigrationSourceId") = migrationSourceId
    d("SKU") = sku
    If qty <> 0 Then
        d("Qty") = qty
    Else
        d("Qty") = ""
    End If
    d("Location") = location
    d("Note") = noteVal
    d("PayloadJson") = payloadJson
    d("Status") = "NEW"
    d("RetryCount") = 0
    d("ErrorCode") = ""
    d("ErrorMessage") = ""
    d("FailedAtUTC") = ""
    Set BuildInboxRowValuesRole = d
End Function

Private Function ShouldStageEventLocallyRole(ByVal eventType As String) As Boolean
    Select Case UCase$(Trim$(eventType))
        Case ROLE_EVENT_TYPE_SHIP, ROLE_EVENT_TYPE_BOX_BUILD, ROLE_EVENT_TYPE_BOX_UNBOX
            ShouldStageEventLocallyRole = True
    End Select
End Function

Private Function AppendInboxRowToLocalStagingRole(ByVal rowValues As Object, _
                                                  ByVal stagingPath As String, _
                                                  ByRef errorMessage As String) As Boolean
    On Error GoTo FailAppend

    Dim fileNum As Integer

    EnsureFolderExistsRole ParentFolderPathRole(stagingPath)
    fileNum = FreeFile
    Open stagingPath For Append As #fileNum
    Print #fileNum, DictionaryToJson(rowValues)
    Close #fileNum
    AppendInboxRowToLocalStagingRole = True
    Exit Function

FailAppend:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    errorMessage = "Local staging append failed: " & Err.Description
End Function

Private Function MergeLocalStagingFileRole(ByVal stagingPath As String, _
                                           ByRef mergedCount As Long, _
                                           ByRef report As String) As Boolean
    On Error GoTo FailMerge

    Dim workingPath As String
    Dim sourcePath As String
    Dim rows As Collection
    Dim rowValues As Object
    Dim eventType As String
    Dim warehouseId As String
    Dim stationId As String
    Dim attempt As Long
    Dim attemptReport As String
    Dim lockMessage As Boolean

    sourcePath = stagingPath
    workingPath = SyncingStagingPathRole(stagingPath)

    If Not FileExistsRole(sourcePath) Then
        MergeLocalStagingFileRole = True
        Exit Function
    End If

    If StrComp(sourcePath, workingPath, vbTextCompare) <> 0 Then
        If FileExistsRole(workingPath) Then
            report = "Local staging sync deferred because a prior .syncing file exists: " & workingPath
            Exit Function
        End If
        Name sourcePath As workingPath
    End If

    Set rows = ReadStagedInboxRowsRole(workingPath, report)
    If rows Is Nothing Then GoTo RestoreWorking
    If rows.Count = 0 Then
        ArchiveStagingFileRole workingPath
        MergeLocalStagingFileRole = True
        Exit Function
    End If

    Set rowValues = rows(1)
    eventType = Trim$(CStr(rowValues("EventType")))
    warehouseId = Trim$(CStr(rowValues("WarehouseId")))
    stationId = Trim$(CStr(rowValues("StationId")))
    If eventType = "" Or warehouseId = "" Or stationId = "" Then
        report = "Local staged inbox file is missing EventType/WarehouseId/StationId: " & workingPath
        GoTo RestoreWorking
    End If
    If Not ShouldStageEventLocallyRole(eventType) Then
        ArchiveStagingFileRole workingPath
        report = "Archived unsupported local staged inbox event type '" & eventType & "': " & workingPath
        MergeLocalStagingFileRole = True
        Exit Function
    End If

    For attempt = 1 To 3
        attemptReport = vbNullString
        If MergeRowsIntoNasInboxRole(rows, eventType, warehouseId, stationId, mergedCount, attemptReport) Then
            ArchiveStagingFileRole workingPath
            MergeLocalStagingFileRole = True
            Exit Function
        End If

        lockMessage = IsLockContentionMessageRole(attemptReport)
        If Not lockMessage Then Exit For
        If attempt < 3 Then WaitSecondsRole 1
    Next attempt

    report = "Local staging sync could not merge " & workingPath & ": " & attemptReport

RestoreWorking:
    RestoreWorkingStagingFileRole workingPath
    Exit Function

FailMerge:
    report = "Local staging merge failed for " & stagingPath & ": " & Err.Description
    On Error Resume Next
    If workingPath <> "" Then RestoreWorkingStagingFileRole workingPath
    On Error GoTo 0
End Function

Private Function MergeRowsIntoNasInboxRole(ByVal rows As Collection, _
                                           ByVal eventType As String, _
                                           ByVal warehouseId As String, _
                                           ByVal stationId As String, _
                                           ByRef mergedCount As Long, _
                                           ByRef report As String) As Boolean
    On Error GoTo FailMergeRows

    Dim wbInbox As Workbook
    Dim lo As ListObject
    Dim rowValues As Variant
    Dim rowDict As Object
    Dim openPaths As Object
    Dim openedTransient As Boolean
    Dim schemaReport As String
    Dim ws As Worksheet
    Dim sheetWasProtected As Boolean

    Set openPaths = CaptureOpenWorkbookPathsRole()
    Set wbInbox = ResolveInboxWorkbookForEventType(eventType, warehouseId, stationId, report)
    If wbInbox Is Nothing Then Exit Function
    If wbInbox.ReadOnly Then
        report = "Inbox workbook is read-only or locked by another Excel session."
        GoTo CleanExit
    End If
    openedTransient = Not WorkbookWasAlreadyOpenRole(openPaths, wbInbox)
    If openedTransient Then HideWorkbookWindowsRole wbInbox

    Select Case UCase$(Trim$(eventType))
        Case ROLE_EVENT_TYPE_RECEIVE
            If Not modProcessor.EnsureReceiveInboxSchema(wbInbox, schemaReport) Then
                report = schemaReport
                GoTo CleanExit
            End If
        Case ROLE_EVENT_TYPE_SHIP, ROLE_EVENT_TYPE_BOX_BUILD, ROLE_EVENT_TYPE_BOX_UNBOX
            If Not modProcessor.EnsureShipInboxSchema(wbInbox, schemaReport) Then
                report = schemaReport
                GoTo CleanExit
            End If
        Case ROLE_EVENT_TYPE_PROD_CONSUME, ROLE_EVENT_TYPE_PROD_COMPLETE, ROLE_EVENT_TYPE_MIGRATION_SEED
            If Not modProcessor.EnsureProductionInboxSchema(wbInbox, schemaReport) Then
                report = schemaReport
                GoTo CleanExit
            End If
        Case Else
            report = "Unsupported event type '" & eventType & "'."
            GoTo CleanExit
    End Select

    Set lo = FindListObjectByNameRole(wbInbox, InboxTableNameRole(eventType))
    If lo Is Nothing Then
        report = "Inbox table not found for event type '" & eventType & "'."
        GoTo CleanExit
    End If

    Set ws = lo.Parent
    sheetWasProtected = ws.ProtectContents
    EnsureWorksheetEditableRole ws, lo.Name

    For Each rowValues In rows
        Set rowDict = rowValues
        If Not InboxContainsEventIdRole(lo, CStr(rowDict("EventID"))) Then
            WriteInboxRowValuesRole lo, rowDict
            mergedCount = mergedCount + 1
        End If
    Next rowValues

    SaveWorkbookRole wbInbox
    MergeRowsIntoNasInboxRole = True

CleanExit:
    On Error Resume Next
    If Not ws Is Nothing Then
        If sheetWasProtected Then RestoreWorksheetProtectionRole ws
    End If
    If openedTransient Then CloseTransientRoleWorkbook wbInbox
    On Error GoTo 0
    Exit Function

FailMergeRows:
    report = Err.Description
    Resume CleanExit
End Function

Private Sub WriteInboxRowValuesRole(ByVal lo As ListObject, ByVal rowValues As Object)
    Dim rowIndex As Long

    rowIndex = lo.ListRows.Add.Index
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "EventID"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "ParentEventId"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "UndoOfEventId"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "EventType"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "CreatedAtUTC"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "WarehouseId"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "StationId"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "UserId"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "MigrationSourceId"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "SKU"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "Qty"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "Location"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "Note"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "PayloadJson"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "Status"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "RetryCount"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "ErrorCode"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "ErrorMessage"
    SetTableRowValueFromDictionaryRole lo, rowIndex, rowValues, "FailedAtUTC"
End Sub

Private Sub SetTableRowValueFromDictionaryRole(ByVal lo As ListObject, _
                                               ByVal rowIndex As Long, _
                                               ByVal rowValues As Object, _
                                               ByVal columnName As String)
    If rowValues Is Nothing Then Exit Sub
    If Not rowValues.Exists(columnName) Then Exit Sub
    SetTableRowValueRole lo, rowIndex, columnName, rowValues(columnName)
End Sub

Private Function InboxContainsEventIdRole(ByVal lo As ListObject, ByVal eventId As String) As Boolean
    Dim rowIndex As Long

    eventId = Trim$(eventId)
    If eventId = "" Then Exit Function
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(GetTableRowValueRole(lo, rowIndex, "EventID"))), eventId, vbTextCompare) = 0 Then
            InboxContainsEventIdRole = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function ReadStagedInboxRowsRole(ByVal filePath As String, ByRef report As String) As Collection
    On Error GoTo FailRead

    Dim rows As Collection
    Dim fileNum As Integer
    Dim lineText As String
    Dim rowValues As Object
    Dim parseReport As String

    Set rows = New Collection
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    Do While Not EOF(fileNum)
        Line Input #fileNum, lineText
        lineText = Trim$(lineText)
        If lineText <> "" Then
            parseReport = vbNullString
            Set rowValues = ParseJsonObjectRole(lineText, parseReport)
            If rowValues Is Nothing Then
                report = "Unable to parse local staged inbox row in " & filePath & ": " & parseReport
                Close #fileNum
                Exit Function
            End If
            rows.Add rowValues
        End If
    Loop
    Close #fileNum
    Set ReadStagedInboxRowsRole = rows
    Exit Function

FailRead:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    report = "Unable to read local staged inbox file " & filePath & ": " & Err.Description
End Function

Private Sub AccumulateBoxPayloadInventoryDeltasRole(ByVal payloadJson As String, _
                                                   ByVal eventType As String, _
                                                   ByVal deltas As Object)
    On Error GoTo CleanExit

    Dim rx As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim objectText As String
    Dim ioType As String
    Dim rowValue As Long
    Dim qtyValue As Double
    Dim deltaValue As Double
    Dim key As String

    If deltas Is Nothing Then Exit Sub
    payloadJson = Trim$(payloadJson)
    If payloadJson = "" Then Exit Sub

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "\{[^{}]*\}"
    Set matches = rx.Execute(payloadJson)

    For Each matchItem In matches
        objectText = CStr(matchItem.Value)
        ioType = UCase$(JsonObjectStringFieldRole(objectText, "IoType"))
        If eventType = ROLE_EVENT_TYPE_SHIP Then
            If ioType <> "" And ioType <> "SHIPPED" Then GoTo NextObject
        ElseIf ioType <> "MADE" And ioType <> "UNMADE" Then
            GoTo NextObject
        End If

        rowValue = CLng(JsonObjectNumberFieldRole(objectText, "ROW"))
        If rowValue <= 0 Then rowValue = CLng(JsonObjectNumberFieldRole(objectText, "Row"))
        If rowValue <= 0 Then GoTo NextObject

        qtyValue = JsonObjectNumberFieldRole(objectText, "Qty")
        If qtyValue <= 0 Then GoTo NextObject

        If eventType = ROLE_EVENT_TYPE_SHIP Then
            deltaValue = -qtyValue
        ElseIf eventType = ROLE_EVENT_TYPE_BOX_UNBOX Or ioType = "UNMADE" Then
            deltaValue = -qtyValue
        Else
            deltaValue = qtyValue
        End If

        key = CStr(rowValue)
        If deltas.Exists(key) Then
            deltas(key) = CDbl(deltas(key)) + deltaValue
        Else
            deltas(key) = deltaValue
        End If
NextObject:
    Next matchItem

CleanExit:
End Sub

Private Sub AccumulateBoxPayloadVersionInventoryDeltasRole(ByVal payloadJson As String, _
                                                          ByVal eventType As String, _
                                                          ByVal packageRow As Long, _
                                                          ByVal deltas As Object)
    On Error GoTo CleanExit

    Dim rx As Object
    Dim matches As Object
    Dim matchItem As Object
    Dim objectText As String
    Dim ioType As String
    Dim rowValue As Long
    Dim qtyValue As Double
    Dim deltaValue As Double
    Dim versionLabel As String

    If deltas Is Nothing Then Exit Sub
    If packageRow <= 0 Then Exit Sub
    payloadJson = Trim$(payloadJson)
    If payloadJson = "" Then Exit Sub

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "\{[^{}]*\}"
    Set matches = rx.Execute(payloadJson)

    For Each matchItem In matches
        objectText = CStr(matchItem.Value)
        ioType = UCase$(JsonObjectStringFieldRole(objectText, "IoType"))
        If eventType = ROLE_EVENT_TYPE_SHIP Then
            If ioType <> "" And ioType <> "SHIPPED" Then GoTo NextObject
        ElseIf ioType <> "MADE" And ioType <> "UNMADE" Then
            GoTo NextObject
        End If

        rowValue = CLng(JsonObjectNumberFieldRole(objectText, "ROW"))
        If rowValue <= 0 Then rowValue = CLng(JsonObjectNumberFieldRole(objectText, "Row"))
        If rowValue <> packageRow Then GoTo NextObject

        versionLabel = NormalizeVersionLabelRole(JsonObjectStringFieldRole(objectText, "BomVersionLabel"))
        If versionLabel = "" Then versionLabel = NormalizeVersionLabelRole(JsonObjectStringFieldRole(objectText, "Version"))
        If versionLabel = "" Then GoTo NextObject

        qtyValue = JsonObjectNumberFieldRole(objectText, "Qty")
        If qtyValue <= 0 Then GoTo NextObject

        If eventType = ROLE_EVENT_TYPE_SHIP Then
            deltaValue = -qtyValue
        ElseIf eventType = ROLE_EVENT_TYPE_BOX_UNBOX Or ioType = "UNMADE" Then
            deltaValue = -qtyValue
        Else
            deltaValue = qtyValue
        End If

        If deltas.Exists(versionLabel) Then
            deltas(versionLabel) = CDbl(deltas(versionLabel)) + deltaValue
        Else
            deltas(versionLabel) = deltaValue
        End If
NextObject:
    Next matchItem

CleanExit:
End Sub

Private Function NormalizeVersionLabelRole(ByVal versionText As String) As String
    versionText = LCase$(Trim$(versionText))
    If versionText = "" Then Exit Function
    If Left$(versionText, 1) = "v" Then versionText = Mid$(versionText, 2)
    If IsNumeric(versionText) Then
        NormalizeVersionLabelRole = "v" & CStr(CLng(versionText))
    Else
        NormalizeVersionLabelRole = "v" & versionText
    End If
End Function

Private Function JsonObjectStringFieldRole(ByVal objectText As String, ByVal fieldName As String) As String
    On Error GoTo CleanExit

    Dim rx As Object
    Dim matches As Object

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = """" & EscapeRegexRole(fieldName) & """\s*:\s*""([^""]*)"""
    Set matches = rx.Execute(objectText)
    If matches.Count > 0 Then JsonObjectStringFieldRole = JsonUnescapeRole(CStr(matches(0).SubMatches(0)))

CleanExit:
End Function

Private Function JsonObjectNumberFieldRole(ByVal objectText As String, ByVal fieldName As String) As Double
    On Error GoTo CleanExit

    Dim rx As Object
    Dim matches As Object
    Dim rawValue As String

    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False
    rx.IgnoreCase = True
    rx.Pattern = """" & EscapeRegexRole(fieldName) & """\s*:\s*(""?-?[0-9]+(?:\.[0-9]+)?""?)"
    Set matches = rx.Execute(objectText)
    If matches.Count > 0 Then
        rawValue = Replace$(CStr(matches(0).SubMatches(0)), """", "")
        If IsNumeric(rawValue) Then JsonObjectNumberFieldRole = CDbl(rawValue)
    End If

CleanExit:
End Function

Private Function JsonUnescapeRole(ByVal textIn As String) As String
    JsonUnescapeRole = textIn
    JsonUnescapeRole = Replace$(JsonUnescapeRole, "\t", vbTab)
    JsonUnescapeRole = Replace$(JsonUnescapeRole, "\n", vbLf)
    JsonUnescapeRole = Replace$(JsonUnescapeRole, "\r", vbCr)
    JsonUnescapeRole = Replace$(JsonUnescapeRole, "\" & Chr$(34), Chr$(34))
    JsonUnescapeRole = Replace$(JsonUnescapeRole, "\\", "\")
End Function

Private Function EscapeRegexRole(ByVal textIn As String) As String
    Dim specials As Variant
    Dim token As Variant

    EscapeRegexRole = textIn
    specials = Array("\", ".", "+", "*", "?", "^", "$", "(", ")", "[", "]", "{", "}", "|")
    For Each token In specials
        EscapeRegexRole = Replace$(EscapeRegexRole, CStr(token), "\" & CStr(token))
    Next token
End Function

Private Function LocalStagingPathRole(ByVal eventType As String, _
                                      ByVal warehouseId As String, _
                                      ByVal stationId As String) As String
    Dim folderPath As String

    folderPath = CombinePathRole(LocalStagingRootRole(), SafePathTokenRole(warehouseId))
    folderPath = CombinePathRole(folderPath, SafePathTokenRole(stationId))
    LocalStagingPathRole = CombinePathRole(folderPath, InboxWorkbookNameRole(eventType, stationId) & ".staging.jsonl")
End Function

Private Function LocalStagingRootRole() As String
    Dim rootPath As String

    rootPath = Trim$(Environ$("LOCALAPPDATA"))
    If rootPath = "" Then rootPath = Trim$(Environ$("TEMP"))
    If rootPath = "" Then rootPath = ThisWorkbook.Path
    LocalStagingRootRole = CombinePathRole(rootPath, LOCAL_STAGING_ROOT_FOLDER)
End Function

Private Function LocalStagingSearchRootRole(ByVal warehouseId As String, ByVal stationId As String) As String
    Dim target As WarehouseTarget

    warehouseId = Trim$(warehouseId)
    stationId = Trim$(stationId)

    If warehouseId = "" Or stationId = "" Then
        On Error Resume Next
        Set target = modNasConnection.GetCurrentTarget()
        If warehouseId = "" And Not target Is Nothing Then warehouseId = Trim$(target.WarehouseId)
        If stationId = "" And Not target Is Nothing Then stationId = Trim$(target.StationId)
        If (warehouseId = "" Or stationId = "") And modConfig.LoadConfig("", "") Then
            If warehouseId = "" Then warehouseId = Trim$(modConfig.GetWarehouseId())
            If stationId = "" Then stationId = Trim$(modConfig.GetStationId())
        End If
        On Error GoTo 0
    End If

    If warehouseId <> "" And stationId <> "" Then
        LocalStagingSearchRootRole = CombinePathRole(LocalStagingRootRole(), SafePathTokenRole(warehouseId))
        LocalStagingSearchRootRole = CombinePathRole(LocalStagingSearchRootRole, SafePathTokenRole(stationId))
    Else
        LocalStagingSearchRootRole = LocalStagingRootRole()
    End If
End Function

Private Function SafePathTokenRole(ByVal tokenText As String) As String
    Dim badChars As Variant
    Dim ch As Variant

    SafePathTokenRole = Trim$(tokenText)
    badChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    For Each ch In badChars
        SafePathTokenRole = Replace$(SafePathTokenRole, CStr(ch), "_")
    Next ch
    If SafePathTokenRole = "" Then SafePathTokenRole = "_"
End Function

Private Function ParentFolderPathRole(ByVal filePath As String) As String
    Dim pos As Long

    filePath = Trim$(Replace$(filePath, "/", "\"))
    pos = InStrRev(filePath, "\")
    If pos > 0 Then ParentFolderPathRole = Left$(filePath, pos - 1)
End Function

Private Function FileExistsRole(ByVal filePath As String) As Boolean
    Dim fso As Object

    filePath = Trim$(filePath)
    If filePath = "" Then Exit Function
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsRole = fso.FileExists(filePath)
    If Err.Number <> 0 Then
        Err.Clear
        FileExistsRole = (Len(Dir$(filePath, vbNormal)) > 0)
    End If
    On Error GoTo 0
End Function

Private Sub CollectLocalStagingFilesRole(ByVal folderPath As String, ByVal files As Object)
    On Error GoTo CleanExit

    Dim fso As Object
    Dim folder As Object
    Dim fileItem As Object
    Dim subFolder As Object
    Dim fileName As String

    If files Is Nothing Then Exit Sub
    If Not FolderExistsRole(folderPath) Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    For Each fileItem In folder.Files
        fileName = LCase$(CStr(fileItem.Name))
        If Right$(fileName, Len(".staging.jsonl")) = ".staging.jsonl" _
           Or Right$(fileName, Len(".syncing")) = ".syncing" Then
            files.Add CStr(fileItem.Path)
        End If
    Next fileItem

    For Each subFolder In folder.SubFolders
        If StrComp(CStr(subFolder.Name), LOCAL_STAGING_ARCHIVE_FOLDER, vbTextCompare) <> 0 Then
            CollectLocalStagingFilesRole CStr(subFolder.Path), files
        End If
    Next subFolder

CleanExit:
End Sub

Private Function SyncingStagingPathRole(ByVal stagingPath As String) As String
    If Right$(LCase$(stagingPath), Len(".syncing")) = ".syncing" Then
        SyncingStagingPathRole = stagingPath
    Else
        SyncingStagingPathRole = stagingPath & ".syncing"
    End If
End Function

Private Sub RestoreWorkingStagingFileRole(ByVal workingPath As String)
    Dim originalPath As String

    If workingPath = "" Then Exit Sub
    If Not FileExistsRole(workingPath) Then Exit Sub
    If Right$(LCase$(workingPath), Len(".syncing")) <> ".syncing" Then Exit Sub
    originalPath = Left$(workingPath, Len(workingPath) - Len(".syncing"))
    If FileExistsRole(originalPath) Then Exit Sub
    Name workingPath As originalPath
End Sub

Private Sub ArchiveStagingFileRole(ByVal workingPath As String)
    On Error GoTo FailArchive

    Dim archiveFolder As String
    Dim archivePath As String
    Dim fso As Object

    If Not FileExistsRole(workingPath) Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    archiveFolder = CombinePathRole(ParentFolderPathRole(workingPath), LOCAL_STAGING_ARCHIVE_FOLDER)
    EnsureFolderExistsRole archiveFolder
    archivePath = CombinePathRole(archiveFolder, Format$(Now, "yyyymmdd_hhnnss") & "_" & fso.GetFileName(workingPath) & ".done")
    Do While FileExistsRole(archivePath)
        archivePath = CombinePathRole(archiveFolder, Format$(Now, "yyyymmdd_hhnnss") & "_" & CreateGuidFallbackRole() & "_" & fso.GetFileName(workingPath) & ".done")
    Loop
    Name workingPath As archivePath
    Exit Sub

FailArchive:
    ' Leave the .syncing file in place so a later run can retry or inspect it.
End Sub

Private Function IsLockContentionMessageRole(ByVal messageText As String) As Boolean
    messageText = LCase$(Trim$(messageText))
    IsLockContentionMessageRole = _
        (InStr(1, messageText, "read-only", vbTextCompare) > 0) Or _
        (InStr(1, messageText, "locked", vbTextCompare) > 0) Or _
        (InStr(1, messageText, "permission denied", vbTextCompare) > 0) Or _
        (InStr(1, messageText, "sharing violation", vbTextCompare) > 0)
End Function

Private Sub WaitSecondsRole(ByVal secondsToWait As Long)
    If secondsToWait <= 0 Then Exit Sub
    Application.Wait Now + TimeSerial(0, 0, secondsToWait)
End Sub

Private Function ParseJsonObjectRole(ByVal jsonText As String, ByRef report As String) As Object
    On Error GoTo FailParse

    Dim pos As Long
    Dim keyText As String
    Dim valueText As Variant
    Dim d As Object

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    pos = 1
    SkipJsonWhitespaceRole jsonText, pos
    If Mid$(jsonText, pos, 1) <> "{" Then
        report = "Expected object start."
        Exit Function
    End If
    pos = pos + 1

    Do
        SkipJsonWhitespaceRole jsonText, pos
        If Mid$(jsonText, pos, 1) = "}" Then
            pos = pos + 1
            Exit Do
        End If
        keyText = ParseJsonStringRole(jsonText, pos, report)
        If report <> "" Then Exit Function
        SkipJsonWhitespaceRole jsonText, pos
        If Mid$(jsonText, pos, 1) <> ":" Then
            report = "Expected ':' after key."
            Exit Function
        End If
        pos = pos + 1
        valueText = ParseJsonValueRole(jsonText, pos, report)
        If report <> "" Then Exit Function
        d(keyText) = valueText
        SkipJsonWhitespaceRole jsonText, pos
        Select Case Mid$(jsonText, pos, 1)
            Case ","
                pos = pos + 1
            Case "}"
                pos = pos + 1
                Exit Do
            Case Else
                report = "Expected ',' or object end."
                Exit Function
        End Select
    Loop While pos <= Len(jsonText)

    Set ParseJsonObjectRole = d
    Exit Function

FailParse:
    report = Err.Description
End Function

Private Function ParseJsonValueRole(ByVal jsonText As String, ByRef pos As Long, ByRef report As String) As Variant
    Dim startPos As Long
    Dim token As String

    SkipJsonWhitespaceRole jsonText, pos
    Select Case Mid$(jsonText, pos, 1)
        Case Chr$(34)
            ParseJsonValueRole = ParseJsonStringRole(jsonText, pos, report)
        Case Else
            startPos = pos
            Do While pos <= Len(jsonText)
                Select Case Mid$(jsonText, pos, 1)
                    Case ",", "}", " ", vbTab, vbCr, vbLf
                        Exit Do
                End Select
                pos = pos + 1
            Loop
            token = Trim$(Mid$(jsonText, startPos, pos - startPos))
            Select Case LCase$(token)
                Case "null"
                    ParseJsonValueRole = ""
                Case "true"
                    ParseJsonValueRole = True
                Case "false"
                    ParseJsonValueRole = False
                Case Else
                    If IsNumeric(token) Then
                        ParseJsonValueRole = CDbl(token)
                    Else
                        ParseJsonValueRole = token
                    End If
            End Select
    End Select
End Function

Private Function ParseJsonStringRole(ByVal jsonText As String, ByRef pos As Long, ByRef report As String) As String
    Dim ch As String
    Dim escaped As Boolean

    If Mid$(jsonText, pos, 1) <> Chr$(34) Then
        report = "Expected string."
        Exit Function
    End If
    pos = pos + 1
    Do While pos <= Len(jsonText)
        ch = Mid$(jsonText, pos, 1)
        pos = pos + 1
        If escaped Then
            Select Case ch
                Case Chr$(34), "\", "/"
                    ParseJsonStringRole = ParseJsonStringRole & ch
                Case "n"
                    ParseJsonStringRole = ParseJsonStringRole & vbLf
                Case "r"
                    ParseJsonStringRole = ParseJsonStringRole & vbCr
                Case "t"
                    ParseJsonStringRole = ParseJsonStringRole & vbTab
                Case Else
                    ParseJsonStringRole = ParseJsonStringRole & ch
            End Select
            escaped = False
        ElseIf ch = "\" Then
            escaped = True
        ElseIf ch = Chr$(34) Then
            Exit Function
        Else
            ParseJsonStringRole = ParseJsonStringRole & ch
        End If
    Loop
    report = "Unterminated string."
End Function

Private Sub SkipJsonWhitespaceRole(ByVal jsonText As String, ByRef pos As Long)
    Do While pos <= Len(jsonText)
        Select Case Mid$(jsonText, pos, 1)
            Case " ", vbTab, vbCr, vbLf
                pos = pos + 1
            Case Else
                Exit Do
        End Select
    Loop
End Sub

Private Function EnsureCurrentRoleWriteAllowed(ByVal requiredCapability As String, _
                                               ByVal requestedUserId As String, _
                                               ByRef resolvedUserId As String, _
                                               ByRef errorMessage As String) As Boolean
    Dim target As WarehouseTarget

    requiredCapability = UCase$(Trim$(requiredCapability))
    requestedUserId = Trim$(requestedUserId)
    If requiredCapability = "" Then
        errorMessage = "Unsupported event type."
        Exit Function
    End If

    If Not modNasConnection.IsCurrentTargetAllowed(True) Then
        errorMessage = "A connected NAS warehouse target is required before posting role events."
        Exit Function
    End If

    If Not modAuth.IsSignedIn() Then
        errorMessage = "Current invSys user is not signed in."
        Exit Function
    End If

    resolvedUserId = Trim$(modAuth.GetCurrentUserId())
    If resolvedUserId = "" Then
        errorMessage = "Current invSys user is not signed in."
        Exit Function
    End If
    If requestedUserId <> "" And StrComp(requestedUserId, resolvedUserId, vbTextCompare) <> 0 Then
        errorMessage = "Requested user does not match the signed-in invSys user."
        Exit Function
    End If

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then
        errorMessage = "A connected NAS warehouse target is required before posting role events."
        Exit Function
    End If
    If Not modAuth.CanPerform(requiredCapability, resolvedUserId, target.WarehouseId, target.StationId, "ROLE_UI", "") Then
        errorMessage = "Current user lacks " & requiredCapability & " capability." & vbCrLf & _
                       "User=" & ValueOrPlaceholderRole(resolvedUserId) & _
                       "; Warehouse=" & ValueOrPlaceholderRole(target.WarehouseId) & _
                       "; Station=" & ValueOrPlaceholderRole(target.StationId) & _
                       "; Auth=" & ValueOrPlaceholderRole(modAuth.GetResolvedAuthWorkbookName())
        Exit Function
    End If

    EnsureCurrentRoleWriteAllowed = True
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
            If wb.ReadOnly Then
                errorMessage = "Inbox workbook is read-only or locked by another Excel session."
                Exit Function
            End If
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
        Set ResolveInboxWorkbookForEventType = Application.Workbooks.Open( _
            Filename:=fullPath, _
            UpdateLinks:=0, _
            ReadOnly:=False, _
            IgnoreReadOnlyRecommended:=True, _
            Notify:=False, _
            AddToMru:=False)
        Application.ScreenUpdating = prevScreenUpdating
        Application.DisplayAlerts = prevAlerts
        alertsSuppressed = False
        If ResolveInboxWorkbookForEventType Is Nothing Then
            errorMessage = "Inbox workbook open failed."
            Exit Function
        End If
        If ResolveInboxWorkbookForEventType.ReadOnly Then
            errorMessage = "Inbox workbook is read-only or locked by another Excel session."
            ResolveInboxWorkbookForEventType.Close SaveChanges:=False
            Set ResolveInboxWorkbookForEventType = Nothing
            Exit Function
        End If
        HideWorkbookWindowsRole ResolveInboxWorkbookForEventType
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
    ExpandConfigPathRole = modConfig.NormalizeFolderPathForRuntime(ExpandConfigPathRole, False)
    If ExpandConfigPathRole = "" Then Exit Function
    Do While Right$(ExpandConfigPathRole, 1) = "\"
        ExpandConfigPathRole = Left$(ExpandConfigPathRole, Len(ExpandConfigPathRole) - 1)
    Loop
End Function

Private Function InboxWorkbookNameRole(ByVal eventType As String, ByVal stationId As String) As String
    Select Case UCase$(Trim$(eventType))
        Case ROLE_EVENT_TYPE_RECEIVE
            InboxWorkbookNameRole = "invSys.Inbox.Receiving." & stationId & ".xlsb"
        Case ROLE_EVENT_TYPE_SHIP, ROLE_EVENT_TYPE_BOX_BUILD, ROLE_EVENT_TYPE_BOX_UNBOX
            InboxWorkbookNameRole = "invSys.Inbox.Shipping." & stationId & ".xlsb"
        Case ROLE_EVENT_TYPE_PROD_CONSUME, ROLE_EVENT_TYPE_PROD_COMPLETE, ROLE_EVENT_TYPE_MIGRATION_SEED
            InboxWorkbookNameRole = "invSys.Inbox.Production." & stationId & ".xlsb"
    End Select
End Function

Private Function InboxTableNameRole(ByVal eventType As String) As String
    Select Case UCase$(Trim$(eventType))
        Case ROLE_EVENT_TYPE_RECEIVE
            InboxTableNameRole = TABLE_INBOX_RECEIVE
        Case ROLE_EVENT_TYPE_SHIP, ROLE_EVENT_TYPE_BOX_BUILD, ROLE_EVENT_TYPE_BOX_UNBOX
            InboxTableNameRole = TABLE_INBOX_SHIP
        Case ROLE_EVENT_TYPE_PROD_CONSUME, ROLE_EVENT_TYPE_PROD_COMPLETE, ROLE_EVENT_TYPE_MIGRATION_SEED
            InboxTableNameRole = TABLE_INBOX_PROD
    End Select
End Function

Private Function CapabilityForEventTypeRole(ByVal eventType As String) As String
    Select Case UCase$(Trim$(eventType))
        Case ROLE_EVENT_TYPE_RECEIVE
            CapabilityForEventTypeRole = "RECEIVE_POST"
        Case ROLE_EVENT_TYPE_SHIP, ROLE_EVENT_TYPE_BOX_BUILD, ROLE_EVENT_TYPE_BOX_UNBOX
            CapabilityForEventTypeRole = "SHIP_POST"
        Case ROLE_EVENT_TYPE_PROD_CONSUME, ROLE_EVENT_TYPE_PROD_COMPLETE
            CapabilityForEventTypeRole = "PROD_POST"
        Case ROLE_EVENT_TYPE_MIGRATION_SEED
            CapabilityForEventTypeRole = "ADMIN_MAINT"
    End Select
End Function

Private Sub EnsureFolderExistsRole(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long
    Dim fso As Object

    folderPath = NormalizeFolderPathRole(folderPath, False)
    If folderPath = "" Then Exit Sub
    If FolderExistsRole(folderPath) Then Exit Sub
    If IsUncShareRootRole(folderPath) Then Exit Sub

    sepPos = InStrRev(folderPath, "\")
    If sepPos > 1 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Not FolderExistsRole(parentPath) Then EnsureFolderExistsRole parentPath
    End If

    If FolderExistsRole(folderPath) Then Exit Sub
    If IsUncPathRole(folderPath) Then
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder folderPath
    Else
        MkDir folderPath
    End If
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

Private Function NormalizeFolderPathRole(ByVal folderPath As String, ByVal withTrailingSlash As Boolean) As String
    NormalizeFolderPathRole = modConfig.NormalizeFolderPathForRuntime(folderPath, withTrailingSlash)
End Function

Private Function FolderExistsRole(ByVal folderPath As String) As Boolean
    Dim fso As Object

    folderPath = NormalizeFolderPathRole(folderPath, False)
    If folderPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsRole = fso.FolderExists(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        FolderExistsRole = (Len(Dir$(folderPath, vbDirectory)) > 0)
    End If
    On Error GoTo 0
End Function

Private Function IsUncPathRole(ByVal folderPath As String) As Boolean
    folderPath = NormalizeFolderPathRole(folderPath, False)
    IsUncPathRole = (Left$(folderPath, 2) = "\\")
End Function

Private Function IsUncShareRootRole(ByVal folderPath As String) As Boolean
    Dim trimmedPath As String
    Dim parts() As String

    trimmedPath = NormalizeFolderPathRole(folderPath, False)
    If Left$(trimmedPath, 2) <> "\\" Then Exit Function

    trimmedPath = Mid$(trimmedPath, 3)
    If trimmedPath = "" Then Exit Function

    parts = Split(trimmedPath, "\")
    IsUncShareRootRole = (UBound(parts) = 1)
End Function

Private Sub SaveWorkbookAsXlsbRole(ByVal wb As Workbook, ByVal fullPath As String)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    If Application.CutCopyMode <> False Then Application.CutCopyMode = False
    On Error GoTo 0
    wb.SaveAs fullPath, 50
End Sub

Private Sub SaveWorkbookRole(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If wb.Path = "" Then Exit Sub
    On Error Resume Next
    If Application.CutCopyMode <> False Then Application.CutCopyMode = False
    On Error GoTo 0
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
    ReactivateQuietOwnerSafeRole
    On Error GoTo 0
End Sub

Private Sub ReactivateQuietOwnerSafeRole()
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!modUiQuiet.ReactivateQuietOwner"
    On Error GoTo 0
End Sub

Private Sub PerfBeginSafeRole(ByVal runId As String, ByVal activityName As String)
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!modPerfLog.PerfBegin", runId, activityName
    On Error GoTo 0
End Sub

Private Sub PerfMarkSafeRole(ByVal runId As String, ByVal segmentName As String, ByVal elapsedMs As Long)
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!modPerfLog.PerfMark", runId, segmentName, elapsedMs
    On Error GoTo 0
End Sub

Private Sub PerfEndSafeRole(ByVal runId As String, ByVal totalMs As Long, ByVal detailText As String)
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!modPerfLog.PerfEnd", runId, totalMs, detailText
    On Error GoTo 0
End Sub

Private Sub CloseTransientRoleWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    If Application.CutCopyMode <> False Then Application.CutCopyMode = False
    HideWorkbookWindowsRole wb
    If Not wb.ReadOnly Then
        If wb.Saved = False Then wb.Save
    End If
    If Application.CutCopyMode <> False Then Application.CutCopyMode = False
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

Private Function GetTableRowValueRole(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim idx As Long

    idx = GetColumnIndexRole(lo, columnName)
    If idx = 0 Then Exit Function
    GetTableRowValueRole = lo.DataBodyRange.Cells(rowIndex, idx).Value
End Function

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

Private Function EventIdListedRole(ByVal eventId As String, ByVal eventIdsCsv As String) As Boolean
    Dim normalizedIds As String

    eventId = Trim$(eventId)
    eventIdsCsv = Replace$(Trim$(eventIdsCsv), " ", "")
    If eventId = "" Or eventIdsCsv = "" Then Exit Function

    normalizedIds = "," & eventIdsCsv & ","
    EventIdListedRole = (InStr(1, normalizedIds, "," & eventId & ",", vbTextCompare) > 0)
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
