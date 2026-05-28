Attribute VB_Name = "modNasConnection"
Option Explicit

Public Enum NasStatusCode
    NAS_OK = 0
    NAS_ROOT_UNREACHABLE = 1
    NAS_ROOT_NO_CONFIG = 2
    NAS_CREDENTIAL_REJECTED = 3
    NAS_ROOT_NOT_IN_SESSION = 4
    WH_RUNTIME_NOT_FOUND = 5
    WH_CONFIG_INVALID = 6
    WH_AUTH_NOT_FOUND = 7
    WH_TARGET_INCOMPLETE = 8
    NAS_TARGET_UNREACHABLE = 9
    WH_NO_TARGET = 10
End Enum

Public Enum WH_SourceType
    WH_SOURCE_NAS = 0
    WH_SOURCE_REMEMBERED = 1
    WH_SOURCE_LOCAL = 2
    WH_SOURCE_FALLBACK = 3
End Enum

#If VBA7 Then
Private Declare PtrSafe Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
#Else
Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
#End If

Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Private Const RESOURCETYPE_DISK As Long = 1
Private Const CONNECT_UPDATE_PROFILE As Long = &H1
Private Const ERROR_ACCESS_DENIED As Long = 5
Private Const ERROR_ALREADY_ASSIGNED As Long = 85
Private Const ERROR_SESSION_CREDENTIAL_CONFLICT As Long = 1219
Private Const ERROR_LOGON_FAILURE As Long = 1326

Private Const SETTINGS_APP As String = "invSys"
Private Const SETTINGS_SECTION_NAS As String = "NAS"
Private Const SETTINGS_SECTION_RUNTIME As String = "Runtime"
Private Const SETTINGS_REMEMBERED_ROOTS As String = "RememberedRoots"
Private Const SETTINGS_REMEMBERED_TARGET As String = "RememberedWarehouseTarget"
Private Const SETTINGS_SELECTED_WAREHOUSE_TARGET As String = "SelectedWarehouseTarget"
Private Const SETTINGS_DELIM As String = "|"
Private Const TARGET_FIELD_DELIM As String = vbTab

Private m_CurrentTarget As WarehouseTarget
Private m_StaleTarget As WarehouseTarget
Private m_SessionRoots As Object
Private m_LastStatus As NasStatusCode
Private m_LastStatusText As String

Public Function ConnectNasRootWithCredentials(ByVal rootPath As String, _
                                              ByVal userName As String, _
                                              ByVal windowsPassword As String) As NasStatusCode
    Dim shareRoot As String
    Dim resultCode As Long

    rootPath = NormalizeFolderNas(rootPath)
    shareRoot = ResolveShareRootNas(rootPath)
    If shareRoot = "" Then
        SetStatusNas NAS_ROOT_UNREACHABLE, "NAS root is not a UNC path."
        ConnectNasRootWithCredentials = NAS_ROOT_UNREACHABLE
        Exit Function
    End If
    If Trim$(userName) = "" Or Len(windowsPassword) = 0 Then
        SetStatusNas NAS_CREDENTIAL_REJECTED, "NAS username or password is blank."
        ConnectNasRootWithCredentials = NAS_CREDENTIAL_REJECTED
        Exit Function
    End If

    resultCode = ConnectShareNas(shareRoot, Trim$(userName), windowsPassword, True)
    Select Case resultCode
        Case 0, ERROR_ALREADY_ASSIGNED
            RememberRoot shareRoot
            RememberRoot rootPath
            SaveRememberedNasUserNas shareRoot, Trim$(userName)
            SetStatusNas NAS_OK, "Connected to NAS root."
            ConnectNasRootWithCredentials = NAS_OK
        Case ERROR_ACCESS_DENIED, ERROR_LOGON_FAILURE
            SetStatusNas NAS_CREDENTIAL_REJECTED, "NAS credential rejected."
            ConnectNasRootWithCredentials = NAS_CREDENTIAL_REJECTED
        Case ERROR_SESSION_CREDENTIAL_CONFLICT
            SetStatusNas NAS_CREDENTIAL_REJECTED, "Windows already has a conflicting NAS session."
            ConnectNasRootWithCredentials = NAS_CREDENTIAL_REJECTED
        Case Else
            SetStatusNas NAS_TARGET_UNREACHABLE, "NAS connection failed. Windows error " & CStr(resultCode) & "."
            ConnectNasRootWithCredentials = NAS_TARGET_UNREACHABLE
    End Select
End Function

Public Function TryRevalidateRememberedRoot(ByVal rootPath As String) As NasStatusCode
    Dim shareRoot As String

    rootPath = NormalizeFolderNas(rootPath)
    If rootPath = "" Then
        SetStatusNas NAS_TARGET_UNREACHABLE, "Remembered NAS root is blank."
        TryRevalidateRememberedRoot = NAS_TARGET_UNREACHABLE
        Exit Function
    End If

    shareRoot = ResolveShareRootNas(rootPath)
    If shareRoot = "" Then
        If FolderExistsNas(rootPath) Then
            RememberRoot rootPath
            SetStatusNas NAS_OK, "Remembered local root is reachable."
            TryRevalidateRememberedRoot = NAS_OK
        Else
            SetStatusNas NAS_TARGET_UNREACHABLE, "Remembered root is unreachable."
            TryRevalidateRememberedRoot = NAS_TARGET_UNREACHABLE
        End If
        Exit Function
    End If

    If FolderExistsNas(rootPath) Then
        RememberRoot shareRoot
        RememberRoot rootPath
        SetStatusNas NAS_OK, "Remembered NAS root is reachable."
        TryRevalidateRememberedRoot = NAS_OK
    Else
        Select Case Err.Number
            Case 5, 1326
                SetStatusNas NAS_CREDENTIAL_REJECTED, "Remembered NAS credential rejected."
                TryRevalidateRememberedRoot = NAS_CREDENTIAL_REJECTED
            Case Else
                SetStatusNas NAS_TARGET_UNREACHABLE, "Remembered NAS root is unreachable."
                TryRevalidateRememberedRoot = NAS_TARGET_UNREACHABLE
        End Select
    End If
End Function

Public Sub ShowWarehouseConnectionPrompt(Optional ByVal reason As String = "")
    Dim frm As frmWarehouseConnection

    Set frm = New frmWarehouseConnection
    frm.InitializeConnectionPrompt reason
    frm.Show vbModal
    Unload frm
End Sub

Public Sub DisconnectNasRoot(ByVal rootPath As String, Optional ByVal disconnectWindowsSession As Boolean = False)
    Dim shareRoot As String

    rootPath = NormalizeFolderNas(rootPath)
    If rootPath = "" Then Exit Sub
    If Not m_CurrentTarget Is Nothing Then
        If SamePathNas(m_CurrentTarget.HubRoot, rootPath) Or SamePathNas(ResolveShareRootNas(m_CurrentTarget.HubRoot), ResolveShareRootNas(rootPath)) Then
            ClearWarehouseTarget
        End If
    End If

    EnsureSessionRootsNas
    If m_SessionRoots.Exists(rootPath) Then m_SessionRoots.Remove rootPath
    shareRoot = ResolveShareRootNas(rootPath)
    If shareRoot <> "" And m_SessionRoots.Exists(shareRoot) Then m_SessionRoots.Remove shareRoot
    If disconnectWindowsSession And shareRoot <> "" Then WNetCancelConnection2 shareRoot, 0, True
End Sub

Public Sub ForgetRoot(ByVal rootPath As String)
    rootPath = NormalizeFolderNas(rootPath)
    If rootPath = "" Then Exit Sub
    SaveRootsTextNas RemovePathFromListNas(GetRootsTextNas(), rootPath)
    If Not m_CurrentTarget Is Nothing Then
        If SamePathNas(m_CurrentTarget.HubRoot, rootPath) Then ForgetTarget m_CurrentTarget.WarehouseId
    End If
End Sub

Public Function ScanNasRoot(ByVal rootPath As String) As Collection
    Dim results As Collection
    Dim childName As String
    Dim childPath As String

    Set results = New Collection
    rootPath = NormalizeFolderNas(rootPath)
    If rootPath = "" Then
        Set ScanNasRoot = results
        Exit Function
    End If
    If Not FolderExistsNas(rootPath) Then
        Set ScanNasRoot = results
        Exit Function
    End If

    If RuntimeLooksCompleteNas(rootPath) Then AddPathIfMissingNas results, rootPath

    On Error GoTo CleanExit
    childName = Dir$(rootPath & "\*", vbDirectory)
    Do While childName <> ""
        If childName <> "." And childName <> ".." Then
            childPath = rootPath & "\" & childName
            If FolderExistsNas(childPath) Then
                If RuntimeLooksCompleteNas(childPath) Then AddPathIfMissingNas results, childPath
            End If
        End If
        childName = Dir$
    Loop

CleanExit:
    Set ScanNasRoot = results
End Function

Public Function SelectWarehouseTarget(ByVal hubRoot As String, _
                                      ByVal runtimeRoot As String, _
                                      ByRef outTarget As WarehouseTarget, _
                                      Optional ByVal stationId As String = "", _
                                      Optional ByVal requireStationInbox As Boolean = False) As NasStatusCode
    Dim configPath As String
    Dim authPath As String
    Dim inboxRoot As String
    Dim whId As String
    Dim whName As String
    Dim resolvedStation As String
    Dim statusCode As NasStatusCode

    Set outTarget = Nothing
    hubRoot = NormalizeFolderNas(hubRoot)
    runtimeRoot = NormalizeFolderNas(runtimeRoot)
    If hubRoot = "" Then hubRoot = runtimeRoot
    If runtimeRoot = "" Then runtimeRoot = hubRoot
    If requireStationInbox And (Trim$(stationId) = "" Or Trim$(stationId) = "*") Then
        SetStatusNas WH_TARGET_INCOMPLETE, "Station-scoped target requires a StationId."
        SelectWarehouseTarget = WH_TARGET_INCOMPLETE
        Exit Function
    End If
    If Not IsRootInSessionNas(hubRoot) And IsUncPathNas(hubRoot) Then
        SetStatusNas NAS_ROOT_NOT_IN_SESSION, "NAS root is not connected in this Excel session."
        SelectWarehouseTarget = NAS_ROOT_NOT_IN_SESSION
        Exit Function
    End If
    If Not FolderExistsNas(runtimeRoot) Then
        SetStatusNas WH_RUNTIME_NOT_FOUND, "Warehouse runtime folder is unreachable."
        SelectWarehouseTarget = WH_RUNTIME_NOT_FOUND
        Exit Function
    End If

    configPath = FindFirstWorkbookNas(runtimeRoot, "*.invsys.config.xls*")
    If configPath = "" Then
        SetStatusNas NAS_ROOT_NO_CONFIG, "Warehouse config workbook was not found."
        SelectWarehouseTarget = NAS_ROOT_NO_CONFIG
        Exit Function
    End If

    statusCode = ReadConfigIdentityNas(configPath, stationId, whId, whName, resolvedStation)
    If statusCode <> NAS_OK Then
        SetStatusNas statusCode, "Warehouse config workbook is invalid."
        SelectWarehouseTarget = statusCode
        Exit Function
    End If
    If requireStationInbox And (resolvedStation = "" Or resolvedStation = "*") Then
        SetStatusNas WH_TARGET_INCOMPLETE, "Station-scoped target requires a live StationId."
        SelectWarehouseTarget = WH_TARGET_INCOMPLETE
        Exit Function
    End If

    authPath = runtimeRoot & "\" & whId & ".invSys.Auth.xlsb"
    If Not WorkbookReadableNas(authPath) Then
        SetStatusNas WH_AUTH_NOT_FOUND, "Warehouse auth workbook was not readable."
        SelectWarehouseTarget = WH_AUTH_NOT_FOUND
        Exit Function
    End If

    inboxRoot = ResolveInboxRootNas(runtimeRoot, whId, resolvedStation)
    If requireStationInbox And inboxRoot = "" Then
        SetStatusNas WH_TARGET_INCOMPLETE, "Station inbox path could not be resolved."
        SelectWarehouseTarget = WH_TARGET_INCOMPLETE
        Exit Function
    End If

    Set outTarget = New WarehouseTarget
    With outTarget
        .WarehouseId = whId
        .WarehouseName = whName
        .StationId = resolvedStation
        .HubRoot = hubRoot
        .RuntimeRoot = runtimeRoot
        .ConfigPath = configPath
        .AuthPath = authPath
        .InboxRoot = inboxRoot
        .SourceType = WH_SOURCE_NAS
        .LastResolvedUTC = Now
    End With

    Set m_CurrentTarget = CloneTargetNas(outTarget)
    Set m_StaleTarget = Nothing
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    modConfig.LoadConfig whId, resolvedStation
    RememberTarget outTarget
    SetStatusNas NAS_OK, "Connected to " & TargetDisplayNas(outTarget) & "."
    SelectWarehouseTarget = NAS_OK
End Function

Public Function SelectWarehouseTargetForAutomation(ByVal hubRoot As String, _
                                                   ByVal runtimeRoot As String, _
                                                   Optional ByVal stationId As String = "", _
                                                   Optional ByVal requireStationInbox As Boolean = False) As String
    On Error GoTo FailSelect

    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode

    statusCode = SelectWarehouseTarget(hubRoot, runtimeRoot, target, stationId, requireStationInbox)
    If statusCode = NAS_OK Then
        SelectWarehouseTargetForAutomation = "OK|" & GetConnectionStatus()
    Else
        SelectWarehouseTargetForAutomation = "FAIL|" & CStr(statusCode) & "|" & GetConnectionStatus()
    End If
    Exit Function

FailSelect:
    SelectWarehouseTargetForAutomation = "FAIL|ERROR|" & Err.Description
End Function

Public Function ResolveWarehouseTarget(ByRef outTarget As WarehouseTarget, ByRef statusCode As NasStatusCode) As Boolean
    Dim remembered As WarehouseTarget
    Dim roots As Collection
    Dim rootPath As Variant
    Dim targets As Collection
    Dim runtimeRoot As Variant
    Dim candidate As WarehouseTarget

    Set outTarget = Nothing
    If Not m_CurrentTarget Is Nothing Then
        Set outTarget = CloneTargetNas(m_CurrentTarget)
        statusCode = NAS_OK
        ResolveWarehouseTarget = True
        Exit Function
    End If

    Set remembered = LoadRememberedTargetNas()
    If Not remembered Is Nothing Then
        If IsUncPathNas(remembered.HubRoot) Then
            If TryRevalidateRememberedRoot(remembered.HubRoot) <> NAS_OK Then
                Set m_StaleTarget = CloneTargetNas(remembered)
                Set outTarget = CloneTargetNas(remembered)
                statusCode = NAS_TARGET_UNREACHABLE
                SetStatusNas statusCode, "NAS unreachable - last known: " & TargetDisplayNas(remembered)
                Exit Function
            End If
        ElseIf Not FolderExistsNas(remembered.HubRoot) Then
            Set m_StaleTarget = CloneTargetNas(remembered)
            Set outTarget = CloneTargetNas(remembered)
            statusCode = NAS_TARGET_UNREACHABLE
            SetStatusNas statusCode, "Remembered target is unreachable."
            Exit Function
        End If
        statusCode = SelectWarehouseTarget(remembered.HubRoot, remembered.RuntimeRoot, candidate, remembered.StationId)
        If statusCode = NAS_OK Then
            candidate.SourceType = IIf(IsUncPathNas(candidate.HubRoot), WH_SOURCE_REMEMBERED, WH_SOURCE_LOCAL)
            Set m_CurrentTarget = CloneTargetNas(candidate)
            Set outTarget = CloneTargetNas(candidate)
            ResolveWarehouseTarget = True
            Exit Function
        End If
        Set m_StaleTarget = CloneTargetNas(remembered)
        Set outTarget = CloneTargetNas(remembered)
        statusCode = NAS_TARGET_UNREACHABLE
        SetStatusNas statusCode, "Remembered target could not be revalidated."
        Exit Function
    End If

    Set roots = GetRememberedRootsNas()
    For Each rootPath In roots
        If IsUncPathNas(CStr(rootPath)) Then
            If TryRevalidateRememberedRoot(CStr(rootPath)) <> NAS_OK Then
                statusCode = NAS_TARGET_UNREACHABLE
                SetStatusNas statusCode, "Remembered NAS root is unreachable."
                Exit Function
            End If
        ElseIf Not FolderExistsNas(CStr(rootPath)) Then
            statusCode = NAS_TARGET_UNREACHABLE
            SetStatusNas statusCode, "Remembered root is unreachable."
            Exit Function
        End If
        Set targets = ScanNasRoot(CStr(rootPath))
        For Each runtimeRoot In targets
            statusCode = SelectWarehouseTarget(CStr(rootPath), CStr(runtimeRoot), candidate)
            If statusCode = NAS_OK Then
                candidate.SourceType = IIf(IsUncPathNas(candidate.HubRoot), WH_SOURCE_REMEMBERED, WH_SOURCE_LOCAL)
                Set m_CurrentTarget = CloneTargetNas(candidate)
                Set outTarget = CloneTargetNas(candidate)
                ResolveWarehouseTarget = True
                Exit Function
            End If
        Next runtimeRoot
    Next rootPath

    If TryResolveOpenOrConfiguredTargetNas(candidate) Then
        Set outTarget = CloneTargetNas(candidate)
        statusCode = NAS_OK
        ResolveWarehouseTarget = True
        Exit Function
    End If

    If TryResolveFallbackTargetNas(candidate) Then
        Set outTarget = CloneTargetNas(candidate)
        statusCode = NAS_OK
        ResolveWarehouseTarget = True
        Exit Function
    End If

    statusCode = WH_NO_TARGET
    SetStatusNas statusCode, "No warehouse target selected."
End Function

Public Function EnsureWarehouseTargetInteractive(Optional ByVal reason As String = "", _
                                                Optional ByVal requireNasTarget As Boolean = False) As Boolean
    Dim target As WarehouseTarget
    Dim statusCode As NasStatusCode

    If ResolveWarehouseTarget(target, statusCode) Then
        If Not requireNasTarget Or target.SourceType <> WH_SOURCE_FALLBACK Then
            EnsureWarehouseTargetInteractive = True
            Exit Function
        End If
    End If

    ShowWarehouseConnectionPrompt reason
    If ResolveWarehouseTarget(target, statusCode) Then
        EnsureWarehouseTargetInteractive = (Not requireNasTarget Or target.SourceType <> WH_SOURCE_FALLBACK)
    End If
End Function

Public Sub ClearWarehouseTarget()
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!modAuth.SignOut"
    If Err.Number <> 0 Then
        Err.Clear
        Application.Run "modAuth.SignOut"
    End If
    On Error GoTo 0
    Set m_CurrentTarget = Nothing
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    SetStatusNas WH_NO_TARGET, "No warehouse target selected."
End Sub

Public Function IsTargetResolved() As Boolean
    IsTargetResolved = Not m_CurrentTarget Is Nothing
End Function

Public Function GetCurrentTarget() As WarehouseTarget
    If m_CurrentTarget Is Nothing Then Exit Function
    Set GetCurrentTarget = CloneTargetNas(m_CurrentTarget)
End Function

Public Function GetKnownWarehouseTargetRoots() As Collection
    Dim result As Collection
    Dim roots As Collection
    Dim rootPath As Variant
    Dim remembered As WarehouseTarget

    Set result = New Collection
    If Not m_CurrentTarget Is Nothing Then
        AddPathIfMissingNas result, m_CurrentTarget.HubRoot
        AddPathIfMissingNas result, m_CurrentTarget.RuntimeRoot
    End If

    Set remembered = LoadRememberedTargetNas()
    If Not remembered Is Nothing Then
        AddPathIfMissingNas result, remembered.HubRoot
        AddPathIfMissingNas result, remembered.RuntimeRoot
    End If

    Set roots = GetRememberedRootsNas()
    For Each rootPath In roots
        AddPathIfMissingNas result, CStr(rootPath)
    Next rootPath

    Set GetKnownWarehouseTargetRoots = result
End Function

Public Function ConnectKnownWarehouseServer(ByRef connectedRoot As String, _
                                           ByRef statusText As String) As Boolean
    Dim roots As Collection
    Dim knownRoots As Collection
    Dim rootPath As Variant
    Dim statusCode As NasStatusCode

    Set roots = New Collection
    AddPathIfMissingNas roots, ResolvePromptDefaultRootNas()

    Set knownRoots = GetKnownWarehouseTargetRoots()
    For Each rootPath In knownRoots
        AddPathIfMissingNas roots, CStr(rootPath)
    Next rootPath

    If roots.Count = 0 Then
        statusText = "No remembered NAS root is available. Use Admin > Add Warehouse Root or setup to save the server path."
        SetStatusNas WH_NO_TARGET, statusText
        Exit Function
    End If

    For Each rootPath In roots
        statusCode = TryRevalidateRememberedRoot(CStr(rootPath))
        If statusCode = NAS_OK Then
            connectedRoot = NormalizeFolderNas(CStr(rootPath))
            statusText = GetConnectionStatus()
            ConnectKnownWarehouseServer = True
            Exit Function
        End If
    Next rootPath

    statusText = GetConnectionStatus()
End Function

Public Function IsWarehouseTargetAllowed(ByVal target As WarehouseTarget, _
                                         Optional ByVal requireNasTarget As Boolean = False) As Boolean
    If target Is Nothing Then Exit Function
    If requireNasTarget And target.SourceType = WH_SOURCE_FALLBACK Then Exit Function
    IsWarehouseTargetAllowed = True
End Function

Public Function IsCurrentTargetAllowed(Optional ByVal requireNasTarget As Boolean = False) As Boolean
    IsCurrentTargetAllowed = IsWarehouseTargetAllowed(m_CurrentTarget, requireNasTarget)
End Function

Public Function SetCurrentTargetSourceTypeForTest(ByVal sourceType As WH_SourceType) As Boolean
    If m_CurrentTarget Is Nothing Then Exit Function
    m_CurrentTarget.SourceType = sourceType
    SetCurrentTargetSourceTypeForTest = True
End Function

Public Sub RememberTarget(ByVal target As WarehouseTarget)
    If target Is Nothing Then Exit Sub
    SaveSetting SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_REMEMBERED_TARGET, SerializeTargetNas(target)
    SaveSetting SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_SELECTED_WAREHOUSE_TARGET, _
                target.WarehouseId & SETTINGS_DELIM & target.StationId & SETTINGS_DELIM & target.RuntimeRoot
    RememberRoot target.HubRoot
End Sub

Public Sub ForgetTarget(ByVal warehouseId As String)
    Dim remembered As WarehouseTarget

    warehouseId = Trim$(warehouseId)
    If warehouseId = "" Then Exit Sub
    Set remembered = LoadRememberedTargetNas()
    If Not remembered Is Nothing Then
        If StrComp(remembered.WarehouseId, warehouseId, vbTextCompare) = 0 Then
            ForgetPersistedTargetNas
        End If
    End If
    If Not m_CurrentTarget Is Nothing Then
        If StrComp(m_CurrentTarget.WarehouseId, warehouseId, vbTextCompare) = 0 Then ClearWarehouseTarget
    End If
End Sub

Public Function IsConnected() As Boolean
    EnsureSessionRootsNas
    IsConnected = (m_SessionRoots.Count > 0)
End Function

Public Function GetConnectionStatus() As String
    If Not m_CurrentTarget Is Nothing Then
        If m_CurrentTarget.SourceType = WH_SOURCE_FALLBACK Then
            GetConnectionStatus = "Local fallback active - " & TargetDisplayNas(m_CurrentTarget)
        Else
            GetConnectionStatus = "Connected - " & TargetDisplayNas(m_CurrentTarget)
        End If
    ElseIf Not m_StaleTarget Is Nothing Then
        GetConnectionStatus = "NAS unreachable - last known: " & TargetDisplayNas(m_StaleTarget)
    ElseIf m_LastStatusText <> "" Then
        GetConnectionStatus = m_LastStatusText
    Else
        GetConnectionStatus = "No warehouse target selected"
    End If
End Function

Public Function GetPromptDefaultRoot() As String
    GetPromptDefaultRoot = ResolvePromptDefaultRootNas()
End Function

Public Function GetRememberedNasUser() As String
    GetRememberedNasUser = ResolveRememberedNasUserNas()
End Function

Private Function TryResolveOpenOrConfiguredTargetNas(ByRef outTarget As WarehouseTarget) As Boolean
    Dim rootPath As String
    Dim statusCode As NasStatusCode

    If Not modConfig.IsLoaded() Then Exit Function
    rootPath = NormalizeFolderNas(modConfig.GetString("PathDataRoot", ""))
    If rootPath = "" Then Exit Function
    RememberRoot rootPath
    statusCode = SelectWarehouseTarget(rootPath, rootPath, outTarget, modConfig.GetStationId())
    If statusCode = NAS_OK Then
        outTarget.SourceType = IIf(IsUncPathNas(outTarget.HubRoot), WH_SOURCE_NAS, WH_SOURCE_LOCAL)
        Set m_CurrentTarget = CloneTargetNas(outTarget)
        TryResolveOpenOrConfiguredTargetNas = True
    End If
End Function

Private Function TryResolveFallbackTargetNas(ByRef outTarget As WarehouseTarget) As Boolean
    Dim rootPath As String
    Dim statusCode As NasStatusCode

    rootPath = NormalizeFolderNas(modRuntimeWorkbooks.ResolveCoreDataRoot("", "WH1"))
    If rootPath = "" Then Exit Function
    RememberRoot rootPath
    statusCode = SelectWarehouseTarget(rootPath, rootPath, outTarget, "")
    If statusCode = NAS_OK Then
        outTarget.SourceType = WH_SOURCE_FALLBACK
        Set m_CurrentTarget = CloneTargetNas(outTarget)
        ForgetPersistedTargetNas
        SaveRootsTextNas RemovePathFromListNas(GetRootsTextNas(), rootPath)
        SetStatusNas NAS_OK, "Local fallback active - " & TargetDisplayNas(outTarget)
        TryResolveFallbackTargetNas = True
    End If
End Function

Private Function ReadConfigIdentityNas(ByVal configPath As String, _
                                       ByVal requestedStation As String, _
                                       ByRef warehouseId As String, _
                                       ByRef warehouseName As String, _
                                       ByRef stationId As String) As NasStatusCode
    Dim wb As Workbook
    Dim openedTransient As Boolean
    Dim loWh As ListObject
    Dim loSt As ListObject
    Dim rowIndex As Long

    On Error GoTo FailRead
    Set wb = FindOpenWorkbookNas(configPath)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(configPath, ReadOnly:=True, UpdateLinks:=False)
        openedTransient = True
    End If

    Set loWh = FindListObjectNas(wb, "tblWarehouseConfig")
    Set loSt = FindListObjectNas(wb, "tblStationConfig")
    If loWh Is Nothing Or loWh.DataBodyRange Is Nothing Then GoTo InvalidConfig
    warehouseId = TableValueNas(loWh, 1, "WarehouseId")
    warehouseName = TableValueNas(loWh, 1, "WarehouseName")
    If warehouseId = "" Then GoTo InvalidConfig
    If warehouseName = "" Then warehouseName = warehouseId

    stationId = Trim$(requestedStation)
    If stationId <> "" And Not loSt Is Nothing Then
        If Not loSt.DataBodyRange Is Nothing Then
            rowIndex = FindStationRowNas(loSt, warehouseId, stationId)
            If rowIndex = 0 Then rowIndex = 1
            stationId = TableValueNas(loSt, rowIndex, "StationId")
        End If
    End If

    ReadConfigIdentityNas = NAS_OK
    GoTo CleanExit

InvalidConfig:
    ReadConfigIdentityNas = WH_CONFIG_INVALID
    GoTo CleanExit

FailRead:
    ReadConfigIdentityNas = WH_CONFIG_INVALID

CleanExit:
    If openedTransient And Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
End Function

Private Function WorkbookReadableNas(ByVal workbookPath As String) As Boolean
    Dim wb As Workbook
    Dim openedTransient As Boolean

    On Error GoTo CleanFail
    If NormalizeFolderNas(workbookPath) = "" Then Exit Function
    Set wb = FindOpenWorkbookNas(workbookPath)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(workbookPath, ReadOnly:=True, UpdateLinks:=False)
        openedTransient = True
    End If
    WorkbookReadableNas = Not wb Is Nothing

CleanExit:
    If openedTransient And Not wb Is Nothing Then
        On Error Resume Next
        wb.Close SaveChanges:=False
        On Error GoTo 0
    End If
    Exit Function

CleanFail:
    WorkbookReadableNas = False
    Resume CleanExit
End Function

Private Function RuntimeLooksCompleteNas(ByVal runtimeRoot As String) As Boolean
    RuntimeLooksCompleteNas = (FindFirstWorkbookNas(runtimeRoot, "*.invsys.config.xls*") <> "") And _
                              (FindFirstWorkbookNas(runtimeRoot, "*.invsys.auth.xls*") <> "")
End Function

Private Function FindFirstWorkbookNas(ByVal folderPath As String, ByVal likePattern As String) As String
    Dim fileName As String

    On Error GoTo CleanFail
    folderPath = NormalizeFolderNas(folderPath)
    If folderPath = "" Then Exit Function
    fileName = Dir$(folderPath & "\" & likePattern, vbNormal)
    If fileName <> "" Then FindFirstWorkbookNas = folderPath & "\" & fileName
    Exit Function

CleanFail:
    FindFirstWorkbookNas = vbNullString
End Function

Private Function ResolveInboxRootNas(ByVal runtimeRoot As String, ByVal warehouseId As String, ByVal stationId As String) As String
    Dim report As String
    Dim inboxPath As String

    If Trim$(stationId) = "" Or Trim$(stationId) = "*" Then
        ResolveInboxRootNas = runtimeRoot
        Exit Function
    End If
    inboxPath = modConfig.ResolveStationInboxPath(warehouseId, stationId, "RECEIVE", runtimeRoot & "\" & warehouseId & ".invSys.Config.xlsb", report)
    If inboxPath <> "" Then
        ResolveInboxRootNas = modDeploymentPaths.GetParentFolderManaged(inboxPath)
    Else
        ResolveInboxRootNas = runtimeRoot
    End If
End Function

Private Function ConnectShareNas(ByVal shareRoot As String, _
                                 ByVal userName As String, _
                                 ByVal windowsPassword As String, _
                                 ByVal updateProfile As Boolean) As Long
    Dim resource As NETRESOURCE

    resource.dwType = RESOURCETYPE_DISK
    resource.lpRemoteName = shareRoot
    ConnectShareNas = WNetAddConnection2(resource, windowsPassword, userName, IIf(updateProfile, CONNECT_UPDATE_PROFILE, 0))
End Function

Private Sub RememberRoot(ByVal rootPath As String)
    Dim rootsText As String

    rootPath = NormalizeFolderNas(rootPath)
    If rootPath = "" Then Exit Sub
    EnsureSessionRootsNas
    If Not m_SessionRoots.Exists(rootPath) Then m_SessionRoots(rootPath) = True

    rootsText = AddPathToListNas(GetRootsTextNas(), rootPath)
    SaveRootsTextNas rootsText
End Sub

Private Function IsRootInSessionNas(ByVal rootPath As String) As Boolean
    Dim shareRoot As String

    EnsureSessionRootsNas
    rootPath = NormalizeFolderNas(rootPath)
    shareRoot = ResolveShareRootNas(rootPath)
    IsRootInSessionNas = m_SessionRoots.Exists(rootPath) Or (shareRoot <> "" And m_SessionRoots.Exists(shareRoot))
End Function

Private Function GetRememberedRootsNas() As Collection
    Dim roots As Collection
    Dim parts() As String
    Dim idx As Long
    Dim rootsText As String

    Set roots = New Collection
    rootsText = GetRootsTextNas()
    If rootsText <> "" Then
        parts = Split(rootsText, SETTINGS_DELIM)
        For idx = LBound(parts) To UBound(parts)
            AddPathIfMissingNas roots, CStr(parts(idx))
        Next idx
    End If
    Set GetRememberedRootsNas = roots
End Function

Private Function GetRootsTextNas() As String
    On Error Resume Next
    GetRootsTextNas = Trim$(GetSetting(SETTINGS_APP, SETTINGS_SECTION_NAS, SETTINGS_REMEMBERED_ROOTS, ""))
    On Error GoTo 0
End Function

Private Sub SaveRootsTextNas(ByVal rootsText As String)
    On Error Resume Next
    SaveSetting SETTINGS_APP, SETTINGS_SECTION_NAS, SETTINGS_REMEMBERED_ROOTS, rootsText
    On Error GoTo 0
End Sub

Private Function AddPathToListNas(ByVal rootsText As String, ByVal rootPath As String) As String
    Dim roots As Collection
    Dim item As Variant

    Set roots = New Collection
    If rootsText <> "" Then
        Dim parts() As String
        Dim idx As Long
        parts = Split(rootsText, SETTINGS_DELIM)
        For idx = LBound(parts) To UBound(parts)
            AddPathIfMissingNas roots, CStr(parts(idx))
        Next idx
    End If
    AddPathIfMissingNas roots, rootPath
    For Each item In roots
        If AddPathToListNas <> "" Then AddPathToListNas = AddPathToListNas & SETTINGS_DELIM
        AddPathToListNas = AddPathToListNas & CStr(item)
    Next item
End Function

Private Function RemovePathFromListNas(ByVal rootsText As String, ByVal rootPath As String) As String
    Dim parts() As String
    Dim idx As Long
    Dim item As String

    If rootsText = "" Then Exit Function
    parts = Split(rootsText, SETTINGS_DELIM)
    For idx = LBound(parts) To UBound(parts)
        item = NormalizeFolderNas(CStr(parts(idx)))
        If item <> "" And Not SamePathNas(item, rootPath) Then
            If RemovePathFromListNas <> "" Then RemovePathFromListNas = RemovePathFromListNas & SETTINGS_DELIM
            RemovePathFromListNas = RemovePathFromListNas & item
        End If
    Next idx
End Function

Private Function LoadRememberedTargetNas() As WarehouseTarget
    Dim packed As String
    Dim parts() As String
    Dim target As WarehouseTarget

    On Error Resume Next
    packed = GetSetting(SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_REMEMBERED_TARGET, "")
    On Error GoTo 0
    If Trim$(packed) = "" Then Exit Function
    parts = Split(packed, TARGET_FIELD_DELIM)
    If UBound(parts) < 8 Then Exit Function

    Set target = New WarehouseTarget
    target.WarehouseId = parts(0)
    target.WarehouseName = parts(1)
    target.StationId = parts(2)
    target.HubRoot = NormalizeFolderNas(parts(3))
    target.RuntimeRoot = NormalizeFolderNas(parts(4))
    target.ConfigPath = parts(5)
    target.AuthPath = parts(6)
    target.InboxRoot = parts(7)
    target.LastResolvedUTC = DateFromSerialNas(parts(8))
    If target.WarehouseId = "" Or target.HubRoot = "" Or target.RuntimeRoot = "" Then Exit Function
    Set LoadRememberedTargetNas = target
End Function

Private Function SerializeTargetNas(ByVal target As WarehouseTarget) As String
    SerializeTargetNas = target.WarehouseId & TARGET_FIELD_DELIM & _
                         target.WarehouseName & TARGET_FIELD_DELIM & _
                         target.StationId & TARGET_FIELD_DELIM & _
                         target.HubRoot & TARGET_FIELD_DELIM & _
                         target.RuntimeRoot & TARGET_FIELD_DELIM & _
                         target.ConfigPath & TARGET_FIELD_DELIM & _
                         target.AuthPath & TARGET_FIELD_DELIM & _
                         target.InboxRoot & TARGET_FIELD_DELIM & _
                         CStr(CDbl(target.LastResolvedUTC))
End Function

Private Sub ForgetPersistedTargetNas()
    On Error Resume Next
    DeleteSetting SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_REMEMBERED_TARGET
    DeleteSetting SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_SELECTED_WAREHOUSE_TARGET
    On Error GoTo 0
End Sub

Private Function CloneTargetNas(ByVal source As WarehouseTarget) As WarehouseTarget
    Dim copy As WarehouseTarget

    If source Is Nothing Then Exit Function
    Set copy = New WarehouseTarget
    copy.WarehouseId = source.WarehouseId
    copy.WarehouseName = source.WarehouseName
    copy.StationId = source.StationId
    copy.HubRoot = source.HubRoot
    copy.RuntimeRoot = source.RuntimeRoot
    copy.ConfigPath = source.ConfigPath
    copy.AuthPath = source.AuthPath
    copy.InboxRoot = source.InboxRoot
    copy.SourceType = source.SourceType
    copy.LastResolvedUTC = source.LastResolvedUTC
    Set CloneTargetNas = copy
End Function

Private Function FolderExistsNas(ByVal folderPath As String) As Boolean
    On Error GoTo CleanFail
    Err.Clear
    folderPath = NormalizeFolderNas(folderPath)
    If folderPath = "" Then Exit Function
    FolderExistsNas = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
    Exit Function

CleanFail:
    FolderExistsNas = False
End Function

Private Function FindOpenWorkbookNas(ByVal fullNameIn As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullNameIn, vbTextCompare) = 0 Then
            Set FindOpenWorkbookNas = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FindListObjectNas(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindListObjectNas = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindListObjectNas Is Nothing Then Exit Function
    Next ws
End Function

Private Function FindStationRowNas(ByVal lo As ListObject, ByVal warehouseId As String, ByVal stationId As String) As Long
    Dim i As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For i = 1 To lo.ListRows.Count
        If (stationId = "" Or StrComp(TableValueNas(lo, i, "StationId"), stationId, vbTextCompare) = 0) _
           And (warehouseId = "" Or StrComp(TableValueNas(lo, i, "WarehouseId"), warehouseId, vbTextCompare) = 0) Then
            FindStationRowNas = i
            Exit Function
        End If
    Next i
End Function

Private Function TableValueNas(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    On Error Resume Next
    TableValueNas = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value))
    On Error GoTo 0
End Function

Private Sub AddPathIfMissingNas(ByVal paths As Collection, ByVal pathText As String)
    Dim item As Variant

    pathText = NormalizeFolderNas(pathText)
    If pathText = "" Then Exit Sub
    For Each item In paths
        If SamePathNas(CStr(item), pathText) Then Exit Sub
    Next item
    paths.Add pathText
End Sub

Private Sub EnsureSessionRootsNas()
    If m_SessionRoots Is Nothing Then
        Set m_SessionRoots = CreateObject("Scripting.Dictionary")
        m_SessionRoots.CompareMode = vbTextCompare
    End If
End Sub

Private Function ResolveShareRootNas(ByVal rootPath As String) As String
    Dim trimmedPath As String
    Dim parts() As String

    trimmedPath = NormalizeFolderNas(rootPath)
    If Left$(trimmedPath, 2) <> "\\" Then Exit Function
    parts = Split(Mid$(trimmedPath, 3), "\")
    If UBound(parts) < 1 Then Exit Function
    ResolveShareRootNas = "\\" & parts(0) & "\" & parts(1)
End Function

Private Function IsUncPathNas(ByVal pathText As String) As Boolean
    IsUncPathNas = (Left$(NormalizeFolderNas(pathText), 2) = "\\")
End Function

Private Function NormalizeFolderNas(ByVal folderPath As String) As String
    NormalizeFolderNas = Trim$(Replace$(folderPath, "/", "\"))
    Do While Len(NormalizeFolderNas) > 3 And Right$(NormalizeFolderNas, 1) = "\"
        NormalizeFolderNas = Left$(NormalizeFolderNas, Len(NormalizeFolderNas) - 1)
    Loop
End Function

Private Function SamePathNas(ByVal leftPath As String, ByVal rightPath As String) As Boolean
    SamePathNas = (StrComp(NormalizeFolderNas(leftPath), NormalizeFolderNas(rightPath), vbTextCompare) = 0)
End Function

Private Function DateFromSerialNas(ByVal valueText As String) As Date
    If IsNumeric(valueText) Then DateFromSerialNas = CDate(CDbl(valueText))
End Function

Private Function TargetDisplayNas(ByVal target As WarehouseTarget) As String
    If target Is Nothing Then Exit Function
    TargetDisplayNas = target.WarehouseId
    If target.WarehouseName <> "" And StrComp(target.WarehouseName, target.WarehouseId, vbTextCompare) <> 0 Then
        TargetDisplayNas = TargetDisplayNas & " (" & target.WarehouseName & ")"
    End If
    If target.RuntimeRoot <> "" Then TargetDisplayNas = TargetDisplayNas & " at " & target.RuntimeRoot
End Function

Private Sub SetStatusNas(ByVal statusCode As NasStatusCode, ByVal statusText As String)
    m_LastStatus = statusCode
    m_LastStatusText = statusText
End Sub

Private Function PromptTextNas(ByVal titleText As String, ByVal reason As String) As String
    PromptTextNas = titleText
    If Trim$(reason) <> "" Then PromptTextNas = reason & vbCrLf & vbCrLf & PromptTextNas
End Function

Private Function ResolvePromptDefaultRootNas() As String
    If Not m_StaleTarget Is Nothing Then ResolvePromptDefaultRootNas = m_StaleTarget.HubRoot
    If ResolvePromptDefaultRootNas = "" Then ResolvePromptDefaultRootNas = GetSetting(SETTINGS_APP, SETTINGS_SECTION_NAS, "ShareRoot", "")
End Function

Private Sub SaveRememberedNasUserNas(ByVal shareRoot As String, ByVal userName As String)
    On Error Resume Next
    SaveSetting SETTINGS_APP, SETTINGS_SECTION_NAS, "ShareRoot", shareRoot
    SaveSetting SETTINGS_APP, SETTINGS_SECTION_NAS, "UserName", userName
    On Error GoTo 0
End Sub

Private Function ResolveRememberedNasUserNas() As String
    On Error Resume Next
    ResolveRememberedNasUserNas = Trim$(GetSetting(SETTINGS_APP, SETTINGS_SECTION_NAS, "UserName", ""))
    On Error GoTo 0
End Function
