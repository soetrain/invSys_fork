Attribute VB_Name = "modRibbonRuntimeStatus"
Option Explicit

Private Const TARGET_DELIM As String = "|"
Private Const SETTINGS_APP As String = "invSys"
Private Const SETTINGS_SECTION_RUNTIME As String = "Runtime"
Private Const SETTINGS_SELECTED_WAREHOUSE_TARGET As String = "SelectedWarehouseTarget"
Private mRibbonUis As Collection
Private mWarehouseTargetsCache As Collection
Private mWarehouseTargetsCacheReady As Boolean

Public Sub RegisterRibbonUi(ByVal ribbon As Object)
    If mRibbonUis Is Nothing Then Set mRibbonUis = New Collection
    If ribbon Is Nothing Then Exit Sub
    On Error Resume Next
    mRibbonUis.Add ribbon
    On Error GoTo 0
End Sub

Public Sub InvalidateCurrentUserRibbons()
    Dim ribbon As Variant
    If mRibbonUis Is Nothing Then Exit Sub

    On Error Resume Next
    For Each ribbon In mRibbonUis
        ribbon.Invalidate
        ribbon.InvalidateControl "btnReceivingCurrentUser"
        ribbon.InvalidateControl "btnShippingCurrentUser"
        ribbon.InvalidateControl "btnProductionCurrentUser"
        ribbon.InvalidateControl "btnAdminCurrentUser"
        ribbon.InvalidateControl "btnRuntimeUser"
    Next ribbon
    On Error GoTo 0
End Sub

Public Function GetStatusLabel(ByVal controlId As String) As String
    EnsureRuntimeStatusConfigLoaded

    Select Case Trim$(controlId)
        Case "btnRuntimeWarehouse"
            GetStatusLabel = "Warehouse: " & ValueOrPlaceholderStatus(modConfig.GetWarehouseId()) & _
                             " | Station: " & ValueOrPlaceholderStatus(modConfig.GetStationId())
        Case "btnRuntimeDataRoot"
            GetStatusLabel = "Data root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathDataRoot", ""))
        Case "btnRuntimeInboxRoot"
            GetStatusLabel = "Inbox root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathInboxRoot", ""))
        Case "btnRuntimeUser"
            GetStatusLabel = "User ID: " & ValueOrPlaceholderStatus(ResolveRuntimeUserStatus())
        Case "btnRuntimeProcessor"
            GetStatusLabel = "Processor: " & ValueOrPlaceholderStatus(modConfig.GetString("ProcessorServiceUserId", "svc_processor"))
        Case "btnRuntimeHqAggregator"
            GetStatusLabel = "HQ aggregator: " & ResolveHqAggregatorLabelStatus()
        Case Else
            GetStatusLabel = "Runtime context"
    End Select
End Function

Public Function GetServerStatusLabel(ByVal controlId As String) As String
    Dim target As WarehouseTarget

    Set target = modNasConnection.GetCurrentTarget()
    If modNasConnection.IsCurrentTargetAllowed(True) And TargetHasUncPathStatus(target) Then
        GetServerStatusLabel = "Server: Connected - Send To " & target.WarehouseId
        If Trim$(target.StationId) <> "" Then GetServerStatusLabel = GetServerStatusLabel & " / " & target.StationId
    ElseIf modNasConnection.HasConnectedUncRoot() Then
        GetServerStatusLabel = "Server: Connected - choose Send To"
    Else
        GetServerStatusLabel = "Server: Not connected"
    End If
End Function

Public Function GetWarehouseTargetCount() As Long
    GetWarehouseTargetCount = GetWarehouseTargetsCachedStatus().Count
    If GetWarehouseTargetCount = 0 Then GetWarehouseTargetCount = 1
End Function

Public Function GetWarehouseTargetLabel(ByVal index As Long) As String
    Dim targets As Collection
    Dim targetText As String

    Set targets = GetWarehouseTargetsCachedStatus()
    If targets.Count = 0 Then
        GetWarehouseTargetLabel = "<no warehouse configs found>"
        Exit Function
    End If
    If index < 0 Or index >= targets.Count Then
        GetWarehouseTargetLabel = TargetLabelStatus(CStr(targets(1)))
        Exit Function
    End If

    targetText = CStr(targets(index + 1))
    GetWarehouseTargetLabel = TargetLabelStatus(targetText)
End Function

Public Function GetSelectedWarehouseTargetIndex() As Long
    Dim targets As Collection
    Dim i As Long
    Dim currentWh As String
    Dim currentSt As String
    Dim targetWh As String
    Dim targetSt As String

    EnsureRuntimeStatusConfigLoaded
    currentWh = Trim$(modConfig.GetWarehouseId())
    currentSt = Trim$(modConfig.GetStationId())
    Set targets = GetWarehouseTargetsCachedStatus()

    For i = 1 To targets.Count
        targetWh = TargetPartStatus(CStr(targets(i)), 0)
        targetSt = TargetPartStatus(CStr(targets(i)), 1)
        If StrComp(targetWh, currentWh, vbTextCompare) = 0 _
           And (currentSt = "" Or StrComp(targetSt, currentSt, vbTextCompare) = 0) Then
            GetSelectedWarehouseTargetIndex = i - 1
            Exit Function
        End If
    Next i

    GetSelectedWarehouseTargetIndex = 0
End Function

Public Sub SelectWarehouseTarget(ByVal selectedIndex As Long)
    Dim targets As Collection
    Dim targetText As String
    Dim targetWh As String
    Dim targetSt As String
    Dim targetRoot As String
    Dim nasTarget As WarehouseTarget

    Set targets = GetWarehouseTargetsCachedStatus()
    If targets.Count = 0 Then
        MsgBox "No warehouse config workbooks were found. Use Admin > Setup Tester Station or Create New Warehouse first.", vbExclamation, "invSys Warehouse Target"
        Exit Sub
    End If
    If selectedIndex < 0 Or selectedIndex >= targets.Count Then Exit Sub

    targetText = CStr(targets(selectedIndex + 1))
    targetWh = TargetPartStatus(targetText, 0)
    targetSt = TargetPartStatus(targetText, 1)
    targetRoot = TargetPartStatus(targetText, 2)

    If targetRoot <> "" Then modRuntimeWorkbooks.SetCoreDataRootOverride targetRoot
    If modConfig.LoadConfig(targetWh, targetSt) Then
        If targetRoot <> "" Then
            Call modNasConnection.SelectWarehouseTarget(targetRoot, targetRoot, nasTarget, targetSt, False)
        End If
        RememberSelectedWarehouseTargetStatus targetText
        MsgBox "Warehouse target selected:" & vbCrLf & vbCrLf & _
               TargetLabelStatus(targetText) & vbCrLf & _
               "Inbox root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathInboxRoot", "")), _
               vbInformation, "invSys Warehouse Target"
    Else
        MsgBox "Warehouse target could not be loaded:" & vbCrLf & _
               TargetLabelStatus(targetText) & vbCrLf & vbCrLf & _
               modConfig.Validate(), vbExclamation, "invSys Warehouse Target"
    End If
End Sub

Public Sub InvalidateWarehouseTargets()
    InvalidateWarehouseTargetsCacheStatus
    InvalidateWarehouseTargetRibbonsStatus
End Sub

Public Function TryApplyRememberedWarehouseTarget() As Boolean
    TryApplyRememberedWarehouseTarget = ApplyRememberedWarehouseTargetStatus()
End Function

Public Sub RefreshRuntimeContext()
    Dim report As String
    Dim configLoaded As Boolean

    InvalidateWarehouseTargetsCacheStatus
    configLoaded = ApplyRememberedWarehouseTargetStatus()
    If Not configLoaded Then configLoaded = modConfig.LoadConfig("", "")

    If configLoaded Then
        report = "Warehouse: " & ValueOrPlaceholderStatus(modConfig.GetWarehouseId()) & vbCrLf & _
                 "Station: " & ValueOrPlaceholderStatus(modConfig.GetStationId()) & vbCrLf & _
                "Data root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathDataRoot", "")) & vbCrLf & _
                "Inbox root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathInboxRoot", "")) & vbCrLf & _
                 "User ID: " & ValueOrPlaceholderStatus(ResolveRuntimeUserStatus()) & vbCrLf & _
                 "Processor: " & ValueOrPlaceholderStatus(modConfig.GetString("ProcessorServiceUserId", "svc_processor")) & vbCrLf & _
                 "HQ aggregator: " & ResolveHqAggregatorLabelStatus()
        MsgBox report, vbInformation, "invSys Runtime Context"
    Else
        MsgBox "Runtime config could not be loaded." & vbCrLf & vbCrLf & modConfig.Validate(), vbExclamation, "invSys Runtime Context"
    End If
End Sub

Private Function GetWarehouseTargetsCachedStatus() As Collection
    If Not mWarehouseTargetsCacheReady Or mWarehouseTargetsCache Is Nothing Then
        Set mWarehouseTargetsCache = BuildWarehouseTargetsStatus()
        mWarehouseTargetsCacheReady = True
    End If
    Set GetWarehouseTargetsCachedStatus = mWarehouseTargetsCache
End Function

Private Sub InvalidateWarehouseTargetsCacheStatus()
    mWarehouseTargetsCacheReady = False
    Set mWarehouseTargetsCache = Nothing
End Sub

Private Sub InvalidateWarehouseTargetRibbonsStatus()
    Dim ribbon As Variant
    If mRibbonUis Is Nothing Then Exit Sub

    On Error Resume Next
    For Each ribbon In mRibbonUis
        ribbon.InvalidateControl "ddReceivingWarehouseTarget"
        ribbon.InvalidateControl "ddShippingWarehouseTarget"
        ribbon.InvalidateControl "ddProductionWarehouseTarget"
        ribbon.InvalidateControl "lblReceivingServerStatus"
        ribbon.InvalidateControl "lblShippingServerStatus"
        ribbon.InvalidateControl "lblProductionServerStatus"
    Next ribbon
    On Error GoTo 0
End Sub

Private Sub RememberSelectedWarehouseTargetStatus(ByVal targetText As String)
    On Error Resume Next
    SaveSetting SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_SELECTED_WAREHOUSE_TARGET, targetText
    On Error GoTo 0
End Sub

Private Function ApplyRememberedWarehouseTargetStatus() As Boolean
    Dim targetText As String
    Dim targetWh As String
    Dim targetSt As String
    Dim targetRoot As String
    Dim nasTarget As WarehouseTarget

    On Error Resume Next
    targetText = GetSetting(SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_SELECTED_WAREHOUSE_TARGET, "")
    On Error GoTo 0
    targetText = Trim$(targetText)
    If targetText = "" Then Exit Function

    targetWh = TargetPartStatus(targetText, 0)
    targetSt = TargetPartStatus(targetText, 1)
    targetRoot = NormalizeFolderForStatus(TargetPartStatus(targetText, 2))
    If targetWh = "" Or targetRoot = "" Then Exit Function
    If Not RuntimeArtifactsExistStatus(targetRoot, targetWh) Then Exit Function
    If Left$(targetRoot, 2) = "\\" Then
        If modNasConnection.TryRevalidateRememberedRoot(targetRoot) <> NAS_OK Then Exit Function
    End If

    modRuntimeWorkbooks.SetCoreDataRootOverride targetRoot
    If Not modConfig.LoadConfig(targetWh, targetSt) Then Exit Function
    Call modNasConnection.SelectWarehouseTarget(targetRoot, targetRoot, nasTarget, targetSt, False)
    ApplyRememberedWarehouseTargetStatus = Not nasTarget Is Nothing
End Function

Private Function BuildWarehouseTargetsStatus() As Collection
    Dim targets As Collection
    Dim seen As Object
    Dim currentRoot As String
    Dim currentWh As String
    Dim currentSt As String
    Dim connected As Boolean
    Dim connectedUnc As Boolean

    Set targets = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    EnsureRuntimeStatusConfigLoaded
    connected = modNasConnection.HasConnectedUncRoot()
    connectedUnc = connected

    AddCurrentWarehouseTargetStatus targets, seen, connectedUnc
    AddRememberedSelectedTargetStatus targets, seen, connectedUnc

    If connected Then
        AddKnownServerConfigTargetsStatus targets, seen, connectedUnc
    Else
        currentWh = Trim$(modConfig.GetWarehouseId())
        currentSt = Trim$(modConfig.GetStationId())
        currentRoot = NormalizeFolderForStatus(modConfig.GetString("PathDataRoot", ""))
        AddWarehouseTargetStatus targets, seen, currentWh, currentSt, currentRoot
        AddOpenConfigTargetsStatus targets, seen
    End If

    If targets.Count = 0 And Not connected Then
        AddConfigTargetsUnderRootStatus targets, seen, modDeploymentPaths.DefaultRuntimeHubRootPath(False)
    End If

    Set BuildWarehouseTargetsStatus = targets
End Function

Private Sub AddCurrentWarehouseTargetStatus(ByVal targets As Collection, _
                                            ByVal seen As Object, _
                                            Optional ByVal requireUncTarget As Boolean = False)
    Dim target As WarehouseTarget

    Set target = modNasConnection.GetCurrentTarget()
    If target Is Nothing Then Exit Sub
    If requireUncTarget And Not TargetHasUncPathStatus(target) Then Exit Sub
    AddWarehouseTargetStatus targets, seen, target.WarehouseId, target.StationId, target.RuntimeRoot
End Sub

Private Sub AddKnownServerConfigTargetsStatus(ByVal targets As Collection, _
                                              ByVal seen As Object, _
                                              Optional ByVal requireUncRoot As Boolean = False)
    Dim roots As Collection
    Dim rootPath As Variant

    Set roots = modNasConnection.GetKnownWarehouseTargetRoots()
    For Each rootPath In roots
        If Not requireUncRoot Or Left$(NormalizeFolderForStatus(CStr(rootPath)), 2) = "\\" Then
            AddConfigTargetsUnderRootStatus targets, seen, CStr(rootPath)
        End If
    Next rootPath
End Sub

Private Function TargetHasUncPathStatus(ByVal target As WarehouseTarget) As Boolean
    If target Is Nothing Then Exit Function
    TargetHasUncPathStatus = _
        (Left$(NormalizeFolderForStatus(target.HubRoot), 2) = "\\") Or _
        (Left$(NormalizeFolderForStatus(target.RuntimeRoot), 2) = "\\")
End Function

Private Sub AddRememberedSelectedTargetStatus(ByVal targets As Collection, _
                                              ByVal seen As Object, _
                                              Optional ByVal requireUncRoot As Boolean = False)
    Dim targetText As String
    Dim targetWh As String
    Dim targetSt As String
    Dim targetRoot As String

    On Error Resume Next
    targetText = GetSetting(SETTINGS_APP, SETTINGS_SECTION_RUNTIME, SETTINGS_SELECTED_WAREHOUSE_TARGET, "")
    On Error GoTo 0
    targetText = Trim$(targetText)
    If targetText = "" Then Exit Sub

    targetWh = TargetPartStatus(targetText, 0)
    targetSt = TargetPartStatus(targetText, 1)
    targetRoot = TargetPartStatus(targetText, 2)
    If requireUncRoot And Left$(NormalizeFolderForStatus(targetRoot), 2) <> "\\" Then Exit Sub
    AddWarehouseTargetStatus targets, seen, targetWh, targetSt, targetRoot
End Sub

Private Sub AddRememberedConfigTargetsStatus(ByVal targets As Collection, ByVal seen As Object)
    Dim roots As Collection
    Dim rootPath As Variant

    Set roots = modRuntimeWorkbooks.GetRememberedWarehouseScanRootsRuntime()
    For Each rootPath In roots
        AddConfigTargetsUnderRootStatus targets, seen, CStr(rootPath)
    Next rootPath
End Sub

Private Sub AddOpenConfigTargetsStatus(ByVal targets As Collection, ByVal seen As Object)
    Dim wb As Workbook
    Dim whId As String
    Dim stId As String
    Dim rootPath As String

    On Error Resume Next
    For Each wb In Application.Workbooks
        If ConfigWorkbookLooksUsableStatus(wb, whId, stId) Then
            rootPath = NormalizeFolderForStatus(wb.Path)
            AddWarehouseTargetStatus targets, seen, whId, stId, rootPath
        End If
    Next wb
    On Error GoTo 0
End Sub

Private Sub AddConfigTargetsUnderRootStatus(ByVal targets As Collection, ByVal seen As Object, ByVal rootPath As String)
    Dim fso As Object
    Dim folderObj As Object
    Dim subFolder As Object
    Dim fileObj As Object

    rootPath = NormalizeFolderForStatus(rootPath)
    If rootPath = "" Then Exit Sub
    If Not FolderExistsStatus(rootPath) Then Exit Sub

    On Error GoTo CleanFail
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObj = fso.GetFolder(rootPath)

    For Each fileObj In folderObj.Files
        If IsConfigFileNameStatus(fileObj.Name) Then AddConfigFileTargetStatus targets, seen, CStr(fileObj.Path)
    Next fileObj

    For Each subFolder In folderObj.SubFolders
        For Each fileObj In subFolder.Files
            If IsConfigFileNameStatus(fileObj.Name) Then AddConfigFileTargetStatus targets, seen, CStr(fileObj.Path)
        Next fileObj
    Next subFolder

CleanFail:
End Sub

Private Sub AddConfigFileTargetStatus(ByVal targets As Collection, ByVal seen As Object, ByVal configPath As String)
    Dim fileName As String
    Dim whId As String
    Dim stId As String
    Dim rootPath As String

    fileName = FileNameFromPathStatus(configPath)
    whId = WarehouseIdFromConfigNameStatus(fileName)
    If whId = "" Then Exit Sub
    stId = "S1"
    rootPath = ParentFolderStatus(configPath)
    AddWarehouseTargetStatus targets, seen, whId, stId, rootPath
End Sub

Private Function ConfigWorkbookLooksUsableStatus(ByVal wb As Workbook, ByRef warehouseId As String, ByRef stationId As String) As Boolean
    Dim loWh As ListObject
    Dim loSt As ListObject

    If wb Is Nothing Then Exit Function
    Set loWh = FindListObjectByNameStatus(wb, "tblWarehouseConfig")
    Set loSt = FindListObjectByNameStatus(wb, "tblStationConfig")
    If loWh Is Nothing Or loSt Is Nothing Then Exit Function
    If loWh.DataBodyRange Is Nothing Then Exit Function

    warehouseId = SafeTableValueStatus(loWh, 1, "WarehouseId")
    If loSt.DataBodyRange Is Nothing Then
        stationId = "S1"
    Else
        stationId = SafeTableValueStatus(loSt, 1, "StationId")
    End If
    If warehouseId = "" Then warehouseId = WarehouseIdFromConfigNameStatus(wb.Name)
    If stationId = "" Then stationId = "S1"
    ConfigWorkbookLooksUsableStatus = (warehouseId <> "")
End Function

Private Sub AddWarehouseTargetStatus(ByVal targets As Collection, _
                                     ByVal seen As Object, _
                                     ByVal warehouseId As String, _
                                     ByVal stationId As String, _
                                     ByVal rootPath As String)
    Dim key As String

    warehouseId = Trim$(warehouseId)
    stationId = Trim$(stationId)
    rootPath = NormalizeFolderForStatus(rootPath)
    If warehouseId = "" Then Exit Sub
    If stationId = "" Then stationId = "S1"

    key = warehouseId & TARGET_DELIM & stationId & TARGET_DELIM & rootPath
    If seen.Exists(key) Then Exit Sub
    seen(key) = True
    targets.Add key
End Sub

Private Sub EnsureRuntimeStatusConfigLoaded()
    If Trim$(modConfig.GetWarehouseId()) = "" Then
        On Error Resume Next
        Call modConfig.LoadConfig("", "")
        On Error GoTo 0
    End If
End Sub

Private Function ResolveHqAggregatorLabelStatus() As String
    Dim sharePointRoot As String

    sharePointRoot = Trim$(modConfig.GetString("PathSharePointRoot", ""))
    If sharePointRoot = "" Then
        ResolveHqAggregatorLabelStatus = "<not configured>"
    Else
        ResolveHqAggregatorLabelStatus = "Admin scheduled aggregation via " & NormalizeFolderForStatus(sharePointRoot) & "\Snapshots"
    End If
End Function

Private Function ResolveRuntimeUserStatus() As String
    If modAuth.IsSignedIn() Then
        ResolveRuntimeUserStatus = Trim$(modAuth.GetCurrentUserId())
    Else
        ResolveRuntimeUserStatus = "<not signed in>"
    End If
End Function

Private Function ValueOrPlaceholderStatus(ByVal valueIn As String) As String
    valueIn = Trim$(valueIn)
    If valueIn = "" Then
        ValueOrPlaceholderStatus = "<not configured>"
    Else
        ValueOrPlaceholderStatus = valueIn
    End If
End Function

Private Function TargetLabelStatus(ByVal targetText As String) As String
    TargetLabelStatus = TargetPartStatus(targetText, 0) & " | " & TargetPartStatus(targetText, 1)
    If TargetPartStatus(targetText, 2) <> "" Then TargetLabelStatus = TargetLabelStatus & " | " & TargetPartStatus(targetText, 2)
End Function

Private Function TargetPartStatus(ByVal targetText As String, ByVal partIndex As Long) As String
    Dim parts() As String

    parts = Split(targetText, TARGET_DELIM)
    If partIndex < LBound(parts) Or partIndex > UBound(parts) Then Exit Function
    TargetPartStatus = Trim$(parts(partIndex))
End Function

Private Function IsConfigFileNameStatus(ByVal fileName As String) As Boolean
    fileName = LCase$(Trim$(fileName))
    IsConfigFileNameStatus = (fileName Like "*.invsys.config.xlsb") Or _
                             (fileName Like "*.invsys.config.xlsm") Or _
                             (fileName Like "*.invsys.config.xlsx")
End Function

Private Function WarehouseIdFromConfigNameStatus(ByVal fileName As String) As String
    Dim markerPos As Long

    markerPos = InStr(1, fileName, ".invSys.Config.", vbTextCompare)
    If markerPos > 1 Then WarehouseIdFromConfigNameStatus = Left$(fileName, markerPos - 1)
End Function

Private Function FileNameFromPathStatus(ByVal filePath As String) As String
    Dim slashPos As Long

    filePath = Trim$(Replace$(filePath, "/", "\"))
    slashPos = InStrRev(filePath, "\")
    If slashPos > 0 Then
        FileNameFromPathStatus = Mid$(filePath, slashPos + 1)
    Else
        FileNameFromPathStatus = filePath
    End If
End Function

Private Function ParentFolderStatus(ByVal pathText As String) As String
    ParentFolderStatus = modDeploymentPaths.GetParentFolderManaged(pathText)
End Function

Private Function FolderExistsStatus(ByVal folderPath As String) As Boolean
    FolderExistsStatus = modDeploymentPaths.FolderExistsManaged(folderPath)
End Function

Private Function RuntimeArtifactsExistStatus(ByVal rootPath As String, ByVal warehouseId As String) As Boolean
    rootPath = NormalizeFolderForStatus(rootPath)
    warehouseId = Trim$(warehouseId)
    If rootPath = "" Or warehouseId = "" Then Exit Function

    RuntimeArtifactsExistStatus = _
        FileExistsStatus(rootPath & "\" & warehouseId & ".invSys.Config.xlsb") And _
        FileExistsStatus(rootPath & "\" & warehouseId & ".invSys.Auth.xlsb")
End Function

Private Function FileExistsStatus(ByVal filePath As String) As Boolean
    FileExistsStatus = modDeploymentPaths.FileExistsManaged(filePath)
End Function

Private Function FindListObjectByNameStatus(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindListObjectByNameStatus = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindListObjectByNameStatus Is Nothing Then Exit Function
    Next ws
End Function

Private Function SafeTableValueStatus(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    On Error Resume Next
    SafeTableValueStatus = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value))
    On Error GoTo 0
End Function

Private Function NormalizeFolderForStatus(ByVal folderPath As String) As String
    NormalizeFolderForStatus = Trim$(Replace$(folderPath, "/", "\"))
    Do While Len(NormalizeFolderForStatus) > 3 And Right$(NormalizeFolderForStatus, 1) = "\"
        NormalizeFolderForStatus = Left$(NormalizeFolderForStatus, Len(NormalizeFolderForStatus) - 1)
    Loop
End Function
