Attribute VB_Name = "modRibbonRuntimeStatus"
Option Explicit

Private Const TARGET_DELIM As String = "|"

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
        Case "btnRuntimeProcessor"
            GetStatusLabel = "Processor: " & ValueOrPlaceholderStatus(modConfig.GetString("ProcessorServiceUserId", "svc_processor"))
        Case "btnRuntimeHqAggregator"
            GetStatusLabel = "HQ aggregator: " & ResolveHqAggregatorLabelStatus()
        Case Else
            GetStatusLabel = "Runtime context"
    End Select
End Function

Public Function GetWarehouseTargetCount() As Long
    GetWarehouseTargetCount = BuildWarehouseTargetsStatus().Count
    If GetWarehouseTargetCount = 0 Then GetWarehouseTargetCount = 1
End Function

Public Function GetWarehouseTargetLabel(ByVal index As Long) As String
    Dim targets As Collection
    Dim targetText As String

    Set targets = BuildWarehouseTargetsStatus()
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
    Set targets = BuildWarehouseTargetsStatus()

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

    Set targets = BuildWarehouseTargetsStatus()
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

Public Sub RefreshRuntimeContext()
    Dim report As String

    If modConfig.LoadConfig("", "") Then
        report = "Warehouse: " & ValueOrPlaceholderStatus(modConfig.GetWarehouseId()) & vbCrLf & _
                 "Station: " & ValueOrPlaceholderStatus(modConfig.GetStationId()) & vbCrLf & _
                 "Data root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathDataRoot", "")) & vbCrLf & _
                 "Inbox root: " & ValueOrPlaceholderStatus(modConfig.GetString("PathInboxRoot", "")) & vbCrLf & _
                 "Processor: " & ValueOrPlaceholderStatus(modConfig.GetString("ProcessorServiceUserId", "svc_processor")) & vbCrLf & _
                 "HQ aggregator: " & ResolveHqAggregatorLabelStatus()
        MsgBox report, vbInformation, "invSys Runtime Context"
    Else
        MsgBox "Runtime config could not be loaded." & vbCrLf & vbCrLf & modConfig.Validate(), vbExclamation, "invSys Runtime Context"
    End If
End Sub

Private Function BuildWarehouseTargetsStatus() As Collection
    Dim targets As Collection
    Dim seen As Object
    Dim currentRoot As String
    Dim currentWh As String
    Dim currentSt As String

    Set targets = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    EnsureRuntimeStatusConfigLoaded
    currentWh = Trim$(modConfig.GetWarehouseId())
    currentSt = Trim$(modConfig.GetStationId())
    currentRoot = NormalizeFolderForStatus(modConfig.GetString("PathDataRoot", ""))
    AddWarehouseTargetStatus targets, seen, currentWh, currentSt, currentRoot

    AddOpenConfigTargetsStatus targets, seen
    AddConfigTargetsUnderRootStatus targets, seen, modDeploymentPaths.DefaultRuntimeHubRootPath(False)
    If currentRoot <> "" Then AddConfigTargetsUnderRootStatus targets, seen, ParentFolderStatus(currentRoot)

    Set BuildWarehouseTargetsStatus = targets
End Function

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
