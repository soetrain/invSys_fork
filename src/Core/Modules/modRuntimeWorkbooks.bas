Attribute VB_Name = "modRuntimeWorkbooks"
Option Explicit

Private mCoreDataRootOverride As String

Private Const SETTINGS_APP As String = "invSys"
Private Const SETTINGS_SECTION_ADMIN As String = "Admin"
Private Const SETTINGS_WAREHOUSE_SCAN_ROOTS As String = "WarehouseScanRoots"
Private Const WAREHOUSE_SCAN_ROOT_DELIMITER As String = "|"

Public Sub SetCoreDataRootOverride(ByVal rootPath As String)
    mCoreDataRootOverride = Trim$(rootPath)
End Sub

Public Sub ClearCoreDataRootOverride()
    mCoreDataRootOverride = vbNullString
End Sub

Public Function GetCoreDataRootOverride() As String
    GetCoreDataRootOverride = Trim$(mCoreDataRootOverride)
End Function

Public Function ResolveCoreDataRoot(Optional ByVal rootPath As String = "", _
                                    Optional ByVal warehouseId As String = "") As String
    Dim resolvedPath As String

    resolvedPath = Trim$(rootPath)
    If resolvedPath = "" Then resolvedPath = Trim$(mCoreDataRootOverride)
    If resolvedPath = "" Then resolvedPath = ResolveConfiguredRuntimeRoot(warehouseId)
    If resolvedPath = "" Then resolvedPath = DefaultRuntimeRoot(ResolveWarehouseIdRuntime(warehouseId))
    If resolvedPath = "" Then resolvedPath = Trim$(CurDir$)

    ResolveCoreDataRoot = NormalizeFolderPath(resolvedPath)
End Function

Public Function OpenOrCreateConfigWorkbookRuntime(Optional ByVal warehouseId As String = "", _
                                                  Optional ByVal stationId As String = "", _
                                                  Optional ByVal rootPath As String = "", _
                                                  Optional ByRef report As String = "") As Workbook
    Dim resolvedWh As String
    Dim targetPath As String
    Dim resolvedRoot As String

    resolvedWh = ResolveWarehouseIdRuntime(warehouseId)
    resolvedRoot = ResolveCoreDataRoot(rootPath, resolvedWh)
    If Trim$(rootPath) = "" Then
        If TryResolveExistingRuntimeRoot(resolvedWh) <> "" Then resolvedRoot = TryResolveExistingRuntimeRoot(resolvedWh)
    End If
    targetPath = BuildCanonicalWorkbookPath(resolvedRoot, resolvedWh, "Config")

    Set OpenOrCreateConfigWorkbookRuntime = OpenOrCreateRuntimeWorkbook( _
        targetPath, "CONFIG", resolvedWh, ResolveStationIdRuntime(stationId), "", report)
End Function

Public Function OpenOrCreateAuthWorkbookRuntime(Optional ByVal warehouseId As String = "", _
                                                Optional ByVal processorServiceUserId As String = "", _
                                                Optional ByVal rootPath As String = "", _
                                                Optional ByRef report As String = "") As Workbook
    Dim resolvedWh As String
    Dim resolvedServiceUser As String
    Dim targetPath As String
    Dim resolvedRoot As String

    resolvedWh = ResolveWarehouseIdRuntime(warehouseId)
    resolvedServiceUser = Trim$(processorServiceUserId)
    If resolvedServiceUser = "" Then resolvedServiceUser = "svc_processor"
    resolvedRoot = ResolveCoreDataRoot(rootPath, resolvedWh)
    If Trim$(rootPath) = "" Then
        If TryResolveExistingRuntimeRoot(resolvedWh) <> "" Then resolvedRoot = TryResolveExistingRuntimeRoot(resolvedWh)
    End If
    targetPath = BuildCanonicalWorkbookPath(resolvedRoot, resolvedWh, "Auth")

    Set OpenOrCreateAuthWorkbookRuntime = OpenOrCreateRuntimeWorkbook( _
        targetPath, "AUTH", resolvedWh, "", resolvedServiceUser, report)
End Function

Public Function OpenFirstRuntimeConfigWorkbook(Optional ByRef report As String = "") As Workbook
    Set OpenFirstRuntimeConfigWorkbook = OpenFirstRuntimeWorkbook("*.invsys.config.xlsb", "CONFIG", report)
End Function

Public Function OpenFirstRuntimeAuthWorkbook(Optional ByRef report As String = "") As Workbook
    Set OpenFirstRuntimeAuthWorkbook = OpenFirstRuntimeWorkbook("*.invsys.auth.xlsb", "AUTH", report)
End Function

Public Function TryResolveExistingRuntimeRoot(Optional ByVal warehouseId As String = "") As String
    Dim resolvedWh As String
    Dim candidateRoot As String
    Dim scanRoot As String
    Dim parentPath As String
    Dim wb As Workbook

    On Error GoTo CleanFail

    resolvedWh = ResolveWarehouseIdRuntime(warehouseId)

    scanRoot = NormalizeFolderPath(Trim$(mCoreDataRootOverride))
    If RuntimeArtifactsExistRuntime(scanRoot, resolvedWh) Then
        TryResolveExistingRuntimeRoot = scanRoot
        Exit Function
    End If
    candidateRoot = FindRuntimeRootUnderParentRuntime(scanRoot, resolvedWh)
    If candidateRoot <> "" Then
        TryResolveExistingRuntimeRoot = candidateRoot
        Exit Function
    End If

    scanRoot = ResolveConfiguredRuntimeRoot(resolvedWh)
    If RuntimeArtifactsExistRuntime(scanRoot, resolvedWh) Then
        TryResolveExistingRuntimeRoot = scanRoot
        Exit Function
    End If
    candidateRoot = FindRuntimeRootUnderParentRuntime(scanRoot, resolvedWh)
    If candidateRoot <> "" Then
        TryResolveExistingRuntimeRoot = candidateRoot
        Exit Function
    End If

    parentPath = GetParentFolder(scanRoot)
    If parentPath <> "" Then
        candidateRoot = FindRuntimeRootUnderParentRuntime(parentPath, resolvedWh)
        If candidateRoot <> "" Then
            TryResolveExistingRuntimeRoot = candidateRoot
            Exit Function
        End If
    End If

    candidateRoot = FindRuntimeRootUnderRememberedRootsRuntime(resolvedWh)
    If candidateRoot <> "" Then
        TryResolveExistingRuntimeRoot = candidateRoot
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, resolvedWh & ".invSys.", vbTextCompare) = 1 Then
            candidateRoot = NormalizeFolderPath(Trim$(wb.Path))
            If RuntimeArtifactsExistRuntime(candidateRoot, resolvedWh) Then
                TryResolveExistingRuntimeRoot = candidateRoot
                Exit Function
            End If
        End If
    Next wb

    candidateRoot = NormalizeFolderPath(DefaultRuntimeRoot(resolvedWh))
    If RuntimeArtifactsExistRuntime(candidateRoot, resolvedWh) Then
        TryResolveExistingRuntimeRoot = candidateRoot
    End If
    Exit Function

CleanFail:
    TryResolveExistingRuntimeRoot = vbNullString
End Function

Private Function FindRuntimeRootUnderRememberedRootsRuntime(ByVal warehouseId As String) As String
    Dim roots As Collection
    Dim rootPath As Variant
    Dim candidateRoot As String

    Set roots = GetRememberedWarehouseScanRootsRuntime()
    For Each rootPath In roots
        If RuntimeArtifactsExistRuntime(CStr(rootPath), warehouseId) Then
            FindRuntimeRootUnderRememberedRootsRuntime = NormalizeFolderPath(CStr(rootPath))
            Exit Function
        End If
        candidateRoot = FindRuntimeRootUnderParentRuntime(CStr(rootPath), warehouseId)
        If candidateRoot <> "" Then
            FindRuntimeRootUnderRememberedRootsRuntime = candidateRoot
            Exit Function
        End If
    Next rootPath
End Function

Public Function GetRememberedWarehouseScanRootsRuntime() As Collection
    Dim roots As Collection
    Dim persistedText As String
    Dim parts() As String
    Dim idx As Long

    Set roots = New Collection
    On Error Resume Next
    persistedText = GetSetting(SETTINGS_APP, SETTINGS_SECTION_ADMIN, SETTINGS_WAREHOUSE_SCAN_ROOTS, "")
    On Error GoTo 0
    If Trim$(persistedText) = "" Then
        Set GetRememberedWarehouseScanRootsRuntime = roots
        Exit Function
    End If

    parts = Split(persistedText, WAREHOUSE_SCAN_ROOT_DELIMITER)
    For idx = LBound(parts) To UBound(parts)
        AddRememberedWarehouseScanRootRuntime roots, CStr(parts(idx))
    Next idx

    Set GetRememberedWarehouseScanRootsRuntime = roots
End Function

Private Sub AddRememberedWarehouseScanRootRuntime(ByVal roots As Collection, ByVal rootPath As String)
    Dim normalizedRoot As String
    Dim item As Variant

    normalizedRoot = NormalizeFolderPath(rootPath)
    If normalizedRoot = "" Then Exit Sub

    For Each item In roots
        If StrComp(CStr(item), normalizedRoot, vbTextCompare) = 0 Then Exit Sub
    Next item
    roots.Add normalizedRoot
End Sub

Private Function OpenOrCreateRuntimeWorkbook(ByVal targetPath As String, _
                                             ByVal workbookKind As String, _
                                             ByVal warehouseId As String, _
                                             ByVal stationId As String, _
                                             ByVal processorServiceUserId As String, _
                                             ByRef report As String) As Workbook
    On Error GoTo FailOpen

    Dim wb As Workbook
    Dim wasCreated As Boolean
    Dim prevEvents As Boolean
    Dim eventsSuppressed As Boolean

    If targetPath = "" Then Exit Function

    Set wb = FindOpenWorkbookByFullName(targetPath)
    If wb Is Nothing Then
        EnsureFolderRecursiveRuntime GetParentFolder(targetPath)
        If Len(Dir$(targetPath)) > 0 Then
            Set wb = Application.Workbooks.Open(targetPath)
        Else
            prevEvents = Application.EnableEvents
            Application.EnableEvents = False
            eventsSuppressed = True
            Set wb = Application.Workbooks.Add(xlWBATWorksheet)
            PrepareWorkbookSurface wb, workbookKind
            wb.SaveAs Filename:=targetPath, FileFormat:=50
            wasCreated = True
            Application.EnableEvents = prevEvents
            eventsSuppressed = False
        End If
    End If

    NormalizeRuntimeWorkbookSheets wb, workbookKind

    Select Case UCase$(workbookKind)
        Case "CONFIG"
            If Not EnsureConfigSchemaRuntime(wb, warehouseId, stationId, report) Then GoTo FailSoft
        Case "AUTH"
            If Not EnsureAuthSchemaRuntime(wb, warehouseId, processorServiceUserId, report) Then GoTo FailSoft
        Case Else
            report = "Unsupported workbook kind: " & workbookKind
            GoTo FailSoft
    End Select

    SaveRuntimeWorkbook wb
    Set OpenOrCreateRuntimeWorkbook = wb
    Exit Function

FailSoft:
    If Len(report) = 0 Then report = workbookKind & " workbook surface failed."
    Exit Function

FailOpen:
    On Error Resume Next
    If eventsSuppressed Then Application.EnableEvents = prevEvents
    On Error GoTo 0
    report = workbookKind & " workbook open/create failed: " & Err.Description
End Function

Private Function OpenFirstRuntimeWorkbook(ByVal likePattern As String, _
                                          ByVal workbookKind As String, _
                                          ByRef report As String) As Workbook
    On Error GoTo FailOpen

    Dim rootPath As String
    Dim fileName As String
    Dim targetPath As String

    rootPath = ResolveCoreDataRoot("", ResolveWarehouseIdRuntime(""))
    If rootPath = "" Then Exit Function

    fileName = Dir$(rootPath & "\*.xlsb")
    Do While fileName <> ""
        If LCase$(fileName) Like LCase$(likePattern) Then
            targetPath = rootPath & "\" & fileName
            Set OpenFirstRuntimeWorkbook = OpenOrCreateRuntimeWorkbook(targetPath, workbookKind, "", "", "svc_processor", report)
            Exit Function
        End If
        fileName = Dir$
    Loop
    Exit Function

FailOpen:
    report = workbookKind & " runtime scan failed: " & Err.Description
End Function

Private Sub PrepareWorkbookSurface(ByVal wb As Workbook, ByVal workbookKind As String)
    Dim wantedSheets As Variant

    Select Case UCase$(workbookKind)
        Case "CONFIG"
            wantedSheets = Array("WarehouseConfig", "StationConfig")
        Case "AUTH"
            wantedSheets = Array("Users", "Capabilities")
        Case Else
            Exit Sub
    End Select

    NormalizeSheetSet wb, wantedSheets
End Sub

Private Sub NormalizeRuntimeWorkbookSheets(ByVal wb As Workbook, ByVal workbookKind As String)
    Select Case UCase$(workbookKind)
        Case "CONFIG"
            NormalizeSheetSet wb, Array("WarehouseConfig", "StationConfig")
        Case "AUTH"
            NormalizeSheetSet wb, Array("Users", "Capabilities")
    End Select
End Sub

Private Sub NormalizeSheetSet(ByVal wb As Workbook, ByVal sheetNames As Variant)
    Dim i As Long
    Dim prevAlerts As Boolean
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub

    For i = LBound(sheetNames) To UBound(sheetNames)
        EnsureNamedWorksheetRuntime wb, CStr(sheetNames(i))
    Next i

    prevAlerts = Application.DisplayAlerts
    Application.DisplayAlerts = False
    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If Not WorksheetNameInSetRuntime(ws.Name, sheetNames) Then ws.Delete
    Next i
    Application.DisplayAlerts = prevAlerts
End Sub

Private Function WorksheetIsBlankRuntime(ByVal ws As Worksheet) As Boolean
    WorksheetIsBlankRuntime = (Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.ListObjects.Count = 0)
End Function

Private Function EnsureNamedWorksheetRuntime(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set EnsureNamedWorksheetRuntime = wb.Worksheets(sheetName)
    On Error GoTo 0

    If EnsureNamedWorksheetRuntime Is Nothing Then
        Set EnsureNamedWorksheetRuntime = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        EnsureNamedWorksheetRuntime.Name = sheetName
    End If
End Function

Private Function WorksheetNameInSetRuntime(ByVal sheetName As String, ByVal sheetNames As Variant) As Boolean
    Dim i As Long

    For i = LBound(sheetNames) To UBound(sheetNames)
        If StrComp(CStr(sheetNames(i)), sheetName, vbTextCompare) = 0 Then
            WorksheetNameInSetRuntime = True
            Exit Function
        End If
    Next i
End Function

Private Function FindOpenWorkbookByFullName(ByVal fullNameIn As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullNameIn, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
End Function

Private Function BuildCanonicalWorkbookPath(ByVal rootPath As String, ByVal warehouseId As String, ByVal workbookType As String) As String
    If rootPath = "" Or warehouseId = "" Then Exit Function
    BuildCanonicalWorkbookPath = rootPath & "\" & warehouseId & ".invSys." & workbookType & ".xlsb"
End Function

Private Function ResolveWarehouseIdRuntime(ByVal warehouseId As String) As String
    ResolveWarehouseIdRuntime = Trim$(warehouseId)
    If ResolveWarehouseIdRuntime = "" Then ResolveWarehouseIdRuntime = "WH1"
End Function

Private Function ResolveStationIdRuntime(ByVal stationId As String) As String
    ResolveStationIdRuntime = Trim$(stationId)
    If ResolveStationIdRuntime = "" Then ResolveStationIdRuntime = "S1"
End Function

Private Function NormalizeFolderPath(ByVal folderPath As String) As String
    folderPath = Trim$(folderPath)
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) = "\" Then
        NormalizeFolderPath = Left$(folderPath, Len(folderPath) - 1)
    Else
        NormalizeFolderPath = folderPath
    End If
End Function

Private Function GetParentFolder(ByVal pathIn As String) As String
    GetParentFolder = modDeploymentPaths.GetParentFolderManaged(pathIn)
End Function

Private Function ResolveConfiguredRuntimeRoot(ByVal warehouseId As String) As String
    On Error Resume Next
    ResolveConfiguredRuntimeRoot = Trim$(GetConfigStringRuntime("PathDataRoot", ""))
    On Error GoTo 0

    If ResolveConfiguredRuntimeRoot <> "" Then ResolveConfiguredRuntimeRoot = NormalizeFolderPath(ResolveConfiguredRuntimeRoot)
    If ResolveConfiguredRuntimeRoot = "" And Trim$(warehouseId) <> "" Then
        ResolveConfiguredRuntimeRoot = NormalizeFolderPath(modDeploymentPaths.DefaultWarehouseRuntimeRootPath(ResolveWarehouseIdRuntime(warehouseId), True))
    End If
End Function

Private Function DefaultRuntimeRoot(ByVal warehouseId As String) As String
    DefaultRuntimeRoot = NormalizeFolderPath(modDeploymentPaths.DefaultWarehouseRuntimeRootPath(ResolveWarehouseIdRuntime(warehouseId), True))
End Function

Private Function RuntimeArtifactsExistRuntime(ByVal rootPath As String, ByVal warehouseId As String) As Boolean
    rootPath = NormalizeFolderPath(rootPath)
    If rootPath = "" Then Exit Function

    RuntimeArtifactsExistRuntime = _
        (Len(Dir$(rootPath & "\" & warehouseId & ".invSys.Config.xlsb", vbNormal)) > 0) And _
        (Len(Dir$(rootPath & "\" & warehouseId & ".invSys.Auth.xlsb", vbNormal)) > 0) And _
        (Len(Dir$(rootPath & "\" & warehouseId & ".invSys.Data.Inventory.xlsb", vbNormal)) > 0)
End Function

Private Function FindRuntimeRootUnderParentRuntime(ByVal parentPath As String, ByVal warehouseId As String) As String
    Dim childName As String
    Dim childPath As String

    On Error GoTo CleanFail

    parentPath = NormalizeFolderPath(parentPath)
    If parentPath = "" Then Exit Function

    childName = Dir$(parentPath & "\*", vbDirectory)
    Do While childName <> ""
        If childName <> "." And childName <> ".." Then
            childPath = parentPath & "\" & childName
            If Len(Dir$(childPath, vbDirectory)) > 0 Then
                If RuntimeArtifactsExistRuntime(childPath, warehouseId) Then
                    FindRuntimeRootUnderParentRuntime = childPath
                    Exit Function
                End If
            End If
        End If
        childName = Dir$
    Loop
    Exit Function

CleanFail:
    FindRuntimeRootUnderParentRuntime = vbNullString
End Function

Private Sub EnsureFolderRecursiveRuntime(ByVal folderPath As String)
    modDeploymentPaths.EnsureFolderRecursiveManaged folderPath
End Sub

Private Sub SaveRuntimeWorkbook(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If Trim$(wb.Path) = "" Then Exit Sub
    wb.Save
End Sub

Private Function EnsureConfigSchemaRuntime(ByVal wb As Workbook, _
                                           ByVal warehouseId As String, _
                                           ByVal stationId As String, _
                                           ByRef report As String) As Boolean
    On Error GoTo FailEnsure

    EnsureConfigSchemaRuntime = CBool(RunRuntimeWorkbookMacro4("modConfig.EnsureConfigSchema", wb, warehouseId, stationId, report))
    If Not EnsureConfigSchemaRuntime And Len(report) = 0 Then report = "EnsureConfigSchema failed."
    Exit Function

FailEnsure:
    If Len(report) = 0 Then report = "EnsureConfigSchema failed: " & Err.Description
End Function

Private Function EnsureAuthSchemaRuntime(ByVal wb As Workbook, _
                                         ByVal warehouseId As String, _
                                         ByVal processorServiceUserId As String, _
                                         ByRef report As String) As Boolean
    On Error GoTo FailEnsure

    EnsureAuthSchemaRuntime = CBool(RunRuntimeWorkbookMacro4("modAuth.EnsureAuthSchema", wb, warehouseId, processorServiceUserId, report))
    If Not EnsureAuthSchemaRuntime And Len(report) = 0 Then report = "EnsureAuthSchema failed."
    Exit Function

FailEnsure:
    If Len(report) = 0 Then report = "EnsureAuthSchema failed: " & Err.Description
End Function

Private Function GetConfigStringRuntime(ByVal key As String, ByVal defaultVal As String) As String
    Dim result As Variant

    On Error GoTo UseDefault
    result = RunRuntimeWorkbookMacro2("modConfig.GetString", key, defaultVal)
    GetConfigStringRuntime = Trim$(CStr(result))
    Exit Function

UseDefault:
    GetConfigStringRuntime = defaultVal
End Function

Private Function RunRuntimeWorkbookMacro2(ByVal macroName As String, _
                                          ByVal arg0 As Variant, _
                                          ByVal arg1 As Variant) As Variant
    On Error GoTo TryUnqualified
    RunRuntimeWorkbookMacro2 = Application.Run("'" & ThisWorkbook.Name & "'!" & macroName, arg0, arg1)
    Exit Function

TryUnqualified:
    Err.Clear
    RunRuntimeWorkbookMacro2 = Application.Run(macroName, arg0, arg1)
End Function

Private Function RunRuntimeWorkbookMacro4(ByVal macroName As String, _
                                          ByVal arg0 As Variant, _
                                          ByVal arg1 As Variant, _
                                          ByVal arg2 As Variant, _
                                          ByVal arg3 As Variant) As Variant
    On Error GoTo TryUnqualified
    RunRuntimeWorkbookMacro4 = Application.Run("'" & ThisWorkbook.Name & "'!" & macroName, arg0, arg1, arg2, arg3)
    Exit Function

TryUnqualified:
    Err.Clear
    RunRuntimeWorkbookMacro4 = Application.Run(macroName, arg0, arg1, arg2, arg3)
End Function
