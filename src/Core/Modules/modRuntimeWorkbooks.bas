Attribute VB_Name = "modRuntimeWorkbooks"
Option Explicit

Private mCoreDataRootOverride As String

Public Sub SetCoreDataRootOverride(ByVal rootPath As String)
    mCoreDataRootOverride = Trim$(rootPath)
End Sub

Public Sub ClearCoreDataRootOverride()
    mCoreDataRootOverride = vbNullString
End Sub

Public Function GetCoreDataRootOverride() As String
    GetCoreDataRootOverride = Trim$(mCoreDataRootOverride)
End Function

Public Function ResolveCoreDataRoot(Optional ByVal rootPath As String = "") As String
    Dim resolvedPath As String

    resolvedPath = Trim$(rootPath)
    If resolvedPath = "" Then resolvedPath = Trim$(mCoreDataRootOverride)
    If resolvedPath = "" Then resolvedPath = Trim$(ThisWorkbook.Path)
    If resolvedPath = "" Then resolvedPath = Trim$(CurDir$)

    ResolveCoreDataRoot = NormalizeFolderPath(resolvedPath)
End Function

Public Function OpenOrCreateConfigWorkbookRuntime(Optional ByVal warehouseId As String = "", _
                                                  Optional ByVal stationId As String = "", _
                                                  Optional ByVal rootPath As String = "", _
                                                  Optional ByRef report As String = "") As Workbook
    Dim resolvedWh As String
    Dim targetPath As String

    resolvedWh = ResolveWarehouseIdRuntime(warehouseId)
    targetPath = BuildCanonicalWorkbookPath(ResolveCoreDataRoot(rootPath), resolvedWh, "Config")

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

    resolvedWh = ResolveWarehouseIdRuntime(warehouseId)
    resolvedServiceUser = Trim$(processorServiceUserId)
    If resolvedServiceUser = "" Then resolvedServiceUser = "svc_processor"
    targetPath = BuildCanonicalWorkbookPath(ResolveCoreDataRoot(rootPath), resolvedWh, "Auth")

    Set OpenOrCreateAuthWorkbookRuntime = OpenOrCreateRuntimeWorkbook( _
        targetPath, "AUTH", resolvedWh, "", resolvedServiceUser, report)
End Function

Public Function OpenFirstRuntimeConfigWorkbook(Optional ByRef report As String = "") As Workbook
    Set OpenFirstRuntimeConfigWorkbook = OpenFirstRuntimeWorkbook("*.invsys.config.xlsb", "CONFIG", report)
End Function

Public Function OpenFirstRuntimeAuthWorkbook(Optional ByRef report As String = "") As Workbook
    Set OpenFirstRuntimeAuthWorkbook = OpenFirstRuntimeWorkbook("*.invsys.auth.xlsb", "AUTH", report)
End Function

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

    Select Case UCase$(workbookKind)
        Case "CONFIG"
            If Not modConfig.EnsureConfigSchema(wb, warehouseId, stationId, report) Then GoTo FailSoft
        Case "AUTH"
            If Not modAuth.EnsureAuthSchema(wb, warehouseId, processorServiceUserId, report) Then GoTo FailSoft
        Case Else
            report = "Unsupported workbook kind: " & workbookKind
            GoTo FailSoft
    End Select

    If wasCreated Then wb.Save
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

    rootPath = ResolveCoreDataRoot()
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

    EnsureSheetSet wb, wantedSheets
End Sub

Private Sub EnsureSheetSet(ByVal wb As Workbook, ByVal sheetNames As Variant)
    Dim requiredCount As Long
    Dim i As Long

    requiredCount = UBound(sheetNames) - LBound(sheetNames) + 1
    Do While wb.Worksheets.Count < requiredCount
        wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count)
    Loop

    For i = 1 To requiredCount
        wb.Worksheets(i).Name = CStr(sheetNames(LBound(sheetNames) + i - 1))
    Next i

    For i = wb.Worksheets.Count To requiredCount + 1 Step -1
        If WorksheetIsBlankRuntime(wb.Worksheets(i)) Then
            Application.DisplayAlerts = False
            wb.Worksheets(i).Delete
            Application.DisplayAlerts = True
        End If
    Next i
End Sub

Private Function WorksheetIsBlankRuntime(ByVal ws As Worksheet) As Boolean
    WorksheetIsBlankRuntime = (Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.ListObjects.Count = 0)
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
    Dim sepPos As Long

    sepPos = InStrRev(pathIn, "\")
    If sepPos > 1 Then GetParentFolder = Left$(pathIn, sepPos - 1)
End Function

Private Sub EnsureFolderRecursiveRuntime(ByVal folderPath As String)
    Dim parentPath As String
    Dim sepPos As Long

    folderPath = NormalizeFolderPath(folderPath)
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    sepPos = InStrRev(folderPath, "\")
    If sepPos > 1 Then
        parentPath = Left$(folderPath, sepPos - 1)
        If Right$(parentPath, 1) = ":" Then parentPath = parentPath & "\"
        If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderRecursiveRuntime parentPath
    End If

    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath
End Sub
