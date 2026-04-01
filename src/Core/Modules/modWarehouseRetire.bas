Attribute VB_Name = "modWarehouseRetire"
Option Explicit

Private Const ARCHIVE_VERSION_RETIRE As String = "1.0"
Private Const ARCHIVE_TEMP_SUFFIX_RETIRE As String = "_tmp"

Private mLastRetireReport As String

Public Enum RetireMigrateOperationMode
    MODE_ARCHIVE_ONLY = 1
    MODE_ARCHIVE_MIGRATE = 2
    MODE_ARCHIVE_RETIRE = 3
    MODE_ARCHIVE_RETIRE_DELETE = 4
End Enum

Public Type RetireMigrateSpec
    SourceWarehouseId As String
    TargetWarehouseId As String
    OperationMode As RetireMigrateOperationMode
    AdminUser As String
    ConfirmedByUser As Boolean
    ArchiveDestPath As String
    PublishTombstone As Boolean
End Type

Public Function ValidateRetireMigrateSpec(ByRef spec As RetireMigrateSpec, _
                                          Optional ByRef report As String = "") As Boolean
    NormalizeRetireMigrateSpec spec

    If spec.SourceWarehouseId = "" Then
        report = "SourceWarehouseId is required."
        Exit Function
    End If

    If Not IsSupportedRetireMigrateMode(spec.OperationMode) Then
        report = "OperationMode is not supported."
        Exit Function
    End If

    If spec.OperationMode = MODE_ARCHIVE_MIGRATE And spec.TargetWarehouseId = "" Then
        report = "TargetWarehouseId is required for MODE_ARCHIVE_MIGRATE."
        Exit Function
    End If

    If spec.TargetWarehouseId <> "" Then
        If StrComp(spec.SourceWarehouseId, spec.TargetWarehouseId, vbTextCompare) = 0 Then
            report = "SourceWarehouseId and TargetWarehouseId must not be the same."
            Exit Function
        End If
    End If

    If Not spec.ConfirmedByUser Then
        report = "ConfirmedByUser must be True before any write operation proceeds."
        Exit Function
    End If

    If Not IsValidLocalPathFormatRetire(spec.ArchiveDestPath) Then
        report = "ArchiveDestPath must be a valid local path format."
        Exit Function
    End If

    report = "OK"
    ValidateRetireMigrateSpec = True
End Function

Public Function RequireReAuth(ByVal requiredRole As String) As Boolean
    Dim gate As frmReAuthGate

    Set gate = New frmReAuthGate
    gate.InitializeGate ResolveRequiredRoleRetire(requiredRole), ResolveCurrentAdminUserRetire()
    gate.Show vbModal
    RequireReAuth = gate.Authenticated
    Unload gate
End Function

Public Function WriteArchivePackage(ByRef spec As RetireMigrateSpec) As Boolean
    Dim report As String
    Dim runtimeRoot As String
    Dim archiveParent As String
    Dim archiveFolderName As String
    Dim tempArchivePath As String
    Dim finalArchivePath As String
    Dim configPath As String
    Dim authPath As String
    Dim inventoryPath As String
    Dim snapshotPath As String
    Dim outboxPath As String
    Dim warehouseName As String
    Dim fileList As Collection
    Dim wbCfg As Workbook
    Dim wbAuth As Workbook

    On Error GoTo FailArchive

    mLastRetireReport = vbNullString
    If Not ValidateRetireMigrateSpec(spec, report) Then GoTo FailSoft

    runtimeRoot = ResolveRuntimeRootRetire(spec.SourceWarehouseId)
    If runtimeRoot = "" Then
        report = "Runtime root could not be resolved."
        GoTo FailSoft
    End If

    archiveParent = NormalizeFolderPathRetire(spec.ArchiveDestPath)
    If archiveParent = "" Then
        report = "ArchiveDestPath could not be resolved."
        GoTo FailSoft
    End If
    EnsureFolderRecursiveRetire Left$(archiveParent, Len(archiveParent) - 1)

    archiveFolderName = BuildArchiveFolderNameRetire(spec.SourceWarehouseId)
    finalArchivePath = Left$(archiveParent, Len(archiveParent) - 1) & "\" & archiveFolderName
    tempArchivePath = finalArchivePath & ARCHIVE_TEMP_SUFFIX_RETIRE
    If FolderExistsRetire(finalArchivePath) Then
        report = "Archive output already exists: " & finalArchivePath
        GoTo FailSoft
    End If
    If FolderExistsRetire(tempArchivePath) Then DeleteFolderRecursiveRetire tempArchivePath
    EnsureFolderRecursiveRetire tempArchivePath

    configPath = runtimeRoot & "\" & spec.SourceWarehouseId & ".invSys.Config.xlsb"
    authPath = runtimeRoot & "\" & spec.SourceWarehouseId & ".invSys.Auth.xlsb"
    inventoryPath = runtimeRoot & "\" & spec.SourceWarehouseId & ".invSys.Data.Inventory.xlsb"
    snapshotPath = ResolveLatestSnapshotPathRetire(runtimeRoot, spec.SourceWarehouseId)
    outboxPath = runtimeRoot & "\" & spec.SourceWarehouseId & ".Outbox.Events.xlsb"

    If Not FileExistsRetire(configPath) Then
        report = "Config workbook not found: " & configPath
        GoTo FailSoft
    End If
    If Not FileExistsRetire(authPath) Then
        report = "Auth workbook not found: " & authPath
        GoTo FailSoft
    End If
    If Not FileExistsRetire(inventoryPath) Then
        report = "Inventory workbook not found: " & inventoryPath
        GoTo FailSoft
    End If
    If snapshotPath = "" Then
        report = "Local snapshot workbook not found."
        GoTo FailSoft
    End If
    If Not FileExistsRetire(outboxPath) Then
        report = "Outbox workbook not found: " & outboxPath
        GoTo FailSoft
    End If

    Set fileList = New Collection

    Set wbCfg = OpenWorkbookReadOnlyRetire(configPath, report)
    If wbCfg Is Nothing Then GoTo FailSoft
    warehouseName = ResolveWarehouseNameRetire(wbCfg, spec.SourceWarehouseId)
    If Not ExportConfigTablesRetire(wbCfg, tempArchivePath & "\config", fileList, report) Then GoTo FailSoft
    CloseWorkbookQuietlyRetire wbCfg
    Set wbCfg = Nothing

    Set wbAuth = OpenWorkbookReadOnlyRetire(authPath, report)
    If wbAuth Is Nothing Then GoTo FailSoft
    If Not SanitizeAuthExport(wbAuth, tempArchivePath & "\auth", fileList, report) Then GoTo FailSoft
    CloseWorkbookQuietlyRetire wbAuth
    Set wbAuth = Nothing

    If Not CopyArtifactToArchiveRetire(inventoryPath, tempArchivePath & "\inventory\" & GetFileNameRetire(inventoryPath), "inventory\" & GetFileNameRetire(inventoryPath), fileList, report) Then GoTo FailSoft
    If Not CopyArtifactToArchiveRetire(snapshotPath, tempArchivePath & "\snapshots\" & GetFileNameRetire(snapshotPath), "snapshots\" & GetFileNameRetire(snapshotPath), fileList, report) Then GoTo FailSoft
    If Not CopyArtifactToArchiveRetire(outboxPath, tempArchivePath & "\outbox\" & GetFileNameRetire(outboxPath), "outbox\" & GetFileNameRetire(outboxPath), fileList, report) Then GoTo FailSoft
    If Not CopyPendingOutboxFolderRetire(runtimeRoot & "\outbox", tempArchivePath & "\outbox", fileList, report) Then GoTo FailSoft

    fileList.Add "manifest.json"
    If Not WriteArchiveManifestRetire(tempArchivePath & "\manifest.json", spec, warehouseName, fileList, report) Then GoTo FailSoft

    Name tempArchivePath As finalArchivePath
    report = "OK|ArchivePath=" & finalArchivePath
    WriteArchivePackage = True
    GoTo CleanExit

FailSoft:
    WriteArchivePackage = False
    If Len(report) = 0 Then report = "WriteArchivePackage failed."
    DeleteFolderRecursiveRetire tempArchivePath
    LogDiagnosticSafeRetire "WAREHOUSE-RETIRE", "Archive failed|WarehouseId=" & spec.SourceWarehouseId & "|Reason=" & report
    GoTo CleanExit

FailArchive:
    report = "WriteArchivePackage failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookQuietlyRetire wbAuth
    CloseWorkbookQuietlyRetire wbCfg
    mLastRetireReport = report
End Function

Public Function SanitizeAuthExport(ByVal authWb As Workbook, _
                                   ByVal exportFolder As String, _
                                   ByRef fileList As Collection, _
                                   Optional ByRef report As String = "") As Boolean
    Dim loUsers As ListObject
    Dim loCaps As ListObject

    On Error GoTo FailSanitize

    If authWb Is Nothing Then
        report = "Auth workbook not resolved."
        Exit Function
    End If

    Set loUsers = authWb.Worksheets("Users").ListObjects("tblUsers")
    Set loCaps = authWb.Worksheets("Capabilities").ListObjects("tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then
        report = "Auth tables were not available for export."
        Exit Function
    End If

    EnsureFolderRecursiveRetire exportFolder
    If Not ExportListObjectCsvRetire(loUsers, exportFolder & "\tblUsers.csv", Array("PinHash", "PIN", "Password", "PasswordHash"), report) Then Exit Function
    fileList.Add "auth\tblUsers.csv"
    If Not ExportListObjectCsvRetire(loCaps, exportFolder & "\tblCapabilities.csv", Empty, report) Then Exit Function
    fileList.Add "auth\tblCapabilities.csv"

    SanitizeAuthExport = True
    Exit Function

FailSanitize:
    report = "SanitizeAuthExport failed: " & Err.Description
End Function

Public Function GetLastWarehouseRetireReport() As String
    GetLastWarehouseRetireReport = mLastRetireReport
End Function

Private Sub NormalizeRetireMigrateSpec(ByRef spec As RetireMigrateSpec)
    spec.SourceWarehouseId = Trim$(spec.SourceWarehouseId)
    spec.TargetWarehouseId = Trim$(spec.TargetWarehouseId)
    spec.AdminUser = Trim$(spec.AdminUser)
    spec.ArchiveDestPath = Trim$(Replace$(spec.ArchiveDestPath, "/", "\"))
End Sub

Private Function IsSupportedRetireMigrateMode(ByVal modeValue As RetireMigrateOperationMode) As Boolean
    Select Case modeValue
        Case MODE_ARCHIVE_ONLY, MODE_ARCHIVE_MIGRATE, MODE_ARCHIVE_RETIRE, MODE_ARCHIVE_RETIRE_DELETE
            IsSupportedRetireMigrateMode = True
    End Select
End Function

Private Function IsValidLocalPathFormatRetire(ByVal pathIn As String) As Boolean
    Dim invalidChars As Variant
    Dim item As Variant
    Dim tailPath As String

    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    If pathIn = "" Then Exit Function
    If Len(pathIn) < 3 Then Exit Function
    If Mid$(pathIn, 2, 1) <> ":" Then Exit Function
    If Mid$(pathIn, 3, 1) <> "\" Then Exit Function

    invalidChars = Array("<", ">", """", "|", "?", "*")
    tailPath = Mid$(pathIn, 4)
    For Each item In invalidChars
        If InStr(1, tailPath, CStr(item), vbBinaryCompare) > 0 Then Exit Function
    Next item

    IsValidLocalPathFormatRetire = True
End Function

Private Function ResolveRequiredRoleRetire(ByVal requiredRole As String) As String
    ResolveRequiredRoleRetire = Trim$(requiredRole)
    If ResolveRequiredRoleRetire = "" Then ResolveRequiredRoleRetire = "ADMIN_MAINT"
End Function

Private Function ResolveCurrentAdminUserRetire() As String
    ResolveCurrentAdminUserRetire = Trim$(Environ$("USERNAME"))
    If ResolveCurrentAdminUserRetire = "" Then ResolveCurrentAdminUserRetire = Trim$(Application.UserName)
End Function

Private Function ResolveRuntimeRootRetire(ByVal warehouseId As String) As String
    ResolveRuntimeRootRetire = NormalizeFolderPathRetire(modRuntimeWorkbooks.ResolveCoreDataRoot("", warehouseId))
    If Right$(ResolveRuntimeRootRetire, 1) = "\" Then
        ResolveRuntimeRootRetire = Left$(ResolveRuntimeRootRetire, Len(ResolveRuntimeRootRetire) - 1)
    End If
End Function

Private Function NormalizeFolderPathRetire(ByVal folderPath As String) As String
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathRetire = folderPath
End Function

Private Function BuildArchiveFolderNameRetire(ByVal warehouseId As String) As String
    BuildArchiveFolderNameRetire = Trim$(warehouseId) & "_archive_" & Format$(Now, "yyyymmdd_hhnnss")
End Function

Private Function ResolveLatestSnapshotPathRetire(ByVal runtimeRoot As String, ByVal warehouseId As String) As String
    Dim canonicalPath As String
    Dim folderSnapshots As String
    Dim candidate As String
    Dim bestStamp As String
    Dim candidateStamp As String

    canonicalPath = runtimeRoot & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    If FileExistsRetire(canonicalPath) Then
        ResolveLatestSnapshotPathRetire = canonicalPath
        Exit Function
    End If

    folderSnapshots = NormalizeFolderPathRetire(runtimeRoot & "\snapshots")
    candidate = Dir$(folderSnapshots & warehouseId & "*.invSys.Snapshot.Inventory.xls*")
    Do While candidate <> ""
        candidateStamp = Format$(FileDateTime(folderSnapshots & candidate), "yyyymmddhhnnss")
        If candidateStamp > bestStamp Then
            bestStamp = candidateStamp
            ResolveLatestSnapshotPathRetire = folderSnapshots & candidate
        End If
        candidate = Dir$
    Loop
End Function

Private Function OpenWorkbookReadOnlyRetire(ByVal workbookPath As String, ByRef report As String) As Workbook
    On Error GoTo FailOpen

    Set OpenWorkbookReadOnlyRetire = FindOpenWorkbookByFullNameRetire(workbookPath)
    If OpenWorkbookReadOnlyRetire Is Nothing Then
        Set OpenWorkbookReadOnlyRetire = Application.Workbooks.Open(workbookPath, False, True)
    End If
    Exit Function

FailOpen:
    report = "OpenWorkbookReadOnlyRetire failed: " & Err.Description
End Function

Private Function FindOpenWorkbookByFullNameRetire(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook

    workbookPath = Trim$(workbookPath)
    If workbookPath = "" Then Exit Function
    For Each wb In Application.Workbooks
        If StrComp(Trim$(wb.FullName), workbookPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByFullNameRetire = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub CloseWorkbookQuietlyRetire(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function ResolveWarehouseNameRetire(ByVal wbCfg As Workbook, ByVal warehouseId As String) As String
    Dim lo As ListObject

    On Error Resume Next
    Set lo = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then
        ResolveWarehouseNameRetire = warehouseId
        Exit Function
    End If

    ResolveWarehouseNameRetire = Trim$(CStr(lo.DataBodyRange.Cells(1, lo.ListColumns("WarehouseName").Index).Value))
    If ResolveWarehouseNameRetire = "" Then ResolveWarehouseNameRetire = warehouseId
End Function

Private Function ExportConfigTablesRetire(ByVal wbCfg As Workbook, _
                                          ByVal exportFolder As String, _
                                          ByRef fileList As Collection, _
                                          ByRef report As String) As Boolean
    Dim loWh As ListObject
    Dim loSt As ListObject

    On Error GoTo FailExport

    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")
    If loWh Is Nothing Or loSt Is Nothing Then
        report = "Config tables were not available for export."
        Exit Function
    End If

    EnsureFolderRecursiveRetire exportFolder
    If Not ExportListObjectCsvRetire(loWh, exportFolder & "\tblWarehouseConfig.csv", Empty, report) Then Exit Function
    fileList.Add "config\tblWarehouseConfig.csv"
    If Not ExportListObjectCsvRetire(loSt, exportFolder & "\tblStationConfig.csv", Empty, report) Then Exit Function
    fileList.Add "config\tblStationConfig.csv"

    ExportConfigTablesRetire = True
    Exit Function

FailExport:
    report = "ExportConfigTablesRetire failed: " & Err.Description
End Function

Private Function ExportListObjectCsvRetire(ByVal lo As ListObject, _
                                           ByVal outputPath As String, _
                                           Optional ByVal blankColumns As Variant, _
                                           Optional ByRef report As String = "") As Boolean
    Dim fileNum As Integer
    Dim rowIndex As Long
    Dim colIndex As Long
    Dim lineText As String
    Dim columnName As String
    Dim shouldBlank As Boolean

    On Error GoTo FailExport

    If lo Is Nothing Then
        report = "ListObject not resolved for CSV export."
        Exit Function
    End If

    EnsureFolderRecursiveRetire GetParentFolderRetire(outputPath)
    fileNum = FreeFile
    Open outputPath For Output As #fileNum

    lineText = vbNullString
    For colIndex = 1 To lo.ListColumns.Count
        If colIndex > 1 Then lineText = lineText & ","
        lineText = lineText & EscapeCsvValueRetire(lo.ListColumns(colIndex).Name)
    Next colIndex
    Print #fileNum, lineText

    If Not lo.DataBodyRange Is Nothing Then
        For rowIndex = 1 To lo.ListRows.Count
            lineText = vbNullString
            For colIndex = 1 To lo.ListColumns.Count
                If colIndex > 1 Then lineText = lineText & ","
                columnName = lo.ListColumns(colIndex).Name
                shouldBlank = ColumnShouldBeBlankRetire(columnName, blankColumns)
                If shouldBlank Then
                    lineText = lineText & EscapeCsvValueRetire("")
                Else
                    lineText = lineText & EscapeCsvValueRetire(CStr(lo.DataBodyRange.Cells(rowIndex, colIndex).Value))
                End If
            Next colIndex
            Print #fileNum, lineText
        Next rowIndex
    End If

    Close #fileNum
    ExportListObjectCsvRetire = True
    Exit Function

FailExport:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    report = "ExportListObjectCsvRetire failed: " & Err.Description
End Function

Private Function ColumnShouldBeBlankRetire(ByVal columnName As String, ByVal blankColumns As Variant) As Boolean
    Dim item As Variant

    On Error GoTo NoColumns
    For Each item In blankColumns
        If StrComp(Trim$(CStr(item)), Trim$(columnName), vbTextCompare) = 0 Then
            ColumnShouldBeBlankRetire = True
            Exit Function
        End If
    Next item
NoColumns:
End Function

Private Function EscapeCsvValueRetire(ByVal valueText As String) As String
    valueText = Replace$(valueText, """", """""")
    If InStr(1, valueText, ",", vbBinaryCompare) > 0 _
       Or InStr(1, valueText, vbCr, vbBinaryCompare) > 0 _
       Or InStr(1, valueText, vbLf, vbBinaryCompare) > 0 Then
        EscapeCsvValueRetire = """" & valueText & """"
    Else
        EscapeCsvValueRetire = valueText
    End If
End Function

Private Function CopyArtifactToArchiveRetire(ByVal sourcePath As String, _
                                             ByVal targetPath As String, _
                                             ByVal relativePath As String, _
                                             ByRef fileList As Collection, _
                                             ByRef report As String) As Boolean
    On Error GoTo FailCopy

    If Not FileExistsRetire(sourcePath) Then
        report = "Artifact not found: " & sourcePath
        Exit Function
    End If

    EnsureFolderRecursiveRetire GetParentFolderRetire(targetPath)
    FileCopy sourcePath, targetPath
    fileList.Add Replace$(relativePath, "/", "\")
    CopyArtifactToArchiveRetire = True
    Exit Function

FailCopy:
    report = "CopyArtifactToArchiveRetire failed for " & sourcePath & ": " & Err.Description
End Function

Private Function CopyPendingOutboxFolderRetire(ByVal sourceFolder As String, _
                                               ByVal targetFolder As String, _
                                               ByRef fileList As Collection, _
                                               ByRef report As String) As Boolean
    Dim fileName As String

    On Error GoTo FailCopy

    sourceFolder = Trim$(Replace$(sourceFolder, "/", "\"))
    If sourceFolder = "" Then
        CopyPendingOutboxFolderRetire = True
        Exit Function
    End If
    If Not FolderExistsRetire(sourceFolder) Then
        CopyPendingOutboxFolderRetire = True
        Exit Function
    End If

    EnsureFolderRecursiveRetire targetFolder
    fileName = Dir$(NormalizeFolderPathRetire(sourceFolder) & "*.*", vbNormal)
    Do While fileName <> ""
        FileCopy NormalizeFolderPathRetire(sourceFolder) & fileName, NormalizeFolderPathRetire(targetFolder) & fileName
        fileList.Add "outbox\" & fileName
        fileName = Dir$
    Loop

    CopyPendingOutboxFolderRetire = True
    Exit Function

FailCopy:
    report = "CopyPendingOutboxFolderRetire failed: " & Err.Description
End Function

Private Function WriteArchiveManifestRetire(ByVal manifestPath As String, _
                                            ByRef spec As RetireMigrateSpec, _
                                            ByVal warehouseName As String, _
                                            ByVal fileList As Collection, _
                                            ByRef report As String) As Boolean
    Dim fileNum As Integer
    Dim i As Long

    On Error GoTo FailWrite

    EnsureFolderRecursiveRetire GetParentFolderRetire(manifestPath)
    fileNum = FreeFile
    Open manifestPath For Output As #fileNum
    Print #fileNum, "{"
    Print #fileNum, "  ""SourceWarehouseId"": """ & EscapeJsonRetire(spec.SourceWarehouseId) & ""","
    Print #fileNum, "  ""WarehouseName"": """ & EscapeJsonRetire(warehouseName) & ""","
    Print #fileNum, "  ""ArchiveTimestampUTC"": """ & EscapeJsonRetire(Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")) & ""","
    Print #fileNum, "  ""OperationMode"": """ & EscapeJsonRetire(OperationModeNameRetire(spec.OperationMode)) & ""","
    Print #fileNum, "  ""AdminUser"": """ & EscapeJsonRetire(spec.AdminUser) & ""","
    Print #fileNum, "  ""ArchiveVersion"": """ & ARCHIVE_VERSION_RETIRE & ""","
    Print #fileNum, "  ""FileList"": ["
    For i = 1 To fileList.Count
        Print #fileNum, "    """ & EscapeJsonRetire(CStr(fileList(i))) & """" & IIf(i < fileList.Count, ",", "")
    Next i
    Print #fileNum, "  ]"
    Print #fileNum, "}"
    Close #fileNum

    WriteArchiveManifestRetire = True
    Exit Function

FailWrite:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    report = "WriteArchiveManifestRetire failed: " & Err.Description
End Function

Private Function EscapeJsonRetire(ByVal textIn As String) As String
    textIn = Replace$(textIn, "\", "\\")
    textIn = Replace$(textIn, Chr$(34), "\" & Chr$(34))
    EscapeJsonRetire = textIn
End Function

Private Function OperationModeNameRetire(ByVal modeValue As RetireMigrateOperationMode) As String
    Select Case modeValue
        Case MODE_ARCHIVE_ONLY
            OperationModeNameRetire = "MODE_ARCHIVE_ONLY"
        Case MODE_ARCHIVE_MIGRATE
            OperationModeNameRetire = "MODE_ARCHIVE_MIGRATE"
        Case MODE_ARCHIVE_RETIRE
            OperationModeNameRetire = "MODE_ARCHIVE_RETIRE"
        Case MODE_ARCHIVE_RETIRE_DELETE
            OperationModeNameRetire = "MODE_ARCHIVE_RETIRE_DELETE"
        Case Else
            OperationModeNameRetire = "MODE_UNKNOWN"
    End Select
End Function

Private Function FileExistsRetire(ByVal filePath As String) As Boolean
    filePath = Trim$(Replace$(filePath, "/", "\"))
    If filePath = "" Then Exit Function
    FileExistsRetire = (Len(Dir$(filePath, vbNormal)) > 0)
End Function

Private Function FolderExistsRetire(ByVal folderPath As String) As Boolean
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) = "\" And Len(folderPath) > 3 Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    FolderExistsRetire = (Len(Dir$(folderPath, vbDirectory)) > 0)
End Function

Private Function GetParentFolderRetire(ByVal pathIn As String) As String
    Dim sepPos As Long

    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    sepPos = InStrRev(pathIn, "\")
    If sepPos > 1 Then GetParentFolderRetire = Left$(pathIn, sepPos - 1)
End Function

Private Function GetFileNameRetire(ByVal fullPath As String) As String
    Dim sepPos As Long

    fullPath = Trim$(Replace$(fullPath, "/", "\"))
    sepPos = InStrRev(fullPath, "\")
    If sepPos > 0 Then
        GetFileNameRetire = Mid$(fullPath, sepPos + 1)
    Else
        GetFileNameRetire = fullPath
    End If
End Function

Private Sub EnsureFolderRecursiveRetire(ByVal folderPath As String)
    Dim parentPath As String

    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    parentPath = GetParentFolderRetire(folderPath)
    If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderRecursiveRetire parentPath

    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
End Sub

Private Sub DeleteFolderRecursiveRetire(ByVal folderPath As String)
    Dim fso As Object

    On Error Resume Next
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then fso.DeleteFolder folderPath, True
    On Error GoTo 0
End Sub

Private Sub LogDiagnosticSafeRetire(ByVal categoryName As String, ByVal detailText As String)
    On Error Resume Next
    modDiagnostics.LogDiagnosticEvent categoryName, detailText
    On Error GoTo 0
End Sub
