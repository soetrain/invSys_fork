Attribute VB_Name = "TestWarehouseRetireArchive"
Option Explicit

Public Function TestWriteArchivePackage_SuccessCreatesAtomicArchive() As Long
    Dim warehouseId As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim archiveFolder As String
    Dim archiveSpec As modWarehouseRetire.RetireMigrateSpec
    Dim detailText As String

    warehouseId = "WHRETARC1"
    runtimeRoot = BuildArchiveTempPathRetire("retire_archive_runtime")
    archiveRoot = BuildArchiveTempPathRetire("retire_archive_output")
    templateRoot = BuildArchiveTempPathRetire("retire_archive_templates")

    On Error GoTo CleanFail
    If Not SetupArchiveRuntimeRetire(warehouseId, runtimeRoot, templateRoot, "admin.archive", "654321") Then GoTo CleanExit

    WriteTextFileRetireArchive runtimeRoot & "\outbox\pending-note.txt", "pending archive artifact"

    archiveSpec.SourceWarehouseId = warehouseId
    archiveSpec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_ONLY
    archiveSpec.AdminUser = "admin.archive"
    archiveSpec.ConfirmedByUser = True
    archiveSpec.ArchiveDestPath = archiveRoot

    If Not modWarehouseRetire.WriteArchivePackage(archiveSpec) Then GoTo CleanExit

    archiveFolder = FindArchiveFolderRetire(archiveRoot, warehouseId)
    If archiveFolder = "" Then GoTo CleanExit
    If FolderExistsRetireArchive(archiveFolder & "_tmp") Then GoTo CleanExit

    If Not AssertArchivePathsRetire(archiveFolder, warehouseId, detailText) Then GoTo CleanExit
    If Not AssertManifestShapeRetire(archiveFolder & "\manifest.json", warehouseId, detailText) Then GoTo CleanExit

    TestWriteArchivePackage_SuccessCreatesAtomicArchive = 1

CleanExit:
    CleanupArchiveRuntimeRetire runtimeRoot, archiveRoot, templateRoot, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestWriteArchivePackage_PartialFailureRollsBackTempArchive() As Long
    Dim warehouseId As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim archiveSpec As modWarehouseRetire.RetireMigrateSpec

    warehouseId = "WHRETARC2"
    runtimeRoot = BuildArchiveTempPathRetire("retire_archive_runtime_fail")
    archiveRoot = BuildArchiveTempPathRetire("retire_archive_output_fail")
    templateRoot = BuildArchiveTempPathRetire("retire_archive_templates_fail")

    On Error GoTo CleanFail
    If Not SetupArchiveRuntimeRetire(warehouseId, runtimeRoot, templateRoot, "admin.archive", "654321") Then GoTo CleanExit

    Kill runtimeRoot & "\" & warehouseId & ".invSys.Data.Inventory.xlsb"

    archiveSpec.SourceWarehouseId = warehouseId
    archiveSpec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_ONLY
    archiveSpec.AdminUser = "admin.archive"
    archiveSpec.ConfirmedByUser = True
    archiveSpec.ArchiveDestPath = archiveRoot

    If modWarehouseRetire.WriteArchivePackage(archiveSpec) Then GoTo CleanExit
    If FindArchiveFolderRetire(archiveRoot, warehouseId) <> "" Then GoTo CleanExit
    If HasArchiveTempFolderRetire(archiveRoot, warehouseId) Then GoTo CleanExit

    TestWriteArchivePackage_PartialFailureRollsBackTempArchive = 1

CleanExit:
    CleanupArchiveRuntimeRetire runtimeRoot, archiveRoot, templateRoot, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestWriteArchivePackage_AuthExportMasksPinHash() As Long
    Dim warehouseId As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim archiveFolder As String
    Dim archiveSpec As modWarehouseRetire.RetireMigrateSpec
    Dim expectedHash As String
    Dim authCsvText As String

    warehouseId = "WHRETARC3"
    runtimeRoot = BuildArchiveTempPathRetire("retire_archive_runtime_auth")
    archiveRoot = BuildArchiveTempPathRetire("retire_archive_output_auth")
    templateRoot = BuildArchiveTempPathRetire("retire_archive_templates_auth")
    expectedHash = modAuth.HashUserCredential("654321")

    On Error GoTo CleanFail
    If Not SetupArchiveRuntimeRetire(warehouseId, runtimeRoot, templateRoot, "admin.archive", "654321") Then GoTo CleanExit

    archiveSpec.SourceWarehouseId = warehouseId
    archiveSpec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_ONLY
    archiveSpec.AdminUser = "admin.archive"
    archiveSpec.ConfirmedByUser = True
    archiveSpec.ArchiveDestPath = archiveRoot

    If Not modWarehouseRetire.WriteArchivePackage(archiveSpec) Then GoTo CleanExit
    archiveFolder = FindArchiveFolderRetire(archiveRoot, warehouseId)
    If archiveFolder = "" Then GoTo CleanExit

    authCsvText = ReadAllTextRetireArchive(archiveFolder & "\auth\tblUsers.csv")
    If authCsvText = "" Then GoTo CleanExit
    If InStr(1, authCsvText, expectedHash, vbTextCompare) > 0 Then GoTo CleanExit
    If InStr(1, authCsvText, "654321", vbTextCompare) > 0 Then GoTo CleanExit
    If InStr(1, authCsvText, "PinHash", vbTextCompare) = 0 Then GoTo CleanExit

    TestWriteArchivePackage_AuthExportMasksPinHash = 1

CleanExit:
    CleanupArchiveRuntimeRetire runtimeRoot, archiveRoot, templateRoot, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function SetupArchiveRuntimeRetire(ByVal warehouseId As String, _
                                           ByVal runtimeRoot As String, _
                                           ByVal templateRoot As String, _
                                           ByVal adminUser As String, _
                                           ByVal passwordText As String) As Boolean
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim wbAuth As Workbook

    On Error GoTo FailSetup

    spec.WarehouseId = warehouseId
    spec.WarehouseName = "Archive Warehouse " & warehouseId
    spec.StationId = "ADM1"
    spec.AdminUser = adminUser
    spec.PathLocal = runtimeRoot
    spec.PathSharePoint = ""

    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    If Not modWarehouseBootstrap.BootstrapWarehouseLocal(spec) Then GoTo CleanExit

    Set wbAuth = Application.Workbooks.Open(runtimeRoot & "\" & warehouseId & ".invSys.Auth.xlsb")
    TestPhase2Helpers.SetUserPinHash wbAuth, adminUser, modAuth.HashUserCredential(passwordText)
    wbAuth.Save

    SetupArchiveRuntimeRetire = True

CleanExit:
    On Error Resume Next
    If Not wbAuth Is Nothing Then wbAuth.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

FailSetup:
    Resume CleanExit
End Function

Private Function AssertArchivePathsRetire(ByVal archiveFolder As String, _
                                          ByVal warehouseId As String, _
                                          ByRef detailText As String) As Boolean
    Dim requiredPaths As Variant
    Dim item As Variant

    requiredPaths = Array( _
        archiveFolder & "\config\tblWarehouseConfig.csv", _
        archiveFolder & "\config\tblStationConfig.csv", _
        archiveFolder & "\auth\tblUsers.csv", _
        archiveFolder & "\auth\tblCapabilities.csv", _
        archiveFolder & "\inventory\" & warehouseId & ".invSys.Data.Inventory.xlsb", _
        archiveFolder & "\snapshots\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb", _
        archiveFolder & "\outbox\" & warehouseId & ".Outbox.Events.xlsb", _
        archiveFolder & "\outbox\pending-note.txt", _
        archiveFolder & "\manifest.json")

    For Each item In requiredPaths
        If Not PathExistsRetireArchive(CStr(item)) Then
            detailText = "Missing archive artifact: " & CStr(item)
            Exit Function
        End If
    Next item

    detailText = "Archive contains config, auth, inventory, snapshot, outbox, and manifest artifacts."
    AssertArchivePathsRetire = True
End Function

Private Function AssertManifestShapeRetire(ByVal manifestPath As String, _
                                           ByVal warehouseId As String, _
                                           ByRef detailText As String) As Boolean
    Dim manifestText As String
    Dim normalizedText As String

    manifestText = ReadAllTextRetireArchive(manifestPath)
    If manifestText = "" Then
        detailText = "Manifest content was empty."
        Exit Function
    End If
    If InStr(1, manifestText, """SourceWarehouseId"": """ & warehouseId & """", vbTextCompare) = 0 Then
        detailText = "Manifest SourceWarehouseId missing."
        Exit Function
    End If
    If InStr(1, manifestText, """ArchiveVersion"": ""1.0""", vbTextCompare) = 0 Then
        detailText = "Manifest ArchiveVersion missing."
        Exit Function
    End If
    If InStr(1, manifestText, """FileList"": [", vbTextCompare) = 0 Then
        detailText = "Manifest FileList missing."
        Exit Function
    End If
    If InStr(1, manifestText, "outbox\\pending-note.txt", vbTextCompare) = 0 Then
        detailText = "Manifest did not include pending outbox artifact."
        Exit Function
    End If
    normalizedText = TrimJsonWhitespaceRetire(manifestText)
    If Left$(normalizedText, 1) <> "{" Or Right$(normalizedText, 1) <> "}" Then
        detailText = "Manifest was not JSON-shaped."
        Exit Function
    End If

    detailText = "Manifest contains required JSON keys and file list entries."
    AssertManifestShapeRetire = True
End Function

Private Function FindArchiveFolderRetire(ByVal archiveRoot As String, ByVal warehouseId As String) As String
    Dim candidate As String

    candidate = Dir$(archiveRoot & "\" & warehouseId & "_archive_*", vbDirectory)
    Do While candidate <> ""
        If candidate <> "." And candidate <> ".." Then
            If InStr(1, candidate, "_tmp", vbTextCompare) = 0 Then
                FindArchiveFolderRetire = archiveRoot & "\" & candidate
                Exit Function
            End If
        End If
        candidate = Dir$
    Loop
End Function

Private Function HasArchiveTempFolderRetire(ByVal archiveRoot As String, ByVal warehouseId As String) As Boolean
    Dim candidate As String

    candidate = Dir$(archiveRoot & "\" & warehouseId & "_archive_*_tmp", vbDirectory)
    Do While candidate <> ""
        If candidate <> "." And candidate <> ".." Then
            HasArchiveTempFolderRetire = True
            Exit Function
        End If
        candidate = Dir$
    Loop
End Function

Private Function PathExistsRetireArchive(ByVal pathIn As String) As Boolean
    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    If pathIn = "" Then Exit Function
    PathExistsRetireArchive = (Len(Dir$(pathIn, vbDirectory)) > 0)
    If Not PathExistsRetireArchive Then PathExistsRetireArchive = (Len(Dir$(pathIn, vbNormal)) > 0)
End Function

Private Function FolderExistsRetireArchive(ByVal folderPath As String) As Boolean
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function
    FolderExistsRetireArchive = (Len(Dir$(folderPath, vbDirectory)) > 0)
End Function

Private Function ReadAllTextRetireArchive(ByVal filePath As String) As String
    Dim fileNum As Integer

    On Error GoTo CleanFail
    If Len(Dir$(filePath, vbNormal)) = 0 Then Exit Function
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    ReadAllTextRetireArchive = Input$(LOF(fileNum), fileNum)

CleanExit:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    Exit Function

CleanFail:
    Resume CleanExit
End Function

Private Function BuildArchiveTempPathRetire(ByVal leafName As String) As String
    BuildArchiveTempPathRetire = Environ$("TEMP") & "\" & leafName & "_" & _
                                 Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 1000))
End Function

Private Function TrimJsonWhitespaceRetire(ByVal textIn As String) As String
    Do While Len(textIn) > 0
        Select Case Left$(textIn, 1)
            Case " ", vbTab, vbCr, vbLf
                textIn = Mid$(textIn, 2)
            Case Else
                Exit Do
        End Select
    Loop

    Do While Len(textIn) > 0
        Select Case Right$(textIn, 1)
            Case " ", vbTab, vbCr, vbLf
                textIn = Left$(textIn, Len(textIn) - 1)
            Case Else
                Exit Do
        End Select
    Loop

    TrimJsonWhitespaceRetire = textIn
End Function

Private Sub WriteTextFileRetireArchive(ByVal filePath As String, ByVal textIn As String)
    Dim fileNum As Integer
    Dim folderPath As String

    folderPath = Left$(filePath, InStrRev(filePath, "\") - 1)
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then MkDir folderPath

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, textIn
    Close #fileNum
End Sub

Private Sub CleanupArchiveRuntimeRetire(ByVal runtimeRoot As String, _
                                        ByVal archiveRoot As String, _
                                        ByVal templateRoot As String, _
                                        ByVal warehouseId As String)
    On Error Resume Next
    CloseWorkbookByNameArchiveRetire warehouseId & ".invSys.Config.xlsb"
    CloseWorkbookByNameArchiveRetire warehouseId & ".invSys.Auth.xlsb"
    CloseWorkbookByNameArchiveRetire warehouseId & ".invSys.Data.Inventory.xlsb"
    CloseWorkbookByNameArchiveRetire warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    CloseWorkbookByNameArchiveRetire warehouseId & ".Outbox.Events.xlsb"
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    DeleteFolderRecursiveArchiveRetire archiveRoot
    DeleteFolderRecursiveArchiveRetire runtimeRoot
    DeleteFolderRecursiveArchiveRetire templateRoot
    On Error GoTo 0
End Sub

Private Sub CloseWorkbookByNameArchiveRetire(ByVal workbookName As String)
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=False
            Exit Sub
        End If
    Next wb
End Sub

Private Sub DeleteFolderRecursiveArchiveRetire(ByVal folderPath As String)
    Dim fso As Object

    On Error Resume Next
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then fso.DeleteFolder folderPath, True
    On Error GoTo 0
End Sub
