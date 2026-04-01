Attribute VB_Name = "TestWarehouseRetireLifecycle"
Option Explicit

Public Function TestRetireSourceWarehouse_WritesRetirementMarker() As Long
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim wbCfg As Workbook
    Dim loWh As ListObject

    warehouseId = "WHRETLC1"
    runtimeBase = BuildTempRootRetireLifecycle("retire_lifecycle_marker")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireLifecycleRuntime(warehouseId, runtimeRoot, templateRoot, "admin.retire", "654321") Then GoTo CleanExit
    If Not SeedRetireLifecycleInventory(warehouseId, runtimeRoot, "admin.retire", 4) Then GoTo CleanExit

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
    spec.AdminUser = "admin.retire"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    If Not modWarehouseRetire.WriteArchivePackage(spec) Then GoTo CleanExit
    If Not modWarehouseRetire.RetireSourceWarehouse(spec) Then GoTo CleanExit

    Set wbCfg = OpenWorkbookIfNeededLifecycle(runtimeRoot & "\" & warehouseId & ".invSys.Config.xlsb")
    If wbCfg Is Nothing Then GoTo CleanExit
    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    If loWh Is Nothing Then GoTo CleanExit
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loWh, 1, "WarehouseStatus")), "RETIRED", vbTextCompare) <> 0 Then GoTo CleanExit
    If Not IsDate(TestPhase2Helpers.GetRowValue(loWh, 1, "RetiredAtUTC")) Then GoTo CleanExit

    TestRetireSourceWarehouse_WritesRetirementMarker = 1

CleanExit:
    CleanupRetireLifecycleRuntime runtimeBase, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRetireSourceWarehouse_WritesValidTombstoneJson() As Long
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim tombstonePath As String
    Dim tombstoneText As String
    Dim normalizedText As String

    warehouseId = "WHRETLC2"
    runtimeBase = BuildTempRootRetireLifecycle("retire_lifecycle_tombstone")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireLifecycleRuntime(warehouseId, runtimeRoot, templateRoot, "admin.retire", "654321") Then GoTo CleanExit
    If Not SeedRetireLifecycleInventory(warehouseId, runtimeRoot, "admin.retire", 5) Then GoTo CleanExit

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
    spec.AdminUser = "admin.retire"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    If Not modWarehouseRetire.WriteArchivePackage(spec) Then GoTo CleanExit
    If Not modWarehouseRetire.RetireSourceWarehouse(spec) Then GoTo CleanExit

    tombstonePath = archiveRoot & "\" & warehouseId & ".tombstone.json"
    tombstoneText = ReadAllTextLifecycle(tombstonePath)
    If tombstoneText = "" Then GoTo CleanExit
    If InStr(1, tombstoneText, """WarehouseId"": """ & warehouseId & """", vbTextCompare) = 0 Then GoTo CleanExit
    If InStr(1, tombstoneText, """RetiredByUser"": ""admin.retire""", vbTextCompare) = 0 Then GoTo CleanExit
    If InStr(1, tombstoneText, """ArchivePath"": """, vbTextCompare) = 0 Then GoTo CleanExit
    normalizedText = TrimJsonWhitespaceLifecycle(tombstoneText)
    If Left$(normalizedText, 1) <> "{" Or Right$(normalizedText, 1) <> "}" Then GoTo CleanExit

    TestRetireSourceWarehouse_WritesValidTombstoneJson = 1

CleanExit:
    CleanupRetireLifecycleRuntime runtimeBase, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestRetireSourceWarehouse_SharePointUnavailableDoesNotBlockRetirement() As Long
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim tombstonePath As String
    Dim report As String

    warehouseId = "WHRETLC3"
    runtimeBase = BuildTempRootRetireLifecycle("retire_lifecycle_sharepoint")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireLifecycleRuntime(warehouseId, runtimeRoot, templateRoot, "admin.retire", "654321") Then GoTo CleanExit
    If Not SeedRetireLifecycleInventory(warehouseId, runtimeRoot, "admin.retire", 6) Then GoTo CleanExit
    If Not SetWarehouseSharePointPathLifecycle(warehouseId, runtimeRoot, "Z:\invSys-unavailable") Then GoTo CleanExit

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
    spec.AdminUser = "admin.retire"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    spec.PublishTombstone = True
    If Not modWarehouseRetire.WriteArchivePackage(spec) Then GoTo CleanExit
    If Not modWarehouseRetire.RetireSourceWarehouse(spec) Then GoTo CleanExit

    report = modWarehouseRetire.GetLastWarehouseRetireReport()
    tombstonePath = archiveRoot & "\" & warehouseId & ".tombstone.json"
    If Len(Dir$(tombstonePath, vbNormal)) = 0 Then GoTo CleanExit
    If InStr(1, report, "PublishWarning=", vbTextCompare) = 0 Then GoTo CleanExit

    TestRetireSourceWarehouse_SharePointUnavailableDoesNotBlockRetirement = 1

CleanExit:
    CleanupRetireLifecycleRuntime runtimeBase, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDeleteLocalRuntime_RejectsWithoutTombstone() As Long
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec

    warehouseId = "WHRETLC4"
    runtimeBase = BuildTempRootRetireLifecycle("retire_lifecycle_delete_no_tombstone")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireLifecycleRuntime(warehouseId, runtimeRoot, templateRoot, "admin.retire", "654321") Then GoTo CleanExit

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
    spec.AdminUser = "admin.retire"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot

    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    If modWarehouseRetire.DeleteLocalRuntime(spec) Then GoTo CleanExit
    If InStr(1, modWarehouseRetire.GetLastWarehouseRetireReport(), "Retirement tombstone not found", vbTextCompare) = 0 Then GoTo CleanExit
    If Len(Dir$(runtimeRoot, vbDirectory)) = 0 Then GoTo CleanExit

    TestDeleteLocalRuntime_RejectsWithoutTombstone = 1

CleanExit:
    CleanupRetireLifecycleRuntime runtimeBase, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestDeleteLocalRuntime_RejectsWithoutConfirmation() As Long
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec

    warehouseId = "WHRETLC5"
    runtimeBase = BuildTempRootRetireLifecycle("retire_lifecycle_delete_unconfirmed")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireLifecycleRuntime(warehouseId, runtimeRoot, templateRoot, "admin.retire", "654321") Then GoTo CleanExit
    If Not SeedRetireLifecycleInventory(warehouseId, runtimeRoot, "admin.retire", 3) Then GoTo CleanExit

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
    spec.AdminUser = "admin.retire"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    If Not modWarehouseRetire.WriteArchivePackage(spec) Then GoTo CleanExit
    If Not modWarehouseRetire.RetireSourceWarehouse(spec) Then GoTo CleanExit

    spec.ConfirmedByUser = False
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    If modWarehouseRetire.DeleteLocalRuntime(spec) Then GoTo CleanExit
    If InStr(1, modWarehouseRetire.GetLastWarehouseRetireReport(), "ConfirmedByUser = True", vbTextCompare) = 0 Then GoTo CleanExit
    If Len(Dir$(runtimeRoot, vbDirectory)) = 0 Then GoTo CleanExit

    TestDeleteLocalRuntime_RejectsWithoutConfirmation = 1

CleanExit:
    CleanupRetireLifecycleRuntime runtimeBase, warehouseId
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function SetupRetireLifecycleRuntime(ByVal warehouseId As String, _
                                             ByVal runtimeRoot As String, _
                                             ByVal templateRoot As String, _
                                             ByVal adminUser As String, _
                                             ByVal passwordText As String) As Boolean
    Dim spec As modWarehouseBootstrap.WarehouseSpec
    Dim wbAuth As Workbook

    On Error GoTo FailSetup

    spec.WarehouseId = warehouseId
    spec.WarehouseName = "Warehouse " & warehouseId
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
    SetupRetireLifecycleRuntime = True

CleanExit:
    On Error Resume Next
    If Not wbAuth Is Nothing Then wbAuth.Close SaveChanges:=False
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    On Error GoTo 0
    Exit Function

FailSetup:
    Resume CleanExit
End Function

Private Function SeedRetireLifecycleInventory(ByVal warehouseId As String, _
                                              ByVal runtimeRoot As String, _
                                              ByVal userId As String, _
                                              ByVal qty As Double) As Boolean
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim report As String

    On Error GoTo FailSeed

    Set wbInv = OpenWorkbookIfNeededLifecycle(runtimeRoot & "\" & warehouseId & ".invSys.Data.Inventory.xlsb")
    If wbInv Is Nothing Then GoTo CleanExit

    Set evt = TestPhase2Helpers.CreateReceiveEvent("EVT-" & warehouseId & "-RETIRE", warehouseId, "ADM1", userId, "SKU-RETIRE-001", qty, "A1", "retire-seed", Now, "seed-inbox")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-" & warehouseId, statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If Not GenerateWarehouseSnapshot(warehouseId, wbInv, runtimeRoot & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb", Nothing, report) Then GoTo CleanExit
    SeedRetireLifecycleInventory = True

CleanExit:
    CloseWorkbookIfOpenLifecycle wbInv
    Exit Function

FailSeed:
    Resume CleanExit
End Function

Private Function SetWarehouseSharePointPathLifecycle(ByVal warehouseId As String, _
                                                     ByVal runtimeRoot As String, _
                                                     ByVal sharePointPath As String) As Boolean
    Dim wbCfg As Workbook

    On Error GoTo FailSet

    Set wbCfg = OpenWorkbookIfNeededLifecycle(runtimeRoot & "\" & warehouseId & ".invSys.Config.xlsb")
    If wbCfg Is Nothing Then GoTo CleanExit
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathSharePointRoot", sharePointPath
    wbCfg.Save
    SetWarehouseSharePointPathLifecycle = True

CleanExit:
    CloseWorkbookIfOpenLifecycle wbCfg
    Exit Function

FailSet:
    Resume CleanExit
End Function

Private Function OpenWorkbookIfNeededLifecycle(ByVal fullPath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set OpenWorkbookIfNeededLifecycle = wb
            Exit Function
        End If
    Next wb

    If Len(Dir$(fullPath, vbNormal)) = 0 Then Exit Function
    Set OpenWorkbookIfNeededLifecycle = Application.Workbooks.Open(fullPath)
End Function

Private Function ReadAllTextLifecycle(ByVal filePath As String) As String
    Dim fileNum As Integer

    On Error GoTo CleanFail
    If Len(Dir$(filePath, vbNormal)) = 0 Then Exit Function
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    ReadAllTextLifecycle = Input$(LOF(fileNum), fileNum)

CleanExit:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    Exit Function

CleanFail:
    Resume CleanExit
End Function

Private Function TrimJsonWhitespaceLifecycle(ByVal textIn As String) As String
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

    TrimJsonWhitespaceLifecycle = textIn
End Function

Private Function BuildTempRootRetireLifecycle(ByVal leafName As String) As String
    BuildTempRootRetireLifecycle = Environ$("TEMP") & "\" & leafName & "_" & _
                                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 1000))
End Function

Private Sub CleanupRetireLifecycleRuntime(ByVal runtimeBase As String, ByVal warehouseId As String)
    On Error Resume Next
    If warehouseId <> "" Then
        CloseWorkbookByNameLifecycle warehouseId & ".invSys.Config.xlsb"
        CloseWorkbookByNameLifecycle warehouseId & ".invSys.Auth.xlsb"
        CloseWorkbookByNameLifecycle warehouseId & ".invSys.Data.Inventory.xlsb"
        CloseWorkbookByNameLifecycle warehouseId & ".invSys.Snapshot.Inventory.xlsb"
        CloseWorkbookByNameLifecycle warehouseId & ".Outbox.Events.xlsb"
    End If
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    DeleteFolderRecursiveLifecycle runtimeBase
    On Error GoTo 0
End Sub

Private Sub CloseWorkbookByNameLifecycle(ByVal workbookName As String)
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=False
            Exit Sub
        End If
    Next wb
End Sub

Private Sub CloseWorkbookIfOpenLifecycle(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub DeleteFolderRecursiveLifecycle(ByVal folderPath As String)
    Dim fso As Object

    On Error Resume Next
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then fso.DeleteFolder folderPath, True
    On Error GoTo 0
End Sub
