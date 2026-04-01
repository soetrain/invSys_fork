Attribute VB_Name = "TestWarehouseRetireMigration"
Option Explicit

Public Function TestMigrateInventoryToTarget_SuccessAppendsInventoryAndTracesSource() As Long
    Dim sourceWh As String
    Dim targetWh As String
    Dim runtimeBase As String
    Dim sourceRoot As String
    Dim targetRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim wbInv As Workbook
    Dim loSku As ListObject
    Dim loLog As ListObject
    Dim rowIndex As Long

    sourceWh = "WHRETMIG1A"
    targetWh = "WHRETMIG1B"
    runtimeBase = BuildTempRootRetireMigration("retire_migrate_success")
    sourceRoot = runtimeBase & "\src"
    targetRoot = runtimeBase & "\tgt"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupMigrationRuntimeRetire(sourceWh, sourceRoot, templateRoot, "admin.migrate", "654321") Then GoTo CleanExit
    If Not SetupMigrationRuntimeRetire(targetWh, targetRoot, templateRoot, "admin.migrate", "654321") Then GoTo CleanExit
    If Not SeedInventoryStateRetire(sourceWh, sourceRoot, "EVT-SRC-001", "admin.migrate", 5, "A1", "source-seed") Then GoTo CleanExit
    If Not SeedInventoryStateRetire(targetWh, targetRoot, "EVT-TGT-001", "admin.migrate", 2, "B1", "target-seed") Then GoTo CleanExit
    If Not AddSourceOnlyAuthUserRetire(sourceWh, sourceRoot) Then GoTo CleanExit

    spec.SourceWarehouseId = sourceWh
    spec.TargetWarehouseId = targetWh
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    spec.AdminUser = "admin.migrate"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    If Not modWarehouseRetire.WriteArchivePackage(spec) Then GoTo CleanExit

    modRuntimeWorkbooks.SetCoreDataRootOverride targetRoot
    If Not modWarehouseRetire.MigrateInventoryToTarget(spec) Then GoTo CleanExit

    Set wbInv = OpenWorkbookIfNeededRetire(targetRoot & "\" & targetWh & ".invSys.Data.Inventory.xlsb")
    If wbInv Is Nothing Then GoTo CleanExit
    Set loSku = wbInv.Worksheets("SkuBalance").ListObjects("tblSkuBalance")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    If loSku Is Nothing Or loLog Is Nothing Then GoTo CleanExit

    rowIndex = FindRowByValueRetire(loSku, "SKU", "SKU-MIG-001")
    If rowIndex = 0 Then GoTo CleanExit
    If CDbl(TestPhase2Helpers.GetRowValue(loSku, rowIndex, "QtyOnHand")) <> 7 Then GoTo CleanExit

    rowIndex = FindLastInventoryLogRowRetire(loLog, "MIGRATION_SEED", sourceWh)
    If rowIndex = 0 Then GoTo CleanExit
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loLog, rowIndex, "MigrationSourceId")), sourceWh, vbTextCompare) <> 0 Then GoTo CleanExit

    TestMigrateInventoryToTarget_SuccessAppendsInventoryAndTracesSource = 1

CleanExit:
    CleanupMigrationScenarioRetire runtimeBase, sourceWh, targetWh
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestMigrateInventoryToTarget_RejectsMissingArchiveManifest() As Long
    Dim targetWh As String
    Dim runtimeBase As String
    Dim targetRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec

    targetWh = "WHRETMIG2B"
    runtimeBase = BuildTempRootRetireMigration("retire_migrate_no_archive")
    targetRoot = runtimeBase & "\tgt"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupMigrationRuntimeRetire(targetWh, targetRoot, templateRoot, "admin.migrate", "654321") Then GoTo CleanExit

    spec.SourceWarehouseId = "WHRETMIG2A"
    spec.TargetWarehouseId = targetWh
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    spec.AdminUser = "admin.migrate"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = runtimeBase & "\archive"

    modRuntimeWorkbooks.SetCoreDataRootOverride targetRoot
    If modWarehouseRetire.MigrateInventoryToTarget(spec) Then GoTo CleanExit
    If InStr(1, modWarehouseRetire.GetLastWarehouseRetireReport(), "Archive manifest not found", vbTextCompare) = 0 Then GoTo CleanExit

    TestMigrateInventoryToTarget_RejectsMissingArchiveManifest = 1

CleanExit:
    CleanupMigrationScenarioRetire runtimeBase, vbNullString, targetWh
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestMigrateInventoryToTarget_RejectsMissingTargetWarehouse() As Long
    Dim sourceWh As String
    Dim runtimeBase As String
    Dim sourceRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec

    sourceWh = "WHRETMIG3A"
    runtimeBase = BuildTempRootRetireMigration("retire_migrate_no_target")
    sourceRoot = runtimeBase & "\src"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupMigrationRuntimeRetire(sourceWh, sourceRoot, templateRoot, "admin.migrate", "654321") Then GoTo CleanExit
    If Not SeedInventoryStateRetire(sourceWh, sourceRoot, "EVT-SRC-003", "admin.migrate", 4, "A1", "source-seed") Then GoTo CleanExit

    spec.SourceWarehouseId = sourceWh
    spec.TargetWarehouseId = "WHRETMIG3B"
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    spec.AdminUser = "admin.migrate"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    If Not modWarehouseRetire.WriteArchivePackage(spec) Then GoTo CleanExit

    modRuntimeWorkbooks.ClearCoreDataRootOverride
    If modWarehouseRetire.MigrateInventoryToTarget(spec) Then GoTo CleanExit
    If InStr(1, modWarehouseRetire.GetLastWarehouseRetireReport(), "Target warehouse runtime not found", vbTextCompare) = 0 Then GoTo CleanExit

    TestMigrateInventoryToTarget_RejectsMissingTargetWarehouse = 1

CleanExit:
    CleanupMigrationScenarioRetire runtimeBase, sourceWh, vbNullString
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestMigrateInventoryToTarget_DoesNotCopyAuthIdentities() As Long
    Dim sourceWh As String
    Dim targetWh As String
    Dim runtimeBase As String
    Dim sourceRoot As String
    Dim targetRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim wbAuth As Workbook
    Dim loUsers As ListObject

    sourceWh = "WHRETMIG4A"
    targetWh = "WHRETMIG4B"
    runtimeBase = BuildTempRootRetireMigration("retire_migrate_auth")
    sourceRoot = runtimeBase & "\src"
    targetRoot = runtimeBase & "\tgt"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupMigrationRuntimeRetire(sourceWh, sourceRoot, templateRoot, "admin.migrate", "654321") Then GoTo CleanExit
    If Not SetupMigrationRuntimeRetire(targetWh, targetRoot, templateRoot, "admin.migrate", "654321") Then GoTo CleanExit
    If Not SeedInventoryStateRetire(sourceWh, sourceRoot, "EVT-SRC-004", "admin.migrate", 3, "A1", "source-seed") Then GoTo CleanExit
    If Not AddSourceOnlyAuthUserRetire(sourceWh, sourceRoot) Then GoTo CleanExit

    spec.SourceWarehouseId = sourceWh
    spec.TargetWarehouseId = targetWh
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    spec.AdminUser = "admin.migrate"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    If Not modWarehouseRetire.WriteArchivePackage(spec) Then GoTo CleanExit

    modRuntimeWorkbooks.SetCoreDataRootOverride targetRoot
    If Not modWarehouseRetire.MigrateInventoryToTarget(spec) Then GoTo CleanExit

    Set wbAuth = OpenWorkbookIfNeededRetire(targetRoot & "\" & targetWh & ".invSys.Auth.xlsb")
    If wbAuth Is Nothing Then GoTo CleanExit
    Set loUsers = wbAuth.Worksheets("Users").ListObjects("tblUsers")
    If loUsers Is Nothing Then GoTo CleanExit
    If FindRowByValueRetire(loUsers, "UserId", "source.only") <> 0 Then GoTo CleanExit

    TestMigrateInventoryToTarget_DoesNotCopyAuthIdentities = 1

CleanExit:
    CleanupMigrationScenarioRetire runtimeBase, sourceWh, targetWh
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Public Function TestMigrateInventoryToTarget_PreservesTargetConfigIdentity() As Long
    Dim sourceWh As String
    Dim targetWh As String
    Dim runtimeBase As String
    Dim sourceRoot As String
    Dim targetRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim wbCfg As Workbook
    Dim loWh As ListObject
    Dim loSt As ListObject

    sourceWh = "WHRETMIG5A"
    targetWh = "WHRETMIG5B"
    runtimeBase = BuildTempRootRetireMigration("retire_migrate_config")
    sourceRoot = runtimeBase & "\src"
    targetRoot = runtimeBase & "\tgt"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupMigrationRuntimeRetire(sourceWh, sourceRoot, templateRoot, "admin.migrate", "654321") Then GoTo CleanExit
    If Not SetupMigrationRuntimeRetire(targetWh, targetRoot, templateRoot, "admin.migrate", "654321") Then GoTo CleanExit
    If Not RenameTargetConfigIdentityRetire(targetWh, targetRoot, "Target Warehouse Name", "ADM1") Then GoTo CleanExit
    If Not SeedInventoryStateRetire(sourceWh, sourceRoot, "EVT-SRC-005", "admin.migrate", 6, "A1", "source-seed") Then GoTo CleanExit

    spec.SourceWarehouseId = sourceWh
    spec.TargetWarehouseId = targetWh
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    spec.AdminUser = "admin.migrate"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    If Not modWarehouseRetire.WriteArchivePackage(spec) Then GoTo CleanExit

    modRuntimeWorkbooks.SetCoreDataRootOverride targetRoot
    If Not modWarehouseRetire.MigrateInventoryToTarget(spec) Then GoTo CleanExit

    Set wbCfg = OpenWorkbookIfNeededRetire(targetRoot & "\" & targetWh & ".invSys.Config.xlsb")
    If wbCfg Is Nothing Then GoTo CleanExit
    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")
    If loWh Is Nothing Or loSt Is Nothing Then GoTo CleanExit
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loWh, 1, "WarehouseId")), targetWh, vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loWh, 1, "WarehouseName")), "Target Warehouse Name", vbTextCompare) <> 0 Then GoTo CleanExit
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loSt, 1, "StationId")), "ADM1", vbTextCompare) <> 0 Then GoTo CleanExit

    TestMigrateInventoryToTarget_PreservesTargetConfigIdentity = 1

CleanExit:
    CleanupMigrationScenarioRetire runtimeBase, sourceWh, targetWh
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function SetupMigrationRuntimeRetire(ByVal warehouseId As String, _
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

    SetupMigrationRuntimeRetire = True

CleanExit:
    On Error Resume Next
    If Not wbAuth Is Nothing Then wbAuth.Close SaveChanges:=False
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    On Error GoTo 0
    Exit Function

FailSetup:
    Resume CleanExit
End Function

Private Function SeedInventoryStateRetire(ByVal warehouseId As String, _
                                          ByVal runtimeRoot As String, _
                                          ByVal eventId As String, _
                                          ByVal userId As String, _
                                          ByVal qty As Double, _
                                          ByVal locationVal As String, _
                                          ByVal noteVal As String) As Boolean
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim report As String

    On Error GoTo FailSeed

    Set wbInv = OpenWorkbookIfNeededRetire(runtimeRoot & "\" & warehouseId & ".invSys.Data.Inventory.xlsb")
    If wbInv Is Nothing Then GoTo CleanExit

    Set evt = TestPhase2Helpers.CreateReceiveEvent(eventId, warehouseId, "ADM1", userId, "SKU-MIG-001", qty, locationVal, noteVal, Now, "seed-inbox")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-" & eventId, statusOut, errorCode, errorMessage) Then GoTo CleanExit
    If Not GenerateWarehouseSnapshot(warehouseId, wbInv, runtimeRoot & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb", Nothing, report) Then GoTo CleanExit
    SeedInventoryStateRetire = True

CleanExit:
    CloseWorkbookIfOpenRetire wbInv
    Exit Function

FailSeed:
    Resume CleanExit
End Function

Private Function AddSourceOnlyAuthUserRetire(ByVal warehouseId As String, ByVal runtimeRoot As String) As Boolean
    Dim wbAuth As Workbook

    On Error GoTo FailAdd

    Set wbAuth = OpenWorkbookIfNeededRetire(runtimeRoot & "\" & warehouseId & ".invSys.Auth.xlsb")
    If wbAuth Is Nothing Then GoTo CleanExit
    TestPhase2Helpers.AddCapability wbAuth, "source.only", "ADMIN_MAINT", warehouseId, "ADM1", "ACTIVE"
    TestPhase2Helpers.SetUserPinHash wbAuth, "source.only", modAuth.HashUserCredential("999999")
    wbAuth.Save
    AddSourceOnlyAuthUserRetire = True

CleanExit:
    CloseWorkbookIfOpenRetire wbAuth
    Exit Function

FailAdd:
    Resume CleanExit
End Function

Private Function RenameTargetConfigIdentityRetire(ByVal warehouseId As String, _
                                                  ByVal runtimeRoot As String, _
                                                  ByVal warehouseName As String, _
                                                  ByVal stationId As String) As Boolean
    Dim wbCfg As Workbook

    On Error GoTo FailRename

    Set wbCfg = OpenWorkbookIfNeededRetire(runtimeRoot & "\" & warehouseId & ".invSys.Config.xlsb")
    If wbCfg Is Nothing Then GoTo CleanExit
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "WarehouseName", warehouseName
    TestPhase2Helpers.SetStationConfigValue wbCfg, "StationId", stationId
    wbCfg.Save
    RenameTargetConfigIdentityRetire = True

CleanExit:
    CloseWorkbookIfOpenRetire wbCfg
    Exit Function

FailRename:
    Resume CleanExit
End Function

Private Function OpenWorkbookIfNeededRetire(ByVal fullPath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set OpenWorkbookIfNeededRetire = wb
            Exit Function
        End If
    Next wb

    If Len(Dir$(fullPath, vbNormal)) = 0 Then Exit Function
    Set OpenWorkbookIfNeededRetire = Application.Workbooks.Open(fullPath)
End Function

Private Function FindRowByValueRetire(ByVal lo As ListObject, ByVal columnName As String, ByVal matchValue As String) As Long
    Dim rowIndex As Long
    Dim colIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    colIndex = lo.ListColumns(columnName).Index
    If colIndex <= 0 Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(rowIndex, colIndex).Value), matchValue, vbTextCompare) = 0 Then
            FindRowByValueRetire = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function FindLastInventoryLogRowRetire(ByVal lo As ListObject, _
                                               ByVal eventType As String, _
                                               ByVal migrationSourceId As String) As Long
    Dim rowIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = lo.ListRows.Count To 1 Step -1
        If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, rowIndex, "EventType")), eventType, vbTextCompare) = 0 Then
            If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, rowIndex, "MigrationSourceId")), migrationSourceId, vbTextCompare) = 0 Then
                FindLastInventoryLogRowRetire = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Private Function BuildTempRootRetireMigration(ByVal leafName As String) As String
    BuildTempRootRetireMigration = Environ$("TEMP") & "\" & leafName & "_" & _
                                   Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 1000))
End Function

Private Sub CleanupMigrationScenarioRetire(ByVal runtimeBase As String, _
                                           ByVal sourceWarehouseId As String, _
                                           ByVal targetWarehouseId As String)
    On Error Resume Next
    If sourceWarehouseId <> "" Then
        CloseWorkbookByNameRetire sourceWarehouseId & ".invSys.Config.xlsb"
        CloseWorkbookByNameRetire sourceWarehouseId & ".invSys.Auth.xlsb"
        CloseWorkbookByNameRetire sourceWarehouseId & ".invSys.Data.Inventory.xlsb"
        CloseWorkbookByNameRetire sourceWarehouseId & ".invSys.Snapshot.Inventory.xlsb"
        CloseWorkbookByNameRetire sourceWarehouseId & ".Outbox.Events.xlsb"
    End If
    If targetWarehouseId <> "" Then
        CloseWorkbookByNameRetire targetWarehouseId & ".invSys.Config.xlsb"
        CloseWorkbookByNameRetire targetWarehouseId & ".invSys.Auth.xlsb"
        CloseWorkbookByNameRetire targetWarehouseId & ".invSys.Data.Inventory.xlsb"
        CloseWorkbookByNameRetire targetWarehouseId & ".invSys.Snapshot.Inventory.xlsb"
        CloseWorkbookByNameRetire targetWarehouseId & ".Outbox.Events.xlsb"
    End If
    CloseWorkbookByNameRetire "invSys.Inbox.Production.ADM1.xlsb"
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    DeleteFolderRecursiveRetireMigration runtimeBase
    On Error GoTo 0
End Sub

Private Sub CloseWorkbookByNameRetire(ByVal workbookName As String)
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=False
            Exit Sub
        End If
    Next wb
End Sub

Private Sub CloseWorkbookIfOpenRetire(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub DeleteFolderRecursiveRetireMigration(ByVal folderPath As String)
    Dim fso As Object

    On Error Resume Next
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then fso.DeleteFolder folderPath, True
    On Error GoTo 0
End Sub
