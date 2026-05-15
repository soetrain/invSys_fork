Attribute VB_Name = "test_RetireMigrate"
Option Explicit

Private mCaseNames() As String
Private mCaseResults() As String
Private mCaseDetails() As String
Private mCaseCount As Long
Private mSummary As String
Private mLastSeedFailure As String

Public Function TestRetireMigrate_EndToEndLifecycle() As Long
    Dim detailText As String

    On Error GoTo FailTest

    ResetRetireMigrateEvidence

    RecordRetireMigrateCase "ArchiveOnly.SourceUntouched", RunArchiveOnlyCase(detailText), detailText
    RecordRetireMigrateCase "ArchiveMigrate.TargetSeededNoIdentityBleed", RunArchiveMigrateCase(detailText), detailText
    RecordRetireMigrateCase "ArchiveRetire.TombstoneAndStatus", RunArchiveRetireCase(detailText), detailText
    RecordRetireMigrateCase "ArchiveRetireDelete.RuntimeRemoved", RunArchiveRetireDeleteCase(detailText), detailText
    RecordRetireMigrateCase "RetiredReuse.Rejected", RunRetiredReuseRejectedCase(detailText), detailText
    RecordRetireMigrateCase "DeleteNoManifest.Rejected", RunDeleteWithoutArchiveManifestCase(detailText), detailText
    RecordRetireMigrateCase "DeleteNoConfirmation.Rejected", RunDeleteWithoutConfirmationCase(detailText), detailText
    RecordRetireMigrateCase "RetireSharePointUnavailable.WarningOnly", RunSharePointUnavailableRetireCase(detailText), detailText

    If AllRetireMigrateCasesPassed() Then
        mSummary = "Retire/migrate lifecycle cases passed for archive-only, migrate, retire, delete, reuse rejection, and safety guards."
        TestRetireMigrate_EndToEndLifecycle = 1
    Else
        mSummary = "One or more retire/migrate lifecycle cases failed."
    End If

    Exit Function

FailTest:
    RecordRetireMigrateCase "Harness.Exception", False, Err.Description
    mSummary = "Retire/migrate lifecycle raised an unexpected exception."
End Function

Public Function GetRetireMigrateContextPacked() As String
    GetRetireMigrateContextPacked = "Summary=" & SafeRetireMigrateText(mSummary)
End Function

Public Function GetRetireMigrateEvidenceRows() As String
    Dim i As Long

    For i = 1 To mCaseCount
        If Len(GetRetireMigrateEvidenceRows) > 0 Then GetRetireMigrateEvidenceRows = GetRetireMigrateEvidenceRows & vbLf
        GetRetireMigrateEvidenceRows = GetRetireMigrateEvidenceRows & _
            mCaseNames(i) & vbTab & mCaseResults(i) & vbTab & mCaseDetails(i)
    Next i
End Function

Private Function RunArchiveOnlyCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim archiveFolder As String
    Dim wbCfg As Workbook
    Dim loWh As ListObject

    warehouseId = "WHRME2E1"
    runtimeBase = BuildRetireMigrateTempRoot("archive_only")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireMigrateRuntime(warehouseId, runtimeRoot, templateRoot, "admin.e2e", "654321") Then
        detailText = "Setup failed."
        GoTo CleanExit
    End If
    If Not SeedRetireMigrateInventory(warehouseId, runtimeRoot, "EVT-ARCHIVE-ONLY", "admin.e2e", 4, "A1", "archive-only-seed", "SKU-RM-001") Then
        detailText = "Inventory seed failed. " & mLastSeedFailure
        GoTo CleanExit
    End If

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_ONLY
    spec.AdminUser = "admin.e2e"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot

    If Not modWarehouseRetire.WriteArchivePackage(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If

    archiveFolder = FindArchiveFolderRetireMigrate(archiveRoot, warehouseId)
    If archiveFolder = "" Then
        detailText = "Archive folder not found."
        GoTo CleanExit
    End If
    If Not AssertArchiveArtifactsRetireMigrate(archiveFolder, warehouseId, detailText) Then GoTo CleanExit
    If Not AssertManifestJsonRetireMigrate(archiveFolder & "\manifest.json", warehouseId, detailText) Then GoTo CleanExit
    If Len(Dir$(runtimeRoot & "\" & warehouseId & ".invSys.Data.Inventory.xlsb", vbNormal)) = 0 Then
        detailText = "Source inventory workbook no longer exists."
        GoTo CleanExit
    End If

    Set wbCfg = OpenWorkbookIfNeededRetireMigrate(runtimeRoot & "\" & warehouseId & ".invSys.Config.xlsb")
    If wbCfg Is Nothing Then
        detailText = "Source config workbook could not be reopened."
        GoTo CleanExit
    End If
    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    If loWh Is Nothing Then
        detailText = "Warehouse config table missing."
        GoTo CleanExit
    End If
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loWh, 1, "WarehouseStatus")), "ACTIVE", vbTextCompare) <> 0 Then
        detailText = "WarehouseStatus changed during archive-only."
        GoTo CleanExit
    End If

    RunArchiveOnlyCase = True
    detailText = "Archive package was complete and manifest-valid; source runtime remained present with WarehouseStatus=ACTIVE."

CleanExit:
    CloseWorkbookIfOpenRetireMigrate wbCfg
    CleanupRetireMigrateScenario runtimeBase, Array(warehouseId)
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunArchiveMigrateCase(ByRef detailText As String) As Boolean
    Dim sourceWh As String
    Dim targetWh As String
    Dim runtimeBase As String
    Dim sourceRoot As String
    Dim targetRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim wbInv As Workbook
    Dim wbAuth As Workbook
    Dim wbCfg As Workbook
    Dim loSku As ListObject
    Dim loLog As ListObject
    Dim loUsers As ListObject
    Dim loWh As ListObject
    Dim loSt As ListObject
    Dim rowIndex As Long

    sourceWh = "WHRME2E2A"
    targetWh = "WHRME2E2B"
    runtimeBase = BuildRetireMigrateTempRoot("archive_migrate")
    sourceRoot = runtimeBase & "\src"
    targetRoot = runtimeBase & "\tgt"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireMigrateRuntime(sourceWh, sourceRoot, templateRoot, "admin.e2e", "654321") Then
        detailText = "Source setup failed."
        GoTo CleanExit
    End If
    If Not SetupRetireMigrateRuntime(targetWh, targetRoot, templateRoot, "admin.e2e", "654321") Then
        detailText = "Target setup failed."
        GoTo CleanExit
    End If
    If Not RenameRetireMigrateTargetConfig(targetWh, targetRoot, "Retire Integration Target", "ADM1") Then
        detailText = "Target config rename failed."
        GoTo CleanExit
    End If
    If Not SeedRetireMigrateInventory(sourceWh, sourceRoot, "EVT-MIG-SRC", "admin.e2e", 5, "A1", "source-seed", "SKU-RM-002") Then
        detailText = "Source seed failed. " & mLastSeedFailure
        GoTo CleanExit
    End If
    If Not SeedRetireMigrateInventory(targetWh, targetRoot, "EVT-MIG-TGT", "admin.e2e", 2, "B1", "target-seed", "SKU-RM-002") Then
        detailText = "Target seed failed. " & mLastSeedFailure
        GoTo CleanExit
    End If
    If Not AddSourceOnlyAuthUserRetireMigrate(sourceWh, sourceRoot) Then
        detailText = "Source-only auth user seed failed."
        GoTo CleanExit
    End If

    spec.SourceWarehouseId = sourceWh
    spec.TargetWarehouseId = targetWh
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    spec.AdminUser = "admin.e2e"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot

    If Not modWarehouseRetire.WriteArchivePackage(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If

    modRuntimeWorkbooks.SetCoreDataRootOverride targetRoot
    If Not modWarehouseRetire.MigrateInventoryToTarget(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If

    Set wbInv = OpenWorkbookIfNeededRetireMigrate(targetRoot & "\" & targetWh & ".invSys.Data.Inventory.xlsb")
    Set wbAuth = OpenWorkbookIfNeededRetireMigrate(targetRoot & "\" & targetWh & ".invSys.Auth.xlsb")
    Set wbCfg = OpenWorkbookIfNeededRetireMigrate(targetRoot & "\" & targetWh & ".invSys.Config.xlsb")
    If wbInv Is Nothing Or wbAuth Is Nothing Or wbCfg Is Nothing Then
        detailText = "Target workbooks were not available after migration."
        GoTo CleanExit
    End If

    Set loSku = wbInv.Worksheets("SkuBalance").ListObjects("tblSkuBalance")
    Set loLog = wbInv.Worksheets("InventoryLog").ListObjects("tblInventoryLog")
    Set loUsers = wbAuth.Worksheets("Users").ListObjects("tblUsers")
    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")
    If loSku Is Nothing Or loLog Is Nothing Or loUsers Is Nothing Or loWh Is Nothing Or loSt Is Nothing Then
        detailText = "One or more target tables were missing after migration."
        GoTo CleanExit
    End If

    rowIndex = FindRowByValueRetireMigrate(loSku, "SKU", "SKU-RM-002")
    If rowIndex = 0 Then
        detailText = "Target SKU row not found."
        GoTo CleanExit
    End If
    If CDbl(TestPhase2Helpers.GetRowValue(loSku, rowIndex, "QtyOnHand")) <> 7 Then
        detailText = "Target inventory qty did not append source state."
        GoTo CleanExit
    End If

    rowIndex = FindLastInventoryLogRowRetireMigrate(loLog, "MIGRATION_SEED", sourceWh)
    If rowIndex = 0 Then
        detailText = "Target inventory log did not record MIGRATION_SEED."
        GoTo CleanExit
    End If
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loLog, rowIndex, "MigrationSourceId")), sourceWh, vbTextCompare) <> 0 Then
        detailText = "MigrationSourceId was not recorded."
        GoTo CleanExit
    End If

    If FindRowByValueRetireMigrate(loUsers, "UserId", "source.only") <> 0 Then
        detailText = "Source auth user was copied into target."
        GoTo CleanExit
    End If
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loWh, 1, "WarehouseId")), targetWh, vbTextCompare) <> 0 Then
        detailText = "Target WarehouseId changed during migration."
        GoTo CleanExit
    End If
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loWh, 1, "WarehouseName")), "Retire Integration Target", vbTextCompare) <> 0 Then
        detailText = "Target WarehouseName changed during migration."
        GoTo CleanExit
    End If
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loSt, 1, "StationId")), "ADM1", vbTextCompare) <> 0 Then
        detailText = "Target StationId changed during migration."
        GoTo CleanExit
    End If

    RunArchiveMigrateCase = True
    detailText = "Target inventory appended source state to QtyOnHand=7, MigrationSourceId was logged, auth was not copied, and target config identity stayed intact."

CleanExit:
    CloseWorkbookIfOpenRetireMigrate wbCfg
    CloseWorkbookIfOpenRetireMigrate wbAuth
    CloseWorkbookIfOpenRetireMigrate wbInv
    CleanupRetireMigrateScenario runtimeBase, Array(sourceWh, targetWh)
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunArchiveRetireCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim sharePointRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim wbCfg As Workbook
    Dim loWh As ListObject
    Dim tombstonePath As String
    Dim publishedTombstonePath As String

    warehouseId = "WHRME2E3"
    runtimeBase = BuildRetireMigrateTempRoot("archive_retire")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    sharePointRoot = runtimeBase & "\sharepoint"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireMigrateRuntime(warehouseId, runtimeRoot, templateRoot, "admin.e2e", "654321") Then
        detailText = "Setup failed."
        GoTo CleanExit
    End If
    If Not SeedRetireMigrateInventory(warehouseId, runtimeRoot, "EVT-RETIRE", "admin.e2e", 6, "A1", "retire-seed", "SKU-RM-003") Then
        detailText = "Inventory seed failed. " & mLastSeedFailure
        GoTo CleanExit
    End If
    If Not SetWarehouseSharePointPathRetireMigrate(warehouseId, runtimeRoot, sharePointRoot) Then
        detailText = "SharePoint path seed failed."
        GoTo CleanExit
    End If

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
    spec.AdminUser = "admin.e2e"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    spec.PublishTombstone = True

    If Not modWarehouseRetire.WriteArchivePackage(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If
    If Not modWarehouseRetire.RetireSourceWarehouse(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If

    tombstonePath = archiveRoot & "\" & warehouseId & ".tombstone.json"
    publishedTombstonePath = sharePointRoot & "\tombstones\" & warehouseId & ".tombstone.json"
    If Len(Dir$(tombstonePath, vbNormal)) = 0 Then
        detailText = "Archive tombstone was not written."
        GoTo CleanExit
    End If
    If Len(Dir$(publishedTombstonePath, vbNormal)) = 0 Then
        detailText = "SharePoint tombstone was not published."
        GoTo CleanExit
    End If

    Set wbCfg = OpenWorkbookIfNeededRetireMigrate(runtimeRoot & "\" & warehouseId & ".invSys.Config.xlsb")
    If wbCfg Is Nothing Then
        detailText = "Config workbook could not be reopened."
        GoTo CleanExit
    End If
    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    If loWh Is Nothing Then
        detailText = "Warehouse config table missing."
        GoTo CleanExit
    End If
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loWh, 1, "WarehouseStatus")), "RETIRED", vbTextCompare) <> 0 Then
        detailText = "WarehouseStatus was not set to RETIRED."
        GoTo CleanExit
    End If
    If Not IsDate(TestPhase2Helpers.GetRowValue(loWh, 1, "RetiredAtUTC")) Then
        detailText = "RetiredAtUTC was not stamped."
        GoTo CleanExit
    End If

    RunArchiveRetireCase = True
    detailText = "Retirement stamped WarehouseStatus=RETIRED, wrote the local tombstone, and published the tombstone to the SharePoint folder."

CleanExit:
    CloseWorkbookIfOpenRetireMigrate wbCfg
    CleanupRetireMigrateScenario runtimeBase, Array(warehouseId)
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunArchiveRetireDeleteCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim tombstonePath As String

    warehouseId = "WHRME2E4"
    runtimeBase = BuildRetireMigrateTempRoot("archive_retire_delete")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireMigrateRuntime(warehouseId, runtimeRoot, templateRoot, "admin.e2e", "654321") Then
        detailText = "Setup failed."
        GoTo CleanExit
    End If
    If Not SeedRetireMigrateInventory(warehouseId, runtimeRoot, "EVT-RETIRE-DELETE", "admin.e2e", 3, "A1", "retire-delete-seed", "SKU-RM-004") Then
        detailText = "Inventory seed failed. " & mLastSeedFailure
        GoTo CleanExit
    End If

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
    spec.AdminUser = "admin.e2e"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot

    If Not modWarehouseRetire.WriteArchivePackage(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If
    If Not modWarehouseRetire.RetireSourceWarehouse(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If

    tombstonePath = archiveRoot & "\" & warehouseId & ".tombstone.json"
    If Len(Dir$(tombstonePath, vbNormal)) = 0 Then
        detailText = "Tombstone did not exist before delete."
        GoTo CleanExit
    End If
    If Not modWarehouseRetire.DeleteLocalRuntime(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If
    If Len(Dir$(runtimeRoot, vbDirectory)) > 0 Then
        detailText = "Runtime root still exists after delete."
        GoTo CleanExit
    End If

    RunArchiveRetireDeleteCase = True
    detailText = "Delete mode only ran after the tombstone existed, and the local runtime folder tree was removed."

CleanExit:
    CleanupRetireMigrateScenario runtimeBase, Array()
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunRetiredReuseRejectedCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim sourceRoot As String
    Dim archiveRoot As String
    Dim sharePointRoot As String
    Dim duplicateRoot As String
    Dim templateRoot As String
    Dim sourceSpec As modWarehouseBootstrap.WarehouseSpec
    Dim duplicateSpec As modWarehouseBootstrap.WarehouseSpec
    Dim retireSpec As modWarehouseRetire.RetireMigrateSpec
    Dim wbCfg As Workbook
    Dim loWh As ListObject
    Dim duplicateReport As String

    warehouseId = "WHRME2E5"
    runtimeBase = BuildRetireMigrateTempRoot("reuse_rejected")
    sourceRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    sharePointRoot = runtimeBase & "\sharepoint"
    duplicateRoot = runtimeBase & "\duplicate"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    sourceSpec.WarehouseId = warehouseId
    sourceSpec.WarehouseName = "Retire Reuse Source"
    sourceSpec.StationId = "ADM1"
    sourceSpec.AdminUser = "admin.e2e"
    sourceSpec.PathLocal = sourceRoot
    sourceSpec.PathSharePoint = sharePointRoot
    modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride templateRoot
    modRuntimeWorkbooks.SetCoreDataRootOverride sourceRoot
    If Not modWarehouseBootstrap.BootstrapWarehouseLocal(sourceSpec) Then
        detailText = modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
        GoTo CleanExit
    End If
    If Not SetUserPinHashRetireMigrate(warehouseId, sourceRoot, "admin.e2e", "654321") Then
        detailText = "Admin auth seed failed."
        GoTo CleanExit
    End If
    If Not SeedRetireMigrateInventory(warehouseId, sourceRoot, "EVT-REUSE", "admin.e2e", 2, "A1", "reuse-seed", "SKU-RM-005") Then
        detailText = "Inventory seed failed. " & mLastSeedFailure
        GoTo CleanExit
    End If
    If Not modWarehouseBootstrap.PublishInitialArtifacts(sourceSpec) Then
        detailText = modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
        GoTo CleanExit
    End If

    retireSpec.SourceWarehouseId = warehouseId
    retireSpec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
    retireSpec.AdminUser = "admin.e2e"
    retireSpec.ConfirmedByUser = True
    retireSpec.ArchiveDestPath = archiveRoot
    retireSpec.PublishTombstone = True
    If Not modWarehouseRetire.WriteArchivePackage(retireSpec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If
    If Not modWarehouseRetire.RetireSourceWarehouse(retireSpec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If

    Set wbCfg = OpenWorkbookIfNeededRetireMigrate(sourceRoot & "\" & warehouseId & ".invSys.Config.xlsb")
    If wbCfg Is Nothing Then
        detailText = "Retired config workbook could not be reopened."
        GoTo CleanExit
    End If
    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    If loWh Is Nothing Then
        detailText = "Retired config table missing."
        GoTo CleanExit
    End If
    If StrComp(CStr(TestPhase2Helpers.GetRowValue(loWh, 1, "WarehouseStatus")), "RETIRED", vbTextCompare) <> 0 Then
        detailText = "Source warehouse was not retired before reuse probe."
        GoTo CleanExit
    End If

    duplicateSpec = sourceSpec
    duplicateSpec.PathLocal = duplicateRoot
    modRuntimeWorkbooks.SetCoreDataRootOverride duplicateRoot
    If Not modWarehouseBootstrap.WarehouseIdExists(warehouseId) Then
        detailText = "WarehouseIdExists did not detect the retired warehouse's published artifacts."
        GoTo CleanExit
    End If
    If modWarehouseBootstrap.BootstrapWarehouseLocal(duplicateSpec) Then
        detailText = "Bootstrap allowed duplicate warehouse reuse."
        GoTo CleanExit
    End If
    duplicateReport = modWarehouseBootstrap.GetLastWarehouseBootstrapReport()
    If InStr(1, duplicateReport, "already exists", vbTextCompare) = 0 Then
        detailText = duplicateReport
        GoTo CleanExit
    End If

    RunRetiredReuseRejectedCase = True
    detailText = "After retirement, the same WarehouseId was still network-visible via published artifacts and duplicate bootstrap was rejected."

CleanExit:
    CloseWorkbookIfOpenRetireMigrate wbCfg
    CleanupRetireMigrateScenario runtimeBase, Array(warehouseId)
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunDeleteWithoutArchiveManifestCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim tombstonePath As String

    warehouseId = "WHRME2E6"
    runtimeBase = BuildRetireMigrateTempRoot("delete_no_manifest")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireMigrateRuntime(warehouseId, runtimeRoot, templateRoot, "admin.e2e", "654321") Then
        detailText = "Setup failed."
        GoTo CleanExit
    End If

    tombstonePath = archiveRoot & "\" & warehouseId & ".tombstone.json"
    WriteTextFileRetireMigrate tombstonePath, "{""WarehouseId"":""" & warehouseId & """}"

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
    spec.AdminUser = "admin.e2e"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot

    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    If modWarehouseRetire.DeleteLocalRuntime(spec) Then
        detailText = "DeleteLocalRuntime succeeded without an archive manifest."
        GoTo CleanExit
    End If
    If InStr(1, modWarehouseRetire.GetLastWarehouseRetireReport(), "Archive manifest not found", vbTextCompare) = 0 Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If
    If Len(Dir$(runtimeRoot, vbDirectory)) = 0 Then
        detailText = "Runtime root was touched even though archive manifest was missing."
        GoTo CleanExit
    End If

    RunDeleteWithoutArchiveManifestCase = True
    detailText = "DeleteLocalRuntime rejected a hand-dropped tombstone when no archive manifest existed, and the runtime folder remained untouched."

CleanExit:
    CleanupRetireMigrateScenario runtimeBase, Array(warehouseId)
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunDeleteWithoutConfirmationCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec

    warehouseId = "WHRME2E7"
    runtimeBase = BuildRetireMigrateTempRoot("delete_no_confirmation")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireMigrateRuntime(warehouseId, runtimeRoot, templateRoot, "admin.e2e", "654321") Then
        detailText = "Setup failed."
        GoTo CleanExit
    End If
    If Not SeedRetireMigrateInventory(warehouseId, runtimeRoot, "EVT-DELETE-NO-CONFIRM", "admin.e2e", 2, "A1", "delete-no-confirm", "SKU-RM-007") Then
        detailText = "Inventory seed failed. " & mLastSeedFailure
        GoTo CleanExit
    End If

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
    spec.AdminUser = "admin.e2e"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    If Not modWarehouseRetire.WriteArchivePackage(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If
    If Not modWarehouseRetire.RetireSourceWarehouse(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If

    spec.ConfirmedByUser = False
    modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot
    If modWarehouseRetire.DeleteLocalRuntime(spec) Then
        detailText = "DeleteLocalRuntime succeeded without ConfirmedByUser."
        GoTo CleanExit
    End If
    If InStr(1, modWarehouseRetire.GetLastWarehouseRetireReport(), "ConfirmedByUser = True", vbTextCompare) = 0 Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If
    If Len(Dir$(runtimeRoot, vbDirectory)) = 0 Then
        detailText = "Runtime root was deleted even though confirmation was missing."
        GoTo CleanExit
    End If

    RunDeleteWithoutConfirmationCase = True
    detailText = "DeleteLocalRuntime rejected the unconfirmed destructive request and left the runtime folder intact."

CleanExit:
    CleanupRetireMigrateScenario runtimeBase, Array(warehouseId)
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Function RunSharePointUnavailableRetireCase(ByRef detailText As String) As Boolean
    Dim warehouseId As String
    Dim runtimeBase As String
    Dim runtimeRoot As String
    Dim archiveRoot As String
    Dim templateRoot As String
    Dim spec As modWarehouseRetire.RetireMigrateSpec
    Dim report As String
    Dim tombstonePath As String

    warehouseId = "WHRME2E8"
    runtimeBase = BuildRetireMigrateTempRoot("retire_sharepoint_unavailable")
    runtimeRoot = runtimeBase & "\runtime"
    archiveRoot = runtimeBase & "\archive"
    templateRoot = runtimeBase & "\templates"

    On Error GoTo CleanFail
    If Not SetupRetireMigrateRuntime(warehouseId, runtimeRoot, templateRoot, "admin.e2e", "654321") Then
        detailText = "Setup failed."
        GoTo CleanExit
    End If
    If Not SeedRetireMigrateInventory(warehouseId, runtimeRoot, "EVT-SP-UNAVAILABLE", "admin.e2e", 8, "A1", "sp-unavailable", "SKU-RM-008") Then
        detailText = "Inventory seed failed. " & mLastSeedFailure
        GoTo CleanExit
    End If
    If Not SetWarehouseSharePointPathRetireMigrate(warehouseId, runtimeRoot, "Z:\invSys-unavailable") Then
        detailText = "SharePoint path seed failed."
        GoTo CleanExit
    End If

    modDiagnostics.ResetDiagnosticCapture

    spec.SourceWarehouseId = warehouseId
    spec.OperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
    spec.AdminUser = "admin.e2e"
    spec.ConfirmedByUser = True
    spec.ArchiveDestPath = archiveRoot
    spec.PublishTombstone = True
    If Not modWarehouseRetire.WriteArchivePackage(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If
    If Not modWarehouseRetire.RetireSourceWarehouse(spec) Then
        detailText = modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanExit
    End If

    report = modWarehouseRetire.GetLastWarehouseRetireReport()
    tombstonePath = archiveRoot & "\" & warehouseId & ".tombstone.json"
    If Len(Dir$(tombstonePath, vbNormal)) = 0 Then
        detailText = "Local tombstone was not written."
        GoTo CleanExit
    End If
    If InStr(1, report, "PublishWarning=", vbTextCompare) = 0 Then
        detailText = "Retirement report did not surface the publish warning."
        GoTo CleanExit
    End If
    If modDiagnostics.GetDiagnosticEventCount() < 2 Then
        detailText = "Expected both a publish warning diagnostic and a final retirement diagnostic."
        GoTo CleanExit
    End If
    If StrComp(modDiagnostics.GetLastDiagnosticCategory(), "WAREHOUSE-RETIRE", vbTextCompare) <> 0 Then
        detailText = "Retirement diagnostics category was not recorded."
        GoTo CleanExit
    End If

    RunSharePointUnavailableRetireCase = True
    detailText = "Retirement completed with a local tombstone, SharePoint failure stayed advisory via PublishWarning, and diagnostics captured the warning."

CleanExit:
    CleanupRetireMigrateScenario runtimeBase, Array(warehouseId)
    Exit Function
CleanFail:
    detailText = Err.Description
    Resume CleanExit
End Function

Private Sub ResetRetireMigrateEvidence()
    mCaseCount = 0
    Erase mCaseNames
    Erase mCaseResults
    Erase mCaseDetails
    mSummary = vbNullString
End Sub

Private Sub RecordRetireMigrateCase(ByVal caseName As String, ByVal passed As Boolean, ByVal detailText As String)
    mCaseCount = mCaseCount + 1
    ReDim Preserve mCaseNames(1 To mCaseCount)
    ReDim Preserve mCaseResults(1 To mCaseCount)
    ReDim Preserve mCaseDetails(1 To mCaseCount)

    mCaseNames(mCaseCount) = Trim$(caseName)
    mCaseResults(mCaseCount) = IIf(passed, "PASS", "FAIL")
    mCaseDetails(mCaseCount) = SafeRetireMigrateText(detailText)
End Sub

Private Function AllRetireMigrateCasesPassed() As Boolean
    Dim i As Long

    AllRetireMigrateCasesPassed = (mCaseCount > 0)
    For i = 1 To mCaseCount
        If StrComp(mCaseResults(i), "PASS", vbTextCompare) <> 0 Then
            AllRetireMigrateCasesPassed = False
            Exit Function
        End If
    Next i
End Function

Private Function SetupRetireMigrateRuntime(ByVal warehouseId As String, _
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
    SetupRetireMigrateRuntime = True

CleanExit:
    CloseWorkbookIfOpenRetireMigrate wbAuth
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    Exit Function

FailSetup:
    Resume CleanExit
End Function

Private Function SetUserPinHashRetireMigrate(ByVal warehouseId As String, _
                                             ByVal runtimeRoot As String, _
                                             ByVal userId As String, _
                                             ByVal passwordText As String) As Boolean
    Dim wbAuth As Workbook

    On Error GoTo CleanFail
    Set wbAuth = OpenWorkbookIfNeededRetireMigrate(runtimeRoot & "\" & warehouseId & ".invSys.Auth.xlsb")
    If wbAuth Is Nothing Then GoTo CleanExit
    TestPhase2Helpers.SetUserPinHash wbAuth, userId, modAuth.HashUserCredential(passwordText)
    wbAuth.Save
    SetUserPinHashRetireMigrate = True

CleanExit:
    CloseWorkbookIfOpenRetireMigrate wbAuth
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function SeedRetireMigrateInventory(ByVal warehouseId As String, _
                                            ByVal runtimeRoot As String, _
                                            ByVal eventId As String, _
                                            ByVal userId As String, _
                                            ByVal qty As Double, _
                                            ByVal locationVal As String, _
                                            ByVal noteVal As String, _
                                            ByVal skuValue As String) As Boolean
    Dim wbInv As Workbook
    Dim evt As Object
    Dim statusOut As String
    Dim errorCode As String
    Dim errorMessage As String
    Dim report As String

    On Error GoTo FailSeed
    mLastSeedFailure = vbNullString

    Set wbInv = OpenWorkbookIfNeededRetireMigrate(runtimeRoot & "\" & warehouseId & ".invSys.Data.Inventory.xlsb")
    If wbInv Is Nothing Then
        mLastSeedFailure = "Inventory workbook could not be opened: " & runtimeRoot & "\" & warehouseId & ".invSys.Data.Inventory.xlsb"
        GoTo CleanExit
    End If
    If Not EnsureRetireMigrateSkuCatalog(wbInv, skuValue) Then
        mLastSeedFailure = "SKU catalog seed failed for " & skuValue & "."
        GoTo CleanExit
    End If

    Set evt = TestPhase2Helpers.CreateReceiveEvent(eventId, warehouseId, "ADM1", userId, skuValue, qty, locationVal, noteVal, Now, "seed-inbox")
    If Not modInventoryApply.ApplyEvent(evt, wbInv, "RUN-" & eventId, statusOut, errorCode, errorMessage) Then
        mLastSeedFailure = "ApplyEvent failed: " & errorCode & " " & errorMessage
        GoTo CleanExit
    End If
    If Not GenerateWarehouseSnapshot(warehouseId, wbInv, runtimeRoot & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb", Nothing, report) Then
        mLastSeedFailure = "GenerateWarehouseSnapshot failed: " & report
        GoTo CleanExit
    End If
    SeedRetireMigrateInventory = True

CleanExit:
    CloseWorkbookIfOpenRetireMigrate wbInv
    Exit Function

FailSeed:
    mLastSeedFailure = "Seed exception: " & Err.Description
    Resume CleanExit
End Function

Private Function EnsureRetireMigrateSkuCatalog(ByVal wbInv As Workbook, ByVal skuValue As String) As Boolean
    Dim loSku As ListObject
    Dim rowIndex As Long
    Dim r As ListRow

    On Error GoTo CleanFail

    skuValue = Trim$(skuValue)
    If wbInv Is Nothing Or skuValue = "" Then Exit Function

    Set loSku = wbInv.Worksheets("SkuCatalog").ListObjects("tblSkuCatalog")
    If loSku Is Nothing Then Exit Function

    rowIndex = FindRowByValueRetireMigrate(loSku, "SKU", skuValue)
    If rowIndex = 0 Then
        loSku.Parent.Unprotect
        Set r = loSku.ListRows.Add
        rowIndex = r.Index
    End If

    SetTableCellIfColumnRetireMigrate loSku, rowIndex, "SKU", skuValue
    SetTableCellIfColumnRetireMigrate loSku, rowIndex, "ITEM_CODE", skuValue
    SetTableCellIfColumnRetireMigrate loSku, rowIndex, "ITEM", skuValue
    SetTableCellIfColumnRetireMigrate loSku, rowIndex, "UOM", "EA"
    loSku.Parent.Protect UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
    EnsureRetireMigrateSkuCatalog = True
    Exit Function

CleanFail:
    EnsureRetireMigrateSkuCatalog = False
End Function

Private Sub SetTableCellIfColumnRetireMigrate(ByVal lo As ListObject, _
                                              ByVal rowIndex As Long, _
                                              ByVal columnName As String, _
                                              ByVal valueOut As Variant)
    Dim columnIndex As Long

    On Error Resume Next
    columnIndex = lo.ListColumns(columnName).Index
    On Error GoTo 0
    If columnIndex <= 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    lo.DataBodyRange.Cells(rowIndex, columnIndex).Value = valueOut
End Sub

Private Function AddSourceOnlyAuthUserRetireMigrate(ByVal warehouseId As String, ByVal runtimeRoot As String) As Boolean
    Dim wbAuth As Workbook

    On Error GoTo CleanFail
    Set wbAuth = OpenWorkbookIfNeededRetireMigrate(runtimeRoot & "\" & warehouseId & ".invSys.Auth.xlsb")
    If wbAuth Is Nothing Then GoTo CleanExit
    TestPhase2Helpers.AddCapability wbAuth, "source.only", "ADMIN_MAINT", warehouseId, "ADM1", "ACTIVE"
    TestPhase2Helpers.SetUserPinHash wbAuth, "source.only", modAuth.HashUserCredential("999999")
    wbAuth.Save
    AddSourceOnlyAuthUserRetireMigrate = True

CleanExit:
    CloseWorkbookIfOpenRetireMigrate wbAuth
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function RenameRetireMigrateTargetConfig(ByVal warehouseId As String, _
                                                 ByVal runtimeRoot As String, _
                                                 ByVal warehouseName As String, _
                                                 ByVal stationId As String) As Boolean
    Dim wbCfg As Workbook

    On Error GoTo CleanFail
    Set wbCfg = OpenWorkbookIfNeededRetireMigrate(runtimeRoot & "\" & warehouseId & ".invSys.Config.xlsb")
    If wbCfg Is Nothing Then GoTo CleanExit
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "WarehouseName", warehouseName
    TestPhase2Helpers.SetStationConfigValue wbCfg, "StationId", stationId
    wbCfg.Save
    RenameRetireMigrateTargetConfig = True

CleanExit:
    CloseWorkbookIfOpenRetireMigrate wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function SetWarehouseSharePointPathRetireMigrate(ByVal warehouseId As String, _
                                                         ByVal runtimeRoot As String, _
                                                         ByVal sharePointPath As String) As Boolean
    Dim wbCfg As Workbook

    On Error GoTo CleanFail
    Set wbCfg = OpenWorkbookIfNeededRetireMigrate(runtimeRoot & "\" & warehouseId & ".invSys.Config.xlsb")
    If wbCfg Is Nothing Then GoTo CleanExit
    TestPhase2Helpers.SetWarehouseConfigValue wbCfg, "PathSharePointRoot", sharePointPath
    wbCfg.Save
    SetWarehouseSharePointPathRetireMigrate = True

CleanExit:
    CloseWorkbookIfOpenRetireMigrate wbCfg
    Exit Function
CleanFail:
    Resume CleanExit
End Function

Private Function AssertArchiveArtifactsRetireMigrate(ByVal archiveFolder As String, _
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
        archiveFolder & "\manifest.json")

    For Each item In requiredPaths
        If Not PathExistsRetireMigrate(CStr(item)) Then
            detailText = "Missing archive artifact: " & CStr(item)
            Exit Function
        End If
    Next item

    AssertArchiveArtifactsRetireMigrate = True
End Function

Private Function AssertManifestJsonRetireMigrate(ByVal manifestPath As String, _
                                                 ByVal warehouseId As String, _
                                                 ByRef detailText As String) As Boolean
    Dim manifestText As String
    Dim normalizedText As String

    manifestText = ReadAllTextRetireMigrate(manifestPath)
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
    normalizedText = TrimJsonWhitespaceRetireMigrate(manifestText)
    If Left$(normalizedText, 1) <> "{" Or Right$(normalizedText, 1) <> "}" Then
        detailText = "Manifest was not JSON-shaped."
        Exit Function
    End If

    AssertManifestJsonRetireMigrate = True
End Function

Private Function OpenWorkbookIfNeededRetireMigrate(ByVal fullPath As String) As Workbook
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, fullPath, vbTextCompare) = 0 Then
            Set OpenWorkbookIfNeededRetireMigrate = wb
            Exit Function
        End If
    Next wb

    If Len(Dir$(fullPath, vbNormal)) = 0 Then Exit Function
    Set OpenWorkbookIfNeededRetireMigrate = Application.Workbooks.Open(fullPath)
End Function

Private Function FindRowByValueRetireMigrate(ByVal lo As ListObject, ByVal columnName As String, ByVal matchValue As String) As Long
    Dim rowIndex As Long
    Dim colIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    colIndex = lo.ListColumns(columnName).Index
    If colIndex <= 0 Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(CStr(lo.DataBodyRange.Cells(rowIndex, colIndex).Value), matchValue, vbTextCompare) = 0 Then
            FindRowByValueRetireMigrate = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function FindLastInventoryLogRowRetireMigrate(ByVal lo As ListObject, _
                                                      ByVal eventType As String, _
                                                      ByVal migrationSourceId As String) As Long
    Dim rowIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    For rowIndex = lo.ListRows.Count To 1 Step -1
        If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, rowIndex, "EventType")), eventType, vbTextCompare) = 0 Then
            If StrComp(CStr(TestPhase2Helpers.GetRowValue(lo, rowIndex, "MigrationSourceId")), migrationSourceId, vbTextCompare) = 0 Then
                FindLastInventoryLogRowRetireMigrate = rowIndex
                Exit Function
            End If
        End If
    Next rowIndex
End Function

Private Function FindArchiveFolderRetireMigrate(ByVal archiveRoot As String, ByVal warehouseId As String) As String
    Dim candidate As String

    candidate = Dir$(archiveRoot & "\" & warehouseId & "_archive_*", vbDirectory)
    Do While candidate <> ""
        If candidate <> "." And candidate <> ".." Then
            If InStr(1, candidate, "_tmp", vbTextCompare) = 0 Then
                FindArchiveFolderRetireMigrate = archiveRoot & "\" & candidate
                Exit Function
            End If
        End If
        candidate = Dir$
    Loop
End Function

Private Function BuildRetireMigrateTempRoot(ByVal leafName As String) As String
    BuildRetireMigrateTempRoot = Environ$("TEMP") & "\retire_migrate_" & leafName & "_" & _
                                 Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(Int(Timer * 1000))
End Function

Private Function PathExistsRetireMigrate(ByVal pathIn As String) As Boolean
    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    If pathIn = "" Then Exit Function
    PathExistsRetireMigrate = (Len(Dir$(pathIn, vbDirectory)) > 0)
    If Not PathExistsRetireMigrate Then PathExistsRetireMigrate = (Len(Dir$(pathIn, vbNormal)) > 0)
End Function

Private Function ReadAllTextRetireMigrate(ByVal filePath As String) As String
    Dim fileNum As Integer

    On Error GoTo CleanFail
    If Len(Dir$(filePath, vbNormal)) = 0 Then Exit Function
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    ReadAllTextRetireMigrate = Input$(LOF(fileNum), fileNum)

CleanExit:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    Exit Function

CleanFail:
    Resume CleanExit
End Function

Private Function TrimJsonWhitespaceRetireMigrate(ByVal textIn As String) As String
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

    TrimJsonWhitespaceRetireMigrate = textIn
End Function

Private Sub WriteTextFileRetireMigrate(ByVal filePath As String, ByVal textIn As String)
    Dim fileNum As Integer
    Dim parentPath As String

    parentPath = GetParentFolderRetireMigrate(filePath)
    EnsureFolderExistsRetireMigrate parentPath

    fileNum = FreeFile
    Open filePath For Output As #fileNum
    Print #fileNum, textIn
    Close #fileNum
End Sub

Private Function GetParentFolderRetireMigrate(ByVal pathIn As String) As String
    Dim sepPos As Long

    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    sepPos = InStrRev(pathIn, "\")
    If sepPos > 1 Then GetParentFolderRetireMigrate = Left$(pathIn, sepPos - 1)
End Function

Private Sub EnsureFolderExistsRetireMigrate(ByVal folderPath As String)
    Dim parentPath As String

    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    parentPath = GetParentFolderRetireMigrate(folderPath)
    If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then
        EnsureFolderExistsRetireMigrate parentPath
    End If
    MkDir folderPath
End Sub

Private Sub CleanupRetireMigrateScenario(ByVal runtimeBase As String, ByVal warehouseIds As Variant)
    Dim item As Variant

    On Error Resume Next
    If IsArray(warehouseIds) Then
        For Each item In warehouseIds
            If Trim$(CStr(item)) <> "" Then
                CloseWorkbookByNameRetireMigrate CStr(item) & ".invSys.Config.xlsb"
                CloseWorkbookByNameRetireMigrate CStr(item) & ".invSys.Auth.xlsb"
                CloseWorkbookByNameRetireMigrate CStr(item) & ".invSys.Data.Inventory.xlsb"
                CloseWorkbookByNameRetireMigrate CStr(item) & ".invSys.Snapshot.Inventory.xlsb"
                CloseWorkbookByNameRetireMigrate CStr(item) & ".Outbox.Events.xlsb"
            End If
        Next item
    End If
    modRuntimeWorkbooks.ClearCoreDataRootOverride
    modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride
    DeleteFolderRecursiveRetireMigrate runtimeBase
    On Error GoTo 0
End Sub

Private Sub CloseWorkbookByNameRetireMigrate(ByVal workbookName As String)
    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If StrComp(wb.Name, workbookName, vbTextCompare) = 0 Then
            wb.Close SaveChanges:=False
            Exit Sub
        End If
    Next wb
End Sub

Private Sub CloseWorkbookIfOpenRetireMigrate(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub DeleteFolderRecursiveRetireMigrate(ByVal folderPath As String)
    Dim fso As Object

    On Error Resume Next
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then fso.DeleteFolder folderPath, True
    On Error GoTo 0
End Sub

Private Function SafeRetireMigrateText(ByVal textIn As String) As String
    SafeRetireMigrateText = Replace$(Replace$(Trim$(textIn), vbCr, " "), vbLf, " ")
End Function
