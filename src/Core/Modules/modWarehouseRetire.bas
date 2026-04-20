Attribute VB_Name = "modWarehouseRetire"
Option Explicit

Private Const ARCHIVE_VERSION_RETIRE As String = "1.0"
Private Const ARCHIVE_TEMP_SUFFIX_RETIRE As String = "_tmp"
Private Const TOMBSTONE_FILE_SUFFIX_RETIRE As String = ".tombstone.json"
Private Const TOMBSTONE_PUBLISH_FOLDER_RETIRE As String = "tombstones"

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

Public Function ValidateRetireMigrateSpecValues(ByVal sourceWarehouseId As String, _
                                                ByVal targetWarehouseId As String, _
                                                ByVal operationMode As Long, _
                                                ByVal adminUser As String, _
                                                ByVal confirmedByUser As Boolean, _
                                                ByVal archiveDestPath As String, _
                                                ByVal publishTombstone As Boolean, _
                                                Optional ByRef report As String = "") As Boolean
    Dim spec As RetireMigrateSpec

    spec.SourceWarehouseId = sourceWarehouseId
    spec.TargetWarehouseId = targetWarehouseId
    spec.OperationMode = operationMode
    spec.AdminUser = adminUser
    spec.ConfirmedByUser = confirmedByUser
    spec.ArchiveDestPath = archiveDestPath
    spec.PublishTombstone = publishTombstone
    ValidateRetireMigrateSpecValues = ValidateRetireMigrateSpec(spec, report)
End Function

Public Function WriteArchivePackageValues(ByVal sourceWarehouseId As String, _
                                          ByVal targetWarehouseId As String, _
                                          ByVal operationMode As Long, _
                                          ByVal adminUser As String, _
                                          ByVal confirmedByUser As Boolean, _
                                          ByVal archiveDestPath As String, _
                                          ByVal publishTombstone As Boolean) As Boolean
    Dim spec As RetireMigrateSpec

    spec.SourceWarehouseId = sourceWarehouseId
    spec.TargetWarehouseId = targetWarehouseId
    spec.OperationMode = operationMode
    spec.AdminUser = adminUser
    spec.ConfirmedByUser = confirmedByUser
    spec.ArchiveDestPath = archiveDestPath
    spec.PublishTombstone = publishTombstone
    WriteArchivePackageValues = WriteArchivePackage(spec)
End Function

Public Function MigrateInventoryToTargetValues(ByVal sourceWarehouseId As String, _
                                               ByVal targetWarehouseId As String, _
                                               ByVal operationMode As Long, _
                                               ByVal adminUser As String, _
                                               ByVal confirmedByUser As Boolean, _
                                               ByVal archiveDestPath As String, _
                                               ByVal publishTombstone As Boolean) As Boolean
    Dim spec As RetireMigrateSpec

    spec.SourceWarehouseId = sourceWarehouseId
    spec.TargetWarehouseId = targetWarehouseId
    spec.OperationMode = operationMode
    spec.AdminUser = adminUser
    spec.ConfirmedByUser = confirmedByUser
    spec.ArchiveDestPath = archiveDestPath
    spec.PublishTombstone = publishTombstone
    MigrateInventoryToTargetValues = MigrateInventoryToTarget(spec)
End Function

Public Function RetireSourceWarehouseValues(ByVal sourceWarehouseId As String, _
                                            ByVal targetWarehouseId As String, _
                                            ByVal operationMode As Long, _
                                            ByVal adminUser As String, _
                                            ByVal confirmedByUser As Boolean, _
                                            ByVal archiveDestPath As String, _
                                            ByVal publishTombstone As Boolean) As Boolean
    Dim spec As RetireMigrateSpec

    spec.SourceWarehouseId = sourceWarehouseId
    spec.TargetWarehouseId = targetWarehouseId
    spec.OperationMode = operationMode
    spec.AdminUser = adminUser
    spec.ConfirmedByUser = confirmedByUser
    spec.ArchiveDestPath = archiveDestPath
    spec.PublishTombstone = publishTombstone
    RetireSourceWarehouseValues = RetireSourceWarehouse(spec)
End Function

Public Function DeleteLocalRuntimeValues(ByVal sourceWarehouseId As String, _
                                         ByVal targetWarehouseId As String, _
                                         ByVal operationMode As Long, _
                                         ByVal adminUser As String, _
                                         ByVal confirmedByUser As Boolean, _
                                         ByVal archiveDestPath As String, _
                                         ByVal publishTombstone As Boolean) As Boolean
    Dim spec As RetireMigrateSpec

    spec.SourceWarehouseId = sourceWarehouseId
    spec.TargetWarehouseId = targetWarehouseId
    spec.OperationMode = operationMode
    spec.AdminUser = adminUser
    spec.ConfirmedByUser = confirmedByUser
    spec.ArchiveDestPath = archiveDestPath
    spec.PublishTombstone = publishTombstone
    DeleteLocalRuntimeValues = DeleteLocalRuntime(spec)
End Function

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

    runtimeRoot = ResolveExistingRuntimeRootRetire(spec.SourceWarehouseId)
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

Public Function MigrateInventoryToTarget(ByRef spec As RetireMigrateSpec) As Boolean
    Dim report As String
    Dim archiveFolder As String
    Dim manifestPath As String
    Dim snapshotPath As String
    Dim targetRoot As String
    Dim targetConfigPath As String
    Dim targetStationId As String
    Dim targetCfgWb As Workbook
    Dim queuedCount As Long
    Dim processedCount As Long
    Dim batchReport As String
    Dim priorRootOverride As String

    On Error GoTo FailMigration

    mLastRetireReport = vbNullString
    If Not ValidateRetireMigrateSpec(spec, report) Then GoTo FailSoft
    If Trim$(spec.TargetWarehouseId) = "" Then
        report = "TargetWarehouseId is required for migration."
        GoTo FailSoft
    End If

    archiveFolder = ResolveLatestArchiveFolderRetire(spec.ArchiveDestPath, spec.SourceWarehouseId)
    manifestPath = archiveFolder & "\manifest.json"
    If archiveFolder = "" Or Not FileExistsRetire(manifestPath) Then
        report = "Archive manifest not found. Run WriteArchivePackage successfully before migration."
        GoTo FailSoft
    End If

    snapshotPath = ResolveArchivedSnapshotPathRetire(archiveFolder, spec.SourceWarehouseId)
    If snapshotPath = "" Then
        report = "Archived snapshot not found for source warehouse."
        GoTo FailSoft
    End If

    targetRoot = ResolveExistingRuntimeRootRetire(spec.TargetWarehouseId)
    If targetRoot = "" Then
        report = "Target warehouse runtime not found: " & spec.TargetWarehouseId
        GoTo FailSoft
    End If
    If TargetWarehouseRetiredRetire(spec.TargetWarehouseId, targetRoot) Then
        report = "Target warehouse is retired and cannot accept migration writes: " & spec.TargetWarehouseId
        GoTo FailSoft
    End If

    targetConfigPath = targetRoot & "\" & spec.TargetWarehouseId & ".invSys.Config.xlsb"
    If Not FileExistsRetire(targetConfigPath) Then
        report = "Target config workbook not found: " & targetConfigPath
        GoTo FailSoft
    End If

    Set targetCfgWb = OpenWorkbookReadOnlyRetire(targetConfigPath, report)
    If targetCfgWb Is Nothing Then GoTo FailSoft
    targetStationId = ResolvePrimaryStationIdRetire(targetCfgWb, spec.TargetWarehouseId)
    If targetStationId = "" Then
        report = "Target station could not be resolved from config workbook."
        GoTo FailSoft
    End If
    CloseWorkbookQuietlyRetire targetCfgWb
    Set targetCfgWb = Nothing

    priorRootOverride = modRuntimeWorkbooks.GetCoreDataRootOverride()
    modRuntimeWorkbooks.SetCoreDataRootOverride targetRoot

    If Not modConfig.LoadConfig(spec.TargetWarehouseId, targetStationId) Then
        report = "Target config load failed: " & modConfig.Validate()
        GoTo FailSoft
    End If
    If Not modAuth.LoadAuth(spec.TargetWarehouseId) Then
        report = "Target auth load failed: " & modAuth.ValidateAuth()
        GoTo FailSoft
    End If
    If Not EnsureMigrationInboxReadyRetire(spec.TargetWarehouseId, targetStationId, targetConfigPath, report) Then GoTo FailSoft
    If Not QueueMigrationSeedEventsRetire(spec, archiveFolder, snapshotPath, targetStationId, queuedCount, report) Then GoTo FailSoft

    If queuedCount > 0 Then
        processedCount = modProcessor.RunBatch(spec.TargetWarehouseId, 0, batchReport)
        If Left$(batchReport, 15) = "RunBatch failed" Then
            report = batchReport
            GoTo FailSoft
        End If
    Else
        processedCount = 0
        batchReport = "No positive on-hand rows found in archived snapshot."
    End If

    report = "OK|ArchivePath=" & archiveFolder & "|Queued=" & CStr(queuedCount) & _
             "|Processed=" & CStr(processedCount) & "|Batch=" & batchReport
    MigrateInventoryToTarget = True
    GoTo CleanExit

FailSoft:
    MigrateInventoryToTarget = False
    If Len(report) = 0 Then report = "MigrateInventoryToTarget failed."
    LogDiagnosticSafeRetire "WAREHOUSE-RETIRE", _
        "Migration failed|Source=" & spec.SourceWarehouseId & "|Target=" & spec.TargetWarehouseId & "|Reason=" & report
    GoTo CleanExit

FailMigration:
    report = "MigrateInventoryToTarget failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookQuietlyRetire targetCfgWb
    RestoreCoreRootOverrideRetire priorRootOverride
    mLastRetireReport = report
End Function

Public Function RetireSourceWarehouse(ByRef spec As RetireMigrateSpec) As Boolean
    Dim report As String
    Dim archiveFolder As String
    Dim manifestPath As String
    Dim sourceRoot As String
    Dim configPath As String
    Dim tombstonePath As String
    Dim sharePointRoot As String
    Dim publishedTombstonePath As String
    Dim publishStatus As String
    Dim publishWarning As String
    Dim retiredAtUtc As Date
    Dim warehouseName As String
    Dim wbCfg As Workbook

    On Error GoTo FailRetire

    mLastRetireReport = vbNullString
    If Not ValidateRetireMigrateSpec(spec, report) Then GoTo FailSoft

    archiveFolder = ResolveLatestArchiveFolderRetire(spec.ArchiveDestPath, spec.SourceWarehouseId)
    manifestPath = archiveFolder & "\manifest.json"
    If archiveFolder = "" Or Not FileExistsRetire(manifestPath) Then
        report = "Archive manifest not found. Run WriteArchivePackage successfully before retirement."
        GoTo FailSoft
    End If

    sourceRoot = ResolveExistingRuntimeRootRetire(spec.SourceWarehouseId)
    If sourceRoot = "" Then
        report = "Source warehouse runtime not found: " & spec.SourceWarehouseId
        GoTo FailSoft
    End If

    configPath = sourceRoot & "\" & spec.SourceWarehouseId & ".invSys.Config.xlsb"
    Set wbCfg = OpenWorkbookEditableRetire(configPath, report)
    If wbCfg Is Nothing Then GoTo FailSoft
    If Not modConfig.EnsureConfigSchema(wbCfg, spec.SourceWarehouseId, "", report) Then GoTo FailSoft

    retiredAtUtc = Now
    warehouseName = ResolveWarehouseNameRetire(wbCfg, spec.SourceWarehouseId)
    If Not StampWarehouseRetiredConfigRetire(wbCfg, spec.SourceWarehouseId, retiredAtUtc, report) Then GoTo FailSoft
    SaveWorkbookQuietlyRetire wbCfg
    CloseWorkbookQuietlyRetire wbCfg
    Set wbCfg = Nothing

    tombstonePath = BuildTombstonePathRetire(spec)
    If Not WriteRetirementTombstoneRetire(tombstonePath, spec, warehouseName, retiredAtUtc, archiveFolder, report) Then GoTo FailSoft

    If spec.PublishTombstone Then
        sharePointRoot = ResolveSharePointRootFromConfigRetire(configPath, spec.SourceWarehouseId)
        If sharePointRoot <> "" Then
            publishedTombstonePath = NormalizeFolderPathRetire(sharePointRoot) & TOMBSTONE_PUBLISH_FOLDER_RETIRE & "\" & spec.SourceWarehouseId & TOMBSTONE_FILE_SUFFIX_RETIRE
            If Not modWarehouseSync.PublishFileToTargetPath(tombstonePath, publishedTombstonePath, publishStatus) Then
                publishWarning = publishStatus
                LogDiagnosticSafeRetire "WAREHOUSE-RETIRE", "Tombstone publish warning|WarehouseId=" & spec.SourceWarehouseId & "|Detail=" & publishStatus
            End If
        Else
            publishWarning = "SharePoint root not configured."
            LogDiagnosticSafeRetire "WAREHOUSE-RETIRE", "Tombstone publish skipped|WarehouseId=" & spec.SourceWarehouseId & "|Detail=" & publishWarning
        End If
    End If

    report = "OK|ArchivePath=" & archiveFolder & "|TombstonePath=" & tombstonePath
    If publishWarning <> "" Then
        report = report & "|PublishWarning=" & publishWarning
    ElseIf publishStatus <> "" Then
        report = report & "|Publish=" & publishStatus
    End If
    RetireSourceWarehouse = True
    LogDiagnosticSafeRetire "WAREHOUSE-RETIRE", "Warehouse retired|WarehouseId=" & spec.SourceWarehouseId & "|ArchivePath=" & archiveFolder & "|Tombstone=" & tombstonePath
    GoTo CleanExit

FailSoft:
    RetireSourceWarehouse = False
    If Len(report) = 0 Then report = "RetireSourceWarehouse failed."
    LogDiagnosticSafeRetire "WAREHOUSE-RETIRE", "Retirement failed|WarehouseId=" & spec.SourceWarehouseId & "|Reason=" & report
    GoTo CleanExit

FailRetire:
    report = "RetireSourceWarehouse failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookQuietlyRetire wbCfg
    mLastRetireReport = report
End Function

Public Function DeleteLocalRuntime(ByRef spec As RetireMigrateSpec) As Boolean
    Dim report As String
    Dim targetRoot As String
    Dim tombstonePath As String
    Dim archiveFolder As String
    Dim manifestPath As String

    On Error GoTo FailDelete

    mLastRetireReport = vbNullString
    NormalizeRetireMigrateSpec spec

    If spec.OperationMode <> MODE_ARCHIVE_RETIRE_DELETE Then
        report = "DeleteLocalRuntime is only allowed for MODE_ARCHIVE_RETIRE_DELETE."
        GoTo FailSoft
    End If
    If Not spec.ConfirmedByUser Then
        report = "DeleteLocalRuntime requires ConfirmedByUser = True."
        GoTo FailSoft
    End If

    tombstonePath = BuildTombstonePathRetire(spec)
    If Not FileExistsRetire(tombstonePath) Then
        report = "Retirement tombstone not found. RetireSourceWarehouse must succeed before delete."
        GoTo FailSoft
    End If

    archiveFolder = ResolveLatestArchiveFolderRetire(spec.ArchiveDestPath, spec.SourceWarehouseId)
    manifestPath = archiveFolder & "\manifest.json"
    If archiveFolder = "" Or Not FileExistsRetire(manifestPath) Then
        report = "Archive manifest not found. WriteArchivePackage must succeed before delete."
        GoTo FailSoft
    End If

    targetRoot = ResolveExistingRuntimeRootRetire(spec.SourceWarehouseId)
    If targetRoot = "" Then targetRoot = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(spec.SourceWarehouseId, False)
    If Not FolderExistsRetire(targetRoot) Then
        report = "Local runtime folder not found: " & targetRoot
        GoTo FailSoft
    End If

    DeleteFolderRecursiveRetire targetRoot
    If FolderExistsRetire(targetRoot) Then
        report = "Local runtime folder could not be deleted: " & targetRoot
        GoTo FailSoft
    End If

    report = "OK|DeletedRoot=" & targetRoot & "|TombstonePath=" & tombstonePath
    DeleteLocalRuntime = True
    LogDiagnosticSafeRetire "WAREHOUSE-RETIRE", "Local runtime deleted|WarehouseId=" & spec.SourceWarehouseId & "|DeletedRoot=" & targetRoot & "|AdminUser=" & spec.AdminUser & "|DeletedAtUTC=" & Format$(Now, "yyyy-mm-dd\Thh:nn:ss\Z")
    GoTo CleanExit

FailSoft:
    DeleteLocalRuntime = False
    If Len(report) = 0 Then report = "DeleteLocalRuntime failed."
    LogDiagnosticSafeRetire "WAREHOUSE-RETIRE", "Delete failed|WarehouseId=" & spec.SourceWarehouseId & "|Reason=" & report
    GoTo CleanExit

FailDelete:
    report = "DeleteLocalRuntime failed: " & Err.Description
    Resume FailSoft

CleanExit:
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

Private Function OpenWorkbookEditableRetire(ByVal workbookPath As String, ByRef report As String) As Workbook
    On Error GoTo FailOpen

    Set OpenWorkbookEditableRetire = FindOpenWorkbookByFullNameRetire(workbookPath)
    If OpenWorkbookEditableRetire Is Nothing Then
        Set OpenWorkbookEditableRetire = Application.Workbooks.Open(workbookPath, False, False)
    ElseIf OpenWorkbookEditableRetire.ReadOnly Then
        report = "Workbook is read-only: " & workbookPath
        Set OpenWorkbookEditableRetire = Nothing
    End If
    Exit Function

FailOpen:
    report = "OpenWorkbookEditableRetire failed: " & Err.Description
End Function

Private Sub SaveWorkbookQuietlyRetire(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If Trim$(wb.Path) = "" Then Exit Sub
    wb.Save
End Sub

Private Function StampWarehouseRetiredConfigRetire(ByVal wbCfg As Workbook, _
                                                   ByVal warehouseId As String, _
                                                   ByVal retiredAtUtc As Date, _
                                                   ByRef report As String) As Boolean
    Dim loWh As ListObject
    Dim rowIndex As Long

    On Error GoTo FailStamp

    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    If loWh Is Nothing Or loWh.DataBodyRange Is Nothing Then
        report = "Warehouse config table not available."
        Exit Function
    End If

    rowIndex = FindRowByValueInListObjectRetire(loWh, "WarehouseId", warehouseId)
    If rowIndex = 0 Then rowIndex = 1
    SetListObjectValueRetire loWh, rowIndex, "WarehouseStatus", "RETIRED"
    SetListObjectValueRetire loWh, rowIndex, "RetiredAtUTC", retiredAtUtc
    StampWarehouseRetiredConfigRetire = True
    Exit Function

FailStamp:
    report = "StampWarehouseRetiredConfigRetire failed: " & Err.Description
End Function

Private Function BuildTombstonePathRetire(ByRef spec As RetireMigrateSpec) As String
    BuildTombstonePathRetire = NormalizeFolderPathRetire(spec.ArchiveDestPath) & spec.SourceWarehouseId & TOMBSTONE_FILE_SUFFIX_RETIRE
End Function

Private Function WriteRetirementTombstoneRetire(ByVal tombstonePath As String, _
                                                ByRef spec As RetireMigrateSpec, _
                                                ByVal warehouseName As String, _
                                                ByVal retiredAtUtc As Date, _
                                                ByVal archivePath As String, _
                                                ByRef report As String) As Boolean
    Dim fileNum As Integer

    On Error GoTo FailWrite

    EnsureFolderRecursiveRetire GetParentFolderRetire(tombstonePath)
    fileNum = FreeFile
    Open tombstonePath For Output As #fileNum
    Print #fileNum, "{"
    Print #fileNum, "  ""WarehouseId"": """ & EscapeJsonRetire(spec.SourceWarehouseId) & ""","
    Print #fileNum, "  ""WarehouseName"": """ & EscapeJsonRetire(warehouseName) & ""","
    Print #fileNum, "  ""RetiredAtUTC"": """ & EscapeJsonRetire(Format$(retiredAtUtc, "yyyy-mm-dd\Thh:nn:ss\Z")) & ""","
    Print #fileNum, "  ""RetiredByUser"": """ & EscapeJsonRetire(spec.AdminUser) & ""","
    Print #fileNum, "  ""OperationMode"": """ & EscapeJsonRetire(OperationModeNameRetire(spec.OperationMode)) & ""","
    Print #fileNum, "  ""ArchivePath"": """ & EscapeJsonRetire(archivePath) & ""","
    Print #fileNum, "  ""MigrationTargetId"": """ & EscapeJsonRetire(ResolveMigrationTargetIdRetire(spec)) & """"
    Print #fileNum, "}"
    Close #fileNum

    WriteRetirementTombstoneRetire = True
    Exit Function

FailWrite:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    report = "WriteRetirementTombstoneRetire failed: " & Err.Description
End Function

Private Function ResolveMigrationTargetIdRetire(ByRef spec As RetireMigrateSpec) As String
    If spec.OperationMode = MODE_ARCHIVE_MIGRATE Then ResolveMigrationTargetIdRetire = spec.TargetWarehouseId
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

Private Function ResolveSharePointRootFromConfigRetire(ByVal configPath As String, ByVal warehouseId As String) As String
    Dim wbCfg As Workbook
    Dim report As String
    Dim loWh As ListObject
    Dim rowIndex As Long

    On Error GoTo CleanFail

    Set wbCfg = OpenWorkbookReadOnlyRetire(configPath, report)
    If wbCfg Is Nothing Then GoTo CleanExit

    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    If loWh Is Nothing Or loWh.DataBodyRange Is Nothing Then GoTo CleanExit
    rowIndex = FindRowByValueInListObjectRetire(loWh, "WarehouseId", warehouseId)
    If rowIndex = 0 Then rowIndex = 1
    ResolveSharePointRootFromConfigRetire = Trim$(CStr(loWh.DataBodyRange.Cells(rowIndex, loWh.ListColumns("PathSharePointRoot").Index).Value))

CleanExit:
    CloseWorkbookQuietlyRetire wbCfg
    Exit Function

CleanFail:
    ResolveSharePointRootFromConfigRetire = vbNullString
    Resume CleanExit
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

Private Function ResolveLatestArchiveFolderRetire(ByVal archiveRoot As String, ByVal sourceWarehouseId As String) As String
    Dim candidate As String
    Dim normalizedRoot As String
    Dim bestName As String

    normalizedRoot = NormalizeFolderPathRetire(archiveRoot)
    If normalizedRoot = "" Then Exit Function

    candidate = Dir$(normalizedRoot & sourceWarehouseId & "_archive_*", vbDirectory)
    Do While candidate <> ""
        If candidate <> "." And candidate <> ".." Then
            If InStr(1, candidate, ARCHIVE_TEMP_SUFFIX_RETIRE, vbTextCompare) = 0 Then
                If FileExistsRetire(normalizedRoot & candidate & "\manifest.json") Then
                    If candidate > bestName Then bestName = candidate
                End If
            End If
        End If
        candidate = Dir$
    Loop

    If bestName <> "" Then ResolveLatestArchiveFolderRetire = normalizedRoot & bestName
End Function

Private Function ResolveArchivedSnapshotPathRetire(ByVal archiveFolder As String, ByVal sourceWarehouseId As String) As String
    Dim canonicalPath As String
    Dim candidate As String

    archiveFolder = NormalizeFolderPathRetire(archiveFolder)
    If archiveFolder = "" Then Exit Function

    canonicalPath = archiveFolder & "snapshots\" & sourceWarehouseId & ".invSys.Snapshot.Inventory.xlsb"
    If FileExistsRetire(canonicalPath) Then
        ResolveArchivedSnapshotPathRetire = canonicalPath
        Exit Function
    End If

    candidate = Dir$(archiveFolder & "snapshots\" & sourceWarehouseId & "*.invSys.Snapshot.Inventory.xls*")
    If candidate <> "" Then ResolveArchivedSnapshotPathRetire = archiveFolder & "snapshots\" & candidate
End Function

Private Function ResolveExistingRuntimeRootRetire(ByVal warehouseId As String) As String
    Dim rootPath As String
    Dim wb As Workbook
    Dim candidateRoot As String
    Dim parentPath As String

    On Error GoTo CleanFail

    rootPath = ResolveRuntimeRootRetire(warehouseId)
    If RuntimeArtifactsExistRetire(rootPath, warehouseId) Then
        ResolveExistingRuntimeRootRetire = rootPath
        Exit Function
    End If

    candidateRoot = Trim$(modRuntimeWorkbooks.GetCoreDataRootOverride())
    If RuntimeArtifactsExistRetire(candidateRoot, warehouseId) Then
        ResolveExistingRuntimeRootRetire = candidateRoot
        Exit Function
    End If
    parentPath = GetParentFolderRetire(candidateRoot)
    If parentPath <> "" Then
        candidateRoot = FindRuntimeRootUnderParentRetire(parentPath, warehouseId)
        If candidateRoot <> "" Then
            ResolveExistingRuntimeRootRetire = candidateRoot
            Exit Function
        End If
    End If

    On Error Resume Next
    candidateRoot = Trim$(modConfig.GetString("PathDataRoot", ""))
    On Error GoTo 0
    If RuntimeArtifactsExistRetire(candidateRoot, warehouseId) Then
        ResolveExistingRuntimeRootRetire = candidateRoot
        Exit Function
    End If
    parentPath = GetParentFolderRetire(candidateRoot)
    If parentPath <> "" Then
        candidateRoot = FindRuntimeRootUnderParentRetire(parentPath, warehouseId)
        If candidateRoot <> "" Then
            ResolveExistingRuntimeRootRetire = candidateRoot
            Exit Function
        End If
    End If

    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, warehouseId & ".invSys.", vbTextCompare) = 1 Then
            candidateRoot = wb.Path
            If RuntimeArtifactsExistRetire(candidateRoot, warehouseId) Then
                ResolveExistingRuntimeRootRetire = candidateRoot
                Exit Function
            End If
        End If
    Next wb

    candidateRoot = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(Trim$(warehouseId), False)
    If RuntimeArtifactsExistRetire(candidateRoot, warehouseId) Then ResolveExistingRuntimeRootRetire = candidateRoot
    Exit Function

CleanFail:
    ResolveExistingRuntimeRootRetire = vbNullString
End Function

Private Function FindRuntimeRootUnderParentRetire(ByVal parentPath As String, ByVal warehouseId As String) As String
    Dim childName As String
    Dim childPath As String

    On Error GoTo CleanFail

    parentPath = NormalizeFolderPathRetire(parentPath)
    If parentPath = "" Then Exit Function

    childName = Dir$(parentPath & "*", vbDirectory)
    Do While childName <> ""
        If childName <> "." And childName <> ".." Then
            childPath = parentPath & childName
            If FolderExistsRetire(childPath) Then
                If RuntimeArtifactsExistRetire(childPath, warehouseId) Then
                    FindRuntimeRootUnderParentRetire = childPath
                    Exit Function
                End If
            End If
        End If
        childName = Dir$
    Loop
    Exit Function

CleanFail:
    FindRuntimeRootUnderParentRetire = vbNullString
End Function

Private Function RuntimeArtifactsExistRetire(ByVal rootPath As String, ByVal warehouseId As String) As Boolean
    rootPath = NormalizeFolderPathRetire(rootPath)
    If rootPath = "" Then Exit Function
    RuntimeArtifactsExistRetire = _
        FileExistsRetire(rootPath & warehouseId & ".invSys.Config.xlsb") And _
        FileExistsRetire(rootPath & warehouseId & ".invSys.Auth.xlsb") And _
        FileExistsRetire(rootPath & warehouseId & ".invSys.Data.Inventory.xlsb")
End Function

Private Function TargetWarehouseRetiredRetire(ByVal targetWarehouseId As String, ByVal targetRoot As String) As Boolean
    Dim sharePointRoot As String

    targetRoot = NormalizeFolderPathRetire(targetRoot)
    If targetRoot = "" Then Exit Function

    If FileExistsRetire(targetRoot & targetWarehouseId & ".retired.json") Then
        TargetWarehouseRetiredRetire = True
        Exit Function
    End If
    If FileExistsRetire(targetRoot & "config\" & targetWarehouseId & ".retired.json") Then
        TargetWarehouseRetiredRetire = True
        Exit Function
    End If

    sharePointRoot = NormalizeFolderPathRetire(modConfig.GetString("PathSharePointRoot", ""))
    If sharePointRoot <> "" Then
        If FileExistsRetire(sharePointRoot & targetWarehouseId & ".retired.json") Then
            TargetWarehouseRetiredRetire = True
            Exit Function
        End If
    End If
End Function

Private Function ResolvePrimaryStationIdRetire(ByVal wbCfg As Workbook, ByVal warehouseId As String) As String
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim rowWarehouse As String

    On Error Resume Next
    Set lo = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")
    On Error GoTo 0
    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        rowWarehouse = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("WarehouseId").Index).Value))
        If rowWarehouse = "" Or StrComp(rowWarehouse, warehouseId, vbTextCompare) = 0 Then
            ResolvePrimaryStationIdRetire = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, lo.ListColumns("StationId").Index).Value))
            If ResolvePrimaryStationIdRetire <> "" Then Exit Function
        End If
    Next rowIndex
End Function

Private Function EnsureMigrationInboxReadyRetire(ByVal targetWarehouseId As String, _
                                                 ByVal targetStationId As String, _
                                                 ByVal targetConfigPath As String, _
                                                 ByRef report As String) As Boolean
    Dim inboxPath As String

    EnsureMigrationInboxReadyRetire = modConfig.EnsureStationInbox( _
        targetWarehouseId, targetStationId, "PRODUCTION", targetConfigPath, inboxPath, report)
End Function

Private Function QueueMigrationSeedEventsRetire(ByRef spec As RetireMigrateSpec, _
                                                ByVal archiveFolder As String, _
                                                ByVal snapshotPath As String, _
                                                ByVal targetStationId As String, _
                                                ByRef queuedCount As Long, _
                                                ByRef report As String) As Boolean
    Dim wbSnap As Workbook
    Dim loSnap As ListObject
    Dim rowIndex As Long
    Dim payloadItems As Collection
    Dim payloadJson As String
    Dim eventId As String
    Dim errorMessage As String
    Dim createdAtUtc As Date

    On Error GoTo FailQueue

    Set wbSnap = OpenWorkbookReadOnlyRetire(snapshotPath, report)
    If wbSnap Is Nothing Then Exit Function

    On Error Resume Next
    Set loSnap = wbSnap.Worksheets("InventorySnapshot").ListObjects("tblInventorySnapshot")
    On Error GoTo FailQueue
    If loSnap Is Nothing Or loSnap.DataBodyRange Is Nothing Then
        report = "Archived snapshot table was not available for migration."
        GoTo CleanExit
    End If

    createdAtUtc = ParseArchiveTimestampRetire(archiveFolder)
    If createdAtUtc = 0 Then createdAtUtc = Now

    For rowIndex = 1 To loSnap.ListRows.Count
        Set payloadItems = BuildMigrationPayloadItemsRetire(loSnap, rowIndex)
        If Not payloadItems Is Nothing Then
            If payloadItems.Count > 0 Then
                payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
                eventId = BuildMigrationEventIdRetire(spec.SourceWarehouseId, spec.TargetWarehouseId, archiveFolder, rowIndex)
                errorMessage = vbNullString
                If Not modRoleEventWriter.QueueMigrationSeedEvent(spec.TargetWarehouseId, targetStationId, spec.AdminUser, payloadJson, spec.SourceWarehouseId, _
                                                                  "Migration seed from " & spec.SourceWarehouseId & " via " & GetFileNameRetire(snapshotPath), _
                                                                  createdAtUtc, Nothing, eventId, errorMessage, eventId) Then
                    report = "QueueMigrationSeedEvent failed at snapshot row " & CStr(rowIndex) & ": " & errorMessage
                    GoTo CleanExit
                End If
                queuedCount = queuedCount + 1
            End If
        End If
    Next rowIndex

    QueueMigrationSeedEventsRetire = True

CleanExit:
    CloseWorkbookQuietlyRetire wbSnap
    Exit Function

FailQueue:
    report = "QueueMigrationSeedEventsRetire failed: " & Err.Description
    Resume CleanExit
End Function

Private Function BuildMigrationPayloadItemsRetire(ByVal loSnap As ListObject, ByVal rowIndex As Long) As Collection
    Dim sku As String
    Dim qtyOnHand As Double
    Dim locationSummary As String
    Dim locationVal As String
    Dim meta As Object
    Dim parsedAny As Boolean

    sku = GetTableTextRetire(loSnap, rowIndex, "SKU")
    If sku = "" Then Exit Function
    qtyOnHand = GetTableDoubleRetire(loSnap, rowIndex, "QtyOnHand")
    If qtyOnHand <= 0 Then Exit Function

    Set meta = BuildSnapshotMetadataRetire(loSnap, rowIndex)
    locationSummary = GetTableTextRetire(loSnap, rowIndex, "LocationSummary")
    locationVal = GetTableTextRetire(loSnap, rowIndex, "LOCATION")

    Set BuildMigrationPayloadItemsRetire = New Collection
    parsedAny = AppendLocationSummaryItemsRetire(BuildMigrationPayloadItemsRetire, sku, locationSummary, meta)
    If Not parsedAny Then
        BuildMigrationPayloadItemsRetire.Add CreateMigrationPayloadItemRetire(sku, qtyOnHand, locationVal, locationSummary, meta)
    End If
End Function

Private Function BuildSnapshotMetadataRetire(ByVal loSnap As ListObject, ByVal rowIndex As Long) As Object
    Dim meta As Object

    Set meta = CreateObject("Scripting.Dictionary")
    meta.CompareMode = vbTextCompare
    meta("ITEM") = GetTableTextRetire(loSnap, rowIndex, "ITEM")
    meta("UOM") = GetTableTextRetire(loSnap, rowIndex, "UOM")
    meta("LOCATION") = GetTableTextRetire(loSnap, rowIndex, "LOCATION")
    meta("DESCRIPTION") = GetTableTextRetire(loSnap, rowIndex, "DESCRIPTION")
    meta("VENDOR(s)") = GetTableTextRetire(loSnap, rowIndex, "VENDOR(s)")
    meta("VENDOR_CODE") = GetTableTextRetire(loSnap, rowIndex, "VENDOR_CODE")
    meta("CATEGORY") = GetTableTextRetire(loSnap, rowIndex, "CATEGORY")
    Set BuildSnapshotMetadataRetire = meta
End Function

Private Function AppendLocationSummaryItemsRetire(ByVal items As Collection, _
                                                  ByVal sku As String, _
                                                  ByVal locationSummary As String, _
                                                  ByVal meta As Object) As Boolean
    Dim normalized As String
    Dim tokens() As String
    Dim token As Variant
    Dim eqPos As Long
    Dim locationVal As String
    Dim qtyText As String
    Dim qtyVal As Double

    locationSummary = Trim$(locationSummary)
    If locationSummary = "" Then Exit Function

    normalized = Replace$(locationSummary, vbCrLf, ";")
    normalized = Replace$(normalized, vbCr, ";")
    normalized = Replace$(normalized, vbLf, ";")
    normalized = Replace$(normalized, ",", ";")
    tokens = Split(normalized, ";")
    For Each token In tokens
        token = Trim$(CStr(token))
        If token <> "" Then
            eqPos = InStr(1, CStr(token), "=", vbTextCompare)
            If eqPos <= 1 Then
                Set items = Nothing
                Exit Function
            End If
            locationVal = Trim$(Left$(CStr(token), eqPos - 1))
            qtyText = Trim$(Mid$(CStr(token), eqPos + 1))
            If locationVal = "" Or Not IsNumeric(qtyText) Then
                Set items = Nothing
                Exit Function
            End If
            qtyVal = CDbl(qtyText)
            If qtyVal > 0 Then
                items.Add CreateMigrationPayloadItemRetire(sku, qtyVal, locationVal, "", meta)
                AppendLocationSummaryItemsRetire = True
            End If
        End If
    Next token
End Function

Private Function CreateMigrationPayloadItemRetire(ByVal sku As String, _
                                                  ByVal qtyVal As Double, _
                                                  ByVal locationVal As String, _
                                                  ByVal noteVal As String, _
                                                  ByVal meta As Object) As Object
    Dim item As Object
    Dim key As Variant

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("SKU") = sku
    item("Qty") = qtyVal
    item("Location") = locationVal
    item("IoType") = "SEED"
    If Trim$(noteVal) <> "" Then item("Note") = noteVal
    If Not meta Is Nothing Then
        For Each key In meta.Keys
            item(CStr(key)) = meta(key)
        Next key
    End If
    Set CreateMigrationPayloadItemRetire = item
End Function

Private Function BuildMigrationEventIdRetire(ByVal sourceWarehouseId As String, _
                                             ByVal targetWarehouseId As String, _
                                             ByVal archiveFolder As String, _
                                             ByVal rowIndex As Long) As String
    BuildMigrationEventIdRetire = "MIG-" & SanitizeTokenRetire(sourceWarehouseId) & _
                                  "-TO-" & SanitizeTokenRetire(targetWarehouseId) & _
                                  "-" & SanitizeTokenRetire(GetFileNameRetire(archiveFolder)) & _
                                  "-" & Format$(rowIndex, "000000")
End Function

Private Function ParseArchiveTimestampRetire(ByVal archiveFolder As String) As Date
    Dim folderName As String
    Dim markerPos As Long
    Dim stampText As String

    folderName = GetFileNameRetire(archiveFolder)
    markerPos = InStrRev(folderName, "_archive_", -1, vbTextCompare)
    If markerPos = 0 Then Exit Function

    stampText = Mid$(folderName, markerPos + Len("_archive_"))
    If Len(stampText) >= 15 Then
        On Error Resume Next
        ParseArchiveTimestampRetire = DateSerial(CInt(Left$(stampText, 4)), CInt(Mid$(stampText, 5, 2)), CInt(Mid$(stampText, 7, 2))) + _
                                      TimeSerial(CInt(Mid$(stampText, 10, 2)), CInt(Mid$(stampText, 12, 2)), CInt(Mid$(stampText, 14, 2)))
        On Error GoTo 0
    End If
End Function

Private Function GetTableTextRetire(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim idx As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    idx = lo.ListColumns(columnName).Index
    If idx <= 0 Then Exit Function
    GetTableTextRetire = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, idx).Value))
End Function

Private Function GetTableDoubleRetire(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Double
    Dim valueIn As Variant

    On Error Resume Next
    valueIn = lo.DataBodyRange.Cells(rowIndex, lo.ListColumns(columnName).Index).Value
    If IsNumeric(valueIn) Then GetTableDoubleRetire = CDbl(valueIn)
    On Error GoTo 0
End Function

Private Function FindRowByValueInListObjectRetire(ByVal lo As ListObject, ByVal columnName As String, ByVal matchValue As String) As Long
    Dim colIndex As Long
    Dim rowIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Function
    On Error Resume Next
    colIndex = lo.ListColumns(columnName).Index
    On Error GoTo 0
    If colIndex <= 0 Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, colIndex).Value)), Trim$(matchValue), vbTextCompare) = 0 Then
            FindRowByValueInListObjectRetire = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Sub SetListObjectValueRetire(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueIn As Variant)
    Dim colIndex As Long

    If lo Is Nothing Or lo.DataBodyRange Is Nothing Then Exit Sub
    On Error Resume Next
    colIndex = lo.ListColumns(columnName).Index
    On Error GoTo 0
    If colIndex <= 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value = valueIn
End Sub

Private Function SanitizeTokenRetire(ByVal valueText As String) As String
    Dim i As Long
    Dim ch As String

    valueText = UCase$(Trim$(valueText))
    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If ch Like "[A-Z0-9]" Then
            SanitizeTokenRetire = SanitizeTokenRetire & ch
        Else
            SanitizeTokenRetire = SanitizeTokenRetire & "_"
        End If
    Next i
End Function

Private Sub RestoreCoreRootOverrideRetire(ByVal priorRootOverride As String)
    If Trim$(priorRootOverride) = "" Then
        modRuntimeWorkbooks.ClearCoreDataRootOverride
    Else
        modRuntimeWorkbooks.SetCoreDataRootOverride priorRootOverride
    End If
End Sub

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
