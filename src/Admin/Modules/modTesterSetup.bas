Attribute VB_Name = "modTesterSetup"
Option Explicit

Private Const TESTER_DEFAULT_WAREHOUSE_ID As String = "WH1"
Private Const TESTER_DEFAULT_STATION_ID As String = "R1"
Private Const TESTER_DEFAULT_OPERATOR_SUFFIX As String = ".Receiving.Operator.xlsm"
Private Const TESTER_SEED_SKU As String = "TEST-SKU-001"
Private Const TESTER_SEED_DESCRIPTION As String = "Test SKU for Confirm Writes"
Private Const TESTER_SEED_LOCATION As String = "A1"
Private Const TESTER_SEED_QTY As Double = 100#

Private mLastTesterSetupReport As String
Private mLastTesterOperatorWorkbookPath As String
Private mLastTesterSharePointRoot As String
Private mTesterSharePointRootOverride As String
Private mTesterSetupProgressSink As Object

Public Type TesterSetupSpec
    UserId As String
    PinHash As String
    WarehouseId As String
    StationId As String
    PathLocal As String
    PathSharePointRoot As String
End Type

Public Function SetupTesterStation(ByRef spec As TesterSetupSpec) As Boolean
    Dim report As String
    Dim rootPath As String
    Dim configPath As String
    Dim inboxPath As String
    Dim runtimeExists As Boolean
    Dim seedReport As String
    Dim authReport As String
    Dim operatorReport As String
    Dim configReport As String
    Dim sharePointReport As String
    Dim operatorPath As String
    Dim priorRootOverride As String
    Dim runtimeArtifactsExist As Boolean
    Dim runtimeCreated As Boolean

    On Error GoTo FailSetup

    ResetTesterSetupState
    NormalizeTesterSetupSpec spec
    ApplyTesterSetupDefaults spec

    If Not ValidateTesterSetupSpec(spec, report) Then GoTo FailSoft

    rootPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False)
    If rootPath = "" Then
        report = "PathLocal could not be resolved."
        GoTo FailSoft
    End If

    spec.PathLocal = rootPath
    configPath = rootPath & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    operatorPath = BuildReceivingOperatorPathTesterSetup(spec)
    mLastTesterOperatorWorkbookPath = operatorPath

    priorRootOverride = modRuntimeWorkbooks.GetCoreDataRootOverride()
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath

    runtimeExists = FolderExistsTesterSetup(rootPath)
    runtimeArtifactsExist = RuntimeArtifactsExistTesterSetup(rootPath, spec.WarehouseId)
    If Not runtimeExists And IsUncPathTesterSetup(rootPath) Then
        report = "Warehouse hub path is not accessible: " & rootPath
        GoTo FailSoft
    End If
    If Not runtimeArtifactsExist Then
        If IsUncPathTesterSetup(rootPath) Then
            PublishTesterSetupProgress "Creating tester runtime in hub..."
        Else
            PublishTesterSetupProgress "Creating tester runtime..."
        End If
        If Not CreateTesterRuntimeArtifactsTesterSetup(spec, rootPath, report) Then
            GoTo FailSoft
        End If
        runtimeExists = True
        runtimeArtifactsExist = True
        runtimeCreated = True
    End If
    EnsureFolderRecursiveTesterSetup rootPath & "\auth"

    If Not modConfig.EnsureStationConfigEntry(spec.WarehouseId, spec.StationId, Environ$("COMPUTERNAME"), _
                                              rootPath & "\inbox\", "RECEIVE", configPath, rootPath, report) Then
        GoTo FailSoft
    End If
    If Not modConfig.EnsureStationInbox(spec.WarehouseId, spec.StationId, "RECEIVE", configPath, inboxPath, report) Then
        GoTo FailSoft
    End If
    If Not modConfig.EnsureStationInbox(spec.WarehouseId, spec.StationId, "SHIP", configPath, inboxPath, report) Then
        GoTo FailSoft
    End If
    If Not modConfig.EnsureStationInbox(spec.WarehouseId, spec.StationId, "PRODUCTION", configPath, inboxPath, report) Then
        GoTo FailSoft
    End If

    PublishTesterSetupProgress "Seeding test data..."
    If Not SeedTesterScenarioInventory(spec, seedReport) Then
        report = seedReport
        GoTo FailSoft
    End If

    PublishTesterSetupProgress "Provisioning tester auth..."
    If Not ProvisionTesterAuth(spec, authReport) Then
        report = authReport
        GoTo FailSoft
    End If

    PublishTesterSetupProgress "Creating workbook..."
    If Not CreateOrVerifyReceivingWorkbook(spec, operatorPath, operatorReport) Then
        report = operatorReport
        GoTo FailSoft
    End If

    If Not UpdateConfigSharePointRoot(spec, configReport) Then
        report = configReport
        GoTo FailSoft
    End If
    If Not EnsureTesterSharePointPackage(spec, sharePointReport) Then
        report = sharePointReport
        GoTo FailSoft
    End If

    report = "OK|WarehouseId=" & spec.WarehouseId & _
             "|StationId=" & spec.StationId & _
             "|UserId=" & spec.UserId & _
             "|Runtime=" & IIf(runtimeCreated, "CREATED", "EXISTING") & _
             "|Inbox=" & inboxPath & _
             "|Seed=" & seedReport & _
             "|Auth=" & authReport & _
             "|Operator=" & operatorPath & _
             "|SharePoint=" & spec.PathSharePointRoot & _
             "|SharePointPackage=" & sharePointReport
    PublishTesterSetupProgress "Done"
    SetupTesterStation = True
    GoTo CleanExit

FailSoft:
    SetupTesterStation = False
    If Len(report) = 0 Then report = "SetupTesterStation failed."
    LogDiagnosticEvent "TESTER-SETUP", report
    GoTo CleanExit

FailSetup:
    report = "SetupTesterStation failed: " & Err.Description
    Resume FailSoft

CleanExit:
    mLastTesterSetupReport = report
    mLastTesterSharePointRoot = spec.PathSharePointRoot
    RestoreCoreRootOverrideTesterSetup priorRootOverride
End Function

Public Function ProvisionTesterAuth(ByRef spec As TesterSetupSpec, _
                                    Optional ByRef report As String = "") As Boolean
    Dim authPath As String
    Dim wbAuth As Workbook
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim userRow As Long
    Dim capabilitiesReady As Boolean
    Dim hashReady As Boolean
    Dim openedTransient As Boolean
    Dim displayName As String

    On Error GoTo FailProvision

    NormalizeTesterSetupSpec spec
    ApplyTesterSetupDefaults spec

    authPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False) & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"
    Set wbAuth = OpenWorkbookForWriteTesterSetup(authPath, openedTransient, report)
    If wbAuth Is Nothing Then GoTo FailSoft
    If Not modAuth.EnsureAuthSchema(wbAuth, spec.WarehouseId, "svc_processor", report) Then GoTo FailSoft

    Set loUsers = FindTableByNameTesterSetup(wbAuth, "tblUsers")
    Set loCaps = FindTableByNameTesterSetup(wbAuth, "tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then
        report = "Auth tables were not available for tester provisioning."
        GoTo FailSoft
    End If

    userRow = EnsureUserRowTesterSetup(loUsers, spec.UserId)
    displayName = spec.UserId
    SetTableCellTesterSetup loUsers, userRow, "UserId", spec.UserId
    SetTableCellTesterSetup loUsers, userRow, "DisplayName", displayName
    SetTableCellTesterSetup loUsers, userRow, "Status", "Active"

    If SafeTrimTesterSetup(GetTableCellTesterSetup(loUsers, userRow, "PinHash")) = "" Then
        SetTableCellTesterSetup loUsers, userRow, "PinHash", spec.PinHash
    End If

    EnsureCapabilityActiveTesterSetup loCaps, spec.UserId, "RECEIVE_POST", spec.WarehouseId, spec.StationId
    EnsureCapabilityActiveTesterSetup loCaps, spec.UserId, "RECEIVE_VIEW", spec.WarehouseId, spec.StationId
    EnsureCapabilityActiveTesterSetup loCaps, spec.UserId, "SHIP_POST", spec.WarehouseId, spec.StationId
    EnsureCapabilityActiveTesterSetup loCaps, spec.UserId, "PROD_POST", spec.WarehouseId, spec.StationId
    EnsureCapabilityActiveTesterSetup loCaps, spec.UserId, "READMODEL_REFRESH", spec.WarehouseId, spec.StationId
    DeactivateCapabilityTesterSetup loCaps, spec.UserId, "ADMIN_MAINT", spec.WarehouseId, spec.StationId

    hashReady = (SafeTrimTesterSetup(GetTableCellTesterSetup(loUsers, userRow, "PinHash")) <> "")
    capabilitiesReady = _
        CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "RECEIVE_POST", spec.WarehouseId, spec.StationId) And _
        CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "RECEIVE_VIEW", spec.WarehouseId, spec.StationId) And _
        CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "SHIP_POST", spec.WarehouseId, spec.StationId) And _
        CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "PROD_POST", spec.WarehouseId, spec.StationId) And _
        CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "READMODEL_REFRESH", spec.WarehouseId, spec.StationId) And _
        Not CapabilityIsActiveTesterSetup(loCaps, spec.UserId, "ADMIN_MAINT", spec.WarehouseId, spec.StationId)

    wbAuth.Save
    ProvisionTesterAuth = (hashReady And capabilitiesReady)
    If ProvisionTesterAuth Then
        report = "OK"
    Else
        report = "Tester auth verification failed after provisioning."
    End If
    GoTo CleanExit

FailSoft:
    ProvisionTesterAuth = False
    If Len(report) = 0 Then report = "ProvisionTesterAuth failed."
    LogDiagnosticEvent "TESTER-SETUP", report
    GoTo CleanExit

FailProvision:
    report = "ProvisionTesterAuth failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookIfTransientTesterSetup wbAuth, openedTransient
End Function

Public Function VerifyReceivingWorkbook(ByVal pathIn As String, _
                                        Optional ByRef report As String = "") As Boolean
    Dim wb As Workbook
    Dim openedTransient As Boolean

    On Error GoTo FailVerify

    pathIn = Trim$(pathIn)
    If pathIn = "" Then
        report = "Receiving workbook path is required."
        GoTo FailSoft
    End If
    If Not FileExistsTesterSetup(pathIn) Then
        report = "Receiving workbook not found: " & pathIn
        GoTo FailSoft
    End If

    Set wb = FindOpenWorkbookByPathTesterSetup(pathIn)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(Filename:=pathIn, UpdateLinks:=0, ReadOnly:=True, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        openedTransient = Not wb Is Nothing
    End If
    If wb Is Nothing Then
        report = "Receiving workbook could not be opened."
        GoTo FailSoft
    End If

    If ReceivingWorkbookHasCanonicalSurfacesTesterSetup(wb) Or ReceivingWorkbookHasAliasSurfacesTesterSetup(wb) Then
        VerifyReceivingWorkbook = True
        report = "OK"
        GoTo CleanExit
    End If

    report = "Receiving workbook is missing required sheets or tables."
    GoTo FailSoft

FailSoft:
    VerifyReceivingWorkbook = False
    If Len(report) = 0 Then report = "VerifyReceivingWorkbook failed."
    LogDiagnosticEvent "TESTER-SETUP", report
    GoTo CleanExit

FailVerify:
    report = "VerifyReceivingWorkbook failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookIfTransientTesterSetup wb, openedTransient
End Function

Public Function DetectSharePointRoot(Optional ByVal warehouseId As String = TESTER_DEFAULT_WAREHOUSE_ID) As String
    Dim candidate As String
    Dim searchRoots As Collection
    Dim rootPath As Variant
    Dim fso As Object
    Dim folderObj As Object
    Dim rootFolder As Object

    On Error GoTo FailDetect
    warehouseId = Trim$(warehouseId)
    If warehouseId = "" Then warehouseId = TESTER_DEFAULT_WAREHOUSE_ID

    If Trim$(mTesterSharePointRootOverride) <> "" Then
        DetectSharePointRoot = NormalizeFolderPathTesterSetup(mTesterSharePointRootOverride, False)
        Exit Function
    End If

    candidate = Trim$(modConfig.GetString("PathSharePointRoot", ""))
    If CandidateLooksLikeSharePointRootTesterSetup(candidate, warehouseId) Then
        DetectSharePointRoot = NormalizeFolderPathTesterSetup(candidate, False)
        Exit Function
    End If

    Set searchRoots = New Collection
    AddUniquePathTesterSetup searchRoots, Environ$("OneDriveCommercial")
    AddUniquePathTesterSetup searchRoots, Environ$("OneDrive")
    AddUniquePathTesterSetup searchRoots, Environ$("USERPROFILE")

    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each rootPath In searchRoots
        candidate = NormalizeFolderPathTesterSetup(CStr(rootPath), False)
        If CandidateLooksLikeSharePointRootTesterSetup(candidate, warehouseId) Then
            DetectSharePointRoot = candidate
            Exit Function
        End If

        If candidate <> "" And FolderExistsTesterSetup(candidate) Then
            On Error Resume Next
            Set rootFolder = fso.GetFolder(candidate)
            On Error GoTo FailDetect
            If Not rootFolder Is Nothing Then
                For Each folderObj In rootFolder.SubFolders
                    If CandidateLooksLikeSharePointRootTesterSetup(folderObj.Path, warehouseId) Then
                        DetectSharePointRoot = NormalizeFolderPathTesterSetup(folderObj.Path, False)
                        Exit Function
                    End If
                Next folderObj
            End If
            Set rootFolder = Nothing
        End If
    Next rootPath
    Exit Function

FailDetect:
    DetectSharePointRoot = vbNullString
End Function

Public Function BrowseForSharePointRoot(Optional ByVal initialPath As String = "") As String
    Const FILE_DIALOG_FOLDER_PICKER As Long = 4
    Dim picker As Object

    On Error GoTo FailBrowse
    Set picker = Application.FileDialog(FILE_DIALOG_FOLDER_PICKER)
    If picker Is Nothing Then Exit Function

    picker.Title = "Select SharePoint Sync Root"
    picker.AllowMultiSelect = False
    If Trim$(initialPath) <> "" And FolderExistsTesterSetup(initialPath) Then
        picker.InitialFileName = NormalizeFolderPathTesterSetup(initialPath, True)
    End If
    If picker.Show <> -1 Then Exit Function
    If picker.SelectedItems.Count = 0 Then Exit Function

    BrowseForSharePointRoot = NormalizeFolderPathTesterSetup(CStr(picker.SelectedItems(1)), False)
    Exit Function

FailBrowse:
    BrowseForSharePointRoot = vbNullString
End Function

Public Function OpenTesterReceivingWorkbook(Optional ByVal workbookPath As String = "") As Boolean
    Dim targetPath As String
    Dim wb As Workbook
    Dim defaultSpec As TesterSetupSpec

    On Error GoTo FailOpen

    targetPath = Trim$(workbookPath)
    If targetPath = "" Then targetPath = mLastTesterOperatorWorkbookPath
    If targetPath = "" Then
        defaultSpec = BuildDefaultTesterSpecTesterSetup()
        targetPath = BuildReceivingOperatorPathTesterSetup(defaultSpec)
    End If
    If targetPath = "" Or Not FileExistsTesterSetup(targetPath) Then Exit Function

    Set wb = FindOpenWorkbookByPathTesterSetup(targetPath)
    If wb Is Nothing Then Set wb = Application.Workbooks.Open(Filename:=targetPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    If wb Is Nothing Then Exit Function
    wb.Activate
    OpenTesterReceivingWorkbook = True
    Exit Function

FailOpen:
    LogDiagnosticEvent "TESTER-SETUP", "OpenTesterReceivingWorkbook failed|Path=" & targetPath & "|Reason=" & Err.Description
End Function

Public Function DeleteTesterStationGenerated(ByRef spec As TesterSetupSpec, _
                                             Optional ByVal dryRun As Boolean = False, _
                                             Optional ByRef report As String = "") As Boolean
    Dim rootPath As String
    Dim sharePointRoot As String
    Dim artifacts As Collection
    Dim artifactPath As Variant
    Dim deletedCount As Long
    Dim missingCount As Long
    Dim failedCount As Long
    Dim detail As String
    Dim deleteStatus As String

    On Error GoTo FailDelete

    NormalizeTesterSetupSpec spec
    ApplyTesterSetupDefaults spec

    rootPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False)
    sharePointRoot = NormalizeFolderPathTesterSetup(spec.PathSharePointRoot, False)
    If rootPath = "" Then
        report = "Warehouse hub path is required."
        GoTo FailSoft
    End If
    If spec.WarehouseId = "" Then
        report = "WarehouseId is required."
        GoTo FailSoft
    End If
    If Not IsTesterWarehouseIdForDeleteTesterSetup(spec.WarehouseId) Then
        report = "Refusing cleanup for non-tester WarehouseId: " & spec.WarehouseId
        GoTo FailSoft
    End If
    If IsUncPathTesterSetup(rootPath) And Not FolderExistsTesterSetup(rootPath) Then
        report = "Warehouse hub path is not accessible: " & rootPath
        GoTo FailSoft
    End If

    Set artifacts = BuildTesterGeneratedArtifactsTesterSetup(spec, rootPath, sharePointRoot)
    For Each artifactPath In artifacts
        deleteStatus = DeleteGeneratedArtifactTesterSetup(CStr(artifactPath), dryRun)
        If Left$(deleteStatus, 8) = "DELETED|" Or Left$(deleteStatus, 7) = "DRYRUN|" Then
            deletedCount = deletedCount + 1
        ElseIf Left$(deleteStatus, 8) = "MISSING|" Then
            missingCount = missingCount + 1
        Else
            failedCount = failedCount + 1
        End If
        If Len(detail) > 0 Then detail = detail & "; "
        detail = detail & deleteStatus
    Next artifactPath

    DeleteEmptyTesterGeneratedFoldersTesterSetup spec, rootPath, sharePointRoot, dryRun, detail

    If failedCount > 0 Then
        report = "Cleanup completed with failures. Deleted=" & CStr(deletedCount) & _
                 "; Missing=" & CStr(missingCount) & _
                 "; Failed=" & CStr(failedCount) & _
                 "; " & detail
        GoTo FailSoft
    End If

    mLastTesterOperatorWorkbookPath = vbNullString
    report = IIf(dryRun, "DRYRUN", "OK") & _
             "|Deleted=" & CStr(deletedCount) & _
             "|Missing=" & CStr(missingCount) & _
             "|WarehouseId=" & spec.WarehouseId & _
             "|Hub=" & rootPath & _
             "|" & detail
    mLastTesterSetupReport = report
    DeleteTesterStationGenerated = True
    LogDiagnosticEvent "TESTER-SETUP", "DeleteTesterStationGenerated|" & report
    Exit Function

FailSoft:
    DeleteTesterStationGenerated = False
    If Len(report) = 0 Then report = "DeleteTesterStationGenerated failed."
    mLastTesterSetupReport = report
    LogDiagnosticEvent "TESTER-SETUP", report
    Exit Function

FailDelete:
    report = "DeleteTesterStationGenerated failed: " & Err.Description
    Resume FailSoft
End Function

Public Function GetLastTesterSetupReport() As String
    GetLastTesterSetupReport = mLastTesterSetupReport
End Function

Public Function GetLastTesterOperatorWorkbookPath() As String
    GetLastTesterOperatorWorkbookPath = mLastTesterOperatorWorkbookPath
End Function

Public Function GetLastTesterSharePointRoot() As String
    GetLastTesterSharePointRoot = mLastTesterSharePointRoot
End Function

Public Sub SetTesterSetupProgressSink(ByVal progressSink As Object)
    Set mTesterSetupProgressSink = progressSink
End Sub

Public Sub ClearTesterSetupProgressSink()
    Set mTesterSetupProgressSink = Nothing
End Sub

Public Sub SetTesterSharePointRootOverride(ByVal rootPath As String)
    mTesterSharePointRootOverride = Trim$(rootPath)
End Sub

Public Sub ClearTesterSharePointRootOverride()
    mTesterSharePointRootOverride = vbNullString
End Sub

Private Function SeedTesterScenarioInventory(ByRef spec As TesterSetupSpec, _
                                             ByRef report As String) As Boolean
    Dim inventoryWb As Workbook
    Dim inventoryPath As String
    Dim openedTransient As Boolean
    Dim seedPayload As Object
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim tempAdminGranted As Boolean
    Dim tempAdminReport As String

    On Error GoTo FailSeed

    Set inventoryWb = modInventoryDomainBridge.ResolveInventoryWorkbookBridge(spec.WarehouseId, Nothing)
    If inventoryWb Is Nothing Then
        inventoryPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False) & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb"
        Set inventoryWb = OpenWorkbookForWriteTesterSetup(inventoryPath, openedTransient, report)
        If inventoryWb Is Nothing Then
            If Len(report) = 0 Then report = "Canonical inventory workbook could not be resolved."
            GoTo FailSoft
        End If
        If Not EnsureInventorySchemaBridge(inventoryWb, report) Then GoTo FailSoft
    End If

    If TesterSeedAlreadyPresentTesterSetup(inventoryWb) Then
        SeedTesterScenarioInventory = True
        report = "SKIPPED"
        GoTo CleanExit
    End If

    If Not EnsureTemporaryAdminCapabilityTesterSetup(spec, tempAdminGranted, tempAdminReport) Then
        report = tempAdminReport
        GoTo FailSoft
    End If

    If Not modConfig.LoadConfig(spec.WarehouseId, spec.StationId) Then
        report = "Config load failed: " & modConfig.Validate()
        GoTo FailSoft
    End If
    If Not modAuth.LoadAuth(spec.WarehouseId) Then
        report = "Auth load failed: " & modAuth.ValidateAuth()
        GoTo FailSoft
    End If

    Set seedPayload = modRoleEventWriter.CreatePayloadItem(1, TESTER_SEED_SKU, TESTER_SEED_QTY, TESTER_SEED_LOCATION, "Tester setup seed", "IMPORT")
    seedPayload("Description") = TESTER_SEED_DESCRIPTION
    seedPayload("Item_Code") = TESTER_SEED_SKU
    seedPayload("Item") = TESTER_SEED_SKU
    seedPayload("UOM") = "EA"
    payloadJson = modRoleEventWriter.BuildPayloadJson(seedPayload)

    If Not modRoleEventWriter.QueueMigrationSeedEvent(spec.WarehouseId, spec.StationId, spec.UserId, payloadJson, "", "TESTER_SETUP_SEED", 0, Nothing, eventIdOut, queueError, "") Then
        report = "QueueMigrationSeedEvent failed: " & queueError
        GoTo FailSoft
    End If

    If modProcessor.RunBatch(spec.WarehouseId, 0, batchReport) < 1 Then
        If Not TesterSeedAlreadyPresentTesterSetup(inventoryWb) Then
            report = "Processor did not apply the tester seed. " & batchReport
            GoTo FailSoft
        End If
    End If

    If Not TesterSeedAlreadyPresentTesterSetup(inventoryWb) Then
        report = "Tester seed SKU was not present after processor run."
        GoTo FailSoft
    End If

    SeedTesterScenarioInventory = True
    report = "SEEDED"
    GoTo CleanExit

FailSoft:
    SeedTesterScenarioInventory = False
    If Len(report) = 0 Then report = "SeedTesterScenarioInventory failed."
    LogDiagnosticEvent "TESTER-SETUP", report
    GoTo CleanExit

FailSeed:
    report = "SeedTesterScenarioInventory failed: " & Err.Description
    Resume FailSoft

CleanExit:
    RevokeAdminCapabilityTesterSetup spec
    CloseWorkbookIfTransientTesterSetup inventoryWb, openedTransient
End Function

Private Function CreateOrVerifyReceivingWorkbook(ByRef spec As TesterSetupSpec, _
                                                 ByVal operatorPath As String, _
                                                 ByRef report As String) As Boolean
    Dim wb As Workbook
    Dim openedTransient As Boolean
    Dim refreshReport As String
    Dim parentFolder As String
    Dim prevEvents As Boolean
    Dim prevDisplayAlerts As Boolean

    On Error GoTo FailCreate
    prevEvents = Application.EnableEvents
    prevDisplayAlerts = Application.DisplayAlerts

    parentFolder = GetParentFolderTesterSetup(operatorPath)
    If parentFolder = "" Then
        report = "Operator workbook parent folder could not be resolved."
        GoTo FailSoft
    End If
    EnsureFolderRecursiveTesterSetup parentFolder

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set wb = FindOpenWorkbookByPathTesterSetup(operatorPath)
    If wb Is Nothing And FileExistsTesterSetup(operatorPath) Then
        Set wb = Application.Workbooks.Open(Filename:=operatorPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        openedTransient = Not wb Is Nothing
    End If
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Add(xlWBATWorksheet)
        openedTransient = Not wb Is Nothing
    End If
    If wb Is Nothing Then
        report = "Operator workbook could not be created."
        GoTo FailSoft
    End If

    If Not modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface(wb, report) Then GoTo FailSoft
    RemoveNonReceivingOperatorSheetsTesterSetup wb
    Call modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wb, spec.WarehouseId, "LOCAL", refreshReport)

    If Trim$(wb.FullName) = "" Then
        wb.SaveAs Filename:=operatorPath, FileFormat:=52
    ElseIf StrComp(wb.FullName, operatorPath, vbTextCompare) <> 0 Then
        wb.SaveAs Filename:=operatorPath, FileFormat:=52
    Else
        wb.Save
    End If

    If Not VerifyReceivingWorkbook(operatorPath, report) Then GoTo FailSoft

    mLastTesterOperatorWorkbookPath = operatorPath
    CreateOrVerifyReceivingWorkbook = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    CreateOrVerifyReceivingWorkbook = False
    If Len(report) = 0 Then report = "CreateOrVerifyReceivingWorkbook failed."
    LogDiagnosticEvent "TESTER-SETUP", report
    GoTo CleanExit

FailCreate:
    report = "CreateOrVerifyReceivingWorkbook failed: " & Err.Description
    Resume FailSoft

CleanExit:
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEvents
    CloseWorkbookIfTransientTesterSetup wb, openedTransient
End Function

Private Sub RemoveNonReceivingOperatorSheetsTesterSetup(ByVal wb As Workbook)
    Dim keepSheets As Object
    Dim i As Long
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Sub
    Set keepSheets = CreateObject("Scripting.Dictionary")
    keepSheets.CompareMode = vbTextCompare
    keepSheets("ReceivedTally") = True
    keepSheets("InventoryManagement") = True
    keepSheets("ReceivedLog") = True

    For i = wb.Worksheets.Count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If Not keepSheets.Exists(ws.Name) Then
            If wb.Worksheets.Count > 1 Then ws.Delete
        End If
    Next i
End Sub

Private Function UpdateConfigSharePointRoot(ByRef spec As TesterSetupSpec, _
                                            ByRef report As String) As Boolean
    Dim configPath As String
    Dim wbCfg As Workbook
    Dim loWh As ListObject
    Dim rowIndex As Long
    Dim openedTransient As Boolean

    On Error GoTo FailUpdate

    configPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False) & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    Set wbCfg = OpenWorkbookForWriteTesterSetup(configPath, openedTransient, report)
    If wbCfg Is Nothing Then GoTo FailSoft
    If Not modConfig.EnsureConfigSchema(wbCfg, spec.WarehouseId, spec.StationId, report) Then GoTo FailSoft

    Set loWh = FindTableByNameTesterSetup(wbCfg, "tblWarehouseConfig")
    If loWh Is Nothing Or loWh.DataBodyRange Is Nothing Then
        report = "Warehouse config table not available."
        GoTo FailSoft
    End If

    rowIndex = FindRowByValueTesterSetup(loWh, "WarehouseId", spec.WarehouseId)
    If rowIndex = 0 Then rowIndex = 1

    SetTableCellTesterSetup loWh, rowIndex, "WarehouseId", spec.WarehouseId
    If SafeTrimTesterSetup(GetTableCellTesterSetup(loWh, rowIndex, "WarehouseName")) = "" Then
        SetTableCellTesterSetup loWh, rowIndex, "WarehouseName", spec.WarehouseId
    End If
    SetTableCellTesterSetup loWh, rowIndex, "PathDataRoot", NormalizeFolderPathTesterSetup(spec.PathLocal, False)
    SetTableCellTesterSetup loWh, rowIndex, "PathSharePointRoot", NormalizeFolderPathTesterSetup(spec.PathSharePointRoot, False)
    wbCfg.Save

    UpdateConfigSharePointRoot = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    UpdateConfigSharePointRoot = False
    If Len(report) = 0 Then report = "UpdateConfigSharePointRoot failed."
    LogDiagnosticEvent "TESTER-SETUP", report
    GoTo CleanExit

FailUpdate:
    report = "UpdateConfigSharePointRoot failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookIfTransientTesterSetup wbCfg, openedTransient
End Function

Private Function ValidateTesterSetupSpec(ByRef spec As TesterSetupSpec, _
                                         ByRef report As String) As Boolean
    If spec.UserId = "" Then
        report = "UserId is required."
        Exit Function
    End If
    If spec.PinHash = "" Then
        report = "PinHash is required."
        Exit Function
    End If
    If spec.WarehouseId = "" Then
        report = "WarehouseId is required."
        Exit Function
    End If
    If spec.StationId = "" Then
        report = "StationId is required."
        Exit Function
    End If
    If spec.PathLocal = "" Then
        report = "PathLocal is required."
        Exit Function
    End If

    ValidateTesterSetupSpec = True
    report = "OK"
End Function

Private Sub NormalizeTesterSetupSpec(ByRef spec As TesterSetupSpec)
    spec.UserId = Trim$(spec.UserId)
    spec.PinHash = Trim$(spec.PinHash)
    spec.WarehouseId = Trim$(spec.WarehouseId)
    spec.StationId = Trim$(spec.StationId)
    spec.PathLocal = Trim$(Replace$(spec.PathLocal, "/", "\"))
    spec.PathSharePointRoot = Trim$(Replace$(spec.PathSharePointRoot, "/", "\"))
End Sub

Private Sub ApplyTesterSetupDefaults(ByRef spec As TesterSetupSpec)
    If spec.WarehouseId = "" Then spec.WarehouseId = TESTER_DEFAULT_WAREHOUSE_ID
    If spec.StationId = "" Then spec.StationId = TESTER_DEFAULT_STATION_ID
    If spec.PathLocal = "" Then spec.PathLocal = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(spec.WarehouseId, False)
    If spec.PathSharePointRoot = "" Then spec.PathSharePointRoot = DetectSharePointRoot(spec.WarehouseId)
    spec.PathLocal = NormalizeFolderPathTesterSetup(spec.PathLocal, False)
    spec.PathSharePointRoot = NormalizeFolderPathTesterSetup(spec.PathSharePointRoot, False)
End Sub

Private Function BuildReceivingOperatorPathTesterSetup(ByRef spec As TesterSetupSpec) As String
    Dim rootPath As String

    rootPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False)
    If rootPath = "" Then Exit Function
    BuildReceivingOperatorPathTesterSetup = rootPath & "\" & spec.WarehouseId & TESTER_DEFAULT_OPERATOR_SUFFIX
End Function

Private Function IsTesterWarehouseIdForDeleteTesterSetup(ByVal warehouseId As String) As Boolean
    warehouseId = UCase$(Trim$(warehouseId))
    If warehouseId = "" Then Exit Function

    IsTesterWarehouseIdForDeleteTesterSetup = _
        (warehouseId = "TESTSTATION") Or _
        (Left$(warehouseId, 4) = "TEST") Or _
        (InStr(1, warehouseId, "TESTER", vbTextCompare) > 0)
End Function

Private Function BuildTesterGeneratedArtifactsTesterSetup(ByRef spec As TesterSetupSpec, _
                                                          ByVal rootPath As String, _
                                                          ByVal sharePointRoot As String) As Collection
    Dim artifacts As Collection
    Dim warehouseId As String
    Dim stationId As String

    Set artifacts = New Collection
    warehouseId = Trim$(spec.WarehouseId)
    stationId = Trim$(spec.StationId)
    rootPath = NormalizeFolderPathTesterSetup(rootPath, False)
    sharePointRoot = NormalizeFolderPathTesterSetup(sharePointRoot, False)

    AddGeneratedArtifactTesterSetup artifacts, rootPath & "\" & warehouseId & ".invSys.Config.xlsb"
    AddGeneratedArtifactTesterSetup artifacts, rootPath & "\" & warehouseId & ".invSys.Auth.xlsb"
    AddGeneratedArtifactTesterSetup artifacts, rootPath & "\" & warehouseId & ".invSys.Data.Inventory.xlsb"
    AddGeneratedArtifactTesterSetup artifacts, rootPath & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
    AddGeneratedArtifactTesterSetup artifacts, rootPath & "\" & warehouseId & ".Outbox.Events.xlsb"
    AddGeneratedArtifactTesterSetup artifacts, rootPath & "\" & warehouseId & TESTER_DEFAULT_OPERATOR_SUFFIX

    If stationId <> "" Then
        AddGeneratedArtifactTesterSetup artifacts, rootPath & "\inbox\invSys.Inbox.Receiving." & stationId & ".xlsb"
        AddGeneratedArtifactTesterSetup artifacts, rootPath & "\inbox\invSys.Inbox.Shipping." & stationId & ".xlsb"
        AddGeneratedArtifactTesterSetup artifacts, rootPath & "\inbox\invSys.Inbox.Production." & stationId & ".xlsb"
    End If

    If sharePointRoot <> "" Then
        AddGeneratedArtifactTesterSetup artifacts, sharePointRoot & "\TesterPackage\" & warehouseId & "\" & warehouseId & ".TesterBundle.zip"
        AddGeneratedArtifactTesterSetup artifacts, sharePointRoot & "\TesterPackage\" & warehouseId & "\README.txt"
        AddGeneratedArtifactTesterSetup artifacts, sharePointRoot & "\Snapshots\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb"
        AddGeneratedArtifactTesterSetup artifacts, sharePointRoot & "\Events\" & warehouseId & ".Outbox.Events.xlsb"
    End If

    Set BuildTesterGeneratedArtifactsTesterSetup = artifacts
End Function

Private Sub AddGeneratedArtifactTesterSetup(ByVal artifacts As Collection, ByVal artifactPath As String)
    artifactPath = Trim$(Replace$(artifactPath, "/", "\"))
    If artifactPath = "" Then Exit Sub
    artifacts.Add artifactPath
End Sub

Private Function DeleteGeneratedArtifactTesterSetup(ByVal artifactPath As String, ByVal dryRun As Boolean) As String
    Dim wb As Workbook
    Dim fso As Object

    On Error GoTo FailDelete
    artifactPath = Trim$(Replace$(artifactPath, "/", "\"))
    If artifactPath = "" Then
        DeleteGeneratedArtifactTesterSetup = "SKIP|EmptyPath"
        Exit Function
    End If
    If Not IsUsableLocalPathTesterSetup(artifactPath) Then
        DeleteGeneratedArtifactTesterSetup = "SKIP|UnusablePath=" & artifactPath
        Exit Function
    End If
    If Not FileExistsTesterSetup(artifactPath) Then
        DeleteGeneratedArtifactTesterSetup = "MISSING|" & artifactPath
        Exit Function
    End If
    If dryRun Then
        DeleteGeneratedArtifactTesterSetup = "DRYRUN|" & artifactPath
        Exit Function
    End If

    Set wb = FindOpenWorkbookByPathTesterSetup(artifactPath)
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFile artifactPath, True
    If FileExistsTesterSetup(artifactPath) Then
        DeleteGeneratedArtifactTesterSetup = "FAILED|" & artifactPath & "|StillExists"
    Else
        DeleteGeneratedArtifactTesterSetup = "DELETED|" & artifactPath
    End If
    Exit Function

FailDelete:
    DeleteGeneratedArtifactTesterSetup = "FAILED|" & artifactPath & "|" & Err.Description
End Function

Private Sub DeleteEmptyTesterGeneratedFoldersTesterSetup(ByRef spec As TesterSetupSpec, _
                                                         ByVal rootPath As String, _
                                                         ByVal sharePointRoot As String, _
                                                         ByVal dryRun As Boolean, _
                                                         ByRef detail As String)
    Dim folderPath As String
    Dim statusText As String

    folderPath = NormalizeFolderPathTesterSetup(sharePointRoot, False)
    If folderPath <> "" Then
        statusText = DeleteEmptyFolderTesterSetup(folderPath & "\TesterPackage\" & spec.WarehouseId, dryRun)
        If statusText <> "" Then detail = detail & "; " & statusText
    End If

    rootPath = NormalizeFolderPathTesterSetup(rootPath, False)
    If rootPath = "" Then Exit Sub
    statusText = DeleteEmptyFolderTesterSetup(rootPath & "\auth", dryRun)
    If statusText <> "" Then detail = detail & "; " & statusText
    statusText = DeleteEmptyFolderTesterSetup(rootPath & "\config", dryRun)
    If statusText <> "" Then detail = detail & "; " & statusText
    statusText = DeleteEmptyFolderTesterSetup(rootPath & "\snapshots", dryRun)
    If statusText <> "" Then detail = detail & "; " & statusText
    statusText = DeleteEmptyFolderTesterSetup(rootPath & "\outbox", dryRun)
    If statusText <> "" Then detail = detail & "; " & statusText
    statusText = DeleteEmptyFolderTesterSetup(rootPath & "\inbox", dryRun)
    If statusText <> "" Then detail = detail & "; " & statusText
End Sub

Private Function DeleteEmptyFolderTesterSetup(ByVal folderPath As String, ByVal dryRun As Boolean) As String
    Dim fso As Object
    Dim folderObj As Object

    On Error GoTo FailDelete
    folderPath = NormalizeFolderPathTesterSetup(folderPath, False)
    If folderPath = "" Then Exit Function
    If Not FolderExistsTesterSetup(folderPath) Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folderObj = fso.GetFolder(folderPath)
    If folderObj.Files.Count > 0 Or folderObj.SubFolders.Count > 0 Then
        DeleteEmptyFolderTesterSetup = "KEEP_FOLDER_NOT_EMPTY|" & folderPath
        Exit Function
    End If
    If dryRun Then
        DeleteEmptyFolderTesterSetup = "DRYRUN_EMPTY_FOLDER|" & folderPath
    Else
        fso.DeleteFolder folderPath, True
        DeleteEmptyFolderTesterSetup = "DELETED_EMPTY_FOLDER|" & folderPath
    End If
    Exit Function

FailDelete:
    DeleteEmptyFolderTesterSetup = "FAILED_EMPTY_FOLDER|" & folderPath & "|" & Err.Description
End Function

Private Function RuntimeArtifactsExistTesterSetup(ByVal rootPath As String, ByVal warehouseId As String) As Boolean
    rootPath = NormalizeFolderPathTesterSetup(rootPath, False)
    warehouseId = Trim$(warehouseId)
    If rootPath = "" Or warehouseId = "" Then Exit Function

    RuntimeArtifactsExistTesterSetup = _
        FileExistsTesterSetup(rootPath & "\" & warehouseId & ".invSys.Config.xlsb") And _
        FileExistsTesterSetup(rootPath & "\" & warehouseId & ".invSys.Auth.xlsb") And _
        FileExistsTesterSetup(rootPath & "\" & warehouseId & ".invSys.Data.Inventory.xlsb")
End Function

Private Function CreateTesterRuntimeArtifactsTesterSetup(ByRef spec As TesterSetupSpec, _
                                                         ByVal rootPath As String, _
                                                         ByRef report As String) As Boolean
    Dim configPath As String
    Dim authPath As String
    Dim inventoryPath As String
    Dim outboxPath As String
    Dim snapshotPath As String
    Dim capabilityOut As String
    Dim wbInventory As Workbook
    Dim wbOutbox As Workbook
    Dim wbConfig As Workbook
    Dim configOpenedTransient As Boolean
    Dim inventoryOpenedTransient As Boolean
    Dim outboxOpenedTransient As Boolean

    On Error GoTo FailCreate

    rootPath = NormalizeFolderPathTesterSetup(rootPath, False)
    If rootPath = "" Then
        report = "Warehouse hub path could not be resolved."
        Exit Function
    End If
    If IsUncPathTesterSetup(rootPath) And Not FolderExistsTesterSetup(rootPath) Then
        report = "Warehouse hub path is not accessible: " & rootPath
        Exit Function
    End If

    EnsureFolderRecursiveTesterSetup rootPath
    EnsureFolderRecursiveTesterSetup rootPath & "\inbox"
    EnsureFolderRecursiveTesterSetup rootPath & "\outbox"
    EnsureFolderRecursiveTesterSetup rootPath & "\snapshots"
    EnsureFolderRecursiveTesterSetup rootPath & "\config"
    EnsureFolderRecursiveTesterSetup rootPath & "\auth"

    configPath = rootPath & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    authPath = rootPath & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"
    inventoryPath = rootPath & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb"
    outboxPath = rootPath & "\" & spec.WarehouseId & ".Outbox.Events.xlsb"
    snapshotPath = rootPath & "\" & spec.WarehouseId & ".invSys.Snapshot.Inventory.xlsb"

    If Not modConfig.EnsureStationConfigEntry(spec.WarehouseId, spec.StationId, spec.UserId, rootPath & "\inbox\", "ADMIN", configPath, rootPath, report) Then GoTo FailSoft
    configOpenedTransient = (FindOpenWorkbookByPathTesterSetup(configPath) Is Nothing)
    Set wbConfig = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime(spec.WarehouseId, spec.StationId, rootPath, report)
    If wbConfig Is Nothing Then GoTo FailSoft

    If Not modAuth.EnsureStationRoleAuth(spec.WarehouseId, spec.StationId, spec.UserId, spec.UserId, "ADMIN", authPath, "svc_processor", capabilityOut, report) Then GoTo FailSoft

    inventoryOpenedTransient = (FindOpenWorkbookByPathTesterSetup(inventoryPath) Is Nothing)
    Set wbInventory = ResolveInventoryWorkbookBridge(spec.WarehouseId)
    If wbInventory Is Nothing Then
        report = "Inventory workbook not resolved."
        GoTo FailSoft
    End If
    If Not EnsureInventorySchemaBridge(wbInventory, report) Then GoTo FailSoft
    If Not wbInventory.ReadOnly Then wbInventory.Save

    If Not GenerateWarehouseSnapshot(spec.WarehouseId, wbInventory, snapshotPath, Nothing, report) Then GoTo FailSoft

    outboxOpenedTransient = (FindOpenWorkbookByPathTesterSetup(outboxPath) Is Nothing)
    Set wbOutbox = ResolveOutboxWorkbook(spec.WarehouseId, Nothing, True)
    If wbOutbox Is Nothing Then
        report = "Outbox workbook not resolved."
        GoTo FailSoft
    End If
    If Not EnsureOutboxSchema(wbOutbox, report) Then GoTo FailSoft
    If Not wbOutbox.ReadOnly Then wbOutbox.Save

    CreateTesterRuntimeArtifactsTesterSetup = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    If Len(report) = 0 Then report = "CreateTesterRuntimeArtifacts failed."
    GoTo CleanExit

FailCreate:
    report = "CreateTesterRuntimeArtifacts failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookIfTransientTesterSetup wbOutbox, outboxOpenedTransient
    CloseWorkbookIfTransientTesterSetup wbInventory, inventoryOpenedTransient
    CloseWorkbookIfTransientTesterSetup wbConfig, configOpenedTransient
End Function

Private Function EnsureTesterSharePointPackage(ByRef spec As TesterSetupSpec, _
                                               ByRef report As String) As Boolean
    Dim sharePointRoot As String
    Dim addinsRoot As String
    Dim sourceAddinsRoot As String
    Dim requiredAddins As Variant
    Dim addinName As Variant
    Dim sourcePath As String
    Dim targetPath As String
    Dim copiedCount As Long

    On Error GoTo FailPackage

    sharePointRoot = NormalizeFolderPathTesterSetup(spec.PathSharePointRoot, False)
    If sharePointRoot = "" Then
        report = "SKIPPED|PathSharePointRoot not configured."
        EnsureTesterSharePointPackage = True
        Exit Function
    End If

    EnsureFolderRecursiveTesterSetup sharePointRoot
    EnsureFolderRecursiveTesterSetup sharePointRoot & "\Addins"
    EnsureFolderRecursiveTesterSetup sharePointRoot & "\Events"
    EnsureFolderRecursiveTesterSetup sharePointRoot & "\Snapshots"
    EnsureFolderRecursiveTesterSetup sharePointRoot & "\TesterPackage"
    EnsureFolderRecursiveTesterSetup sharePointRoot & "\TesterPackage\" & spec.WarehouseId

    sourceAddinsRoot = ResolveCurrentAddinsRootTesterSetup()
    If sourceAddinsRoot = "" Then
        report = "Current add-ins folder could not be resolved for tester package publish."
        Exit Function
    End If

    addinsRoot = sharePointRoot & "\Addins"
    requiredAddins = RequiredTesterAddinNamesTesterSetup()
    For Each addinName In requiredAddins
        sourcePath = sourceAddinsRoot & "\" & CStr(addinName)
        targetPath = addinsRoot & "\" & CStr(addinName)
        If Not FileExistsTesterSetup(sourcePath) Then
            report = "Source add-in missing: " & sourcePath
            Exit Function
        End If
        If CopyFileIfNeededTesterSetup(sourcePath, targetPath) Then copiedCount = copiedCount + 1
    Next addinName

    report = "OK|Addins=" & addinsRoot & "|Copied=" & CStr(copiedCount)
    EnsureTesterSharePointPackage = True
    Exit Function

FailPackage:
    report = "EnsureTesterSharePointPackage failed: " & Err.Description
End Function

Private Function IsUncPathTesterSetup(ByVal pathIn As String) As Boolean
    pathIn = NormalizeFolderPathTesterSetup(pathIn, False)
    IsUncPathTesterSetup = (Left$(pathIn, 2) = "\\")
End Function

Private Sub PublishTesterSetupProgress(ByVal stepText As String)
    On Error Resume Next
    If Not mTesterSetupProgressSink Is Nothing Then
        CallByName mTesterSetupProgressSink, "UpdateSetupProgress", VbMethod, Trim$(stepText)
    End If
    On Error GoTo 0
End Sub

Private Function BuildDefaultTesterSpecTesterSetup() As TesterSetupSpec
    Dim spec As TesterSetupSpec
    spec.WarehouseId = TESTER_DEFAULT_WAREHOUSE_ID
    spec.StationId = TESTER_DEFAULT_STATION_ID
    spec.PathLocal = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(TESTER_DEFAULT_WAREHOUSE_ID, False)
    BuildDefaultTesterSpecTesterSetup = spec
End Function

Private Function CandidateLooksLikeSharePointRootTesterSetup(ByVal rootPath As String, ByVal warehouseId As String) As Boolean
    rootPath = NormalizeFolderPathTesterSetup(rootPath, False)
    If rootPath = "" Then Exit Function
    If Not IsUsableLocalPathTesterSetup(rootPath) Then Exit Function

    If FolderExistsTesterSetup(rootPath & "\Addins") Then
        If FileExistsTesterSetup(rootPath & "\Addins\invSys.Core.xlam") _
           And FileExistsTesterSetup(rootPath & "\Addins\invSys.Receiving.xlam") Then
            CandidateLooksLikeSharePointRootTesterSetup = True
            Exit Function
        End If
    End If

    If warehouseId <> "" Then
        If FileExistsTesterSetup(rootPath & "\TesterPackage\" & warehouseId & "\" & warehouseId & ".TesterBundle.zip") Then
            CandidateLooksLikeSharePointRootTesterSetup = True
        End If
    End If
End Function

Private Sub AddUniquePathTesterSetup(ByVal items As Collection, ByVal pathIn As String)
    Dim candidate As Variant
    pathIn = NormalizeFolderPathTesterSetup(pathIn, False)
    If pathIn = "" Then Exit Sub
    For Each candidate In items
        If StrComp(CStr(candidate), pathIn, vbTextCompare) = 0 Then Exit Sub
    Next candidate
    items.Add pathIn
End Sub

Private Function EnsureTemporaryAdminCapabilityTesterSetup(ByRef spec As TesterSetupSpec, _
                                                           ByRef tempGranted As Boolean, _
                                                           ByRef report As String) As Boolean
    Dim authPath As String
    Dim wbAuth As Workbook
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim userRow As Long
    Dim capRow As Long
    Dim openedTransient As Boolean

    On Error GoTo FailEnsure

    authPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False) & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"
    Set wbAuth = OpenWorkbookForWriteTesterSetup(authPath, openedTransient, report)
    If wbAuth Is Nothing Then GoTo FailSoft
    If Not modAuth.EnsureAuthSchema(wbAuth, spec.WarehouseId, "svc_processor", report) Then GoTo FailSoft

    Set loUsers = FindTableByNameTesterSetup(wbAuth, "tblUsers")
    Set loCaps = FindTableByNameTesterSetup(wbAuth, "tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then
        report = "Auth tables were not available for temporary admin grant."
        GoTo FailSoft
    End If

    userRow = EnsureUserRowTesterSetup(loUsers, spec.UserId)
    SetTableCellTesterSetup loUsers, userRow, "UserId", spec.UserId
    SetTableCellTesterSetup loUsers, userRow, "DisplayName", spec.UserId
    SetTableCellTesterSetup loUsers, userRow, "Status", "Active"
    If SafeTrimTesterSetup(GetTableCellTesterSetup(loUsers, userRow, "PinHash")) = "" Then
        SetTableCellTesterSetup loUsers, userRow, "PinHash", spec.PinHash
    End If

    capRow = EnsureCapabilityRowTesterSetup(loCaps, spec.UserId, "ADMIN_MAINT", spec.WarehouseId, spec.StationId)
    SetTableCellTesterSetup loCaps, capRow, "Status", "ACTIVE"
    wbAuth.Save

    tempGranted = True
    EnsureTemporaryAdminCapabilityTesterSetup = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    EnsureTemporaryAdminCapabilityTesterSetup = False
    If Len(report) = 0 Then report = "EnsureTemporaryAdminCapabilityTesterSetup failed."
    GoTo CleanExit

FailEnsure:
    report = "EnsureTemporaryAdminCapabilityTesterSetup failed: " & Err.Description
    Resume FailSoft

CleanExit:
    CloseWorkbookIfTransientTesterSetup wbAuth, openedTransient
End Function

Private Sub RevokeAdminCapabilityTesterSetup(ByRef spec As TesterSetupSpec)
    Dim authPath As String
    Dim wbAuth As Workbook
    Dim loCaps As ListObject
    Dim openedTransient As Boolean
    Dim report As String

    On Error Resume Next
    authPath = NormalizeFolderPathTesterSetup(spec.PathLocal, False) & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"
    Set wbAuth = OpenWorkbookForWriteTesterSetup(authPath, openedTransient, report)
    If wbAuth Is Nothing Then GoTo CleanExit
    Set loCaps = FindTableByNameTesterSetup(wbAuth, "tblCapabilities")
    If loCaps Is Nothing Then GoTo CleanExit
    DeactivateCapabilityTesterSetup loCaps, spec.UserId, "ADMIN_MAINT", spec.WarehouseId, spec.StationId
    wbAuth.Save
CleanExit:
    CloseWorkbookIfTransientTesterSetup wbAuth, openedTransient
    On Error GoTo 0
End Sub

Private Function TesterSeedAlreadyPresentTesterSetup(ByVal inventoryWb As Workbook) As Boolean
    Dim loSku As ListObject
    Dim rowIndex As Long
    Dim qtyVal As Variant
    Dim loInv As ListObject

    If inventoryWb Is Nothing Then Exit Function

    Set loSku = FindTableByNameTesterSetup(inventoryWb, "tblSkuBalance")
    If Not loSku Is Nothing Then
        rowIndex = FindRowByValueTesterSetup(loSku, "SKU", TESTER_SEED_SKU)
        If rowIndex > 0 Then
            qtyVal = GetTableCellTesterSetup(loSku, rowIndex, "QtyOnHand")
            If IsNumeric(qtyVal) Then
                TesterSeedAlreadyPresentTesterSetup = (CDbl(qtyVal) > 0)
                If TesterSeedAlreadyPresentTesterSetup Then Exit Function
            Else
                TesterSeedAlreadyPresentTesterSetup = True
                Exit Function
            End If
        End If
    End If

    Set loInv = FindTableByNameTesterSetup(inventoryWb, "invSys")
    If Not loInv Is Nothing Then
        rowIndex = FindRowByValueTesterSetup(loInv, "ITEM_CODE", TESTER_SEED_SKU)
        If rowIndex > 0 Then
            qtyVal = GetTableCellTesterSetup(loInv, rowIndex, "TOTAL INV")
            If IsNumeric(qtyVal) Then TesterSeedAlreadyPresentTesterSetup = (CDbl(qtyVal) > 0)
        End If
    End If
End Function

Private Function ReceivingWorkbookHasCanonicalSurfacesTesterSetup(ByVal wb As Workbook) As Boolean
    ReceivingWorkbookHasCanonicalSurfacesTesterSetup = _
        WorksheetHasTableTesterSetup(wb, "ReceivedTally", "ReceivedTally") And _
        WorksheetHasTableTesterSetup(wb, "InventoryManagement", "invSys") And _
        WorksheetHasTableTesterSetup(wb, "ReceivedLog", "ReceivedLog")
End Function

Private Function ReceivingWorkbookHasAliasSurfacesTesterSetup(ByVal wb As Workbook) As Boolean
    ReceivingWorkbookHasAliasSurfacesTesterSetup = _
        WorksheetExistsTesterSetup(wb, "tblReceiving") And _
        WorksheetExistsTesterSetup(wb, "tblReadModel") And _
        WorksheetExistsTesterSetup(wb, "tblStatus")
End Function

Private Function WorksheetHasTableTesterSetup(ByVal wb As Workbook, ByVal sheetName As String, ByVal tableName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    On Error Resume Next
    WorksheetHasTableTesterSetup = Not (ws.ListObjects(tableName) Is Nothing)
    On Error GoTo 0
End Function

Private Function WorksheetExistsTesterSetup(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetExistsTesterSetup = Not ws Is Nothing
    On Error GoTo 0
End Function

Private Function OpenWorkbookForWriteTesterSetup(ByVal workbookPath As String, _
                                                 ByRef openedTransient As Boolean, _
                                                 ByRef report As String) As Workbook
    Dim wb As Workbook

    workbookPath = Trim$(workbookPath)
    If workbookPath = "" Then
        report = "Workbook path is required."
        Exit Function
    End If

    Set wb = FindOpenWorkbookByPathTesterSetup(workbookPath)
    If wb Is Nothing Then
        If Not FileExistsTesterSetup(workbookPath) Then
            report = "Workbook not found: " & workbookPath
            Exit Function
        End If
        Set wb = Application.Workbooks.Open(Filename:=workbookPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
        openedTransient = Not wb Is Nothing
    End If
    If wb Is Nothing Then
        report = "Workbook could not be opened: " & workbookPath
        Exit Function
    End If
    If wb.ReadOnly Then
        report = "Workbook is read-only: " & workbookPath
        If openedTransient Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            On Error GoTo 0
        End If
        Exit Function
    End If

    Set OpenWorkbookForWriteTesterSetup = wb
End Function

Private Sub CloseWorkbookIfTransientTesterSetup(ByVal wb As Workbook, ByVal openedTransient As Boolean)
    If Not openedTransient Then Exit Sub
    If wb Is Nothing Then Exit Sub

    On Error Resume Next
    If Not wb.ReadOnly Then
        If wb.Saved = False Then wb.Save
    End If
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function FindOpenWorkbookByPathTesterSetup(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook

    workbookPath = Trim$(workbookPath)
    If workbookPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(Trim$(wb.FullName), workbookPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByPathTesterSetup = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FindTableByNameTesterSetup(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    On Error Resume Next
    For Each ws In wb.Worksheets
        Set FindTableByNameTesterSetup = ws.ListObjects(tableName)
        If Not FindTableByNameTesterSetup Is Nothing Then Exit Function
    Next ws
    On Error GoTo 0
End Function

Private Function EnsureUserRowTesterSetup(ByVal lo As ListObject, ByVal userId As String) As Long
    EnsureUserRowTesterSetup = FindRowByValueTesterSetup(lo, "UserId", userId)
    If EnsureUserRowTesterSetup > 0 Then Exit Function

    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        EnsureUserRowTesterSetup = 1
    ElseIf lo.ListRows.Count = 1 And SafeTrimTesterSetup(GetTableCellTesterSetup(lo, 1, "UserId")) = "" Then
        EnsureUserRowTesterSetup = 1
    Else
        EnsureUserRowTesterSetup = lo.ListRows.Add.Index
    End If
End Function

Private Function EnsureCapabilityRowTesterSetup(ByVal lo As ListObject, _
                                                ByVal userId As String, _
                                                ByVal capability As String, _
                                                ByVal warehouseId As String, _
                                                ByVal stationId As String) As Long
    EnsureCapabilityRowTesterSetup = FindCapabilityRowTesterSetup(lo, userId, capability, warehouseId, stationId)
    If EnsureCapabilityRowTesterSetup > 0 Then Exit Function

    If lo.DataBodyRange Is Nothing Then
        lo.ListRows.Add
        EnsureCapabilityRowTesterSetup = 1
    ElseIf lo.ListRows.Count = 1 And SafeTrimTesterSetup(GetTableCellTesterSetup(lo, 1, "UserId")) = "" _
        And SafeTrimTesterSetup(GetTableCellTesterSetup(lo, 1, "Capability")) = "" Then
        EnsureCapabilityRowTesterSetup = 1
    Else
        EnsureCapabilityRowTesterSetup = lo.ListRows.Add.Index
    End If

    SetTableCellTesterSetup lo, EnsureCapabilityRowTesterSetup, "UserId", userId
    SetTableCellTesterSetup lo, EnsureCapabilityRowTesterSetup, "Capability", capability
    SetTableCellTesterSetup lo, EnsureCapabilityRowTesterSetup, "WarehouseId", warehouseId
    SetTableCellTesterSetup lo, EnsureCapabilityRowTesterSetup, "StationId", stationId
End Function

Private Sub EnsureCapabilityActiveTesterSetup(ByVal lo As ListObject, _
                                              ByVal userId As String, _
                                              ByVal capability As String, _
                                              ByVal warehouseId As String, _
                                              ByVal stationId As String)
    Dim rowIndex As Long

    rowIndex = EnsureCapabilityRowTesterSetup(lo, userId, capability, warehouseId, stationId)
    If rowIndex = 0 Then Exit Sub
    SetTableCellTesterSetup lo, rowIndex, "Status", "ACTIVE"
End Sub

Private Sub DeactivateCapabilityTesterSetup(ByVal lo As ListObject, _
                                            ByVal userId As String, _
                                            ByVal capability As String, _
                                            ByVal warehouseId As String, _
                                            ByVal stationId As String)
    Dim rowIndex As Long

    rowIndex = FindCapabilityRowTesterSetup(lo, userId, capability, warehouseId, stationId)
    If rowIndex = 0 Then Exit Sub
    SetTableCellTesterSetup lo, rowIndex, "Status", "INACTIVE"
End Sub

Private Function CapabilityIsActiveTesterSetup(ByVal lo As ListObject, _
                                               ByVal userId As String, _
                                               ByVal capability As String, _
                                               ByVal warehouseId As String, _
                                               ByVal stationId As String) As Boolean
    Dim rowIndex As Long

    rowIndex = FindCapabilityRowTesterSetup(lo, userId, capability, warehouseId, stationId)
    If rowIndex = 0 Then Exit Function
    CapabilityIsActiveTesterSetup = (StrComp(SafeTrimTesterSetup(GetTableCellTesterSetup(lo, rowIndex, "Status")), "ACTIVE", vbTextCompare) = 0)
End Function

Private Function FindCapabilityRowTesterSetup(ByVal lo As ListObject, _
                                              ByVal userId As String, _
                                              ByVal capability As String, _
                                              ByVal warehouseId As String, _
                                              ByVal stationId As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrimTesterSetup(GetTableCellTesterSetup(lo, i, "UserId")), userId, vbTextCompare) = 0 _
           And StrComp(UCase$(SafeTrimTesterSetup(GetTableCellTesterSetup(lo, i, "Capability"))), UCase$(capability), vbTextCompare) = 0 _
           And StrComp(SafeTrimTesterSetup(GetTableCellTesterSetup(lo, i, "WarehouseId")), warehouseId, vbTextCompare) = 0 _
           And StrComp(SafeTrimTesterSetup(GetTableCellTesterSetup(lo, i, "StationId")), stationId, vbTextCompare) = 0 Then
            FindCapabilityRowTesterSetup = i
            Exit Function
        End If
    Next i
End Function

Private Function FindRowByValueTesterSetup(ByVal lo As ListObject, ByVal columnName As String, ByVal expectedValue As String) As Long
    Dim i As Long
    Dim columnIndex As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    columnIndex = GetColumnIndexTesterSetup(lo, columnName)
    If columnIndex = 0 Then Exit Function

    For i = 1 To lo.ListRows.Count
        If StrComp(SafeTrimTesterSetup(lo.DataBodyRange.Cells(i, columnIndex).Value), expectedValue, vbTextCompare) = 0 Then
            FindRowByValueTesterSetup = i
            Exit Function
        End If
    Next i
End Function

Private Function GetColumnIndexTesterSetup(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexTesterSetup = i
            Exit Function
        End If
    Next i
End Function

Private Function GetTableCellTesterSetup(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim columnIndex As Long

    columnIndex = GetColumnIndexTesterSetup(lo, columnName)
    If columnIndex = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Function
    GetTableCellTesterSetup = lo.DataBodyRange.Cells(rowIndex, columnIndex).Value
End Function

Private Sub SetTableCellTesterSetup(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueOut As Variant)
    Dim columnIndex As Long

    columnIndex = GetColumnIndexTesterSetup(lo, columnName)
    If columnIndex = 0 Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If rowIndex < 1 Or rowIndex > lo.ListRows.Count Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, columnIndex).Value = valueOut
End Sub

Private Function NormalizeFolderPathTesterSetup(ByVal folderPath As String, Optional ByVal withTrailingSlash As Boolean = False) As String
    NormalizeFolderPathTesterSetup = modConfig.NormalizeFolderPathForRuntime(folderPath, withTrailingSlash)
End Function

Private Function FolderExistsTesterSetup(ByVal folderPath As String) As Boolean
    Dim fso As Object

    folderPath = NormalizeFolderPathTesterSetup(folderPath, False)
    If folderPath = "" Then Exit Function
    If Not IsUsableLocalPathTesterSetup(folderPath) Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsTesterSetup = fso.FolderExists(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        FolderExistsTesterSetup = (Len(Dir$(folderPath, vbDirectory)) > 0)
    End If
    On Error GoTo 0
End Function

Private Function FileExistsTesterSetup(ByVal filePath As String) As Boolean
    Dim fso As Object

    filePath = Trim$(Replace$(filePath, "/", "\"))
    If filePath = "" Then Exit Function
    If Not IsUsableLocalPathTesterSetup(filePath) Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsTesterSetup = fso.FileExists(filePath)
    If Err.Number <> 0 Then
        Err.Clear
        FileExistsTesterSetup = (Len(Dir$(filePath, vbNormal)) > 0)
    End If
    On Error GoTo 0
End Function

Private Function RequiredTesterAddinNamesTesterSetup() As Variant
    RequiredTesterAddinNamesTesterSetup = Array( _
        "invSys.Core.xlam", _
        "invSys.Receiving.xlam", _
        "invSys.Shipping.xlam", _
        "invSys.Production.xlam", _
        "invSys.Admin.xlam")
End Function

Private Function ResolveCurrentAddinsRootTesterSetup() As String
    Dim candidate As String

    candidate = NormalizeFolderPathTesterSetup(ThisWorkbook.Path, False)
    If CandidateHasRequiredTesterAddinsTesterSetup(candidate) Then
        ResolveCurrentAddinsRootTesterSetup = candidate
        Exit Function
    End If

    candidate = NormalizeFolderPathTesterSetup(Environ$("USERPROFILE") & "\source\repos\invSys_fork\deploy\current", False)
    If CandidateHasRequiredTesterAddinsTesterSetup(candidate) Then
        ResolveCurrentAddinsRootTesterSetup = candidate
        Exit Function
    End If

    candidate = NormalizeFolderPathTesterSetup(modConfig.GetString("PathSharePointRoot", ""), False)
    If CandidateHasRequiredTesterAddinsTesterSetup(candidate & "\Addins") Then
        ResolveCurrentAddinsRootTesterSetup = NormalizeFolderPathTesterSetup(candidate & "\Addins", False)
    End If
End Function

Private Function CandidateHasRequiredTesterAddinsTesterSetup(ByVal folderPath As String) As Boolean
    Dim addinName As Variant
    Dim requiredAddins As Variant

    folderPath = NormalizeFolderPathTesterSetup(folderPath, False)
    If folderPath = "" Then Exit Function
    If Not FolderExistsTesterSetup(folderPath) Then Exit Function

    requiredAddins = RequiredTesterAddinNamesTesterSetup()
    For Each addinName In requiredAddins
        If Not FileExistsTesterSetup(folderPath & "\" & CStr(addinName)) Then Exit Function
    Next addinName

    CandidateHasRequiredTesterAddinsTesterSetup = True
End Function

Private Function CopyFileIfNeededTesterSetup(ByVal sourcePath As String, ByVal targetPath As String) As Boolean
    sourcePath = Trim$(Replace$(sourcePath, "/", "\"))
    targetPath = Trim$(Replace$(targetPath, "/", "\"))
    If sourcePath = "" Or targetPath = "" Then Exit Function
    If Not FileExistsTesterSetup(sourcePath) Then Exit Function
    If FileExistsTesterSetup(targetPath) Then
        If SafeFileLenTesterSetup(sourcePath) = SafeFileLenTesterSetup(targetPath) Then Exit Function
        Kill targetPath
    End If
    EnsureFolderRecursiveTesterSetup GetParentFolderTesterSetup(targetPath)
    FileCopy sourcePath, targetPath
    CopyFileIfNeededTesterSetup = True
End Function

Private Function SafeFileLenTesterSetup(ByVal filePath As String) As Long
    On Error Resume Next
    SafeFileLenTesterSetup = FileLen(filePath)
    On Error GoTo 0
End Function

Private Sub EnsureFolderRecursiveTesterSetup(ByVal folderPath As String)
    Dim parentPath As String

    folderPath = NormalizeFolderPathTesterSetup(folderPath, False)
    If folderPath = "" Then Exit Sub
    If FolderExistsTesterSetup(folderPath) Then Exit Sub

    parentPath = GetParentFolderTesterSetup(folderPath)
    If parentPath <> "" And Not FolderExistsTesterSetup(parentPath) Then EnsureFolderRecursiveTesterSetup parentPath
    MkDir folderPath
End Sub

Private Function GetParentFolderTesterSetup(ByVal pathIn As String) As String
    Dim slashPos As Long

    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    If pathIn = "" Then Exit Function
    slashPos = InStrRev(pathIn, "\")
    If slashPos = 3 And Mid$(pathIn, 2, 2) = ":\" Then
        GetParentFolderTesterSetup = Left$(pathIn, 3)
    ElseIf slashPos > 1 Then
        GetParentFolderTesterSetup = Left$(pathIn, slashPos - 1)
    End If
End Function

Private Function IsUsableLocalPathTesterSetup(ByVal pathIn As String) As Boolean
    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    If pathIn = "" Then Exit Function
    If InStr(1, pathIn, "://", vbTextCompare) > 0 Then Exit Function
    If Left$(pathIn, 8) = "https:\\" Then Exit Function
    If Left$(pathIn, 7) = "http:\\" Then Exit Function
    If InStr(pathIn, "*") > 0 Or InStr(pathIn, "?") > 0 Then Exit Function
    IsUsableLocalPathTesterSetup = True
End Function

Private Function SafeTrimTesterSetup(ByVal valueIn As Variant) As String
    On Error Resume Next
    SafeTrimTesterSetup = Trim$(CStr(valueIn))
    On Error GoTo 0
End Function

Private Sub RestoreCoreRootOverrideTesterSetup(ByVal priorRootOverride As String)
    If Trim$(priorRootOverride) = "" Then
        modRuntimeWorkbooks.ClearCoreDataRootOverride
    Else
        modRuntimeWorkbooks.SetCoreDataRootOverride priorRootOverride
    End If
End Sub

Private Sub ResetTesterSetupState()
    mLastTesterSetupReport = vbNullString
    mLastTesterOperatorWorkbookPath = vbNullString
    mLastTesterSharePointRoot = vbNullString
End Sub
