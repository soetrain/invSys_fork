Attribute VB_Name = "modWarehouseBootstrap"
Option Explicit

Private Const BOOTSTRAP_SHAREPOINT_CONFIG_FOLDER As String = "Config"
Private Const BOOTSTRAP_SHAREPOINT_CONFIG_JSON_SUFFIX As String = ".config.json"
Private Const BOOTSTRAP_SHAREPOINT_CONFIG_WORKBOOK_SUFFIX As String = ".invSys.Config.xlsb"
Private Const BOOTSTRAP_TEMPLATE_FOLDER_NAME As String = "templates"
Private Const BOOTSTRAP_INVENTORY_TEMPLATE_FILE As String = "invSys.Data.Inventory.template.xlsb"
Private Const BOOTSTRAP_RECEIVING_OPERATOR_SUFFIX As String = ".Receiving.Operator.xlsm"
Private Const BOOTSTRAP_LOCAL_OPERATOR_ROOT As String = "invSys\OperatorWorkbooks"
Private Const BOOTSTRAP_SEED_SOURCE_ID As String = "CREATE_WAREHOUSE_DEMO_SEED"

Private mBootstrapTemplateRootOverride As String
Private mLastBootstrapReport As String
Private mLastBootstrapOperatorWorkbookPath As String

Public Type WarehouseSpec
    WarehouseId As String
    WarehouseName As String
    StationId As String
    AdminUser As String
    PathLocal As String
    PathSharePoint As String
End Type

Public Function ValidateWarehouseSpecValues(ByVal warehouseId As String, _
                                            ByVal warehouseName As String, _
                                            ByVal stationId As String, _
                                            ByVal adminUser As String, _
                                            ByVal pathLocal As String, _
                                            ByVal pathSharePoint As String, _
                                            Optional ByRef report As String = "") As Boolean
    Dim spec As WarehouseSpec

    spec.WarehouseId = warehouseId
    spec.WarehouseName = warehouseName
    spec.StationId = stationId
    spec.AdminUser = adminUser
    spec.PathLocal = pathLocal
    spec.PathSharePoint = pathSharePoint
    ValidateWarehouseSpecValues = ValidateWarehouseSpec(spec, report)
End Function

Public Function BootstrapWarehouseLocalValues(ByVal warehouseId As String, _
                                              ByVal warehouseName As String, _
                                              ByVal stationId As String, _
                                              ByVal adminUser As String, _
                                              ByVal pathLocal As String, _
                                              ByVal pathSharePoint As String) As Boolean
    Dim spec As WarehouseSpec

    spec.WarehouseId = warehouseId
    spec.WarehouseName = warehouseName
    spec.StationId = stationId
    spec.AdminUser = adminUser
    spec.PathLocal = pathLocal
    spec.PathSharePoint = pathSharePoint
    BootstrapWarehouseLocalValues = BootstrapWarehouseLocal(spec)
End Function

Public Function PublishInitialArtifactsValues(ByVal warehouseId As String, _
                                              ByVal warehouseName As String, _
                                              ByVal stationId As String, _
                                              ByVal adminUser As String, _
                                              ByVal pathLocal As String, _
                                              ByVal pathSharePoint As String) As Boolean
    Dim spec As WarehouseSpec

    spec.WarehouseId = warehouseId
    spec.WarehouseName = warehouseName
    spec.StationId = stationId
    spec.AdminUser = adminUser
    spec.PathLocal = pathLocal
    spec.PathSharePoint = pathSharePoint
    PublishInitialArtifactsValues = PublishInitialArtifacts(spec)
End Function

Public Function ValidateWarehouseSpec(ByRef spec As WarehouseSpec, _
                                      Optional ByRef report As String = "") As Boolean
    NormalizeWarehouseSpec spec

    If spec.WarehouseId = "" Then
        report = "WarehouseId is required."
        Exit Function
    End If

    If Not IsValidWarehouseIdBootstrap(spec.WarehouseId) Then
        report = "WarehouseId must contain only letters, digits, hyphens, and underscores."
        Exit Function
    End If

    If spec.PathLocal <> "" Then
        spec.PathLocal = modDeploymentPaths.NormalizeManagedFolderPath(spec.PathLocal, False)
        If spec.PathLocal = "" Then
            report = "Warehouse hub path could not be resolved."
            Exit Function
        End If
    End If

    If spec.PathSharePoint <> "" Then
        spec.PathSharePoint = modDeploymentPaths.NormalizeManagedFolderPath(spec.PathSharePoint, False)
        If spec.PathSharePoint = "" Then
            report = "SharePoint root could not be resolved."
            Exit Function
        End If
    End If

    report = "OK"
    ValidateWarehouseSpec = True
End Function

Public Function WarehouseIdExists(ByVal warehouseId As String) As Boolean
    warehouseId = Trim$(warehouseId)
    If warehouseId = "" Then Exit Function

    If LocalWarehouseIdExistsBootstrap(warehouseId) Then
        WarehouseIdExists = True
        Exit Function
    End If

    WarehouseIdExists = SharePointWarehouseIdExistsBootstrap(warehouseId)
End Function

Public Function WarehouseArtifactsExistAtPath(ByVal warehouseId As String, ByVal rootPath As String) As Boolean
    rootPath = NormalizeFolderPathBootstrap(rootPath)
    If Right$(rootPath, 1) = "\" Then rootPath = Left$(rootPath, Len(rootPath) - 1)
    WarehouseArtifactsExistAtPath = RuntimeArtifactsExistBootstrap(rootPath, warehouseId)
End Function

Public Function BootstrapWarehouseLocal(ByRef spec As WarehouseSpec) As Boolean
    Dim rootPath As String
    Dim priorRootOverride As String
    Dim report As String
    Dim configPath As String
    Dim authPath As String
    Dim inventoryPath As String
    Dim capabilityOut As String
    Dim wbCfg As Workbook
    Dim wbInventory As Workbook
    Dim wbOutbox As Workbook
    Dim createdRoot As Boolean
    Dim rootExisted As Boolean
    Dim inboxPath As String
    Dim operatorPath As String
    Dim operatorReport As String
    Dim seedReport As String

    On Error GoTo FailBootstrap

    mLastBootstrapReport = vbNullString
    mLastBootstrapOperatorWorkbookPath = vbNullString
    If Not ValidateWarehouseSpec(spec, report) Then GoTo FailSoft
    If Trim$(spec.AdminUser) = "" Then
        report = "AdminUser is required."
        GoTo FailSoft
    End If

    rootPath = ResolveBootstrapRootPath(spec)
    If rootPath = "" Then
        report = "Warehouse hub path could not be resolved."
        GoTo FailSoft
    End If
    spec.PathLocal = rootPath

    If WarehouseIdExists(spec.WarehouseId) And Not RuntimeArtifactsExistBootstrap(rootPath, spec.WarehouseId) Then
        report = "WarehouseId already exists in the configured warehouse catalog: " & spec.WarehouseId
        GoTo FailSoft
    End If

    If RuntimeArtifactsExistBootstrap(rootPath, spec.WarehouseId) Then
        report = "Warehouse runtime artifacts already exist at hub path: " & rootPath
        GoTo FailSoft
    End If

    rootExisted = FolderExistsBootstrap(rootPath)
    If Not rootExisted Then
        EnsureFolderRecursiveBootstrap rootPath
        createdRoot = True
    ElseIf Not FolderExistsBootstrap(rootPath) Then
        report = "Warehouse hub path is not accessible: " & rootPath
        GoTo FailSoft
    End If

    EnsureFolderRecursiveBootstrap rootPath & "\inbox"
    EnsureFolderRecursiveBootstrap rootPath & "\outbox"
    EnsureFolderRecursiveBootstrap rootPath & "\snapshots"
    EnsureFolderRecursiveBootstrap rootPath & "\config"

    configPath = rootPath & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    authPath = rootPath & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"
    inventoryPath = rootPath & "\" & spec.WarehouseId & ".invSys.Data.Inventory.xlsb"

    priorRootOverride = modRuntimeWorkbooks.GetCoreDataRootOverride()
    modRuntimeWorkbooks.SetCoreDataRootOverride rootPath

    If Not CopyInventoryTemplateBootstrap(spec.WarehouseId, inventoryPath, report) Then GoTo FailSoft

    If Not modConfig.EnsureStationConfigEntry(spec.WarehouseId, spec.StationId, spec.AdminUser, rootPath & "\inbox\", "ADMIN", configPath, rootPath, report) Then GoTo FailSoft
    Set wbCfg = modRuntimeWorkbooks.OpenOrCreateConfigWorkbookRuntime(spec.WarehouseId, spec.StationId, rootPath, report)
    If wbCfg Is Nothing Then GoTo FailSoft
    If Not StampBootstrapConfigWorkbook(wbCfg, spec, report) Then GoTo FailSoft
    If Not modConfig.EnsureStationConfigEntry(spec.WarehouseId, spec.StationId, spec.AdminUser, rootPath & "\inbox\", "RECEIVE", configPath, rootPath, report) Then GoTo FailSoft
    If Not modConfig.EnsureStationInbox(spec.WarehouseId, spec.StationId, "RECEIVE", configPath, inboxPath, report) Then GoTo FailSoft

    If Not modAuth.EnsureStationRoleAuth(spec.WarehouseId, spec.StationId, spec.AdminUser, spec.AdminUser, "ADMIN", authPath, "svc_processor", capabilityOut, report) Then GoTo FailSoft
    If StrComp(capabilityOut, "ADMIN_MAINT", vbTextCompare) <> 0 Then
        report = "Admin capability was not provisioned."
        GoTo FailSoft
    End If
    If Not modAuth.EnsureStationRoleAuth(spec.WarehouseId, spec.StationId, spec.AdminUser, spec.AdminUser, "RECEIVE", authPath, "svc_processor", capabilityOut, report) Then GoTo FailSoft
    If StrComp(capabilityOut, "RECEIVE_POST", vbTextCompare) <> 0 Then
        report = "Receiving capability was not provisioned."
        GoTo FailSoft
    End If

    Set wbInventory = ResolveInventoryWorkbookBridge(spec.WarehouseId)
    If wbInventory Is Nothing Then
        report = "Inventory workbook not resolved after template stamp."
        GoTo FailSoft
    End If
    If Not GenerateWarehouseSnapshot(spec.WarehouseId, wbInventory, rootPath & "\" & spec.WarehouseId & ".invSys.Snapshot.Inventory.xlsb", Nothing, report) Then GoTo FailSoft

    Set wbOutbox = ResolveOutboxWorkbook(spec.WarehouseId, Nothing, True)
    If wbOutbox Is Nothing Then
        report = "Outbox workbook not resolved."
        GoTo FailSoft
    End If
    If Not EnsureOutboxSchema(wbOutbox, report) Then GoTo FailSoft
    SaveWorkbookIfWritableBootstrap wbOutbox

    If Not modConfig.LoadConfig(spec.WarehouseId, spec.StationId) Then
        report = "Config load failed: " & modConfig.Validate()
        GoTo FailSoft
    End If
    If Not modAuth.LoadAuth(spec.WarehouseId) Then
        report = "Auth load failed: " & modAuth.ValidateAuth()
        GoTo FailSoft
    End If
    If Not modAuth.HasProvisionedCapabilityForSystem("ADMIN_MAINT", spec.AdminUser, spec.WarehouseId, spec.StationId) Then
        report = "Admin user was not granted ADMIN_MAINT."
        GoTo FailSoft
    End If
    If Not modAuth.HasProvisionedCapabilityForSystem("RECEIVE_POST", spec.AdminUser, spec.WarehouseId, spec.StationId) Then
        report = "Admin user was not granted RECEIVE_POST."
        GoTo FailSoft
    End If

    If Not SeedBootstrapDemoInventory(spec, seedReport) Then
        report = seedReport
        GoTo FailSoft
    End If

    CloseWorkbookIfOpenBootstrap wbOutbox
    Set wbOutbox = Nothing
    CloseWorkbookIfOpenBootstrap wbInventory
    Set wbInventory = Nothing
    CloseWorkbookIfOpenBootstrap wbCfg
    Set wbCfg = Nothing

    operatorPath = BuildReceivingOperatorPathBootstrap(spec)
    If Not PrepareReceivingOperatorRuntimeFilesBootstrap(spec, operatorPath, operatorReport) Then
        report = operatorReport
        GoTo FailSoft
    End If
    If Not CreateOrVerifyReceivingOperatorWorkbookBootstrap(spec, operatorPath, operatorReport) Then
        report = operatorReport
        GoTo FailSoft
    End If

    modRuntimeWorkbooks.RememberWarehouseScanRootRuntime rootPath
    BootstrapWarehouseLocal = True
    report = "OK|Hub=" & rootPath & "|Inbox=" & inboxPath & "|Seed=" & seedReport & "|Operator=" & operatorPath
    GoTo CleanExit

FailSoft:
    BootstrapWarehouseLocal = False
    If Len(report) = 0 Then report = "BootstrapWarehouseLocal failed."
    CloseWorkbookIfOpenBootstrap wbOutbox
    Set wbOutbox = Nothing
    CloseWorkbookIfOpenBootstrap wbInventory
    Set wbInventory = Nothing
    CloseWorkbookIfOpenBootstrap wbCfg
    Set wbCfg = Nothing
    If createdRoot Then DeleteFolderRecursiveBootstrap rootPath
    LogDiagnosticSafeBootstrap "WAREHOUSE-BOOTSTRAP", "Warehouse bootstrap failed|WarehouseId=" & spec.WarehouseId & "|Root=" & rootPath & "|Reason=" & report
    GoTo CleanExit

FailBootstrap:
    report = "BootstrapWarehouseLocal failed: " & Err.Description
    Resume FailSoft

CleanExit:
    mLastBootstrapReport = report
    CloseWorkbookIfOpenBootstrap wbOutbox
    CloseWorkbookIfOpenBootstrap wbInventory
    CloseWorkbookIfOpenBootstrap wbCfg
    RestoreCoreRootOverrideBootstrap priorRootOverride
End Function

Public Function PublishInitialArtifacts(ByRef spec As WarehouseSpec) As Boolean
    Dim report As String
    Dim rootPath As String
    Dim sharePointRoot As String
    Dim localConfigPath As String
    Dim localDiscoveryPath As String
    Dim publishedConfigPath As String
    Dim publishedDiscoveryPath As String
    Dim configStatus As String
    Dim discoveryStatus As String

    On Error GoTo FailPublish

    mLastBootstrapReport = vbNullString
    If Not ValidateWarehouseSpec(spec, report) Then GoTo FailSoft

    rootPath = ResolveBootstrapRootPath(spec)
    If rootPath = "" Then
        report = "Warehouse hub path could not be resolved."
        GoTo FailSoft
    End If
    spec.PathLocal = rootPath

    sharePointRoot = NormalizeFolderPathBootstrap(spec.PathSharePoint)
    If sharePointRoot = "" Then sharePointRoot = NormalizeFolderPathBootstrap(modConfig.GetString("PathSharePointRoot", ""))
    If sharePointRoot = "" Then
        report = "SharePoint root not configured."
        GoTo FailSoft
    End If
    spec.PathSharePoint = Left$(sharePointRoot, Len(sharePointRoot) - 1)

    localConfigPath = rootPath & "\" & spec.WarehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_WORKBOOK_SUFFIX
    If Not FileExistsBootstrap(localConfigPath) Then
        report = "Local config artifact not found: " & localConfigPath
        GoTo FailSoft
    End If

    localDiscoveryPath = rootPath & "\config\" & spec.WarehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_JSON_SUFFIX
    If Not WriteBootstrapDiscoveryFile(spec, localConfigPath, localDiscoveryPath, report) Then GoTo FailSoft

    publishedConfigPath = sharePointRoot & spec.WarehouseId & "\" & GetFileNameBootstrap(localConfigPath)
    publishedDiscoveryPath = sharePointRoot & spec.WarehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_JSON_SUFFIX

    If Not modWarehouseSync.PublishFileToTargetPath(localConfigPath, publishedConfigPath, configStatus) Then
        report = configStatus
        GoTo FailSoft
    End If
    If Not modWarehouseSync.PublishFileToTargetPath(localDiscoveryPath, publishedDiscoveryPath, discoveryStatus) Then
        report = "Config=" & configStatus & "|Discovery=" & discoveryStatus
        GoTo FailSoft
    End If

    report = "OK|Config=" & configStatus & "|Discovery=" & discoveryStatus
    PublishInitialArtifacts = True
    GoTo CleanExit

FailSoft:
    PublishInitialArtifacts = False
    If Len(report) = 0 Then report = "PublishInitialArtifacts failed."
    LogDiagnosticSafeBootstrap "WAREHOUSE-BOOTSTRAP", _
        "Initial publish failed|WarehouseId=" & spec.WarehouseId & "|Root=" & sharePointRoot & "|Reason=" & report
    GoTo CleanExit

FailPublish:
    report = "PublishInitialArtifacts failed: " & Err.Description
    Resume FailSoft

CleanExit:
    mLastBootstrapReport = report
End Function

Public Function GetLastWarehouseBootstrapReport() As String
    GetLastWarehouseBootstrapReport = mLastBootstrapReport
End Function

Public Function GetLastWarehouseOperatorWorkbookPath() As String
    GetLastWarehouseOperatorWorkbookPath = mLastBootstrapOperatorWorkbookPath
End Function

Public Sub SetWarehouseBootstrapTemplateRootOverride(ByVal rootPath As String)
    mBootstrapTemplateRootOverride = Trim$(rootPath)
End Sub

Public Sub ClearWarehouseBootstrapTemplateRootOverride()
    mBootstrapTemplateRootOverride = vbNullString
End Sub

Private Sub NormalizeWarehouseSpec(ByRef spec As WarehouseSpec)
    spec.WarehouseId = Trim$(spec.WarehouseId)
    spec.WarehouseName = Trim$(spec.WarehouseName)
    spec.StationId = Trim$(spec.StationId)
    spec.AdminUser = Trim$(spec.AdminUser)
    spec.PathLocal = Trim$(spec.PathLocal)
    spec.PathSharePoint = Trim$(spec.PathSharePoint)
End Sub

Private Function IsValidWarehouseIdBootstrap(ByVal warehouseId As String) As Boolean
    warehouseId = Trim$(warehouseId)
    If warehouseId = "" Then Exit Function

    IsValidWarehouseIdBootstrap = Not (warehouseId Like "*[!A-Za-z0-9_-]*")
End Function

Private Function LocalWarehouseIdExistsBootstrap(ByVal warehouseId As String) As Boolean
    LocalWarehouseIdExistsBootstrap = FolderExistsBootstrap(modDeploymentPaths.DefaultWarehouseRuntimeRootPath(warehouseId, False))
End Function

Private Function SharePointWarehouseIdExistsBootstrap(ByVal warehouseId As String) As Boolean
    Dim sharePointRoot As String
    Dim normalizedRoot As String
    Dim configFolder As String

    sharePointRoot = Trim$(modConfig.GetString("PathSharePointRoot", ""))
    If sharePointRoot = "" Then Exit Function

    On Error GoTo SkipUnavailable
    normalizedRoot = NormalizeFolderPathBootstrap(sharePointRoot)
    If normalizedRoot = "" Or Not FolderExistsBootstrap(normalizedRoot) Then
        Err.Raise vbObjectError + 7385, "modWarehouseBootstrap.SharePointWarehouseIdExistsBootstrap", "SharePoint root is not reachable."
    End If
    configFolder = normalizedRoot & BOOTSTRAP_SHAREPOINT_CONFIG_FOLDER & "\"
    SharePointWarehouseIdExistsBootstrap = _
        FileExistsBootstrap(normalizedRoot & warehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_JSON_SUFFIX) Or _
        FileExistsBootstrap(configFolder & warehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_JSON_SUFFIX) Or _
        FileExistsBootstrap(configFolder & warehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_WORKBOOK_SUFFIX) Or _
        FileExistsBootstrap(normalizedRoot & warehouseId & "\" & warehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_WORKBOOK_SUFFIX)
    Exit Function

SkipUnavailable:
    LogDiagnosticSafeBootstrap "WAREHOUSE-BOOTSTRAP", _
        "SharePoint collision check skipped|WarehouseId=" & warehouseId & _
        "|Root=" & sharePointRoot & "|Reason=" & Err.Description
    SharePointWarehouseIdExistsBootstrap = False
End Function

Private Function RuntimeArtifactsExistBootstrap(ByVal rootPath As String, ByVal warehouseId As String) As Boolean
    rootPath = NormalizeFolderPathBootstrap(rootPath)
    If Right$(rootPath, 1) = "\" Then rootPath = Left$(rootPath, Len(rootPath) - 1)
    warehouseId = Trim$(warehouseId)
    If rootPath = "" Or warehouseId = "" Then Exit Function

    RuntimeArtifactsExistBootstrap = _
        FileExistsBootstrap(rootPath & "\" & warehouseId & ".invSys.Config.xlsb") Or _
        FileExistsBootstrap(rootPath & "\" & warehouseId & ".invSys.Auth.xlsb") Or _
        FileExistsBootstrap(rootPath & "\" & warehouseId & ".invSys.Data.Inventory.xlsb") Or _
        FileExistsBootstrap(rootPath & "\" & warehouseId & ".invSys.Snapshot.Inventory.xlsb") Or _
        FileExistsBootstrap(rootPath & "\" & warehouseId & ".Outbox.Events.xlsb") Or _
        FileExistsBootstrap(rootPath & "\" & warehouseId & BOOTSTRAP_RECEIVING_OPERATOR_SUFFIX)
End Function

Private Function FolderExistsBootstrap(ByVal folderPath As String) As Boolean
    FolderExistsBootstrap = modDeploymentPaths.FolderExistsManaged(folderPath)
End Function

Private Function FileExistsBootstrap(ByVal filePath As String) As Boolean
    FileExistsBootstrap = modDeploymentPaths.FileExistsManaged(filePath)
End Function

Private Function NormalizeFolderPathBootstrap(ByVal folderPath As String) As String
    NormalizeFolderPathBootstrap = modDeploymentPaths.NormalizeManagedFolderPath(folderPath, True)
End Function

Private Function ResolveBootstrapRootPath(ByRef spec As WarehouseSpec) As String
    Dim resolvedPath As String

    resolvedPath = Trim$(spec.PathLocal)
    If resolvedPath = "" Then resolvedPath = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(spec.WarehouseId, False)
    ResolveBootstrapRootPath = NormalizeFolderPathBootstrap(resolvedPath)
    If Right$(ResolveBootstrapRootPath, 1) = "\" Then
        ResolveBootstrapRootPath = Left$(ResolveBootstrapRootPath, Len(ResolveBootstrapRootPath) - 1)
    End If
End Function

Private Function ResolveTemplateRootBootstrap() As String
    Dim basePath As String

    basePath = Trim$(mBootstrapTemplateRootOverride)
    If basePath <> "" Then
        ResolveTemplateRootBootstrap = NormalizeFolderPathBootstrap(basePath)
        Exit Function
    End If

    basePath = GetParentFolderBootstrap(ThisWorkbook.Path)
    If basePath = "" Then basePath = ThisWorkbook.Path
    ResolveTemplateRootBootstrap = NormalizeFolderPathBootstrap(basePath & "\" & BOOTSTRAP_TEMPLATE_FOLDER_NAME)
End Function

Private Function WriteBootstrapDiscoveryFile(ByRef spec As WarehouseSpec, _
                                             ByVal localConfigPath As String, _
                                             ByVal discoveryPath As String, _
                                             ByRef report As String) As Boolean
    Dim fileNum As Integer
    Dim warehouseName As String

    On Error GoTo FailWrite

    warehouseName = spec.WarehouseName
    If Trim$(warehouseName) = "" Then warehouseName = spec.WarehouseId
    EnsureFolderForFileBootstrap discoveryPath

    fileNum = FreeFile
    Open discoveryPath For Output As #fileNum
    Print #fileNum, "{"
    Print #fileNum, "  ""WarehouseId"": """ & EscapeJsonBootstrap(spec.WarehouseId) & ""","
    Print #fileNum, "  ""WarehouseName"": """ & EscapeJsonBootstrap(warehouseName) & ""","
    Print #fileNum, "  ""StationId"": """ & EscapeJsonBootstrap(spec.StationId) & ""","
    Print #fileNum, "  ""ConfigArtifact"": """ & EscapeJsonBootstrap(spec.WarehouseId & "\" & GetFileNameBootstrap(localConfigPath)) & """"
    Print #fileNum, "}"
    Close #fileNum

    WriteBootstrapDiscoveryFile = True
    Exit Function

FailWrite:
    On Error Resume Next
    If fileNum <> 0 Then Close #fileNum
    On Error GoTo 0
    report = "WriteBootstrapDiscoveryFile failed: " & Err.Description
End Function

Private Function ResolveInventoryTemplatePathBootstrap() As String
    ResolveInventoryTemplatePathBootstrap = ResolveTemplateRootBootstrap() & BOOTSTRAP_INVENTORY_TEMPLATE_FILE
End Function

Private Function CopyInventoryTemplateBootstrap(ByVal warehouseId As String, _
                                                ByVal targetPath As String, _
                                                ByRef report As String) As Boolean
    Dim templatePath As String

    templatePath = ResolveInventoryTemplatePathBootstrap()
    If Not EnsureInventoryTemplateExistsBootstrap(templatePath, report) Then Exit Function

    EnsureFolderForFileBootstrap targetPath
    FileCopy templatePath, targetPath
    CopyInventoryTemplateBootstrap = True
End Function

Private Function EnsureInventoryTemplateExistsBootstrap(ByVal templatePath As String, _
                                                        ByRef report As String) As Boolean
    Dim templateRoot As String
    Dim wbTemplate As Workbook
    Dim priorEvents As Boolean
    Dim eventsSuppressed As Boolean

    On Error GoTo FailEnsure
    If FileExistsBootstrap(templatePath) Then
        EnsureInventoryTemplateExistsBootstrap = True
        Exit Function
    End If

    templateRoot = GetParentFolderBootstrap(templatePath)
    EnsureFolderRecursiveBootstrap templateRoot

    priorEvents = Application.EnableEvents
    Application.EnableEvents = False
    eventsSuppressed = True
    Set wbTemplate = Application.Workbooks.Add(xlWBATWorksheet)
    If wbTemplate Is Nothing Then
        report = "Inventory template workbook could not be created."
        GoTo CleanExit
    End If
    If Not EnsureInventorySchemaBridge(wbTemplate, report) Then GoTo CleanExit
    wbTemplate.SaveAs Filename:=templatePath, FileFormat:=50
    wbTemplate.Save
    EnsureInventoryTemplateExistsBootstrap = True

CleanExit:
    On Error Resume Next
    If eventsSuppressed Then Application.EnableEvents = priorEvents
    On Error GoTo 0
    CloseWorkbookIfOpenBootstrap wbTemplate
    Exit Function

FailEnsure:
    report = "EnsureInventoryTemplateExistsBootstrap failed: " & Err.Description
    Resume CleanExit
End Function

Private Function StampBootstrapConfigWorkbook(ByVal wbCfg As Workbook, _
                                              ByRef spec As WarehouseSpec, _
                                              ByRef report As String) As Boolean
    Dim loWh As ListObject
    Dim loSt As ListObject

    On Error GoTo FailStamp

    Set loWh = wbCfg.Worksheets("WarehouseConfig").ListObjects("tblWarehouseConfig")
    Set loSt = wbCfg.Worksheets("StationConfig").ListObjects("tblStationConfig")
    If loWh Is Nothing Or loSt Is Nothing Then
        report = "Config tables were not available after bootstrap."
        Exit Function
    End If

    SetTableCellByColumnBootstrap loWh, 1, "WarehouseId", spec.WarehouseId
    SetTableCellByColumnBootstrap loWh, 1, "WarehouseName", IIf(Trim$(spec.WarehouseName) = "", spec.WarehouseId, spec.WarehouseName)
    SetTableCellByColumnBootstrap loWh, 1, "PathDataRoot", spec.PathLocal
    SetTableCellByColumnBootstrap loWh, 1, "PathSharePointRoot", spec.PathSharePoint

    SetTableCellByColumnBootstrap loSt, 1, "StationId", spec.StationId
    SetTableCellByColumnBootstrap loSt, 1, "WarehouseId", spec.WarehouseId
    SetTableCellByColumnBootstrap loSt, 1, "StationName", spec.AdminUser
    SetTableCellByColumnBootstrap loSt, 1, "PathInboxRoot", spec.PathLocal & "\inbox\"
    SetTableCellByColumnBootstrap loSt, 1, "RoleDefault", "RECEIVE"

    SaveWorkbookIfWritableBootstrap wbCfg
    StampBootstrapConfigWorkbook = True
    Exit Function

FailStamp:
    report = "StampBootstrapConfigWorkbook failed: " & Err.Description
End Function

Private Sub SetTableCellByColumnBootstrap(ByVal lo As ListObject, _
                                          ByVal rowIndex As Long, _
                                          ByVal columnName As String, _
                                          ByVal valueOut As Variant)
    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    idx = lo.ListColumns(columnName).Index
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueOut
End Sub

Private Function BuildReceivingOperatorPathBootstrap(ByRef spec As WarehouseSpec) As String
    Dim operatorRoot As String

    operatorRoot = ResolveLocalOperatorRootBootstrap(spec.WarehouseId, spec.StationId)
    If operatorRoot = "" Then operatorRoot = ResolveBootstrapRootPath(spec)
    If operatorRoot = "" Then Exit Function
    BuildReceivingOperatorPathBootstrap = operatorRoot & "\" & spec.WarehouseId & BOOTSTRAP_RECEIVING_OPERATOR_SUFFIX
End Function

Private Function ResolveLocalOperatorRootBootstrap(ByVal warehouseId As String, ByVal stationId As String) As String
    Dim documentsRoot As String
    Dim userRoot As String
    Dim whSegment As String
    Dim stSegment As String
    Dim localRoot As String

    userRoot = Trim$(Environ$("USERPROFILE"))
    If userRoot = "" Then userRoot = Trim$(Environ$("HOMEDRIVE") & Environ$("HOMEPATH"))
    If userRoot = "" Then Exit Function

    whSegment = SafePathSegmentBootstrap(warehouseId)
    stSegment = SafePathSegmentBootstrap(stationId)
    If whSegment = "" Then Exit Function
    If stSegment = "" Then stSegment = "S1"

    documentsRoot = userRoot & "\Documents"
    localRoot = documentsRoot & "\" & BOOTSTRAP_LOCAL_OPERATOR_ROOT & "\" & whSegment & "\" & stSegment
    ResolveLocalOperatorRootBootstrap = NormalizeFolderPathBootstrap(localRoot)
    If Right$(ResolveLocalOperatorRootBootstrap, 1) = "\" Then
        ResolveLocalOperatorRootBootstrap = Left$(ResolveLocalOperatorRootBootstrap, Len(ResolveLocalOperatorRootBootstrap) - 1)
    End If
End Function

Private Function PrepareReceivingOperatorRuntimeFilesBootstrap(ByRef spec As WarehouseSpec, _
                                                               ByVal operatorPath As String, _
                                                               ByRef report As String) As Boolean
    Dim hubRoot As String
    Dim operatorRoot As String
    Dim configSource As String
    Dim authSource As String
    Dim configTarget As String
    Dim authTarget As String

    On Error GoTo FailPrepare

    hubRoot = ResolveBootstrapRootPath(spec)
    operatorRoot = GetParentFolderBootstrap(operatorPath)
    If hubRoot = "" Or operatorRoot = "" Then
        report = "Operator workbook runtime paths could not be resolved."
        Exit Function
    End If

    EnsureFolderRecursiveBootstrap operatorRoot
    configSource = hubRoot & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    authSource = hubRoot & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"
    configTarget = operatorRoot & "\" & spec.WarehouseId & ".invSys.Config.xlsb"
    authTarget = operatorRoot & "\" & spec.WarehouseId & ".invSys.Auth.xlsb"

    If Not FileExistsBootstrap(configSource) Then
        report = "Warehouse config source not found for operator workbook: " & configSource
        Exit Function
    End If
    If Not FileExistsBootstrap(authSource) Then
        report = "Warehouse auth source not found for operator workbook: " & authSource
        Exit Function
    End If

    If Not CopyBootstrapFileIfDifferent(configSource, configTarget, report) Then Exit Function
    If Not CopyBootstrapFileIfDifferent(authSource, authTarget, report) Then Exit Function

    PrepareReceivingOperatorRuntimeFilesBootstrap = True
    report = "OK"
    Exit Function

FailPrepare:
    report = "PrepareReceivingOperatorRuntimeFiles failed: " & Err.Description
End Function

Private Function CopyBootstrapFileIfDifferent(ByVal sourcePath As String, _
                                              ByVal targetPath As String, _
                                              ByRef report As String) As Boolean
    On Error GoTo FailCopy

    sourcePath = Trim$(sourcePath)
    targetPath = Trim$(targetPath)
    If sourcePath = "" Or targetPath = "" Then
        report = "Copy source or target path was blank."
        Exit Function
    End If
    If StrComp(sourcePath, targetPath, vbTextCompare) = 0 Then
        CopyBootstrapFileIfDifferent = True
        report = "OK"
        Exit Function
    End If

    EnsureFolderForFileBootstrap targetPath
    If FileExistsBootstrap(targetPath) Then Kill targetPath
    FileCopy sourcePath, targetPath
    CopyBootstrapFileIfDifferent = True
    report = "OK"
    Exit Function

FailCopy:
    report = "CopyBootstrapFileIfDifferent failed: " & sourcePath & " -> " & targetPath & ": " & Err.Description
End Function

Private Function CreateOrVerifyReceivingOperatorWorkbookBootstrap(ByRef spec As WarehouseSpec, _
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

    parentFolder = GetParentFolderBootstrap(operatorPath)
    If parentFolder = "" Then
        report = "Operator workbook parent folder could not be resolved."
        GoTo FailSoft
    End If
    EnsureFolderRecursiveBootstrap parentFolder

    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set wb = FindOpenWorkbookByPathBootstrap(operatorPath)
    If wb Is Nothing And FileExistsBootstrap(operatorPath) Then
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
    RemoveNonReceivingOperatorSheetsBootstrap wb
    Call modOperatorReadModel.RefreshInventoryReadModelForWorkbook(wb, spec.WarehouseId, "LOCAL", refreshReport)

    If Trim$(wb.FullName) = "" Then
        wb.SaveAs Filename:=operatorPath, FileFormat:=52
    ElseIf StrComp(wb.FullName, operatorPath, vbTextCompare) <> 0 Then
        wb.SaveAs Filename:=operatorPath, FileFormat:=52
    Else
        wb.Save
    End If

    mLastBootstrapOperatorWorkbookPath = operatorPath
    CreateOrVerifyReceivingOperatorWorkbookBootstrap = True
    report = "OK"
    GoTo CleanExit

FailSoft:
    CreateOrVerifyReceivingOperatorWorkbookBootstrap = False
    If Len(report) = 0 Then report = "CreateOrVerifyReceivingOperatorWorkbook failed."
    GoTo CleanExit

FailCreate:
    report = "CreateOrVerifyReceivingOperatorWorkbook failed: " & Err.Description
    Resume FailSoft

CleanExit:
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEvents
    CloseWorkbookIfOpenBootstrap wb
End Function

Private Function SeedBootstrapDemoInventory(ByRef spec As WarehouseSpec, ByRef report As String) As Boolean
    Dim inventoryWb As Workbook
    Dim payloadItems As Collection
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim processedCount As Long

    On Error GoTo FailSeed

    Set inventoryWb = ResolveInventoryWorkbookBridge(spec.WarehouseId)
    If inventoryWb Is Nothing Then
        report = "Inventory workbook could not be resolved for demo seed."
        GoTo FailSoft
    End If
    If BootstrapDemoSeedAlreadyPresent(inventoryWb) Then
        SeedBootstrapDemoInventory = True
        report = "SKIPPED"
        Exit Function
    End If

    Set payloadItems = BuildBootstrapDemoPayload()
    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        report = "Demo seed payload was empty."
        GoTo FailSoft
    End If

    If Not modRoleEventWriter.QueueMigrationSeedEvent(spec.WarehouseId, spec.StationId, spec.AdminUser, payloadJson, "", BOOTSTRAP_SEED_SOURCE_ID, 0, Nothing, eventIdOut, queueError, "") Then
        report = "QueueMigrationSeedEvent failed: " & queueError
        GoTo FailSoft
    End If

    processedCount = modProcessor.RunBatch(spec.WarehouseId, 0, batchReport)
    If processedCount < 1 Then
        If Not BootstrapDemoSeedAlreadyPresent(inventoryWb) Then
            report = "Processor did not apply demo inventory seed. " & batchReport
            GoTo FailSoft
        End If
    End If

    If Not BootstrapDemoSeedAlreadyPresent(inventoryWb) Then
        report = "Demo inventory seed was not present after processor run."
        GoTo FailSoft
    End If

    SeedBootstrapDemoInventory = True
    report = "SEEDED"
    Exit Function

FailSoft:
    SeedBootstrapDemoInventory = False
    If Len(report) = 0 Then report = "SeedBootstrapDemoInventory failed."
    Exit Function

FailSeed:
    report = "SeedBootstrapDemoInventory failed: " & Err.Description
    Resume FailSoft
End Function

Private Function BuildBootstrapDemoPayload() As Collection
    Dim items As Collection

    Set items = New Collection
    items.Add BuildBootstrapDemoPayloadItem(1, "DEMO-RAW-BLACK-TEA", "Black Tea", "lbs", "CLEARVIEW", "Loose black tea for receiving test.", "Tea Importers", "TEA-001", "raw", 500#)
    items.Add BuildBootstrapDemoPayloadItem(2, "DEMO-RAW-CARDAMOM", "Cardamom", "lbs", "CLEARVIEW", "Cardamom for receiving test.", "Spice House", "SPICE-001", "raw", 50#)
    items.Add BuildBootstrapDemoPayloadItem(3, "DEMO-FG-CLASSIC-CHAI", "Classic Chai Concentrate", "gal", "CLEARVIEW", "Finished good receiving test item.", "Internal", "FG-001", "shippable", 25#)
    Set BuildBootstrapDemoPayload = items
End Function

Private Function BuildBootstrapDemoPayloadItem(ByVal rowVal As Long, _
                                               ByVal sku As String, _
                                               ByVal itemName As String, _
                                               ByVal uom As String, _
                                               ByVal locationVal As String, _
                                               ByVal description As String, _
                                               ByVal vendorName As String, _
                                               ByVal vendorCode As String, _
                                               ByVal category As String, _
                                               ByVal qty As Double) As Object
    Dim item As Object

    Set item = modRoleEventWriter.CreatePayloadItem(rowVal, sku, qty, locationVal, BOOTSTRAP_SEED_SOURCE_ID, "IMPORT")
    item("Description") = description
    item("ITEM_CODE") = sku
    item("Item") = itemName
    item("UOM") = uom
    item("VENDOR(s)") = vendorName
    item("VENDOR_CODE") = vendorCode
    item("CATEGORY") = category
    Set BuildBootstrapDemoPayloadItem = item
End Function

Private Function BootstrapDemoSeedAlreadyPresent(ByVal inventoryWb As Workbook) As Boolean
    BootstrapDemoSeedAlreadyPresent = InventoryWorkbookHasSkuBootstrap(inventoryWb, "DEMO-RAW-BLACK-TEA") _
        And InventoryWorkbookHasSkuBootstrap(inventoryWb, "DEMO-RAW-CARDAMOM") _
        And InventoryWorkbookHasSkuBootstrap(inventoryWb, "DEMO-FG-CLASSIC-CHAI")
End Function

Private Function InventoryWorkbookHasSkuBootstrap(ByVal inventoryWb As Workbook, ByVal sku As String) As Boolean
    Dim lo As ListObject
    Dim rowIndex As Long

    Set lo = FindTableByNameBootstrap(inventoryWb, "tblSkuBalance")
    If lo Is Nothing Then Set lo = FindTableByNameBootstrap(inventoryWb, "invSys")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        If TableCellEqualsBootstrap(lo, rowIndex, "SKU", sku) _
           Or TableCellEqualsBootstrap(lo, rowIndex, "ITEM_CODE", sku) Then
            InventoryWorkbookHasSkuBootstrap = True
            Exit Function
        End If
    Next rowIndex
End Function

Private Function FindTableByNameBootstrap(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindTableByNameBootstrap = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindTableByNameBootstrap Is Nothing Then Exit Function
    Next ws
End Function

Private Function TableCellEqualsBootstrap(ByVal lo As ListObject, _
                                          ByVal rowIndex As Long, _
                                          ByVal columnName As String, _
                                          ByVal expectedValue As String) As Boolean
    Dim idx As Long

    If lo Is Nothing Then Exit Function
    On Error Resume Next
    idx = lo.ListColumns(columnName).Index
    On Error GoTo 0
    If idx = 0 Then Exit Function
    TableCellEqualsBootstrap = (StrComp(Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, idx).Value)), expectedValue, vbTextCompare) = 0)
End Function

Private Sub RemoveNonReceivingOperatorSheetsBootstrap(ByVal wb As Workbook)
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

Private Function FindOpenWorkbookByPathBootstrap(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook

    workbookPath = Trim$(workbookPath)
    If workbookPath = "" Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(Trim$(wb.FullName), workbookPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByPathBootstrap = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub SaveWorkbookIfWritableBootstrap(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    If wb.ReadOnly Then Exit Sub
    If Trim$(wb.Path) = "" Then Exit Sub
    wb.Save
End Sub

Private Sub CloseWorkbookIfOpenBootstrap(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub
    On Error Resume Next
    If Not wb.ReadOnly Then
        If wb.Saved = False Then wb.Save
    End If
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Sub RestoreCoreRootOverrideBootstrap(ByVal priorRootOverride As String)
    If Trim$(priorRootOverride) = "" Then
        modRuntimeWorkbooks.ClearCoreDataRootOverride
    Else
        modRuntimeWorkbooks.SetCoreDataRootOverride priorRootOverride
    End If
End Sub

Private Function GetParentFolderBootstrap(ByVal pathIn As String) As String
    GetParentFolderBootstrap = modDeploymentPaths.GetParentFolderManaged(pathIn)
End Function

Private Sub EnsureFolderForFileBootstrap(ByVal filePath As String)
    Dim folderPath As String

    folderPath = GetParentFolderBootstrap(filePath)
    If folderPath <> "" Then EnsureFolderRecursiveBootstrap folderPath
End Sub

Private Function GetFileNameBootstrap(ByVal fullPath As String) As String
    Dim sepPos As Long

    fullPath = Trim$(Replace$(fullPath, "/", "\"))
    If fullPath = "" Then Exit Function

    sepPos = InStrRev(fullPath, "\")
    If sepPos > 0 Then
        GetFileNameBootstrap = Mid$(fullPath, sepPos + 1)
    Else
        GetFileNameBootstrap = fullPath
    End If
End Function

Private Function SafePathSegmentBootstrap(ByVal segmentText As String) As String
    Dim invalidChars As Variant
    Dim item As Variant

    segmentText = Trim$(segmentText)
    invalidChars = Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
    For Each item In invalidChars
        segmentText = Replace$(segmentText, CStr(item), "_")
    Next item
    SafePathSegmentBootstrap = segmentText
End Function

Private Sub EnsureFolderRecursiveBootstrap(ByVal folderPath As String)
    modDeploymentPaths.EnsureFolderRecursiveManaged folderPath
End Sub

Private Sub DeleteFolderRecursiveBootstrap(ByVal folderPath As String)
    Dim fso As Object

    On Error Resume Next
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then fso.DeleteFolder folderPath, True
    On Error GoTo 0
End Sub

Private Sub LogDiagnosticSafeBootstrap(ByVal categoryName As String, ByVal detailText As String)
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.Name & "'!modPerfLog.LogDiagnostic", categoryName, detailText
    On Error GoTo 0
End Sub

Private Function EscapeJsonBootstrap(ByVal textIn As String) As String
    textIn = Replace$(textIn, "\", "\\")
    textIn = Replace$(textIn, Chr$(34), "\" & Chr$(34))
    EscapeJsonBootstrap = textIn
End Function
