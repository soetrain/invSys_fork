Attribute VB_Name = "modWarehouseBootstrap"
Option Explicit

Private Const BOOTSTRAP_LOCAL_ROOT As String = "C:\invSys"
Private Const BOOTSTRAP_SHAREPOINT_CONFIG_FOLDER As String = "Config"
Private Const BOOTSTRAP_SHAREPOINT_CONFIG_JSON_SUFFIX As String = ".config.json"
Private Const BOOTSTRAP_SHAREPOINT_CONFIG_WORKBOOK_SUFFIX As String = ".invSys.Config.xlsb"
Private Const BOOTSTRAP_TEMPLATE_FOLDER_NAME As String = "templates"
Private Const BOOTSTRAP_INVENTORY_TEMPLATE_FILE As String = "invSys.Data.Inventory.template.xlsb"

Private mBootstrapTemplateRootOverride As String
Private mLastBootstrapReport As String

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

    On Error GoTo FailBootstrap

    mLastBootstrapReport = vbNullString
    If Not ValidateWarehouseSpec(spec, report) Then GoTo FailSoft
    If Trim$(spec.AdminUser) = "" Then
        report = "AdminUser is required."
        GoTo FailSoft
    End If

    rootPath = ResolveBootstrapRootPath(spec)
    If rootPath = "" Then
        report = "PathLocal could not be resolved."
        GoTo FailSoft
    End If
    spec.PathLocal = rootPath

    If WarehouseIdExists(spec.WarehouseId) Then
        report = "WarehouseId already exists."
        GoTo FailSoft
    End If
    If FolderExistsBootstrap(rootPath) Then
        report = "Local warehouse root already exists: " & rootPath
        GoTo FailSoft
    End If

    EnsureFolderRecursiveBootstrap rootPath
    createdRoot = True
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

    If Not modAuth.EnsureStationRoleAuth(spec.WarehouseId, spec.StationId, spec.AdminUser, spec.AdminUser, "ADMIN", authPath, "svc_processor", capabilityOut, report) Then GoTo FailSoft
    If StrComp(capabilityOut, "ADMIN_MAINT", vbTextCompare) <> 0 Then
        report = "Admin capability was not provisioned."
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
    If Not modAuth.CanPerform("ADMIN_MAINT", spec.AdminUser, spec.WarehouseId, spec.StationId, "BOOTSTRAP", "WAREHOUSE-BOOTSTRAP") Then
        report = "Admin user was not granted ADMIN_MAINT."
        GoTo FailSoft
    End If

    BootstrapWarehouseLocal = True
    report = "OK"
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
    LogDiagnosticSafeBootstrap "WAREHOUSE-BOOTSTRAP", "Local bootstrap failed|WarehouseId=" & spec.WarehouseId & "|Root=" & rootPath & "|Reason=" & report
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
        report = "PathLocal could not be resolved."
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
    LocalWarehouseIdExistsBootstrap = FolderExistsBootstrap(BOOTSTRAP_LOCAL_ROOT & "\" & warehouseId)
End Function

Private Function SharePointWarehouseIdExistsBootstrap(ByVal warehouseId As String) As Boolean
    Dim sharePointRoot As String
    Dim configFolder As String

    sharePointRoot = Trim$(modConfig.GetString("PathSharePointRoot", ""))
    If sharePointRoot = "" Then Exit Function

    On Error GoTo SkipUnavailable
    configFolder = NormalizeFolderPathBootstrap(sharePointRoot) & BOOTSTRAP_SHAREPOINT_CONFIG_FOLDER & "\"
    SharePointWarehouseIdExistsBootstrap = _
        FileExistsBootstrap(NormalizeFolderPathBootstrap(sharePointRoot) & warehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_JSON_SUFFIX) Or _
        FileExistsBootstrap(configFolder & warehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_JSON_SUFFIX) Or _
        FileExistsBootstrap(configFolder & warehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_WORKBOOK_SUFFIX) Or _
        FileExistsBootstrap(NormalizeFolderPathBootstrap(sharePointRoot) & warehouseId & "\" & warehouseId & BOOTSTRAP_SHAREPOINT_CONFIG_WORKBOOK_SUFFIX)
    Exit Function

SkipUnavailable:
    LogDiagnosticSafeBootstrap "WAREHOUSE-BOOTSTRAP", _
        "SharePoint collision check skipped|WarehouseId=" & warehouseId & _
        "|Root=" & sharePointRoot & "|Reason=" & Err.Description
    SharePointWarehouseIdExistsBootstrap = False
End Function

Private Function FolderExistsBootstrap(ByVal folderPath As String) As Boolean
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) = "\" And Len(folderPath) > 3 Then folderPath = Left$(folderPath, Len(folderPath) - 1)

    FolderExistsBootstrap = (Len(Dir$(folderPath, vbDirectory)) > 0)
End Function

Private Function FileExistsBootstrap(ByVal filePath As String) As Boolean
    filePath = Trim$(Replace$(filePath, "/", "\"))
    If filePath = "" Then Exit Function

    FileExistsBootstrap = (Len(Dir$(filePath, vbNormal)) > 0)
End Function

Private Function NormalizeFolderPathBootstrap(ByVal folderPath As String) As String
    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Function
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    NormalizeFolderPathBootstrap = folderPath
End Function

Private Function ResolveBootstrapRootPath(ByRef spec As WarehouseSpec) As String
    Dim resolvedPath As String

    resolvedPath = Trim$(spec.PathLocal)
    If resolvedPath = "" Then resolvedPath = BOOTSTRAP_LOCAL_ROOT & "\" & spec.WarehouseId
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
    SetTableCellByColumnBootstrap loSt, 1, "RoleDefault", "ADMIN"

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
    Dim sepPos As Long

    pathIn = Trim$(Replace$(pathIn, "/", "\"))
    sepPos = InStrRev(pathIn, "\")
    If sepPos > 1 Then GetParentFolderBootstrap = Left$(pathIn, sepPos - 1)
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

Private Sub EnsureFolderRecursiveBootstrap(ByVal folderPath As String)
    Dim parentPath As String

    folderPath = Trim$(Replace$(folderPath, "/", "\"))
    If folderPath = "" Then Exit Sub
    If Len(Dir$(folderPath, vbDirectory)) > 0 Then Exit Sub

    parentPath = GetParentFolderBootstrap(folderPath)
    If parentPath <> "" And Len(Dir$(parentPath, vbDirectory)) = 0 Then EnsureFolderRecursiveBootstrap parentPath

    On Error Resume Next
    MkDir folderPath
    On Error GoTo 0
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
