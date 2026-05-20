Attribute VB_Name = "modAdmin"
Option Explicit

Sub Admin_Click()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    Call modAdminConsole.OpenAdminConsole(targetWb, report)
End Sub

Sub Set_CurrentUser()
    modRoleEventWriter.PromptSetCurrentUser
End Sub

Sub Open_CreateDeleteUser()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered("", report) Then
        MsgBox "Current invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & vbCrLf & _
               report, vbExclamation, "invSys Admin"
        Exit Sub
    End If
    frmCreateDeleteUser.Show
End Sub

Sub Open_CreateWarehouse()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered("", report) Then
        MsgBox "Current invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & vbCrLf & _
               report, vbExclamation, "invSys Admin"
        Exit Sub
    End If
    frmCreateWarehouse.Show
End Sub

Sub Admin_SetupTesterStation_Click()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered("", report) Then
        MsgBox "Current invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & vbCrLf & _
               report, vbExclamation, "invSys Admin"
        Exit Sub
    End If
    frmSetupTesterStation.Show
End Sub

Sub Open_SetupTesterStation()
    Admin_SetupTesterStation_Click
End Sub

Sub Open_LastTesterWorkbook()
    If modTesterSetup.OpenTesterReceivingWorkbook("") Then
        MsgBox "Tester receiving workbook opened. Use Refresh Inventory, then run Confirm Writes.", vbInformation, "invSys Admin"
    Else
        MsgBox "No tester receiving workbook is available in this Excel session. Run Generate Test Warehouse first.", vbExclamation, "invSys Admin"
    End If
End Sub

Sub Open_WarehouseDirectory()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    PromptForWarehouseDirectoryRootIfNeeded
    If modAdminConsole.OpenWarehouseDirectory(targetWb, report) Then
        MsgBox "Warehouse directory refreshed.", vbInformation, "invSys Admin"
    Else
        If Len(Trim$(report)) = 0 Then report = "Warehouse directory could not be opened."
        MsgBox report, vbExclamation, "invSys Admin"
    End If
End Sub

Sub Add_WarehouseDirectoryRoot()
    Dim report As String
    Dim targetWb As Workbook
    Dim rootPath As String

    rootPath = InputBox("Enter a NAS/server warehouse hub folder or a specific warehouse runtime folder to include in Admin warehouse scans.", _
                        "invSys Admin - Warehouse Root", _
                        "\\100.84.136.19\invSysWH1")
    rootPath = Trim$(rootPath)
    If rootPath = "" Then Exit Sub

    modAdminConsole.RememberWarehouseScanRoot rootPath
    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If modAdminConsole.OpenWarehouseDirectory(targetWb, report) Then
        MsgBox "Warehouse root remembered and directory refreshed.", vbInformation, "invSys Admin"
    Else
        If Len(Trim$(report)) = 0 Then report = "Warehouse directory could not be opened."
        MsgBox report, vbExclamation, "invSys Admin"
    End If
End Sub

Sub Seed_DemoInventory()
    Dim warehouseId As String
    Dim stationId As String
    Dim userId As String
    Dim report As String

    If Not ResolveSeedInventoryContext(warehouseId, stationId, userId, report) Then
        MsgBox report, vbExclamation, "invSys Admin"
        Exit Sub
    End If

    If SeedDemoInventoryForWarehouse(warehouseId, stationId, userId, report) Then
        MsgBox report, vbInformation, "invSys Admin"
    Else
        MsgBox report, vbExclamation, "invSys Admin"
    End If
End Sub

Private Function ResolveSeedInventoryContext(ByRef warehouseId As String, _
                                             ByRef stationId As String, _
                                             ByRef userId As String, _
                                             ByRef report As String) As Boolean
    Dim warehouseOptions As Collection
    Dim runtimeRoot As String
    Dim formReport As String

    warehouseId = Trim$(modConfig.GetWarehouseId())
    stationId = Trim$(modConfig.GetStationId())
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetString("WarehouseId", ""))
    If stationId = "" Then stationId = Trim$(modConfig.GetString("StationId", "S1"))

    userId = Trim$(modRoleEventWriter.ResolveCurrentUserId())
    If userId = "" Then userId = Trim$(Application.UserName)

    Set warehouseOptions = modAdminConsole.GetWarehouseDirectoryOptions(Nothing, formReport)
    If warehouseOptions Is Nothing Then
        report = "No warehouse configs were found. Use Add Warehouse Root or View Warehouses first."
        Exit Function
    End If
    If warehouseOptions.Count = 0 Then
        If Trim$(formReport) = "" Or StrComp(formReport, "OK", vbTextCompare) = 0 Then
            formReport = "No warehouse configs were found. Use Add Warehouse Root or View Warehouses first."
        End If
        report = formReport
        Exit Function
    End If

    frmSeedInventory.Configure warehouseOptions, warehouseId, stationId, userId
    frmSeedInventory.Show
    If Not frmSeedInventory.Accepted Then
        report = "Seed inventory cancelled."
        Unload frmSeedInventory
        Exit Function
    End If

    warehouseId = Trim$(frmSeedInventory.SelectedWarehouseId)
    stationId = Trim$(frmSeedInventory.SelectedStationId)
    runtimeRoot = Trim$(frmSeedInventory.SelectedRuntimeRoot)
    userId = Trim$(frmSeedInventory.SelectedUserId)
    Unload frmSeedInventory

    If warehouseId = "" Then
        report = "WarehouseId is required."
        Exit Function
    End If
    If stationId = "" Then stationId = "S1"
    If userId = "" Then
        report = "Admin user is required."
        Exit Function
    End If
    If runtimeRoot <> "" Then modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot

    If Not modConfig.LoadConfig(warehouseId, stationId) Then
        report = "Config load failed: " & modConfig.Validate()
        Exit Function
    End If

    ResolveSeedInventoryContext = True
End Function

Private Function SeedDemoInventoryForWarehouse(ByVal warehouseId As String, _
                                               ByVal stationId As String, _
                                               ByVal userId As String, _
                                               ByRef report As String) As Boolean
    Dim payloadItems As Collection
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim processedCount As Long
    Dim inboxReport As String

    On Error GoTo FailSeed

    If Not EnsureDemoStationInboxes(warehouseId, stationId, inboxReport) Then
        report = inboxReport
        Exit Function
    End If

    Set payloadItems = BuildAdminDemoInventoryPayload()
    payloadJson = modRoleEventWriter.BuildPayloadJsonFromCollection(payloadItems)
    If payloadJson = "" Or payloadJson = "[]" Then
        report = "Demo inventory payload was empty."
        Exit Function
    End If

    If Not modRoleEventWriter.QueueMigrationSeedEvent(warehouseId, stationId, userId, payloadJson, _
                                                      "ADMIN_DEMO_INVENTORY", "Admin demo inventory seed", _
                                                      0, Nothing, eventIdOut, queueError, "") Then
        report = "Seed event could not be queued: " & queueError & vbCrLf & _
                 "Use Users & Roles to grant ADMIN_MAINT to '" & userId & "' for " & warehouseId & " / " & stationId & "."
        Exit Function
    End If

    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport)
    If processedCount < 1 Then
        report = "Seed event was queued but not applied. " & batchReport
        Exit Function
    End If

    report = "Demo inventory seeded." & vbCrLf & _
             "Warehouse: " & warehouseId & vbCrLf & _
             "Applied events: " & CStr(processedCount) & vbCrLf & _
             "Processor: " & batchReport & vbCrLf & _
             "Now click Refresh Inventory in Receiving."
    SeedDemoInventoryForWarehouse = True
    Exit Function

FailSeed:
    report = "SeedDemoInventory failed: " & Err.Description
End Function

Private Function EnsureDemoStationInboxes(ByVal warehouseId As String, _
                                          ByVal stationId As String, _
                                          ByRef report As String) As Boolean
    Dim inboxPath As String
    Dim stepReport As String

    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "RECEIVE", "", inboxPath, stepReport) Then
        report = "Receiving inbox could not be created or repaired: " & stepReport
        Exit Function
    End If

    inboxPath = ""
    stepReport = ""
    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "SHIP", "", inboxPath, stepReport) Then
        report = "Shipping inbox could not be created or repaired: " & stepReport
        Exit Function
    End If

    inboxPath = ""
    stepReport = ""
    If Not modConfig.EnsureStationInbox(warehouseId, stationId, "PRODUCTION", "", inboxPath, stepReport) Then
        report = "Production inbox could not be created or repaired: " & stepReport
        Exit Function
    End If

    report = "OK"
    EnsureDemoStationInboxes = True
End Function

Private Function BuildAdminDemoInventoryPayload() As Collection
    Dim rows As Collection
    Dim item As Object

    Set rows = New Collection

    Set item = modRoleEventWriter.CreatePayloadItem(1, "DEMO-RAW-BLACK-TEA", 100, "NAS-A1", "Admin demo seed", "IMPORT")
    item("ITEM_CODE") = "DEMO-RAW-BLACK-TEA"
    item("ITEM") = "Black Tea Base"
    item("UOM") = "LB"
    item("DESCRIPTION") = "Demo raw black tea for receiving tests"
    item("VENDOR(s)") = "Demo Vendor"
    item("CATEGORY") = "Raw Material"
    rows.Add item

    Set item = modRoleEventWriter.CreatePayloadItem(2, "DEMO-SPICE-CARDAMOM", 25, "NAS-A2", "Admin demo seed", "IMPORT")
    item("ITEM_CODE") = "DEMO-SPICE-CARDAMOM"
    item("ITEM") = "Cardamom Pods"
    item("UOM") = "LB"
    item("DESCRIPTION") = "Demo spice inventory for receiving tests"
    item("VENDOR(s)") = "Demo Vendor"
    item("CATEGORY") = "Spice"
    rows.Add item

    Set item = modRoleEventWriter.CreatePayloadItem(3, "DEMO-PKG-TIN", 48, "NAS-P1", "Admin demo seed", "IMPORT")
    item("ITEM_CODE") = "DEMO-PKG-TIN"
    item("ITEM") = "Retail Tea Tin"
    item("UOM") = "EA"
    item("DESCRIPTION") = "Demo packaging item for picker tests"
    item("VENDOR(s)") = "Demo Vendor"
    item("CATEGORY") = "Packaging"
    rows.Add item

    Set BuildAdminDemoInventoryPayload = rows
End Function

Private Sub PromptForWarehouseDirectoryRootIfNeeded()
    Dim rootPath As String

    If modAdminConsole.HasRememberedWarehouseScanRoots() Then Exit Sub
    rootPath = InputBox("Optional: enter a NAS/server warehouse root to include in this warehouse scan. Leave blank to scan only local/open warehouse configs.", _
                        "invSys Admin - Warehouse Root", _
                        "\\100.84.136.19\invSysWH1")
    rootPath = Trim$(rootPath)
    If rootPath <> "" Then modAdminConsole.RememberWarehouseScanRoot rootPath
End Sub

Sub Verify_AddinsPublished()
    Dim report As String
    Dim detail As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If modAddinsPublish.VerifyAddinsPublished() Then
        MsgBox "All required add-ins are published." & vbCrLf & modAddinsPublish.GetLastAddinsPublishReport(), vbInformation, "invSys Admin"
    Else
        detail = modAddinsPublish.GetLastAddinsPublishReport()
        If Len(detail) = 0 Then detail = "One or more required add-ins are missing or zero-byte."
        If InStr(1, detail, "PathSharePointRoot is not configured", vbTextCompare) > 0 Then
            detail = detail & vbCrLf & _
                     "Use Create New Warehouse or Setup Tester Station to choose the locally synced invSys SharePoint root first."
        End If
        MsgBox "Add-ins publish verification failed." & vbCrLf & detail, vbExclamation, "invSys Admin"
    End If
End Sub

Sub Export_LoadedPackageReport()
    Dim report As String
    Dim pathOut As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    If modPackageDiagnostics.ExportLoadedPackageReport("", "", "", pathOut, report) Then
        MsgBox "Loaded package report written to:" & vbCrLf & pathOut, vbInformation, "invSys Admin"
    Else
        If Len(Trim$(report)) = 0 Then report = "Loaded package report export failed."
        MsgBox report, vbExclamation, "invSys Admin"
    End If
End Sub

Sub Admin_RetireMigrateWarehouse_Click()
    Dim report As String
    Dim targetWb As Workbook

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    frmRetireMigrateWarehouse.Show
End Sub

Sub Open_RetireMigrateWarehouse()
    Admin_RetireMigrateWarehouse_Click
End Sub

Public Sub Scheduler_RunWarehouseBatch()
    PublishSchedulerResult modAdminConsole.RunScheduledWarehouseBatchForAutomation("", 0)
End Sub

Public Sub Scheduler_RunWarehousePublish()
    PublishSchedulerResult modAdminConsole.RunScheduledWarehousePublishForAutomation("", "")
End Sub

Public Sub Scheduler_RunHQAggregation()
    PublishSchedulerResult modAdminConsole.RunScheduledHQAggregationForAutomation("", "")
End Sub

Private Sub PublishSchedulerResult(ByVal resultText As String)
    Debug.Print resultText
    On Error Resume Next
    Application.StatusBar = resultText
    On Error GoTo 0
End Sub

Public Function ResolveInteractiveAdminWorkbook(Optional ByVal allowAddinFallback As Boolean = True) As Workbook
    Set ResolveInteractiveAdminWorkbook = modAdminWorkbookTarget.ResolveAdminTargetWorkbook(Nothing, ThisWorkbook, allowAddinFallback)
End Function

''''''''''''''''''''''''''''''''''''
' This module contains administrative functions for the application.
' It includes functions to manage user accounts, roles, and permissions. yada yada
' It also includes functions to manage application settings and configurations.
' The functions in this module are used by the frmAdminControls form to perform administrative tasks.
''''''''''''''''''''''''''''''''''''
