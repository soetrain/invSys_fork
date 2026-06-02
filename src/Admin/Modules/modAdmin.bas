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
    Dim csvPath As String
    Dim payload As Collection

    csvPath = ResolveAdminDemoInventoryCsvPath()
    If csvPath <> "" Then
        Set payload = BuildAdminDemoInventoryPayloadFromCsv(csvPath)
        If Not payload Is Nothing Then
            If payload.Count > 0 Then
                Set BuildAdminDemoInventoryPayload = payload
                Exit Function
            End If
        End If
    End If

    Set BuildAdminDemoInventoryPayload = BuildAdminDemoInventoryFallbackPayload()
End Function

Private Function BuildAdminDemoInventoryFallbackPayload() As Collection
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

    Set BuildAdminDemoInventoryFallbackPayload = rows
End Function

Private Function BuildAdminDemoInventoryPayloadFromCsv(ByVal csvPath As String) As Collection
    On Error GoTo FailCsv

    Dim fso As Object
    Dim textStream As Object
    Dim headerLine As String
    Dim fields As Collection
    Dim headers As Object
    Dim rows As Collection
    Dim lineText As String
    Dim item As Object
    Dim rowVal As Long
    Dim sku As String
    Dim itemName As String
    Dim uom As String
    Dim location As String
    Dim category As String
    Dim qty As Double

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(csvPath) Then Exit Function

    Set textStream = fso.OpenTextFile(csvPath, 1, False)
    If textStream.AtEndOfStream Then GoTo CleanExit

    headerLine = textStream.ReadLine
    Set headers = CsvHeaderMapAdmin(ParseCsvLineAdmin(headerLine))
    Set rows = New Collection

    Do While Not textStream.AtEndOfStream
        lineText = textStream.ReadLine
        If Trim$(lineText) = "" Then GoTo NextLine

        Set fields = ParseCsvLineAdmin(lineText)
        sku = CsvFieldAdmin(fields, headers, "ITEM_CODE")
        itemName = CsvFieldAdmin(fields, headers, "ITEM")
        If sku = "" And itemName = "" Then GoTo NextLine
        If sku = "" Then sku = itemName

        rowVal = CLng(Val(CsvFieldAdmin(fields, headers, "ROW")))
        If rowVal <= 0 Then rowVal = rows.Count + 1
        uom = CsvFieldAdmin(fields, headers, "UOM")
        location = CsvFieldAdmin(fields, headers, "LOCATION")
        category = CsvFieldAdmin(fields, headers, "CATEGORY")
        qty = ResolveDemoSeedQuantityAdmin(category, CsvFieldAdmin(fields, headers, "PHASE"), uom)

        Set item = modRoleEventWriter.CreatePayloadItem(rowVal, sku, qty, location, "Admin CSV demo inventory seed", "IMPORT")
        item("ITEM_CODE") = sku
        item("ITEM") = itemName
        item("UOM") = uom
        item("LOCATION") = location
        item("DESCRIPTION") = CsvFieldAdmin(fields, headers, "DESCRIPTION")
        item("VENDOR(s)") = CsvFieldAdmin(fields, headers, "VENDOR(s)")
        item("VENDOR_CODE") = CsvFieldAdmin(fields, headers, "VENDOR_CODE")
        item("CATEGORY") = category
        If CsvFieldAdmin(fields, headers, "SUBSTITUTION") <> "" Then item("SUBSTITUTION") = CsvFieldAdmin(fields, headers, "SUBSTITUTION")
        If CsvFieldAdmin(fields, headers, "PHASE") <> "" Then item("PHASE") = CsvFieldAdmin(fields, headers, "PHASE")
        If CsvFieldAdmin(fields, headers, "ASSIGNEE") <> "" Then item("ASSIGNEE") = CsvFieldAdmin(fields, headers, "ASSIGNEE")
        rows.Add item
NextLine:
    Loop

    Set BuildAdminDemoInventoryPayloadFromCsv = rows

CleanExit:
    On Error Resume Next
    If Not textStream Is Nothing Then textStream.Close
    On Error GoTo 0
    Exit Function

FailCsv:
    Resume CleanExit
End Function

Private Function ResolveAdminDemoInventoryCsvPath() As String
    Dim fso As Object
    Dim candidates As Collection
    Dim basePath As String
    Dim parentPath As String
    Dim candidate As Variant

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set candidates = New Collection

    basePath = ThisWorkbook.Path
    If basePath <> "" Then
        candidates.Add basePath & "\assets\inv.sample.data.csv"
        parentPath = fso.GetParentFolderName(basePath)
        If parentPath <> "" Then
            candidates.Add parentPath & "\assets\inv.sample.data.csv"
            parentPath = fso.GetParentFolderName(parentPath)
            If parentPath <> "" Then candidates.Add parentPath & "\assets\inv.sample.data.csv"
        End If
    End If

    On Error Resume Next
    candidates.Add CurDir$ & "\assets\inv.sample.data.csv"
    candidates.Add CurDir$ & "\..\assets\inv.sample.data.csv"
    On Error GoTo 0

    For Each candidate In candidates
        If fso.FileExists(CStr(candidate)) Then
            ResolveAdminDemoInventoryCsvPath = CStr(candidate)
            Exit Function
        End If
    Next candidate
End Function

Private Function CsvHeaderMapAdmin(ByVal headers As Collection) As Object
    Dim result As Object
    Dim i As Long
    Dim headerText As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    For i = 1 To headers.Count
        headerText = Trim$(CStr(headers(i)))
        If i = 1 Then headerText = Replace$(headerText, ChrW$(&HFEFF), "")
        If headerText <> "" Then result(headerText) = i
    Next i
    Set CsvHeaderMapAdmin = result
End Function

Private Function CsvFieldAdmin(ByVal fields As Collection, ByVal headers As Object, ByVal headerName As String) As String
    Dim idx As Long

    If fields Is Nothing Then Exit Function
    If headers Is Nothing Then Exit Function
    If Not headers.Exists(headerName) Then Exit Function
    idx = CLng(headers(headerName))
    If idx <= 0 Or idx > fields.Count Then Exit Function
    CsvFieldAdmin = Trim$(CStr(fields(idx)))
End Function

Private Function ParseCsvLineAdmin(ByVal lineText As String) As Collection
    Dim result As Collection
    Dim i As Long
    Dim ch As String
    Dim current As String
    Dim inQuotes As Boolean

    Set result = New Collection
    For i = 1 To Len(lineText)
        ch = Mid$(lineText, i, 1)
        If ch = """" Then
            If inQuotes And i < Len(lineText) And Mid$(lineText, i + 1, 1) = """" Then
                current = current & """"
                i = i + 1
            Else
                inQuotes = Not inQuotes
            End If
        ElseIf ch = "," And Not inQuotes Then
            result.Add current
            current = ""
        Else
            current = current & ch
        End If
    Next i
    result.Add current
    Set ParseCsvLineAdmin = result
End Function

Private Function ResolveDemoSeedQuantityAdmin(ByVal category As String, ByVal phase As String, ByVal uom As String) As Double
    Dim keyText As String

    keyText = LCase$(Trim$(category & " " & phase & " " & uom))
    If InStr(1, keyText, "shippable", vbTextCompare) > 0 Then
        ResolveDemoSeedQuantityAdmin = 24#
    ElseIf InStr(1, keyText, "sell", vbTextCompare) > 0 Then
        ResolveDemoSeedQuantityAdmin = 36#
    ElseIf InStr(1, keyText, "packaging.ship", vbTextCompare) > 0 Then
        ResolveDemoSeedQuantityAdmin = 250#
    ElseIf InStr(1, keyText, "packaging", vbTextCompare) > 0 Then
        ResolveDemoSeedQuantityAdmin = 150#
    ElseIf InStr(1, keyText, "oil", vbTextCompare) > 0 Then
        ResolveDemoSeedQuantityAdmin = 8#
    ElseIf InStr(1, keyText, "spice", vbTextCompare) > 0 Then
        ResolveDemoSeedQuantityAdmin = 25#
    ElseIf InStr(1, keyText, "ingredient", vbTextCompare) > 0 Then
        ResolveDemoSeedQuantityAdmin = 120#
    ElseIf InStr(1, keyText, "tea", vbTextCompare) > 0 Then
        ResolveDemoSeedQuantityAdmin = 200#
    ElseIf InStr(1, keyText, "lbs", vbTextCompare) > 0 Or InStr(1, keyText, "lb", vbTextCompare) > 0 Then
        ResolveDemoSeedQuantityAdmin = 80#
    ElseIf InStr(1, keyText, "ft", vbTextCompare) > 0 Then
        ResolveDemoSeedQuantityAdmin = 1000#
    Else
        ResolveDemoSeedQuantityAdmin = 50#
    End If
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
