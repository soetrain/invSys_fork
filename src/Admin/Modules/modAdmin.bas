Attribute VB_Name = "modAdmin"
Option Explicit

Private Const ADMIN_DEMO_INVENTORY_QTY As Double = 1000#
Private mSeedCallbackAutomationEnabled As Boolean
Private mSeedCallbackAutomationWarehouseId As String
Private mSeedCallbackAutomationStationId As String
Private mSeedCallbackAutomationUserId As String
Private mSeedCallbackAutomationAction As String
Private mSeedCallbackAutomationUploadPath As String
Private mSeedCallbackAutomationDeleteConfirmed As Boolean
Private mSeedCallbackLastResult As String

Sub Admin_Click()
    modAdminConsoleActivity.OpenFromRibbon
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
    modAdminCreateEntry.OpenFromRibbon
End Sub

Sub Admin_SetupTesterStation_Click()
    modTesterPresentation.ShowTesterSetup ResolveInteractiveAdminWorkbook()
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
    modAdminDirectoryActivity.OpenFromRibbon targetWb
End Sub

Sub Open_Settings()
    Dim observation As cAdminSettingsEntry
    Set observation = New cAdminSettingsEntry
    observation.OpenSettings
End Sub

Public Function AdminSettingsFormInitializeSmokeForWorkbook(ByVal operatorWb As Workbook) As String
    On Error GoTo FailSmoke

    Dim frm As frmAdminSettings
    Dim detail As String
    Dim configRows As Variant
    Dim expectedRowCount As Long

    Set frm = New frmAdminSettings
    detail = frm.TestInitializeConfigEditor()
    configRows = modConfig.GetConfigEditorRows()
    If Not IsArray(configRows) Then
        Err.Raise vbObjectError + 7319, "AdminSettingsFormInitializeSmokeForWorkbook", _
                  "Admin config schema rows could not be resolved. " & detail
    End If
    expectedRowCount = UBound(configRows, 1)
    If InStr(1, detail, "Rows=" & CStr(expectedRowCount), vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 7320, "AdminSettingsFormInitializeSmokeForWorkbook", _
                  "Admin config editor did not load all " & CStr(expectedRowCount) & " config keys. " & detail
    End If
    AdminSettingsFormInitializeSmokeForWorkbook = "OK|" & detail

CleanExit:
    On Error Resume Next
    If Not frm Is Nothing Then Unload frm
    Set frm = Nothing
    On Error GoTo 0
    Exit Function

FailSmoke:
    AdminSettingsFormInitializeSmokeForWorkbook = "FAIL|" & CStr(Err.Number) & "|" & Err.Description
    Resume CleanExit
End Function

Sub Add_WarehouseDirectoryRoot()
    modAdminRootActivity.AddFromRibbon
End Sub

Sub Seed_DemoInventory()
    Dim warehouseId As String
    Dim stationId As String
    Dim userId As String
    Dim report As String
    Dim stage As String
    Dim actionName As String
    Dim uploadPath As String
    Dim runtimeRoot As String
    Dim storedPath As String
    Dim deleteConfirmed As Boolean
    Dim succeeded As Boolean

    On Error GoTo FailSeedCallback

    stage = "context resolution"
    mSeedCallbackLastResult = ""
    If Not ResolveSeedInventoryContext(warehouseId, stationId, userId, actionName, uploadPath, runtimeRoot, report) Then
        If Trim$(report) = "" Then
            mSeedCallbackLastResult = "CANCEL|"
        Else
            If Not mSeedCallbackAutomationEnabled Then MsgBox report, vbExclamation, "invSys Admin"
            mSeedCallbackLastResult = "FAIL|" & report
        End If
        GoTo CleanExit
    End If

    If actionName = modAdminInventorySeed.DEMO_ACTION_DELETE Then
        If mSeedCallbackAutomationEnabled Then
            deleteConfirmed = mSeedCallbackAutomationDeleteConfirmed
        Else
            deleteConfirmed = (MsgBox( _
                "Delete all active DEMO- inventory from this warehouse?" & vbCrLf & vbCrLf & _
                "This posts audited adjustments that deplete demo entities; event history is retained.", _
                vbQuestion Or vbYesNo Or vbDefaultButton2, "invSys Admin - Demo Inventory") = vbYes)
        End If
        If Not deleteConfirmed Then
            report = "Delete Demo Inventory cancelled."
            mSeedCallbackLastResult = "CANCEL|" & report
            GoTo CleanExit
        End If
    ElseIf actionName = modAdminInventorySeed.DEMO_ACTION_DELETE_DATA_SET Then
        If mSeedCallbackAutomationEnabled Then
            deleteConfirmed = mSeedCallbackAutomationDeleteConfirmed
        Else
            deleteConfirmed = (MsgBox( _
                "Delete the selected uploaded demo inventory data set?" & vbCrLf & vbCrLf & _
                "This deletes only the stored CSV definition. It does not change inventory already seeded from it.", _
                vbQuestion Or vbYesNo Or vbDefaultButton2, "invSys Admin - Demo Inventory") = vbYes)
        End If
        If Not deleteConfirmed Then
            report = "Delete Data Set cancelled."
            mSeedCallbackLastResult = "CANCEL|" & report
            GoTo CleanExit
        End If
    ElseIf actionName = modAdminInventorySeed.DEMO_ACTION_UPLOAD _
           Or actionName = modAdminInventorySeed.DEMO_ACTION_UPLOAD_DATA_SET Then
        If uploadPath = "" And Not mSeedCallbackAutomationEnabled Then uploadPath = PromptDemoInventoryCsvPath()
        If uploadPath = "" Then
            report = "Upload Data Set cancelled."
            mSeedCallbackLastResult = "CANCEL|" & report
            GoTo CleanExit
        End If
    End If

    stage = LCase$(actionName) & " action"
    Select Case actionName
        Case modAdminInventorySeed.DEMO_ACTION_SEED
            succeeded = modAdminSeedObservation.Seed(warehouseId, stationId, userId, runtimeRoot, uploadPath, report)
        Case modAdminInventorySeed.DEMO_ACTION_DELETE
            succeeded = modAdminInventorySeed.DeleteDemoInventoryForWarehouse( _
                warehouseId, stationId, userId, report)
        Case modAdminInventorySeed.DEMO_ACTION_UPLOAD
            succeeded = modAdminSeedObservation.Seed(warehouseId, stationId, userId, runtimeRoot, uploadPath, report)
        Case modAdminInventorySeed.DEMO_ACTION_UPLOAD_DATA_SET
            succeeded = modAdminInventorySeed.ImportDemoInventoryDataSet( _
                runtimeRoot, uploadPath, storedPath, report)
        Case modAdminInventorySeed.DEMO_ACTION_DELETE_DATA_SET
            succeeded = modAdminInventorySeed.DeleteDemoInventoryDataSet( _
                runtimeRoot, uploadPath, report)
        Case Else
            report = "Choose a demo inventory or data set action."
    End Select

    If succeeded Then
        mSeedCallbackLastResult = "OK|" & report
        If Not mSeedCallbackAutomationEnabled Then MsgBox report, vbInformation, "invSys Admin"
    Else
        mSeedCallbackLastResult = "FAIL|" & report
        If Not mSeedCallbackAutomationEnabled Then MsgBox report, vbExclamation, "invSys Admin"
    End If

CleanExit:
    mSeedCallbackAutomationEnabled = False
    mSeedCallbackAutomationWarehouseId = ""
    mSeedCallbackAutomationStationId = ""
    mSeedCallbackAutomationUserId = ""
    mSeedCallbackAutomationAction = ""
    mSeedCallbackAutomationUploadPath = ""
    mSeedCallbackAutomationDeleteConfirmed = False
    Exit Sub

FailSeedCallback:
    report = "Seed Demo Inventory failed at " & stage & "." & vbCrLf & _
             "Error " & CStr(Err.Number) & vbCrLf & _
             "Source: " & SanitizeSeedCallbackErrorText(Err.Source) & vbCrLf & _
             SanitizeSeedCallbackErrorText(Err.Description)
    mSeedCallbackLastResult = "FAIL|" & report
    If Not mSeedCallbackAutomationEnabled Then MsgBox report, vbExclamation, "invSys Admin"
    Resume CleanExit
End Sub

Public Sub SetSeedInventorySelectionForAutomation(ByVal warehouseId As String, _
                                                  ByVal stationId As String, _
                                                  ByVal userId As String)
    mSeedCallbackAutomationWarehouseId = Trim$(warehouseId)
    mSeedCallbackAutomationStationId = Trim$(stationId)
    mSeedCallbackAutomationUserId = Trim$(userId)
    mSeedCallbackAutomationAction = modAdminInventorySeed.DEMO_ACTION_SEED
    mSeedCallbackAutomationUploadPath = ""
    mSeedCallbackAutomationDeleteConfirmed = False
    mSeedCallbackAutomationEnabled = True
End Sub

Public Sub SetDemoInventorySelectionForAutomation(ByVal warehouseId As String, _
                                                  ByVal stationId As String, _
                                                  ByVal userId As String, _
                                                  ByVal actionName As String, _
                                                  Optional ByVal uploadPath As String = "", _
                                                  Optional ByVal deleteConfirmed As Boolean = False)
    mSeedCallbackAutomationWarehouseId = Trim$(warehouseId)
    mSeedCallbackAutomationStationId = Trim$(stationId)
    mSeedCallbackAutomationUserId = Trim$(userId)
    mSeedCallbackAutomationAction = UCase$(Trim$(actionName))
    mSeedCallbackAutomationUploadPath = Trim$(uploadPath)
    mSeedCallbackAutomationDeleteConfirmed = deleteConfirmed
    mSeedCallbackAutomationEnabled = True
End Sub

Public Function GetLastDemoInventoryCallbackResultForAutomation() As String
    GetLastDemoInventoryCallbackResultForAutomation = mSeedCallbackLastResult
End Function

Public Function RunDemoInventoryActionCallbackForAutomation(ByVal warehouseId As String, _
                                                            ByVal stationId As String, _
                                                            ByVal userId As String, _
                                                            ByVal actionName As String, _
                                                            Optional ByVal uploadPath As String = "", _
                                                            Optional ByVal deleteConfirmed As Boolean = False) As String
    SetDemoInventorySelectionForAutomation warehouseId, stationId, userId, _
        actionName, uploadPath, deleteConfirmed
    Seed_DemoInventory
    RunDemoInventoryActionCallbackForAutomation = mSeedCallbackLastResult
End Function

Public Function DemoInventoryFormContractForAutomation() As String
    DemoInventoryFormContractForAutomation = frmSeedInventory.TestDemoInventoryActionContract()
    Unload frmSeedInventory
End Function

Public Function InventoryEditSelectionContractForAutomation() As String
    On Error GoTo Failed

    InventoryEditSelectionContractForAutomation = _
        frmAddInventoryItem.TestEditItemComboSelectionContract()
    Unload frmAddInventoryItem
    Exit Function

Failed:
    InventoryEditSelectionContractForAutomation = "FAIL|Error=" & CStr(Err.Number) & _
        " " & Err.Description
    On Error Resume Next
    Unload frmAddInventoryItem
    On Error GoTo 0
End Function

Public Function InventoryAddVisibilityDropdownContractForAutomation() As String
    On Error GoTo Failed

    Dim payloadItem As Object
    Dim systemKey As String
    Dim exactEntityCreate As Boolean
    Dim formEvidence As String

    systemKey = modRoleEventWriter.CreateSystemKey()
    Set payloadItem = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        systemKey, "SKU-ADMIN-ADD-CONTRACT", 1, "CLEARVIEW", "GOOD", "", _
        "Admin Add Item contract evidence")
    exactEntityCreate = (Trim$(CStr(payloadItem("System_Key"))) = systemKey And _
        UCase$(Trim$(CStr(payloadItem("IoType")))) = "CREATE")
    formEvidence = frmAddInventoryItem.TestAddItemVisibilityDropdownContract()
    InventoryAddVisibilityDropdownContractForAutomation = formEvidence & _
        "|ExactEntityCreate=" & CStr(exactEntityCreate)
    Unload frmAddInventoryItem
    Exit Function

Failed:
    InventoryAddVisibilityDropdownContractForAutomation = "FAIL|Error=" & _
        CStr(Err.Number) & " " & Err.Description
    On Error Resume Next
    Unload frmAddInventoryItem
    On Error GoTo 0
End Function

Public Function InventoryDeleteContractForAutomation() As String
    On Error GoTo Failed

    Dim formEvidence As String
    Dim countedItem As Object
    Dim utilityItem As Object
    Dim exactKey As Boolean
    Dim utilityZeroDelta As Boolean

    formEvidence = frmAddInventoryItem.TestInventoryDeleteActionContract()
    Set countedItem = BuildInventoryRetirementPayloadItemAdmin( _
        "SYS-DELETE-COUNTED", "SKU-HONEY", 25, "CLEARVIEW", "GOOD", "", "contract")
    Set utilityItem = BuildInventoryRetirementPayloadItemAdmin( _
        "SYS-DELETE-UTILITY", "SKU-WATER", 0, "CLEARVIEW", "GOOD", "", "contract")
    exactKey = (CStr(countedItem("System_Key")) = "SYS-DELETE-COUNTED" And _
        CDbl(countedItem("Qty")) = -25 And _
        CStr(countedItem("InventoryState")) = "RETIRED")
    utilityZeroDelta = (CStr(utilityItem("System_Key")) = "SYS-DELETE-UTILITY" And _
        CDbl(utilityItem("Qty")) = 0 And _
        CStr(utilityItem("IoType")) = "RETIRE")

    InventoryDeleteContractForAutomation = formEvidence & _
        "|ExactKey=" & CStr(exactKey) & _
        "|UtilityZeroDelta=" & CStr(utilityZeroDelta)
    Unload frmAddInventoryItem
    Exit Function

Failed:
    InventoryDeleteContractForAutomation = "FAIL|Error=" & CStr(Err.Number) & _
        " " & Err.Description
    On Error Resume Next
    Unload frmAddInventoryItem
    On Error GoTo 0
End Function

Public Function InventoryWorksheetContractForAutomation(ByVal operatorWb As Workbook) As String
    On Error GoTo Failed

    InventoryWorksheetContractForAutomation = _
        frmAddInventoryItem.TestInventoryWorksheetActionContract(operatorWb)
    Unload frmAddInventoryItem
    Exit Function

Failed:
    InventoryWorksheetContractForAutomation = "FAIL|Error=" & CStr(Err.Number) & _
        " " & Err.Description
    On Error Resume Next
    modAdminInventoryWorksheet.EndInventoryWorksheetAutomation
    Unload frmAddInventoryItem
    On Error GoTo 0
End Function

Public Function NextInventoryCodeOrdinalForWorksheet(ByVal warehouseId As String) As Long
    NextInventoryCodeOrdinalForWorksheet = NextInventoryCodeOrdinalAdmin(warehouseId)
End Function

Public Function GenerateInventoryItemCodeForWorksheet(ByVal warehouseId As String, _
                                                      ByVal codeOrdinal As Long) As String
    GenerateInventoryItemCodeForWorksheet = GenerateInventorySkuAdmin(warehouseId, codeOrdinal)
End Function

Public Function InventoryItemCatalogContainsForWorksheet(ByVal warehouseId As String, _
                                                         ByVal sku As String) As Boolean
    Dim catalogItems As Collection
    Dim item As Variant

    sku = Trim$(sku)
    If sku = "" Then Exit Function
    Set catalogItems = LoadInventoryCatalogItemsAdmin(warehouseId)
    For Each item In catalogItems
        If StrComp(modInventoryUiText.CatalogItemTextAdmin(item, "SKU"), sku, vbTextCompare) = 0 Then
            InventoryItemCatalogContainsForWorksheet = True
            Exit Function
        End If
    Next item
End Function

Public Function ApplyInventoryWorksheetRecordForWarehouse(ByVal warehouseId As String, _
                                                          ByVal stationId As String, _
                                                          ByVal userId As String, _
                                                          ByVal record As Object, _
                                                          ByRef report As String, _
                                                          Optional ByVal interactiveContext As String = "", _
                                                          Optional ByRef observedEvent As String = "") As Boolean
    Dim customFields As Object
    Dim actionName As String
    Dim sku As String
    Dim itemName As String
    Dim uom As String
    Dim locationValue As String
    Dim categoryValue As String
    Dim vendorValue As String
    Dim qty As Double
    Dim editReason As String
    Dim actionSucceeded As Boolean
    Dim qtyReport As String
    Dim editStamp As String, receiptRow As Long

    On Error GoTo Failed
    observedEvent = ""
    If record Is Nothing Then report = "Inventory worksheet record is unavailable.": Exit Function
    actionName = UCase$(Trim$(CStr(record("Action"))))
    sku = Trim$(CStr(record("ItemCode")))
    itemName = Trim$(CStr(record("ItemName")))
    uom = Trim$(CStr(record("UOM")))
    locationValue = Trim$(CStr(record("Location")))
    categoryValue = Trim$(CStr(record("Category")))
    vendorValue = Trim$(CStr(record("Vendors")))
    qty = CDbl(record("Quantity"))
    editReason = Trim$(CStr(record("EditReason")))
    Set customFields = record("CustomFields")
    If modInventoryItemFields.IsNonCountedCustomFieldsAdmin(customFields) Then
        If categoryValue = "" Then categoryValue = modInventoryItemFields.NonCountedItemKindAdmin(customFields)
        If vendorValue = "" And modInventoryItemFields.NonCountedItemKindAdmin(customFields) = "UTILITY" Then vendorValue = "Utility"
    End If

    If actionName = "ADD" Then
        If interactiveContext <> "" Then receiptRow = CLng(record("TableRow"))
        ApplyInventoryWorksheetRecordForWarehouse = AddInventoryItemForWarehouse( _
            warehouseId, stationId, userId, sku, itemName, _
            uom, locationValue, qty, CStr(record("Description")), vendorValue, _
            CStr(record("VendorCode")), categoryValue, CStr(record("ExternalCode")), _
            CStr(record("Picture")), customFields, report, interactiveContext, _
            CStr(record("CreateEventId")), receiptRow, observedEvent)
        Exit Function
    End If
    If actionName <> "EDIT" Then report = "Inventory worksheet Action must be ADD or EDIT.": Exit Function
    If interactiveContext <> "" Then
        ApplyInventoryWorksheetRecordForWarehouse = modInventoryEditCommands.Execute(interactiveContext, CLng(record("TableRow")), report, observedEvent)
        Exit Function
    End If

    editStamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    customFields("LAST_EDIT_REASON") = editReason
    customFields("LAST_EDIT_AT") = editStamp
    customFields("LAST_EDIT_USER") = userId
    customFields("EDIT_HISTORY_APPEND") = editStamp & " | User=" & userId & _
        " | SKU=" & sku & " | Reason=" & editReason
    actionSucceeded = UpdateInventoryItemCatalogForWarehouse(warehouseId, sku, itemName, _
        uom, locationValue, CStr(record("Description")), vendorValue, _
        CStr(record("VendorCode")), categoryValue, CStr(record("ExternalCode")), _
        CStr(record("Picture")), customFields, report, interactiveContext, userId, stationId)
    If actionSucceeded And Not modInventoryItemFields.IsNonCountedCustomFieldsAdmin(customFields) And _
            CBool(record("HasQuantity")) Then
        If SetInventoryQuantityForWarehouse(warehouseId, stationId, userId, _
                sku, itemName, uom, locationValue, qty, _
                editReason, qtyReport, interactiveContext) Then
            report = report & vbCrLf & vbCrLf & qtyReport
        Else
            report = "Inventory item catalog fields were updated, but set qty failed." & vbCrLf & qtyReport
            actionSucceeded = False
        End If
    End If
    ApplyInventoryWorksheetRecordForWarehouse = actionSucceeded
    Exit Function

Failed:
    report = "Inventory worksheet row apply failed: " & Err.Description
End Function

Sub Add_InventoryItem()
    Dim warehouseId As String
    Dim stationId As String
    Dim userId As String
    Dim report As String
    Dim targetWb As Workbook
    Dim sku As String
    Dim codeOrdinal As Long
    Dim defaultLocation As String
    Dim catalogItems As Object
    Dim addForm As frmAddInventoryItem
    Dim accepted As Boolean
    Dim isEdit As Boolean
    Dim deleteRequested As Boolean
    Dim deleteReason As String
    Dim formSku As String
    Dim formItemName As String
    Dim formUom As String
    Dim formLocation As String
    Dim formQty As Double
    Dim formDescription As String
    Dim formVendorName As String
    Dim formVendorCode As String
    Dim formCategory As String
    Dim formExternalCode As String
    Dim formImagePath As String
    Dim formEditReason As String
    Dim formCustomFields As Object
    Dim editStamp As String
    Dim actionSucceeded As Boolean
    Dim interactiveContext As String, sourceValues As Variant
    Dim activity As cProvisioningActivity

    Set targetWb = ResolveInteractiveAdminWorkbook(False)
    If targetWb Is Nothing Then
        MsgBox "Open the saved Admin/operator workbook before adding or editing inventory items.", _
               vbExclamation, "invSys Admin"
        Exit Sub
    End If

    interactiveContext = modInventoryCommands.Capture(targetWb.Name)
    Set activity = New cProvisioningActivity
    activity.Initialize targetWb, "ADMIN_INVENTORY_", "Inventory"
    sourceValues = modInventoryCommands.SourceValues(interactiveContext)
    If Not IsArray(sourceValues) Then
        MsgBox "Sign in to the current source warehouse as an authorized Admin and reopen Add/Edit Inventory.", vbExclamation, "invSys Admin"
        modInventoryCommands.Release interactiveContext
        activity.OpeningFailed
        Exit Sub
    End If
    warehouseId = CStr(sourceValues(0)): stationId = CStr(sourceValues(1)): userId = CStr(sourceValues(2))

    codeOrdinal = NextInventoryCodeOrdinalAdmin(warehouseId)
    sku = GenerateInventorySkuAdmin(warehouseId, codeOrdinal)
    defaultLocation = Trim$(modConfig.GetString("DefaultLocation", ""))
    Set catalogItems = LoadInventoryCatalogItemsAdmin(warehouseId)

    If Not modInventoryCommands.CheckForm(interactiveContext, report) Then GoTo CleanExit
    Set addForm = New frmAddInventoryItem
    addForm.ObserveWith activity
    addForm.Configure warehouseId, stationId, userId, sku, defaultLocation
    addForm.SetOperatorWorkbook targetWb
    addForm.SetInteractiveContext interactiveContext
    LoadCatalogItemsIntoAddInventoryForm addForm, catalogItems

    Do
        addForm.Show vbModal

        On Error GoTo FormUnavailable
        accepted = addForm.Accepted
        On Error GoTo 0
        If Not accepted Then Exit Do
        If Not modInventoryCommands.CheckForm(interactiveContext, report) Then Exit Do

        On Error GoTo FormUnavailable
        isEdit = addForm.EditMode
        deleteRequested = addForm.DeleteRequested
        deleteReason = addForm.DeleteReason
        formSku = addForm.GeneratedSku
        formItemName = addForm.ItemName
        formUom = addForm.Uom
        formLocation = addForm.LocationValue
        formQty = addForm.StartingQty
        formDescription = addForm.DescriptionValue
        formVendorName = addForm.VendorName
        formVendorCode = addForm.VendorCode
        formCategory = addForm.Category
        formExternalCode = addForm.ExternalCode
        formImagePath = addForm.ImagePath
        formEditReason = addForm.EditReason
        Set formCustomFields = addForm.CustomFields
        On Error GoTo 0

        report = vbNullString
        actionSucceeded = False
        If deleteRequested Then
            actionSucceeded = RetireInventoryItemForWarehouse(warehouseId, stationId, userId, _
                formSku, formItemName, deleteReason, report, interactiveContext)
        ElseIf isEdit Then
            editStamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
            formCustomFields("LAST_EDIT_REASON") = formEditReason
            formCustomFields("LAST_EDIT_AT") = editStamp
            formCustomFields("LAST_EDIT_USER") = userId
            formCustomFields("EDIT_HISTORY_APPEND") = editStamp & " | User=" & userId & " | SKU=" & formSku & " | Reason=" & formEditReason
            actionSucceeded = UpdateInventoryItemCatalogForWarehouse(warehouseId, formSku, formItemName, _
                                                                     formUom, formLocation, _
                                                                     formDescription, formVendorName, _
                                                                     formVendorCode, formCategory, _
                                                                     formExternalCode, formImagePath, _
                                                                     formCustomFields, report, interactiveContext)
            If actionSucceeded And Not modInventoryItemFields.IsNonCountedCustomFieldsAdmin(formCustomFields) Then
                Dim qtyReport As String
                If SetInventoryQuantityForWarehouse(warehouseId, stationId, userId, formSku, _
                                                    formItemName, formUom, formLocation, formQty, formEditReason, qtyReport, interactiveContext) Then
                    report = report & vbCrLf & vbCrLf & qtyReport
                Else
                    report = "Inventory item catalog fields were updated, but set qty failed." & vbCrLf & qtyReport
                    actionSucceeded = False
                End If
            End If
        Else
            actionSucceeded = AddInventoryItemForWarehouse(warehouseId, stationId, userId, _
                                                           formSku, formItemName, _
                                                           formUom, formLocation, _
                                                           formQty, formDescription, _
                                                           formVendorName, formVendorCode, _
                                                           formCategory, formExternalCode, _
                                                           formImagePath, formCustomFields, report, interactiveContext)
        End If

        If actionSucceeded Then
            MsgBox report, vbInformation, "invSys Admin"
            If Not modInventoryCommands.CheckForm(interactiveContext, report) Then Exit Do
            codeOrdinal = NextInventoryCodeOrdinalAdmin(warehouseId)
            sku = GenerateInventorySkuAdmin(warehouseId, codeOrdinal)
            Set catalogItems = LoadInventoryCatalogItemsAdmin(warehouseId)
            If Not modInventoryCommands.CheckForm(interactiveContext, report) Then Exit Do
            addForm.Configure warehouseId, stationId, userId, sku, defaultLocation
            addForm.SetOperatorWorkbook targetWb
            addForm.SetInteractiveContext interactiveContext
            LoadCatalogItemsIntoAddInventoryForm addForm, catalogItems
        Else
            MsgBox report, vbExclamation, "invSys Admin"
            If Not modInventoryCommands.CheckForm(interactiveContext, report) Then Exit Do
        End If
    Loop

CleanExit:
    On Error Resume Next
    If Not activity Is Nothing Then activity.OpeningFailed
    Unload addForm
    modInventoryCommands.Release interactiveContext
    On Error GoTo 0
    Exit Sub

FormUnavailable:
    report = Err.Description
    On Error GoTo 0
    If Trim$(report) <> "" Then
        If InStr(1, report, "callee", vbTextCompare) = 0 _
           And InStr(1, report, "connections are invalid", vbTextCompare) = 0 Then
            MsgBox report, vbExclamation, "invSys Admin"
        End If
    End If
    Resume CleanExit
End Sub

Public Function AddInventoryItemForWarehouse(ByVal warehouseId As String, _
                                             ByVal stationId As String, _
                                             ByVal userId As String, _
                                             ByVal sku As String, _
                                             ByVal itemName As String, _
                                             ByVal uom As String, _
                                             ByVal locationVal As String, _
                                             ByVal qty As Double, _
                                             Optional ByVal description As String = "", _
                                             Optional ByVal vendorName As String = "", _
                                             Optional ByVal vendorCode As String = "", _
                                             Optional ByVal category As String = "", _
                                             Optional ByVal externalCode As String = "", _
                                             Optional ByVal imagePath As String = "", _
                                             Optional ByVal customFields As Object = Nothing, _
                                             Optional ByRef report As String = "", _
                                             Optional ByVal interactiveContext As String = "", _
                                             Optional ByVal creationEventId As String = "", _
                                             Optional ByVal stagingRow As Long = 0, _
                                             Optional ByRef observedEvent As String = "") As Boolean
    Dim payloadJson As String
    Dim eventIdOut As String
    Dim queueError As String
    Dim batchReport As String
    Dim processedCount As Long
    Dim inboxReport As String
    Dim systemKey As String
    Dim nonCounted As Boolean, applied As Boolean
    Dim sourceValues As Variant
    Dim receiptFound As Boolean, receiptApplied As Boolean, receiptPayload As String, receiptKey As String

    On Error GoTo FailAdd
    observedEvent = ""

    warehouseId = Trim$(warehouseId)
    stationId = Trim$(stationId)
    userId = Trim$(userId)
    sku = Trim$(sku)
    itemName = Trim$(itemName)
    uom = Trim$(uom)
    locationVal = Trim$(locationVal)

    If warehouseId = "" Then report = "WarehouseId is required.": Exit Function
    If stationId = "" Then stationId = modStationIdentity.CurrentComputerStationId()
    If userId = "" Then report = "Admin user is required.": Exit Function
    If sku = "" Then report = "SKU is required.": Exit Function
    If itemName = "" Then itemName = sku
    If uom = "" Then report = "UOM is required.": Exit Function
    nonCounted = modInventoryItemFields.IsNonCountedCustomFieldsAdmin(customFields)
    If nonCounted Then
        qty = 0#
    ElseIf qty < 0 Then
        report = "Starting quantity cannot be negative."
        Exit Function
    End If

    If interactiveContext <> "" Then
        sourceValues = modInventoryCommands.SourceValues(interactiveContext)
        If Not IsArray(sourceValues) Then Exit Function
        If sourceValues(0) <> warehouseId Or sourceValues(1) <> stationId Or sourceValues(2) <> userId Then
            report = "The captured inventory command source differs. Reopen Add/Edit Inventory.": Exit Function
        End If
        If Not modInventoryCommands.CheckForm(interactiveContext, report) Then Exit Function
    Else
        If Not modAuth.LoadAuth(warehouseId) Then report = "Provisioned authority could not be loaded.": Exit Function
        If Not modAuth.HasProvisionedCapabilityForSystem("ADMIN_MAINT", userId, warehouseId, stationId) Then
            report = "The provisioned user lacks ADMIN_MAINT for this warehouse and station.": Exit Function
        End If
    End If
    eventIdOut = creationEventId
    If creationEventId <> "" Then
        If interactiveContext = "" Or stagingRow < 1 Then report = "A captured worksheet command is required for receipt retry.": Exit Function
        If Not modInventoryCreateReceipts.Resolve(interactiveContext, warehouseId, stationId, userId, _
                eventIdOut, receiptFound, receiptKey, receiptPayload, receiptApplied, report) Then Exit Function
    End If
    If Not receiptApplied Then
    If Not EnsureDemoStationInboxes(warehouseId, stationId, inboxReport) Then
        report = inboxReport
        Exit Function
    End If
    End If

    If interactiveContext <> "" Then
        If Not modInventoryCommands.CheckForm(interactiveContext, report) Then Exit Function
    End If
    If Not receiptFound Then
        systemKey = modRoleEventWriter.CreateSystemKey()
    Else
        systemKey = receiptKey
    End If
    payloadJson = BuildInventoryCreatePayloadAdmin(systemKey, sku, itemName, uom, locationVal, qty, _
        description, vendorName, vendorCode, category, externalCode, imagePath, customFields)
    If payloadJson = "" Or payloadJson = "[]" Then
        report = "Inventory item payload was empty."
        Exit Function
    End If

    If receiptFound Then
        If Not modInventoryCreateReceipts.PayloadMatches(payloadJson, receiptPayload) Then
            report = "The pending row differs from its exact queued create. Reconcile it before retrying.": Exit Function
        End If
        observedEvent = eventIdOut & vbTab & "Submitted"
        If receiptApplied Then
            batchReport = "Exact previously saved event independently reconciled."
            GoTo CreateSaved
        End If
    End If
    If interactiveContext <> "" Then
        If Not modInventoryCommands.BeginCreate(interactiveContext, payloadJson, report) Then Exit Function
    End If
    If Not receiptFound Then
    If Not modInventoryCreateWire.QueueCreate(warehouseId, stationId, userId, payloadJson, _
                                             "Admin add inventory item " & sku, eventIdOut, queueError, observedEvent) Then
        If interactiveContext <> "" Then modInventoryCommands.Release interactiveContext
        report = "Inventory item event could not be queued: " & queueError & vbCrLf & _
                 "Use Users & Roles to grant ADMIN_MAINT to '" & userId & "' for " & warehouseId & " / " & stationId & "."
        Exit Function
    End If
    End If
    If stagingRow > 0 Then
        If Not modAdminInventoryWorksheet.StoreQueuedCreateReceipt(interactiveContext, stagingRow, sku, _
                eventIdOut, systemKey, report) Then GoTo Unconfirmed
    End If

    If interactiveContext <> "" Then
        If Not modInventoryCommands.BeginCreateSave(interactiveContext, report) Then GoTo Unconfirmed
    End If
    processedCount = modProcessor.RunBatch(warehouseId, 0, batchReport, eventIdOut, applied, True)
    If Not applied Then
        report = "Inventory item event queued; exact saved application is unconfirmed." & vbCrLf & batchReport
        GoTo Unconfirmed
    End If
    If interactiveContext <> "" Then
        If Not modInventoryCommands.ConfirmCreate(interactiveContext, eventIdOut, report) Then
            report = "Partial result: exact event application was saved, but command confirmation failed." & _
                     vbCrLf & "Event: " & eventIdOut & vbCrLf & report
            Exit Function
        End If
    End If

CreateSaved:
    report = "Inventory item added. Exact event application was saved." & vbCrLf & _
             "Event: " & eventIdOut & vbCrLf & _
             "Warehouse: " & warehouseId & vbCrLf & _
             "SKU: " & sku & vbCrLf & _
             "Item: " & itemName & vbCrLf & _
             "System_Key: " & systemKey & vbCrLf & _
             "Starting quantity: " & IIf(nonCounted, "not counted", CStr(qty)) & vbCrLf & _
             "Processor: " & batchReport & vbCrLf & _
             "Refresh inventory in any open role workbook to see the new item."
    AddInventoryItemForWarehouse = True
    Exit Function

Unconfirmed:
    If interactiveContext <> "" Then modInventoryCommands.Release interactiveContext
    If eventIdOut <> "" Then report = report & vbCrLf & "Event: " & eventIdOut & ". Application is unconfirmed."
    Exit Function
FailAdd:
    report = "AddInventoryItem failed: " & Err.Description
    GoTo Unconfirmed
End Function

Private Function BuildInventoryCreatePayloadAdmin(ByVal systemKey As String, ByVal sku As String, _
        ByVal itemName As String, ByVal uom As String, ByVal locationVal As String, ByVal qty As Double, _
        ByVal description As String, ByVal vendorName As String, ByVal vendorCode As String, _
        ByVal category As String, ByVal externalCode As String, ByVal imagePath As String, ByVal customFields As Object) As String
    Dim item As Object, customKey As Variant, nonCounted As Boolean
    nonCounted = modInventoryItemFields.IsNonCountedCustomFieldsAdmin(customFields)
    Set item = CreateObject("Scripting.Dictionary"): item.CompareMode = vbTextCompare
    item("ITEM_CODE") = sku
    item("ITEM") = itemName
    item("UOM") = uom
    item("LOCATION") = locationVal
    item("TOTAL INV") = qty
    item("QtyAvailable") = qty
    item("DESCRIPTION") = description
    item("VENDOR(s)") = vendorName
    item("VENDOR_CODE") = vendorCode
    item("CATEGORY") = category
    item("CATALOG_STATE") = "ACTIVE"
    item("EXTERNAL_CODE") = externalCode
    AddPictureReferencesToPayloadAdmin item, imagePath
    If nonCounted Then
        item("TRACK_QTY") = "FALSE"
        item("ITEM_KIND") = modInventoryItemFields.NonCountedItemKindAdmin(customFields)
    End If
    If Not customFields Is Nothing Then
        For Each customKey In customFields.Keys
            If Trim$(CStr(customKey)) <> "" Then item(Trim$(CStr(customKey))) = customFields(customKey)
        Next customKey
    End If
    BuildInventoryCreatePayloadAdmin = modInventoryCreateWire.BuildPayload( _
        systemKey, sku, qty, locationVal, "GOOD", "", "Admin add inventory item", _
        item.Keys, modAdminInventoryCreate.PrimitiveFieldValues(item))
End Function

Public Function ValidateWorksheetCreateReceipt(ByVal record As Object, ByVal context As String, _
        ByVal eventId As String, ByRef report As String) As Boolean
    Dim values As Variant, found As Boolean, applied As Boolean, systemKey As String, payload As String
    Dim expected As String, category As String, vendor As String, customFields As Object
    values = modInventoryCommands.SourceValues(context)
    If Not IsArray(values) Then Exit Function
    If Not modInventoryCreateReceipts.Resolve(context, CStr(values(0)), CStr(values(1)), CStr(values(2)), _
            eventId, found, systemKey, payload, applied, report) Then Exit Function
    If Not found Then
        If InventoryItemCatalogContainsForWorksheet(CStr(values(0)), CStr(record("ItemCode"))) Then
            report = "The prepared Item Code now exists without its exact create receipt.": Exit Function
        End If
        ValidateWorksheetCreateReceipt = True: Exit Function
    End If
    category = CStr(record("Category")): vendor = CStr(record("Vendors"))
    Set customFields = record("CustomFields")
    If modInventoryItemFields.IsNonCountedCustomFieldsAdmin(customFields) Then
        If category = "" Then category = modInventoryItemFields.NonCountedItemKindAdmin(customFields)
        If vendor = "" And modInventoryItemFields.NonCountedItemKindAdmin(customFields) = "UTILITY" Then vendor = "Utility"
    End If
    expected = BuildInventoryCreatePayloadAdmin(systemKey, CStr(record("ItemCode")), CStr(record("ItemName")), _
        CStr(record("UOM")), CStr(record("Location")), CDbl(record("Quantity")), CStr(record("Description")), _
        vendor, CStr(record("VendorCode")), category, CStr(record("ExternalCode")), CStr(record("Picture")), customFields)
    ValidateWorksheetCreateReceipt = modInventoryCreateReceipts.PayloadMatches(expected, payload)
    If Not ValidateWorksheetCreateReceipt Then report = "The pending row differs from its exact saved create receipt."
End Function

Public Function RetireInventoryItemForWarehouse(ByVal warehouseId As String, _
        ByVal stationId As String, ByVal userId As String, ByVal sku As String, _
        ByVal itemName As String, ByVal reason As String, Optional ByRef report As String = "", _
        Optional ByVal interactiveContext As String = "") As Boolean
    Dim path As String, wb As Workbook, openedHere As Boolean, retiredCount As Long
    Dim eventId As String, retirementReport As String, catalogReport As String
    On Error GoTo Failed
    warehouseId = Trim$(warehouseId): stationId = Trim$(stationId): userId = Trim$(userId)
    sku = Trim$(sku): reason = Trim$(reason)
    If warehouseId = "" Or userId = "" Or sku = "" Or reason = "" Then
        report = "Warehouse, provisioned user, exact Item Code and deletion reason are required.": Exit Function
    End If
    If stationId = "" Then stationId = modStationIdentity.CurrentComputerStationId()
    If Not modInventoryRetirementCommands.RetireActiveEntities(warehouseId, stationId, userId, _
            sku, reason, retiredCount, eventId, retirementReport, interactiveContext) Then
        report = retirementReport: Exit Function
    End If
    If interactiveContext <> "" Then
        If Not modInventoryCommands.CheckForm(interactiveContext, report) Then GoTo Partial
    Else
        If Not modAuth.LoadAuth(warehouseId) Then GoTo Partial
        If Not modAuth.HasProvisionedCapabilityForSystem("ADMIN_MAINT", userId, warehouseId, stationId) Then GoTo Partial
    End If
    path = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouseId)
    Set wb = FindOpenWorkbookByFullNameAdmin(path)
    If wb Is Nothing Then
        Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=False, AddToMru:=False)
        openedHere = True
    End If
    If wb.ReadOnly Or Not wb.Saved Then report = "Inventory workbook is read-only or has unsaved changes.": GoTo Partial
    If Not MarkInventoryCatalogRetiredAdmin(wb, sku, userId, reason, catalogReport, interactiveContext) Then
        report = catalogReport: GoTo Partial
    End If
    report = "Inventory item deleted from managed inventory." & vbCrLf & _
        "Exact entities retired: " & CStr(retiredCount) & vbCrLf & "Catalog and event history were retained."
    If eventId <> "" Then report = report & vbCrLf & "Event: " & eventId
    RetireInventoryItemForWarehouse = True
    GoTo CleanExit
Partial:
    If eventId <> "" Then report = "Partial result: exact retirement event application was saved; catalog confirmation is incomplete." & _
        vbCrLf & "Event: " & eventId & vbCrLf & report
CleanExit:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function
Failed:
    report = "Inventory retirement failed: " & Err.Description
    GoTo Partial
End Function

Private Function BuildInventoryRetirementPayloadItemAdmin(ByVal systemKey As String, _
                                                          ByVal sku As String, _
                                                          ByVal qtyOnHand As Double, _
                                                          ByVal locationValue As String, _
                                                          ByVal conditionValue As String, _
                                                          ByVal attributesJson As String, _
                                                          ByVal reason As String) As Object
    Dim item As Object

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("System_Key") = systemKey
    item("SKU") = sku
    item("ITEM_CODE") = sku
    item("Qty") = IIf(qtyOnHand > 0, -qtyOnHand, 0)
    item("Location") = locationValue
    item("Condition") = conditionValue
    item("AttributesJson") = attributesJson
    item("IoType") = "RETIRE"
    item("InventoryState") = "RETIRED"
    item("Reason") = reason
    item("Note") = "Admin delete inventory item. Reason: " & reason
    Set BuildInventoryRetirementPayloadItemAdmin = item
End Function

Private Function MarkInventoryCatalogRetiredAdmin(ByVal wb As Workbook, ByVal sku As String, _
        ByVal userId As String, ByVal reason As String, ByRef report As String, _
        Optional ByVal interactiveContext As String = "") As Boolean
    Dim lo As ListObject, rowIndex As Long, stamp As String, fields As Object, field As Variant
    Dim history As String, attempted As Boolean
    On Error GoTo Failed
    If wb.ReadOnly Or Not wb.Saved Then report = "Catalog source is read-only or has unsaved changes.": Exit Function
    Set lo = FindListObjectByNameAdminLocal(wb, "tblSkuCatalog")
    If lo Is Nothing Then report = "tblSkuCatalog was not found.": Exit Function
    rowIndex = FindCatalogRowBySkuAdmin(lo, sku)
    If rowIndex = 0 Then report = "The captured catalog selection is unavailable.": Exit Function
    stamp = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Set fields = CreateObject("Scripting.Dictionary")
    fields.Add "CATALOG_STATE", "RETIRED": fields.Add "RETIRED_AT", stamp
    fields.Add "RETIRED_BY", userId: fields.Add "RETIRE_REASON", reason
    history = CatalogCellAdmin(lo, rowIndex, "EDIT_HISTORY")
    If history <> "" Then history = history & vbLf
    fields.Add "EDIT_HISTORY", history & stamp & " | User=" & userId & " | SKU=" & sku & " | Delete reason=" & reason
    If interactiveContext <> "" Then
        If Not modInventoryCommands.BeginCatalog(interactiveContext, sku, report) Then Exit Function
        If Not modInventoryCommands.BindCatalogWriter(interactiveContext, wb.Name, report) Then Exit Function
        For Each field In fields.Keys
            If Not modInventoryCommands.PlanCatalogField(interactiveContext, sku, CStr(field), CStr(fields(field)), report) Then Exit Function
        Next field
        If Not modInventoryCommands.CheckForm(interactiveContext, report) Then Exit Function
    End If
    WriteInventoryCatalogFieldsAdmin lo, rowIndex, fields
    If interactiveContext <> "" Then
        If Not modInventoryCommands.BeginCatalogSave(interactiveContext, report) Then Exit Function
    End If
    attempted = True: wb.Save
    If Not wb.Saved Then GoTo Unconfirmed
    If interactiveContext <> "" Then
        If Not modInventoryCommands.ConfirmCatalog(interactiveContext, report) Then GoTo Unconfirmed
    End If
    report = "Catalog item retained with CATALOG_STATE=RETIRED."
    MarkInventoryCatalogRetiredAdmin = True
    Exit Function
Unconfirmed:
    report = "Partial result: catalog save returned; persisted retirement fields are unconfirmed." & vbCrLf & report
    Exit Function
Failed:
    report = "Catalog retirement failed: " & Err.Description
    If attempted Then report = "Partial result: catalog save was attempted; persistence is unconfirmed." & vbCrLf & report
End Function

Private Function AddInventoryQuantityForWarehouse(ByVal warehouseId As String, _
                                                  ByVal stationId As String, _
                                                  ByVal userId As String, _
                                                  ByVal sku As String, _
                                                  ByVal itemName As String, _
                                                  ByVal uom As String, _
                                                  ByVal locationVal As String, _
                                                  ByVal qty As Double, _
                                                  ByRef report As String) As Boolean
    If qty <= 0 Then report = "Added quantity must be positive.": Exit Function
    AddInventoryQuantityForWarehouse = modInventoryQuantityCommands.SetSingleEntityQuantity( _
        warehouseId, stationId, userId, sku, ResolveInventoryQtyOnHandAdmin(warehouseId, sku) + qty, _
        "Admin add inventory quantity", report)
End Function

Private Function SetInventoryQuantityForWarehouse(ByVal warehouseId As String, _
                                                  ByVal stationId As String, _
                                                  ByVal userId As String, _
                                                  ByVal sku As String, _
                                                  ByVal itemName As String, _
                                                  ByVal uom As String, _
                                                  ByVal locationVal As String, _
                                                  ByVal targetQty As Double, _
                                                  ByVal editReason As String, _
                                                  ByRef report As String, _
                                                  Optional ByVal interactiveContext As String = "") As Boolean
    If targetQty < 0 Then report = "Set qty cannot be negative.": Exit Function
    If Trim$(editReason) = "" Then report = "Why the edit is required before changing quantity.": Exit Function
    If interactiveContext <> "" Then
        If Not modInventoryCommands.CheckForm(interactiveContext, report) Then Exit Function
    End If
    If targetQty > 0 And Not InventorySkuHasManagedEntityAdmin(warehouseId, sku) Then
        SetInventoryQuantityForWarehouse = CreateFirstInventoryEntityForCatalogItemAdmin( _
            warehouseId, stationId, userId, sku, itemName, uom, _
            locationVal, targetQty, editReason, report, interactiveContext)
    Else
        SetInventoryQuantityForWarehouse = modInventoryQuantityCommands.SetSingleEntityQuantity( _
            warehouseId, stationId, userId, sku, targetQty, editReason, report, interactiveContext)
    End If
End Function

Private Function CreateFirstInventoryEntityForCatalogItemAdmin(ByVal warehouseId As String, _
                                                                ByVal stationId As String, _
                                                                ByVal userId As String, _
                                                                ByVal sku As String, _
                                                                ByVal itemName As String, _
                                                                ByVal uom As String, _
                                                                ByVal locationVal As String, _
                                                                ByVal targetQty As Double, _
                                                                ByVal editReason As String, _
                                                                ByRef report As String, _
                                                                Optional ByVal interactiveContext As String = "") As Boolean
    CreateFirstInventoryEntityForCatalogItemAdmin = modAdminInventoryCreate.CompleteCatalogItem( _
        warehouseId, stationId, userId, sku, itemName, uom, locationVal, targetQty, _
        editReason, report, interactiveContext)
End Function

Public Function UpdateInventoryItemCatalogForWarehouse(ByVal warehouseId As String, _
                                                       ByVal sku As String, _
                                                       ByVal itemName As String, _
                                                       ByVal uom As String, _
                                                       ByVal locationVal As String, _
                                                       Optional ByVal description As String = "", _
                                                       Optional ByVal vendorName As String = "", _
                                                       Optional ByVal vendorCode As String = "", _
                                                       Optional ByVal category As String = "", _
                                                       Optional ByVal externalCode As String = "", _
                                                       Optional ByVal imagePath As String = "", _
                                                       Optional ByVal customFields As Object = Nothing, _
                                                       Optional ByRef report As String = "", _
                                                       Optional ByVal interactiveContext As String = "", _
                                                       Optional ByVal headlessUser As String = "", _
                                                       Optional ByVal headlessStation As String = "") As Boolean
    Dim path As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim lo As ListObject
    Dim rowIndex As Long
    Dim fields As Object, field As Variant, saveAttempted As Boolean, savedToDisk As Boolean

    On Error GoTo FailUpdate

    warehouseId = Trim$(warehouseId)
    sku = Trim$(sku)
    itemName = Trim$(itemName)
    uom = Trim$(uom)

    If warehouseId = "" Then report = "WarehouseId is required.": Exit Function
    If sku = "" Then report = "SKU is required.": Exit Function
    If itemName = "" Then report = "Item name is required.": Exit Function
    If uom = "" Then report = "UOM is required.": Exit Function

    If interactiveContext = "" Then
        If Trim$(headlessUser) = "" Or Trim$(headlessStation) = "" Then report = "An explicit provisioned user and station are required.": Exit Function
        If Not modAuth.LoadAuth(warehouseId) Then report = "Provisioned authority could not be loaded.": Exit Function
        If Not modAuth.HasProvisionedCapabilityForSystem("ADMIN_MAINT", headlessUser, warehouseId, headlessStation) Then report = "The provisioned user lacks ADMIN_MAINT for this warehouse and station.": Exit Function
    End If

    path = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouseId)
    If path = "" Then report = "Inventory workbook path could not be resolved for " & warehouseId & ".": Exit Function

    Set wb = FindOpenWorkbookByFullNameAdmin(path)
    If Not wb Is Nothing Then
        If Not wb.Saved Then
            report = "Inventory workbook has unsaved changes. Save or close it before editing the catalog."
            Exit Function
        End If
    Else
        If Len(Dir$(path, vbNormal)) = 0 Then
            report = "Inventory workbook was not found: " & path
            Exit Function
        End If
        Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=False, AddToMru:=False)
        openedHere = True
    End If
    If wb.ReadOnly Then
        report = "Inventory workbook is open read-only. Close other copies and try again."
        GoTo CleanExit
    End If

    Set lo = FindListObjectByNameAdminLocal(wb, "tblSkuCatalog")
    If lo Is Nothing Then
        report = "tblSkuCatalog was not found in " & wb.Name & "."
        GoTo CleanExit
    End If
    rowIndex = FindCatalogRowBySkuAdmin(lo, sku)
    If rowIndex <= 0 Then
        report = "Inventory item was not found in catalog: " & sku
        GoTo CleanExit
    End If

    Set fields = PlanInventoryCatalogFieldsAdmin(lo, rowIndex, Array(sku, itemName, uom, _
        locationVal, description, vendorName, vendorCode, category, externalCode, imagePath), customFields)
    If interactiveContext <> "" Then
        If Not modInventoryCommands.BeginCatalog(interactiveContext, sku, report) Then GoTo CleanExit
        If Not modInventoryCommands.BindCatalogWriter(interactiveContext, wb.Name, report) Then GoTo CleanExit
        For Each field In fields.Keys
            If Not modInventoryCommands.PlanCatalogField(interactiveContext, sku, CStr(field), _
                    CStr(fields(field)), report) Then GoTo CleanExit
        Next field
        If Not modInventoryCommands.CheckForm(interactiveContext, report) Then GoTo CleanExit
    End If
    WriteInventoryCatalogFieldsAdmin lo, rowIndex, fields
    If interactiveContext <> "" Then
        If Not modInventoryCommands.BeginCatalogSave(interactiveContext, report) Then GoTo CleanExit
    End If
    saveAttempted = True
    wb.Save
    savedToDisk = wb.Saved
    If Not savedToDisk Then
        report = "Partial result: catalog save was attempted but is unconfirmed. Reopen and verify before retrying."
        GoTo CleanExit
    End If
    If interactiveContext <> "" Then
        If Not modInventoryCommands.ConfirmCatalog(interactiveContext, report) Then
            report = "Partial result: catalog save returned, but persisted fields could not be confirmed. Reopen and verify before retrying." & vbCrLf & report
            GoTo CleanExit
        End If
    End If

    report = "Inventory item updated." & vbCrLf & _
             "Warehouse: " & warehouseId & vbCrLf & _
             "SKU: " & sku & vbCrLf & _
             "Item: " & itemName & vbCrLf & _
             "Refresh inventory in any open role workbook to see the updated catalog fields."
    UpdateInventoryItemCatalogForWarehouse = True

CleanExit:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Exit Function

FailUpdate:
    report = "UpdateInventoryItemCatalog failed: " & Err.Description
    If saveAttempted Then report = "Partial result: catalog save was attempted; persisted outcome is unconfirmed. Reopen and verify before retrying." & vbCrLf & report
    Resume CleanExit
End Function

Private Function PlanInventoryCatalogFieldsAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, _
        ByVal values As Variant, ByVal customFields As Object) As Object
    Dim fields As Object, names As Variant, data As Variant, headers() As String
    Dim customNames As Variant, customValues As Variant, planned As Variant, i As Long
    customNames = Array(): customValues = Array()
    If Not customFields Is Nothing Then
        customNames = customFields.Keys: customValues = customFields.Items
    End If
    ReDim headers(0 To lo.ListColumns.Count - 1)
    For i = 1 To lo.ListColumns.Count: headers(i - 1) = CStr(lo.ListColumns(i).Name): Next i
    planned = modInventoryCatalogWire.PlanFields(values, customNames, customValues, _
        CatalogCellAdmin(lo, rowIndex, "EDIT_HISTORY"), headers)
    names = planned(0): data = planned(1)
    Set fields = CreateObject("Scripting.Dictionary"): fields.CompareMode = vbTextCompare
    For i = LBound(names) To UBound(names)
        fields(CStr(names(i))) = data(i)
    Next i
    Set PlanInventoryCatalogFieldsAdmin = fields
End Function

Private Sub WriteInventoryCatalogFieldsAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal fields As Object)
    Dim events As Boolean, field As Variant, number As Long, description As String
    events = Application.EnableEvents
    On Error GoTo Failed
    Application.EnableEvents = False
    SetSheetProtectionAdminLocal lo.Parent, False
    For Each field In fields.Keys
        SetCatalogValueAdmin lo, rowIndex, CStr(field), fields(field)
    Next field
    SetSheetProtectionAdminLocal lo.Parent, True
    Application.EnableEvents = events
    Exit Sub
Failed:
    number = Err.Number: description = Err.Description
    On Error Resume Next
    SetSheetProtectionAdminLocal lo.Parent, True
    Application.EnableEvents = events
    On Error GoTo 0
    Err.Raise number, , description
End Sub

Private Sub AppendCatalogEditHistoryAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal entryText As String)
    Dim existingText As String

    entryText = Trim$(entryText)
    If entryText = "" Then Exit Sub

    existingText = CatalogCellAdmin(lo, rowIndex, "EDIT_HISTORY")
    If existingText <> "" Then
        SetCatalogValueAdmin lo, rowIndex, "EDIT_HISTORY", existingText & vbLf & entryText
    Else
        SetCatalogValueAdmin lo, rowIndex, "EDIT_HISTORY", entryText
    End If
End Sub

Private Function GenerateInventorySkuAdmin(ByVal warehouseId As String, ByVal codeOrdinal As Long) As String
    Dim codePart As String
    Dim whPart As String

    whPart = UCase$(Trim$(warehouseId))
    whPart = Replace(whPart, " ", "")
    whPart = Replace(whPart, "-", "")
    If Len(whPart) > 4 Then whPart = Left$(whPart, 4)
    If whPart = "" Then whPart = "INV"

    codePart = Base36Admin(codeOrdinal)
    Do While Len(codePart) < 6
        codePart = "0" & codePart
    Loop
    GenerateInventorySkuAdmin = "ITM-" & whPart & "-" & codePart
End Function

Private Function LoadInventoryCatalogItemsAdmin(ByVal warehouseId As String) As Collection
    Dim path As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim lo As ListObject
    Dim loEntities As ListObject
    Dim rowIndex As Long
    Dim item As Object
    Dim result As Collection

    On Error GoTo CleanExit
    Set result = New Collection
    path = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouseId)
    If path <> "" Then
        Set wb = FindOpenWorkbookByFullNameAdmin(path)
        If wb Is Nothing Then
            If Len(Dir$(path, vbNormal)) > 0 Then
                Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)
                openedHere = True
            End If
        End If
    End If

    If Not wb Is Nothing Then
        Set lo = FindListObjectByNameAdminLocal(wb, "tblSkuCatalog")
        Set loEntities = FindListObjectByNameAdminLocal(wb, "tblInventoryEntities")
        If Not lo Is Nothing Then
            If Not lo.DataBodyRange Is Nothing Then
                For rowIndex = 1 To lo.ListRows.Count
                    If CatalogCellAdmin(lo, rowIndex, "SKU") <> "" And _
                            StrComp(CatalogCellAdmin(lo, rowIndex, "CATALOG_STATE"), _
                                    "RETIRED", vbTextCompare) <> 0 Then
                        Set item = CreateObject("Scripting.Dictionary")
                        item.CompareMode = vbTextCompare
                        item("SKU") = CatalogCellAdmin(lo, rowIndex, "SKU")
                        item("ITEM") = CatalogCellAdmin(lo, rowIndex, "ITEM")
                        item("UOM") = CatalogCellAdmin(lo, rowIndex, "UOM")
                        item("LOCATION") = CatalogCellAdmin(lo, rowIndex, "LOCATION")
                        item("DESCRIPTION") = CatalogCellAdmin(lo, rowIndex, "DESCRIPTION")
                        item("VENDOR(s)") = CatalogCellAdmin(lo, rowIndex, "VENDOR(s)")
                        item("VENDOR_CODE") = CatalogCellAdmin(lo, rowIndex, "VENDOR_CODE")
                        item("CATEGORY") = CatalogCellAdmin(lo, rowIndex, "CATEGORY")
                        item("CATALOG_STATE") = CatalogCellAdmin(lo, rowIndex, "CATALOG_STATE")
                        item("EXTERNAL_CODE") = CatalogCellAdmin(lo, rowIndex, "EXTERNAL_CODE")
                        item("IMAGE_PATH") = CombinedPictureReferencesAdmin(lo, rowIndex)
                        item("TRACK_QTY") = CatalogCellAdmin(lo, rowIndex, "TRACK_QTY")
                        item("ITEM_KIND") = CatalogCellAdmin(lo, rowIndex, "ITEM_KIND")
                        item("QTY_ON_HAND") = ResolveQtyOnHandFromEntitiesAdmin(loEntities, item("SKU"))
                        result.Add item
                    End If
                Next rowIndex
            End If
        End If
    End If

CleanExit:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Set LoadInventoryCatalogItemsAdmin = result
End Function

Private Sub LoadCatalogItemsIntoAddInventoryForm(ByVal targetForm As frmAddInventoryItem, ByVal catalogItems As Object)
    Dim item As Variant

    If targetForm Is Nothing Then Exit Sub
    If catalogItems Is Nothing Then Exit Sub
    For Each item In catalogItems
        targetForm.AddCatalogItem modInventoryUiText.CatalogItemTextAdmin(item, "SKU"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "ITEM"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "UOM"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "LOCATION"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "DESCRIPTION"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "VENDOR(s)"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "VENDOR_CODE"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "CATEGORY"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "EXTERNAL_CODE"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "IMAGE_PATH"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "TRACK_QTY"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "ITEM_KIND"), _
                                  modInventoryUiText.CatalogItemTextAdmin(item, "QTY_ON_HAND")
    Next item
End Sub

Private Function ResolveInventoryQtyOnHandAdmin(ByVal warehouseId As String, ByVal sku As String) As Double
    Dim path As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim loEntities As ListObject

    On Error GoTo CleanExit
    warehouseId = Trim$(warehouseId)
    sku = Trim$(sku)
    If warehouseId = "" Or sku = "" Then Exit Function

    path = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouseId)
    If path = "" Then Exit Function
    Set wb = FindOpenWorkbookByFullNameAdmin(path)
    If wb Is Nothing Then
        If Len(Dir$(path, vbNormal)) = 0 Then Exit Function
        Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)
        openedHere = True
    End If
    Set loEntities = FindListObjectByNameAdminLocal(wb, "tblInventoryEntities")
    ResolveInventoryQtyOnHandAdmin = ResolveQtyOnHandFromEntitiesAdmin(loEntities, sku)

CleanExit:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Private Function InventorySkuHasManagedEntityAdmin(ByVal warehouseId As String, _
                                                   ByVal sku As String) As Boolean
    Dim path As String
    Dim wb As Workbook
    Dim openedHere As Boolean
    Dim loEntities As ListObject
    Dim cSystemKey As Long
    Dim cSku As Long
    Dim rowIndex As Long

    On Error GoTo CleanExit
    warehouseId = Trim$(warehouseId)
    sku = Trim$(sku)
    If warehouseId = "" Or sku = "" Then Exit Function

    path = modProcessor.ResolveInventoryWorkbookPathForAutomation(warehouseId)
    If path = "" Then Exit Function
    Set wb = FindOpenWorkbookByFullNameAdmin(path)
    If wb Is Nothing Then
        If Len(Dir$(path, vbNormal)) = 0 Then Exit Function
        Set wb = Application.Workbooks.Open(path, UpdateLinks:=False, ReadOnly:=True, AddToMru:=False)
        openedHere = True
    End If
    Set loEntities = FindListObjectByNameAdminLocal(wb, "tblInventoryEntities")
    If loEntities Is Nothing Or loEntities.DataBodyRange Is Nothing Then GoTo CleanExit
    cSystemKey = ColumnIndexAdminLocal(loEntities, "System_Key")
    cSku = ColumnIndexAdminLocal(loEntities, "SKU")
    If cSystemKey = 0 Or cSku = 0 Then GoTo CleanExit
    For rowIndex = 1 To loEntities.ListRows.Count
        If Trim$(CStr(loEntities.DataBodyRange.Cells(rowIndex, cSystemKey).Value)) <> "" And _
           StrComp(Trim$(CStr(loEntities.DataBodyRange.Cells(rowIndex, cSku).Value)), _
                   sku, vbTextCompare) = 0 Then
            InventorySkuHasManagedEntityAdmin = True
            Exit For
        End If
    Next rowIndex

CleanExit:
    On Error Resume Next
    If openedHere And Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
End Function

Private Function ResolveQtyOnHandFromEntitiesAdmin(ByVal loEntities As ListObject, _
                                                   ByVal sku As String) As Double
    Dim cSystemKey As Long
    Dim cSku As Long
    Dim cQty As Long
    Dim rowIndex As Long

    If loEntities Is Nothing Or loEntities.DataBodyRange Is Nothing Then Exit Function
    sku = Trim$(sku)
    If sku = "" Then Exit Function
    cSystemKey = ColumnIndexAdminLocal(loEntities, "System_Key")
    cSku = ColumnIndexAdminLocal(loEntities, "SKU")
    cQty = ColumnIndexAdminLocal(loEntities, "QtyOnHand")
    If cSystemKey = 0 Or cSku = 0 Or cQty = 0 Then Exit Function
    For rowIndex = 1 To loEntities.ListRows.Count
        If Trim$(CStr(loEntities.DataBodyRange.Cells(rowIndex, cSystemKey).Value)) <> "" And _
           StrComp(Trim$(CStr(loEntities.DataBodyRange.Cells(rowIndex, cSku).Value)), _
                   sku, vbTextCompare) = 0 Then
            ResolveQtyOnHandFromEntitiesAdmin = ResolveQtyOnHandFromEntitiesAdmin + _
                CDbl(Val(CStr(loEntities.DataBodyRange.Cells(rowIndex, cQty).Value)))
        End If
    Next rowIndex
End Function

Private Function ResolveQtyOnHandFromBalanceAdmin(ByVal loBalance As ListObject, ByVal sku As String) As Double
    Dim cSku As Long
    Dim cQty As Long
    Dim rowIndex As Long

    If loBalance Is Nothing Then Exit Function
    If loBalance.DataBodyRange Is Nothing Then Exit Function
    sku = Trim$(sku)
    If sku = "" Then Exit Function
    cSku = ColumnIndexAdminLocal(loBalance, "SKU")
    cQty = ColumnIndexAdminLocal(loBalance, "QtyOnHand")
    If cSku = 0 Or cQty = 0 Then Exit Function
    For rowIndex = 1 To loBalance.ListRows.Count
        If StrComp(Trim$(CStr(loBalance.DataBodyRange.Cells(rowIndex, cSku).Value)), sku, vbTextCompare) = 0 Then
            ResolveQtyOnHandFromBalanceAdmin = CDbl(Val(CStr(loBalance.DataBodyRange.Cells(rowIndex, cQty).Value)))
            Exit Function
        End If
    Next rowIndex
End Function

Private Function Base36Admin(ByVal valueIn As Long) As String
    Const DIGITS As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim n As Long
    Dim result As String

    If valueIn <= 0 Then
        Base36Admin = "0"
        Exit Function
    End If
    n = valueIn
    Do While n > 0
        result = Mid$(DIGITS, (n Mod 36) + 1, 1) & result
        n = n \ 36
    Loop
    Base36Admin = result
End Function

Private Function ResolveSeedInventoryContext(ByRef warehouseId As String, _
                                             ByRef stationId As String, _
                                             ByRef userId As String, _
                                             ByRef actionName As String, _
                                             ByRef uploadPath As String, _
                                             ByRef runtimeRoot As String, _
                                             ByRef report As String) As Boolean
    Dim warehouseOptions As Collection
    Dim formReport As String
    Dim item As Variant
    Dim selectionFound As Boolean

    warehouseId = Trim$(modConfig.GetWarehouseId())
    stationId = Trim$(modConfig.GetStationId())
    If warehouseId = "" Then warehouseId = Trim$(modConfig.GetString("WarehouseId", ""))
    If stationId = "" Then stationId = Trim$(modConfig.GetString("StationId", modStationIdentity.CurrentComputerStationId()))

    userId = Trim$(modRoleEventWriter.ResolveCurrentUserId())
    If userId = "" Then userId = Trim$(Application.UserName)

    Set warehouseOptions = modAdminConsole.GetWarehouseDirectoryOptions(Nothing, formReport, True)
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

    If mSeedCallbackAutomationEnabled Then
        warehouseId = mSeedCallbackAutomationWarehouseId
        stationId = mSeedCallbackAutomationStationId
        userId = mSeedCallbackAutomationUserId
        actionName = UCase$(Trim$(mSeedCallbackAutomationAction))
        uploadPath = Trim$(mSeedCallbackAutomationUploadPath)
        If actionName = "" Then actionName = modAdminInventorySeed.DEMO_ACTION_SEED
        If stationId = "" Then stationId = modStationIdentity.CurrentComputerStationId()
        For Each item In warehouseOptions
            If StrComp(Trim$(CStr(item(1))), warehouseId, vbTextCompare) = 0 _
               And StrComp(Trim$(CStr(item(2))), stationId, vbTextCompare) = 0 Then
                runtimeRoot = Trim$(CStr(item(3)))
                selectionFound = True
                Exit For
            End If
        Next item
        If Not selectionFound Then
            report = "The selected warehouse/station is not available in the current target."
            Exit Function
        End If
    Else
        frmSeedInventory.Configure warehouseOptions, warehouseId, stationId, userId
        frmSeedInventory.Show
        If Not frmSeedInventory.Accepted Then
            report = ""
            Unload frmSeedInventory
            Exit Function
        End If

        warehouseId = Trim$(frmSeedInventory.SelectedWarehouseId)
        stationId = Trim$(frmSeedInventory.SelectedStationId)
        runtimeRoot = Trim$(frmSeedInventory.SelectedRuntimeRoot)
        userId = Trim$(frmSeedInventory.SelectedUserId)
        actionName = UCase$(Trim$(frmSeedInventory.SelectedAction))
        uploadPath = Trim$(frmSeedInventory.SelectedUploadPath)
        Unload frmSeedInventory
    End If

    If warehouseId = "" Then
        report = "WarehouseId is required."
        Exit Function
    End If
    If stationId = "" Then stationId = modStationIdentity.CurrentComputerStationId()
    If userId = "" Then
        report = "Admin user is required."
        Exit Function
    End If
    Select Case actionName
        Case modAdminInventorySeed.DEMO_ACTION_SEED, _
             modAdminInventorySeed.DEMO_ACTION_DELETE, _
             modAdminInventorySeed.DEMO_ACTION_UPLOAD, _
             modAdminInventorySeed.DEMO_ACTION_UPLOAD_DATA_SET, _
             modAdminInventorySeed.DEMO_ACTION_DELETE_DATA_SET
        Case Else
            report = "Choose a demo inventory action."
            Exit Function
    End Select
    If runtimeRoot <> "" Then modRuntimeWorkbooks.SetCoreDataRootOverride runtimeRoot

    If Not modConfig.LoadConfig(warehouseId, stationId) Then
        report = "Config load failed: " & modConfig.Validate()
        Exit Function
    End If

    ResolveSeedInventoryContext = True
End Function

Private Function PromptDemoInventoryCsvPath() As String
    Dim selectedPath As Variant

    selectedPath = Application.GetOpenFilename( _
        FileFilter:="CSV files (*.csv),*.csv", _
        Title:="Upload Demo Inventory CSV")
    If VarType(selectedPath) = vbBoolean Then Exit Function
    PromptDemoInventoryCsvPath = Trim$(CStr(selectedPath))
End Function

Private Function SanitizeSeedCallbackErrorText(ByVal valueText As String) As String
    valueText = Replace$(Replace$(Trim$(valueText), vbCr, " "), vbLf, " ")
    Do While InStr(1, valueText, "  ", vbBinaryCompare) > 0
        valueText = Replace$(valueText, "  ", " ")
    Loop
    If InStr(1, valueText, "\", vbBinaryCompare) > 0 Then
        valueText = "<redacted-path>"
    ElseIf Len(valueText) > 240 Then
        valueText = Left$(valueText, 240)
    End If
    If valueText = "" Then valueText = "<none>"
    SanitizeSeedCallbackErrorText = valueText
End Function

Private Function NextInventoryCodeOrdinalAdmin(ByVal warehouseId As String) As Long
    ' Preserve the existing greenfield Item Code seed; it is not entity identity.
    NextInventoryCodeOrdinalAdmin = CLng((DateDiff("s", DateSerial(2026, 1, 1), Now) Mod 900000) + 10000)
End Function

Private Function FindOpenWorkbookByFullNameAdmin(ByVal fullName As String) As Workbook
    Dim wb As Workbook

    fullName = LCase$(Trim$(fullName))
    If fullName = "" Then Exit Function
    For Each wb In Application.Workbooks
        If LCase$(Trim$(wb.FullName)) = fullName Then
            Set FindOpenWorkbookByFullNameAdmin = wb
            Exit Function
        End If
    Next wb
End Function

Private Function FindListObjectByNameAdminLocal(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        On Error Resume Next
        Set FindListObjectByNameAdminLocal = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not FindListObjectByNameAdminLocal Is Nothing Then Exit Function
    Next ws
End Function

Private Function ColumnIndexAdminLocal(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim i As Long

    If lo Is Nothing Then Exit Function
    For i = 1 To lo.ListColumns.Count
        If StrComp(lo.ListColumns(i).Name, columnName, vbTextCompare) = 0 Then
            ColumnIndexAdminLocal = i
            Exit Function
        End If
    Next i
End Function

Private Function FindCatalogRowBySkuAdmin(ByVal lo As ListObject, ByVal sku As String) As Long
    Dim cSku As Long
    Dim rowIndex As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    sku = Trim$(sku)
    If sku = "" Then Exit Function

    cSku = ColumnIndexAdminLocal(lo, "SKU")
    If cSku = 0 Then cSku = ColumnIndexAdminLocal(lo, "ITEM_CODE")
    If cSku = 0 Then Exit Function

    For rowIndex = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, cSku).Value)), sku, vbTextCompare) = 0 Then
            FindCatalogRowBySkuAdmin = rowIndex
            Exit Function
        End If
    Next rowIndex
End Function

Private Function CatalogCellAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As String
    Dim idx As Long

    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    idx = ColumnIndexAdminLocal(lo, columnName)
    If idx = 0 Then Exit Function
    CatalogCellAdmin = Trim$(CStr(lo.DataBodyRange.Cells(rowIndex, idx).Value))
End Function

Private Sub SetCatalogValueAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String, ByVal valueIn As Variant)
    Dim idx As Long

    If lo Is Nothing Then Exit Sub
    If rowIndex <= 0 Then Exit Sub
    columnName = NormalizeCatalogColumnNameAdmin(columnName)
    If columnName = "" Then Exit Sub
    idx = EnsureCatalogColumnAdmin(lo, columnName)
    If idx = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, idx).Value = valueIn
End Sub

Private Function EnsureCatalogColumnAdmin(ByVal lo As ListObject, ByVal columnName As String) As Long
    If lo Is Nothing Then Exit Function
    columnName = NormalizeCatalogColumnNameAdmin(columnName)
    If columnName = "" Then Exit Function
    EnsureCatalogColumnAdmin = ColumnIndexAdminLocal(lo, columnName)
    If EnsureCatalogColumnAdmin > 0 Then Exit Function
    EnsureCatalogColumnAdmin = lo.ListColumns.Add.Index
    lo.ListColumns(EnsureCatalogColumnAdmin).Name = columnName
End Function

Private Function NormalizeCatalogColumnNameAdmin(ByVal columnName As String) As String
    NormalizeCatalogColumnNameAdmin = modInventoryCatalogWire.NormalizeHeader(columnName)
End Function

Private Sub SetSheetProtectionAdminLocal(ByVal ws As Worksheet, ByVal protectAfter As Boolean)
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If protectAfter Then
        ws.Protect UserInterfaceOnly:=True
    ElseIf ws.ProtectContents Then
        ws.Unprotect
    End If
    On Error GoTo 0
End Sub

Private Sub AddPictureReferencesToPayloadAdmin(ByVal item As Object, ByVal imagePath As String)
    Dim refs As Collection
    Dim i As Long

    If item Is Nothing Then Exit Sub
    Set refs = ParsePictureReferencesAdmin(imagePath)
    If refs.Count = 0 Then
        item("IMAGE_PATH") = ""
        Exit Sub
    End If
    item("IMAGE_PATH") = refs(1)
    For i = 2 To refs.Count
        item("IMAGE_PATH_" & CStr(i)) = refs(i)
    Next i
End Sub

Private Sub SetPictureReferencesInCatalogAdmin(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal imagePath As String)
    Dim refs As Collection
    Dim i As Long
    Dim columnName As String

    Set refs = ParsePictureReferencesAdmin(imagePath)
    If refs.Count > 0 Then
        SetCatalogValueAdmin lo, rowIndex, "IMAGE_PATH", refs(1)
    Else
        SetCatalogValueAdmin lo, rowIndex, "IMAGE_PATH", ""
    End If
    For i = 2 To 12
        columnName = "IMAGE_PATH_" & CStr(i)
        If i <= refs.Count Then
            SetCatalogValueAdmin lo, rowIndex, columnName, refs(i)
        ElseIf ColumnIndexAdminLocal(lo, columnName) > 0 Then
            SetCatalogValueAdmin lo, rowIndex, columnName, ""
        End If
    Next i
End Sub

Private Function CombinedPictureReferencesAdmin(ByVal lo As ListObject, ByVal rowIndex As Long) As String
    Dim parts As Collection
    Dim i As Long
    Dim valueText As String
    Dim result As String

    Set parts = New Collection
    valueText = CatalogCellAdmin(lo, rowIndex, "IMAGE_PATH")
    If valueText <> "" Then parts.Add valueText
    For i = 2 To 12
        valueText = CatalogCellAdmin(lo, rowIndex, "IMAGE_PATH_" & CStr(i))
        If valueText <> "" Then parts.Add valueText
    Next i
    For i = 1 To parts.Count
        If result <> "" Then result = result & "; "
        result = result & CStr(parts(i))
    Next i
    CombinedPictureReferencesAdmin = result
End Function

Private Function ParsePictureReferencesAdmin(ByVal imagePath As String) As Collection
    Dim result As New Collection, values As Variant, value As Variant
    values = modInventoryCatalogWire.PictureValues(imagePath)
    For Each value In values
        result.Add CStr(value)
    Next value
    Set ParsePictureReferencesAdmin = result
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

    If Not modRoleEventWriter.QueueInventoryCreateEvent(warehouseId, stationId, userId, payloadJson, _
                                                        "Admin demo inventory seed", _
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
             "Now refresh inventory in any open role workbook, or reopen the picker to auto-refresh."
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

    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), "DEMO-RAW-BLACK-TEA", ADMIN_DEMO_INVENTORY_QTY, _
        "NAS-A1", "GOOD", "", "Admin demo seed")
    item("ITEM_CODE") = "DEMO-RAW-BLACK-TEA"
    item("ITEM") = "Black Tea Base"
    item("UOM") = "LB"
    item("TOTAL INV") = ADMIN_DEMO_INVENTORY_QTY
    item("QtyAvailable") = ADMIN_DEMO_INVENTORY_QTY
    item("DESCRIPTION") = "Demo raw black tea for receiving tests"
    item("VENDOR(s)") = "Demo Vendor"
    item("CATEGORY") = "Raw Material"
    rows.Add item

    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), "DEMO-SPICE-CARDAMOM", ADMIN_DEMO_INVENTORY_QTY, _
        "NAS-A2", "GOOD", "", "Admin demo seed")
    item("ITEM_CODE") = "DEMO-SPICE-CARDAMOM"
    item("ITEM") = "Cardamom Pods"
    item("UOM") = "LB"
    item("TOTAL INV") = ADMIN_DEMO_INVENTORY_QTY
    item("QtyAvailable") = ADMIN_DEMO_INVENTORY_QTY
    item("DESCRIPTION") = "Demo spice inventory for receiving tests"
    item("VENDOR(s)") = "Demo Vendor"
    item("CATEGORY") = "Spice"
    rows.Add item

    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
        modRoleEventWriter.CreateSystemKey(), "DEMO-PKG-TIN", ADMIN_DEMO_INVENTORY_QTY, _
        "NAS-P1", "GOOD", "", "Admin demo seed")
    item("ITEM_CODE") = "DEMO-PKG-TIN"
    item("ITEM") = "Retail Tea Tin"
    item("UOM") = "EA"
    item("TOTAL INV") = ADMIN_DEMO_INVENTORY_QTY
    item("QtyAvailable") = ADMIN_DEMO_INVENTORY_QTY
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

        uom = CsvFieldAdmin(fields, headers, "UOM")
        location = CsvFieldAdmin(fields, headers, "LOCATION")
        category = CsvFieldAdmin(fields, headers, "CATEGORY")
        qty = Val(CsvFieldAdmin(fields, headers, "QUANTITY"))
        If qty <= 0 Then qty = ResolveDemoSeedQuantityAdmin(category, CsvFieldAdmin(fields, headers, "PHASE"), uom)

        Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem( _
            modRoleEventWriter.CreateSystemKey(), sku, qty, location, "GOOD", "", _
            "Admin CSV demo inventory seed")
        item("ITEM_CODE") = sku
        item("ITEM") = itemName
        item("UOM") = uom
        item("LOCATION") = location
        item("TOTAL INV") = qty
        item("QtyAvailable") = qty
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

Private Function NormalizeAdminDemoInventoryRow(ByVal rowVal As Long, ByVal sku As String, ByVal fallbackRow As Long) As Long
    Dim demoRow As Long

    demoRow = AdminDemoInventoryRowForSku(sku)
    If demoRow > 0 Then
        NormalizeAdminDemoInventoryRow = demoRow
    ElseIf rowVal > 0 Then
        NormalizeAdminDemoInventoryRow = rowVal
    Else
        NormalizeAdminDemoInventoryRow = fallbackRow
    End If
End Function

Private Function AdminDemoInventoryRowForSku(ByVal sku As String) As Long
    Select Case UCase$(Trim$(sku))
        Case "DEMO-RAW-BLACK-TEA"
            AdminDemoInventoryRowForSku = 9001
        Case "DEMO-RAW-FILTERED-WATER"
            AdminDemoInventoryRowForSku = 9002
        Case "DEMO-RAW-CARDAMOM", "DEMO-SPICE-CARDAMOM"
            AdminDemoInventoryRowForSku = 9003
        Case "DEMO-FG-CLASSIC-CHAI"
            AdminDemoInventoryRowForSku = 9016
        Case "DEMO-PKG-TIN"
            AdminDemoInventoryRowForSku = 9021
    End Select
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
    ResolveDemoSeedQuantityAdmin = ADMIN_DEMO_INVENTORY_QTY
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
    modAdminVerifyAddins.VerifyAddinsPublished
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

Public Sub Admin_AggregateGlobalSnapshot_Click()
    Dim report As String
    Dim targetWb As Workbook

    If Not modRoleUiAccess.RequireCurrentUserCapabilityCached("ADMIN_MAINT", "", report) Then Exit Sub

    Set targetWb = ResolveInteractiveAdminWorkbook()
    Call modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface(targetWb, report)
    frmAggregationSources.ObserveEntry: frmAggregationSources.Show
End Sub

' D13 packaged form initialization smoke. This is not a Ribbon command and does
' not select a source, connect to a server, or alter warehouse state.
Public Function Admin_AggregationSourcesFormSmokeForAutomation() As String
    Admin_AggregationSourcesFormSmokeForAutomation = frmAggregationSources.AggregationSourcesSmokeForAutomation
End Function

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
' Shared administrative helpers used by the active Admin ribbon actions and forms.
''''''''''''''''''''''''''''''''''''
