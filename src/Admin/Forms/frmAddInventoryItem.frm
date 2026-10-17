VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddInventoryItem
   Caption         =   "invSys Admin - Add/Edit Inventory Items"
   ClientHeight    =   6100
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6900
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddInventoryItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mBtnAddField As MSForms.CommandButton
Attribute mBtnAddField.VB_VarHelpID = -1
Private WithEvents mBtnRemoveField As MSForms.CommandButton
Attribute mBtnRemoveField.VB_VarHelpID = -1
Private WithEvents mBtnOK As MSForms.CommandButton
Attribute mBtnOK.VB_VarHelpID = -1
Private WithEvents mBtnCancel As MSForms.CommandButton
Attribute mBtnCancel.VB_VarHelpID = -1
Private WithEvents mBtnAddMode As MSForms.CommandButton
Attribute mBtnAddMode.VB_VarHelpID = -1
Private WithEvents mBtnEditMode As MSForms.CommandButton
Attribute mBtnEditMode.VB_VarHelpID = -1
Private WithEvents mBtnDeleteItem As MSForms.CommandButton
Attribute mBtnDeleteItem.VB_VarHelpID = -1
Private WithEvents mBtnCreateInventoryTable As MSForms.CommandButton
Attribute mBtnCreateInventoryTable.VB_VarHelpID = -1
Private WithEvents mBtnUploadInventoryTable As MSForms.CommandButton
Attribute mBtnUploadInventoryTable.VB_VarHelpID = -1
Private WithEvents mCmbEditItem As MSForms.ComboBox
Attribute mCmbEditItem.VB_VarHelpID = -1
Private WithEvents mLstEditItemResults As MSForms.ListBox
Attribute mLstEditItemResults.VB_VarHelpID = -1
Private WithEvents mCmbUom As MSForms.ComboBox
Attribute mCmbUom.VB_VarHelpID = -1
Private WithEvents mTxtQty As MSForms.ComboBox
Attribute mTxtQty.VB_VarHelpID = -1
Private WithEvents mTxtImagePath As MSForms.TextBox
Attribute mTxtImagePath.VB_VarHelpID = -1

Private mLblTitle As MSForms.Label
Private mLblContext As MSForms.Label
Private mLblEditItem As MSForms.Label
Private mLblItemName As MSForms.Label
Private mLblUom As MSForms.Label
Private mLblQty As MSForms.Label
Private mLblLocation As MSForms.Label
Private mLblDescription As MSForms.Label
Private mLblCategory As MSForms.Label
Private mLblVendorName As MSForms.Label
Private mLblVendorCode As MSForms.Label
Private mLblExternalCode As MSForms.Label
Private mLblImagePath As MSForms.Label
Private mLblEditReason As MSForms.Label
Private mLblCustomName As MSForms.Label
Private mLblCustomValue As MSForms.Label
Private mLblGenerated As MSForms.Label
Private mLblStatus As MSForms.Label
Private mTxtItemName As MSForms.TextBox
Private mCmbLocation As MSForms.ComboBox
Private mTxtDescription As MSForms.TextBox
Private mCmbCategory As MSForms.ComboBox
Private mTxtVendorName As MSForms.TextBox
Private mTxtVendorCode As MSForms.TextBox
Private mTxtExternalCode As MSForms.TextBox
Private mTxtEditReason As MSForms.TextBox
Private mTxtCustomName As MSForms.TextBox
Private mTxtCustomValue As MSForms.TextBox
Private mLstCustomFields As MSForms.ListBox

Private mAccepted As Boolean
Private mGeneratedSku As String
Private mWarehouseId As String
Private mStationId As String
Private mUserId As String
Private mAnchors As Object
Private mResizeInitialized As Boolean
Private mCatalogItems As Object
Private mEditMode As Boolean
Private mSelectedEditSku As String
Private mLoading As Boolean
Private mFilteringEditItems As Boolean
Private mPreviousUom As String
Private mInitStep As String
Private mAllowUomPrompt As Boolean
Private mImagePlaceholderActive As Boolean
Private mOperatorWorkbook As Workbook
Private mInteractiveContext As String
Private mLastInventoryTableName As String
Private mLastInventoryWorksheetReport As String
Private mConfiguredDefaultLocation As String
Private mDeleteRequested As Boolean
Private mDeleteReason As String
Private mDeleteAutomation As Boolean

Private Const ANCHOR_LEFT As Long = 1
Private Const ANCHOR_TOP As Long = 2
Private Const ANCHOR_RIGHT As Long = 4
Private Const ANCHOR_BOTTOM As Long = 8
Private Const ADD_UOM_OPTION As String = "+ Add UOM..."
Private Const IMAGE_PATH_PLACEHOLDER As String = "Paste picture file path(s) or URL(s); separate multiple pictures with ;"
Private Const QTY_OPTION_UTILITY As String = "Utility"
Private Const QTY_OPTION_SERVICE As String = "Service"
Private Const QTY_OPTION_NOT_COUNTED As String = "Not counted"

Public Property Get Accepted() As Boolean
    Accepted = mAccepted
End Property

Public Property Get GeneratedSku() As String
    If mEditMode And mSelectedEditSku <> "" Then
        GeneratedSku = mSelectedEditSku
    Else
        GeneratedSku = mGeneratedSku
    End If
End Property

Public Property Get EditMode() As Boolean
    EditMode = mEditMode
End Property

Public Property Get DeleteRequested() As Boolean
    DeleteRequested = mDeleteRequested
End Property

Public Property Get DeleteReason() As String
    DeleteReason = mDeleteReason
End Property

Public Property Get ItemName() As String
    ItemName = Trim$(CStr(mTxtItemName.Value))
End Property

Public Property Get Uom() As String
    Uom = Trim$(CStr(mCmbUom.Value))
End Property

Public Property Get StartingQty() As Double
    If NonCountedItem Then Exit Property
    StartingQty = CDbl(Val(CStr(mTxtQty.Value)))
End Property

Public Property Get NonCountedItem() As Boolean
    NonCountedItem = QuantityModeIsNonCounted(QuantityModeText())
End Property

Public Property Get LocationValue() As String
    LocationValue = Trim$(CStr(mCmbLocation.Value))
End Property

Public Property Get DescriptionValue() As String
    DescriptionValue = Trim$(CStr(mTxtDescription.Value))
End Property

Public Property Get Category() As String
    Category = Trim$(CStr(mCmbCategory.Value))
End Property

Public Property Get VendorName() As String
    VendorName = Trim$(CStr(mTxtVendorName.Value))
End Property

Public Property Get VendorCode() As String
    VendorCode = Trim$(CStr(mTxtVendorCode.Value))
End Property

Public Property Get ExternalCode() As String
    ExternalCode = Trim$(CStr(mTxtExternalCode.Value))
End Property

Public Property Get ImagePath() As String
    If mImagePlaceholderActive Then Exit Property
    ImagePath = Trim$(CStr(mTxtImagePath.Value))
End Property

Public Property Get EditReason() As String
    If mTxtEditReason Is Nothing Then Exit Property
    EditReason = Trim$(CStr(mTxtEditReason.Value))
End Property

Public Property Get CustomFields() As Object
    Dim result As Object
    Dim i As Long

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    If NonCountedItem Then
        result("TRACK_QTY") = "FALSE"
        result("ITEM_KIND") = NonCountedItemKind()
    Else
        result("TRACK_QTY") = "TRUE"
        result("ITEM_KIND") = "INVENTORY"
    End If
    If Not mLstCustomFields Is Nothing Then
        For i = 0 To mLstCustomFields.ListCount - 1
            If Trim$(CStr(mLstCustomFields.List(i, 0))) <> "" Then
                result(Trim$(CStr(mLstCustomFields.List(i, 0)))) = Trim$(CStr(mLstCustomFields.List(i, 1)))
            End If
        Next i
    End If
    Set CustomFields = result
End Property

Public Sub TestSetQuantityMode(ByVal quantityMode As String)
    If mTxtQty Is Nothing Then EnsureControls
    mTxtQty.Value = quantityMode
    ApplyQuantityModeState
End Sub

Public Function TestEditItemComboSelectionContract() As String
    On Error GoTo Failed

    Dim comboSelected As Boolean
    Dim fieldsLoaded As Boolean
    Dim utilityReady As Boolean
    Dim validationReady As Boolean
    Dim fields As Object
    Dim testStage As String

    testStage = "Configure"
    Configure "WH-TEST", "S1", "admin", "NEW-SKU", "CLEARVIEW"
    testStage = "AddCatalogItem"
    AddCatalogItem "DEMO-RAW-FILTERED-WATER", "Filtered Water", "LB", _
        "Clearview", "clean double filtered water", "Utility", "", "UTILITY", _
        "", "", "TRUE", "INVENTORY", "19400"
    testStage = "EditMode"
    mEditMode = True
    ApplyModeLayout

    testStage = "SelectCombo"
    mLoading = True
    mCmbEditItem.ListIndex = 0
    mLoading = False
    mCmbEditItem_Change

    testStage = "VerifySelection"
    comboSelected = (mSelectedEditSku = "DEMO-RAW-FILTERED-WATER")
    fieldsLoaded = (ItemName = "Filtered Water" And Uom = "LB" And _
                    DescriptionValue = "clean double filtered water")
    TestSetQuantityMode QTY_OPTION_UTILITY
    testStage = "CustomFields"
    Set fields = Me.CustomFields
    utilityReady = (NonCountedItem And StartingQty = 0 And _
                    CStr(fields("TRACK_QTY")) = "FALSE" And _
                    CStr(fields("ITEM_KIND")) = "UTILITY")
    testStage = "Validate"
    validationReady = ValidateForm()

    TestEditItemComboSelectionContract = IIf(comboSelected And fieldsLoaded And _
        utilityReady And validationReady, "OK", "FAIL") & _
        "|ComboSelected=" & CStr(comboSelected) & _
        "|FieldsLoaded=" & CStr(fieldsLoaded) & _
        "|UtilityReady=" & CStr(utilityReady) & _
        "|ValidationReady=" & CStr(validationReady)
    Exit Function

Failed:
    TestEditItemComboSelectionContract = "FAIL|Stage=" & testStage & _
        "|Error=" & CStr(Err.Number) & " " & Err.Description
End Function

Public Function TestAddItemVisibilityDropdownContract() As String
    On Error GoTo Failed

    Dim locationDropdown As Boolean
    Dim categoryDropdown As Boolean
    Dim submitHandler As Boolean
    Dim zeroQtyAccepted As Boolean
    Dim negativeRejected As Boolean
    Dim testStage As String

    testStage = "Configure"
    Configure "WH-TEST", "S1", "admin", "ITM-WHTE-000999", "CLEARVIEW"
    testStage = "CatalogOptions"
    AddCatalogItem "SKU-RAW", "Raw Material", "LB", "SECONDARY", _
        "", "", "", "RAW", "", "", "TRUE", "INVENTORY", "10"
    AddCatalogItem "SKU-FINISHED", "Finished Good", "EA", "CLEARVIEW", _
        "", "", "", "FINISHED", "", "", "TRUE", "INVENTORY", "5"

    locationDropdown = (TypeName(mCmbLocation) = "ComboBox" And _
        ComboContainsValue(mCmbLocation, "CLEARVIEW") And _
        ComboContainsValue(mCmbLocation, "SECONDARY"))
    categoryDropdown = (TypeName(mCmbCategory) = "ComboBox" And _
        ComboContainsValue(mCmbCategory, "RAW") And _
        ComboContainsValue(mCmbCategory, "FINISHED"))

    testStage = "SubmitHandler"
    mTxtItemName.Value = "Honey"
    mCmbUom.Value = "LB"
    mTxtQty.Value = "0"
    mCmbLocation.Value = "CLEARVIEW"
    mCmbCategory.Value = "RAW"
    mBtnOK_Click
    submitHandler = (mAccepted And ItemName = "Honey" And _
        LocationValue = "CLEARVIEW" And Category = "RAW")
    zeroQtyAccepted = (submitHandler And StartingQty = 0)

    testStage = "NegativeValidation"
    mAccepted = False
    mTxtQty.Value = "-1"
    mBtnOK_Click
    negativeRejected = (Not mAccepted And _
        InStr(1, mLblStatus.Caption, "cannot be negative", vbTextCompare) > 0)

    TestAddItemVisibilityDropdownContract = IIf(locationDropdown And categoryDropdown And _
        submitHandler And zeroQtyAccepted And negativeRejected, "OK", "FAIL") & _
        "|SubmitHandler=" & CStr(submitHandler) & _
        "|ZeroQtyAccepted=" & CStr(zeroQtyAccepted) & _
        "|NegativeRejected=" & CStr(negativeRejected) & _
        "|LocationDropdown=" & CStr(locationDropdown) & _
        "|CategoryDropdown=" & CStr(categoryDropdown)
    Exit Function

Failed:
    TestAddItemVisibilityDropdownContract = "FAIL|Stage=" & testStage & _
        "|Error=" & CStr(Err.Number) & " " & Err.Description
End Function

Public Function TestInventoryDeleteActionContract() As String
    On Error GoTo Failed

    Dim deleteVisible As Boolean
    Dim deleteEnabled As Boolean
    Dim deleteHandler As Boolean
    Dim testStage As String

    testStage = "Configure"
    Configure "WH-TEST", "S1", "admin", "ITM-WHTE-000999", "CLEARVIEW"
    AddCatalogItem "SKU-HONEY", "Honey", "LB", "CLEARVIEW", _
        "", "", "", "RAW", "", "", "TRUE", "INVENTORY", "25"
    testStage = "EditMode"
    mEditMode = True
    ApplyModeLayout
    mLoading = True
    mCmbEditItem.ListIndex = 0
    mLoading = False
    mCmbEditItem_Change
    deleteVisible = mBtnDeleteItem.Visible
    deleteEnabled = mBtnDeleteItem.Enabled

    testStage = "DeleteHandler"
    mDeleteAutomation = True
    mBtnDeleteItem_Click
    mDeleteAutomation = False
    deleteHandler = (mAccepted And mDeleteRequested And _
        mSelectedEditSku = "SKU-HONEY" And mDeleteReason = "Packaged deletion contract")

    TestInventoryDeleteActionContract = IIf(deleteVisible And deleteEnabled And _
        deleteHandler, "OK", "FAIL") & _
        "|DeleteVisible=" & CStr(deleteVisible) & _
        "|DeleteEnabled=" & CStr(deleteEnabled) & _
        "|DeleteHandler=" & CStr(deleteHandler)
    Exit Function

Failed:
    mDeleteAutomation = False
    TestInventoryDeleteActionContract = "FAIL|Stage=" & testStage & _
        "|Error=" & CStr(Err.Number) & " " & Err.Description
End Function

Public Function TestInventoryWorksheetActionContract(ByVal operatorWb As Workbook) As String
    On Error GoTo Failed

    Dim tableCreated As Boolean
    Dim rowsPrepared As Boolean
    Dim tableSelected As Boolean
    Dim capturedWorkbook As Boolean
    Dim evidence As String
    Dim testStage As String

    testStage = "Configure"
    Configure "WH-TEST", "S1", "admin", "ITM-WHTE-000999", "CLEARVIEW"
    SetOperatorWorkbook operatorWb
    capturedWorkbook = Not mOperatorWorkbook Is Nothing
    modAdminInventoryWorksheet.BeginInventoryWorksheetAutomation "EXISTING-SKU"

    testStage = "CreateHandler"
    mBtnCreateInventoryTable_Click
    tableCreated = (mLastInventoryTableName <> "" And _
        modAdminInventoryWorksheet.CountInventoryWorksheetTables(operatorWb) = 1)
    testStage = "Populate"
    rowsPrepared = modAdminInventoryWorksheet.PopulateInventoryWorksheetContractRowsForTest( _
        operatorWb, mLastInventoryTableName)
    tableSelected = modAdminInventoryWorksheet.SelectInventoryWorksheetTableForTest( _
        operatorWb, mLastInventoryTableName)
    testStage = "UploadHandler"
    mBtnUploadInventoryTable_Click
    evidence = modAdminInventoryWorksheet.LastInventoryWorksheetAutomationReport()
    modAdminInventoryWorksheet.EndInventoryWorksheetAutomation

    If tableCreated And rowsPrepared And tableSelected And capturedWorkbook And _
            InStr(1, evidence, "Preflight=True", vbTextCompare) > 0 And _
            InStr(1, evidence, "Utility=True", vbTextCompare) > 0 And _
            InStr(1, evidence, "ExactEdit=True", vbTextCompare) > 0 Then
        TestInventoryWorksheetActionContract = "OK|" & evidence & _
            "|CapturedWorkbook=True|CreateHandler=True|UploadHandler=True"
    Else
        TestInventoryWorksheetActionContract = "FAIL|TableCreated=" & CStr(tableCreated) & _
            "|RowsPrepared=" & CStr(rowsPrepared) & _
            "|TableSelected=" & CStr(tableSelected) & _
            "|CapturedWorkbook=" & CStr(capturedWorkbook) & _
            "|Evidence=" & evidence
    End If
    Exit Function

Failed:
    modAdminInventoryWorksheet.EndInventoryWorksheetAutomation
    TestInventoryWorksheetActionContract = "FAIL|Stage=" & testStage & _
        "|Error=" & CStr(Err.Number) & " " & Err.Description
End Function

Public Sub SetOperatorWorkbook(ByVal wb As Workbook)
    On Error GoTo Unavailable
    If wb Is Nothing Then Exit Sub
    If wb.Worksheets.Count < 1 Then Exit Sub
    Set mOperatorWorkbook = wb
Unavailable:
End Sub

Public Sub SetInteractiveContext(ByVal context As String)
    mInteractiveContext = context
End Sub

Private Function CurrentContextAllowed() As Boolean
    Dim report As String
    CurrentContextAllowed = modInventoryCommands.CheckForm(mInteractiveContext, report)
    If Not CurrentContextAllowed Then mLblStatus.Caption = report
End Function

Public Sub Configure(ByVal warehouseId As Variant, _
                     ByVal stationId As Variant, _
                     ByVal userId As Variant, _
                     ByVal generatedSku As Variant, _
                     ByVal defaultLocation As Variant)
    On Error GoTo FailConfigure

    mLoading = True
    mAllowUomPrompt = False
    mInitStep = "EnsureControls"
    EnsureControls
    mInitStep = "Set scalar context"
    mWarehouseId = SafeFormText(warehouseId)
    mStationId = SafeFormText(stationId)
    mUserId = SafeFormText(userId)
    mGeneratedSku = SafeFormText(generatedSku)
    mDeleteRequested = False
    mDeleteReason = ""
    mConfiguredDefaultLocation = SafeFormText(defaultLocation)
    Set mCatalogItems = New Collection
    mAccepted = False
    mEditMode = False
    mSelectedEditSku = ""

    mInitStep = "Load UOM options"
    LoadUomOptions
    mInitStep = "Load edit item options"
    LoadEditItemOptions
    mInitStep = "Reset field values"
    mTxtItemName.Value = ""
    mCmbUom.Value = "EA"
    mPreviousUom = "EA"
    mTxtQty.Value = "1"
    LoadInventoryDimensionOptions
    mCmbLocation.Value = mConfiguredDefaultLocation
    mTxtDescription.Value = ""
    mCmbCategory.Value = ""
    mTxtVendorName.Value = ""
    mTxtVendorCode.Value = ""
    mTxtExternalCode.Value = ""
    If Not mTxtEditReason Is Nothing Then mTxtEditReason.Value = ""
    ApplyQuantityModeState
    ShowImagePathPlaceholder
    mTxtCustomName.Value = ""
    mTxtCustomValue.Value = ""
    mLstCustomFields.Clear
    mInitStep = "Refresh generated label"
    RefreshGeneratedLabel
    mInitStep = "Apply mode layout"
    ApplyModeLayout
    mLblStatus.Caption = ""
    mLoading = False
    mInitStep = ""
    Exit Sub

FailConfigure:
    mLoading = False
    Err.Raise Err.Number, "frmAddInventoryItem.Configure", _
              "Add Inventory Item form failed during " & mInitStep & ": " & Err.Description
End Sub

Private Function SafeFormText(ByVal valueIn As Variant) As String
    If IsError(valueIn) Then Exit Function
    If IsNull(valueIn) Then Exit Function
    If IsEmpty(valueIn) Then Exit Function
    SafeFormText = Trim$(CStr(valueIn))
End Function

Public Sub AddCatalogItem(ByVal sku As String, _
                          ByVal itemName As String, _
                          ByVal uomValue As String, _
                          ByVal locationValue As String, _
                          ByVal descriptionValue As String, _
                          ByVal vendorName As String, _
                          ByVal vendorCode As String, _
                          ByVal categoryValue As String, _
                          ByVal externalCodeValue As String, _
                          ByVal imagePathValue As String, _
                          Optional ByVal trackQtyValue As String = "", _
                          Optional ByVal itemKindValue As String = "", _
                          Optional ByVal qtyOnHandValue As String = "")
    Dim item As Object
    Dim selectedUom As String

    EnsureControls
    If mCatalogItems Is Nothing Then Set mCatalogItems = New Collection
    sku = Trim$(sku)
    If sku = "" Then Exit Sub
    selectedUom = Trim$(CStr(mCmbUom.Value))

    Set item = CreateObject("Scripting.Dictionary")
    item.CompareMode = vbTextCompare
    item("SKU") = sku
    item("ITEM") = Trim$(itemName)
    item("UOM") = UCase$(Trim$(uomValue))
    item("LOCATION") = Trim$(locationValue)
    item("DESCRIPTION") = Trim$(descriptionValue)
    item("VENDOR(s)") = Trim$(vendorName)
    item("VENDOR_CODE") = Trim$(vendorCode)
    item("CATEGORY") = Trim$(categoryValue)
    item("EXTERNAL_CODE") = Trim$(externalCodeValue)
    item("IMAGE_PATH") = Trim$(imagePathValue)
    item("TRACK_QTY") = UCase$(Trim$(trackQtyValue))
    item("ITEM_KIND") = UCase$(Trim$(itemKindValue))
    item("QTY_ON_HAND") = Trim$(qtyOnHandValue)
    mCatalogItems.Add item

    mLoading = True
    LoadUomOptions
    LoadInventoryDimensionOptions
    If selectedUom <> "" Then
        mCmbUom.Value = selectedUom
    ElseIf Not mEditMode Then
        mCmbUom.Value = "EA"
    End If
    LoadEditItemOptions
    mLoading = False
End Sub

Private Sub UserForm_Initialize()
    mLoading = True
End Sub

Private Sub UserForm_Activate()
    mAllowUomPrompt = True
    If mAnchors Is Nothing Then InitializeAddInventoryAnchors
    If Not mResizeInitialized Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeInitialized = True
    End If
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Sub UserForm_Layout()
    If mAnchors Is Nothing Then Exit Sub
    mAnchors.ResizeControls
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        mAccepted = False
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub UserForm_Terminate()
    Set mAnchors = Nothing
    Set mOperatorWorkbook = Nothing
End Sub

Private Sub EnsureControls()
    If Not mBtnOK Is Nothing Then Exit Sub

    Me.Caption = "invSys Admin - Add/Edit Inventory Items"
    Me.Width = 575
    Me.Height = 665

    Set mBtnAddMode = AddButton("btnAddMode", 14, 10, 118, 24, "Add Item Mode")
    Set mBtnEditMode = AddButton("btnEditMode", 138, 10, 118, 24, "Edit Item")
    Set mBtnDeleteItem = AddButton("btnDeleteItem", 262, 10, 100, 24, "Delete Item")

    Set mLblTitle = AddLabel("lblTitle", 14, 44, 530, 20, "Add inventory item")
    mLblTitle.Font.Bold = True
    Set mLblContext = AddLabel("lblContext", 14, 66, 530, 18, "Fill the required fields. Internal item code is generated by invSys.")
    Set mLblGenerated = AddLabel("lblGenerated", 14, 90, 530, 32, "")
    mLblGenerated.WordWrap = True
    Set mLblEditItem = AddLabel("lblEditItem", 14, 128, 126, 18, "Inventory item")
    Set mCmbEditItem = AddCombo("cmbEditItem", 146, 124, 392, 22)
    mCmbEditItem.ColumnCount = 2
    mCmbEditItem.ColumnWidths = "360 pt;0 pt"
    mCmbEditItem.Style = fmStyleDropDownCombo
    mCmbEditItem.MatchEntry = fmMatchEntryNone
    mCmbEditItem.AutoWordSelect = False
    mCmbEditItem.ListRows = 12
    Set mLstEditItemResults = AddListBox("lstEditItemResults", 146, 148, 392, 96)
    mLstEditItemResults.ColumnCount = 2
    mLstEditItemResults.ColumnWidths = "360 pt;0 pt"
    mLstEditItemResults.Visible = False

    Set mLblItemName = AddLabel("lblItemName", 14, 158, 126, 18, "Item name *")
    Set mTxtItemName = AddTextBox("txtItemName", 146, 154, 392, 22)
    Set mLblUom = AddLabel("lblUom", 14, 190, 126, 18, "UOM *")
    Set mCmbUom = AddCombo("cmbUom", 146, 186, 120, 22)
    LoadUomOptions
    Set mLblQty = AddLabel("lblQty", 288, 190, 92, 18, "Starting qty *")
    Set mTxtQty = AddCombo("txtQty", 386, 186, 152, 22)
    LoadQuantityOptions

    Set mLblLocation = AddLabel("lblLocation", 14, 222, 126, 18, "Default location")
    Set mCmbLocation = AddCombo("cmbLocation", 146, 218, 120, 22)
    mCmbLocation.Style = fmStyleDropDownCombo
    mCmbLocation.MatchEntry = fmMatchEntryComplete
    Set mLblCategory = AddLabel("lblCategory", 288, 222, 92, 18, "Category")
    Set mCmbCategory = AddCombo("cmbCategory", 386, 218, 152, 22)
    mCmbCategory.Style = fmStyleDropDownCombo
    mCmbCategory.MatchEntry = fmMatchEntryComplete

    Set mLblDescription = AddLabel("lblDescription", 14, 254, 126, 18, "Description")
    Set mTxtDescription = AddTextBox("txtDescription", 146, 250, 392, 22)
    Set mLblVendorName = AddLabel("lblVendorName", 14, 286, 126, 18, "Vendor(s)")
    Set mTxtVendorName = AddTextBox("txtVendorName", 146, 282, 392, 22)
    Set mLblVendorCode = AddLabel("lblVendorCode", 14, 318, 126, 18, "Vendor code")
    Set mTxtVendorCode = AddTextBox("txtVendorCode", 146, 314, 120, 22)
    Set mLblExternalCode = AddLabel("lblExternalCode", 288, 318, 92, 18, "External code")
    Set mTxtExternalCode = AddTextBox("txtExternalCode", 386, 314, 152, 22)
    Set mLblImagePath = AddLabel("lblImagePath", 14, 350, 126, 18, "Picture path/URL")
    Set mTxtImagePath = AddTextBox("txtImagePath", 146, 346, 392, 22)
    ShowImagePathPlaceholder

    Set mLblEditReason = AddLabel("lblEditReason", 14, 382, 126, 18, "Why the edit? *")
    Set mTxtEditReason = AddTextBox("txtEditReason", 146, 378, 392, 22)
    mLblEditReason.Visible = False
    mTxtEditReason.Visible = False

    Set mLblCustomName = AddLabel("lblCustomName", 14, 382, 126, 18, "Additional field")
    Set mTxtCustomName = AddTextBox("txtCustomName", 146, 378, 144, 22)
    Set mLblCustomValue = AddLabel("lblCustomValue", 298, 382, 42, 18, "Value")
    Set mTxtCustomValue = AddTextBox("txtCustomValue", 342, 378, 132, 22)
    Set mBtnAddField = AddButton("btnAddField", 484, 377, 54, 24, "Add")

    Set mLstCustomFields = AddListBox("lstCustomFields", 146, 408, 328, 134)
    mLstCustomFields.ColumnCount = 2
    mLstCustomFields.ColumnWidths = "130 pt;190 pt"
    Set mBtnRemoveField = AddButton("btnRemoveField", 484, 446, 54, 24, "Remove")

    Set mLblStatus = AddLabel("lblStatus", 146, 550, 328, 28, "")
    mLblStatus.ForeColor = 255
    Set mBtnCreateInventoryTable = AddButton("btnCreateInventoryTable", 14, 594, 154, 28, "Create Inventory Table")
    Set mBtnUploadInventoryTable = AddButton("btnUploadInventoryTable", 174, 594, 192, 28, "Upload Selected Inventory Table")
    Set mBtnOK = AddButton("btnOK", 374, 594, 78, 28, "Add Item")
    Set mBtnCancel = AddButton("btnCancel", 460, 594, 78, 28, "Cancel")

    ApplyModeLayout
End Sub

Private Sub InitializeAddInventoryAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me, 575, 665

    mAnchors.Add mBtnAddMode, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mBtnEditMode, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mBtnDeleteItem, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mBtnCreateInventoryTable, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnUploadInventoryTable, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mLblContext, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblGenerated, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mCmbEditItem, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstEditItemResults, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtItemName, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtDescription, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtVendorName, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtImagePath, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtCustomValue, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnAddField, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstCustomFields, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnRemoveField, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblStatus, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnOK, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnCancel, ANCHOR_RIGHT Or ANCHOR_BOTTOM
End Sub

Private Sub LoadUomOptions()
    Dim seen As Object
    Dim item As Variant
    Dim uomValue As String

    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    mCmbUom.Clear
    AddUomOption seen, "EA"
    AddUomOption seen, "CS"
    AddUomOption seen, "LB"
    AddUomOption seen, "OZ"
    AddUomOption seen, "GAL"
    AddUomOption seen, "L"
    AddUomOption seen, "ML"
    AddUomOption seen, "KG"
    If Not mCatalogItems Is Nothing Then
        For Each item In mCatalogItems
            uomValue = CatalogField(item, "UOM")
            If uomValue <> "" Then AddUomOption seen, uomValue
        Next item
    End If
    mCmbUom.AddItem ADD_UOM_OPTION
End Sub

Private Sub AddUomOption(ByVal seen As Object, ByVal uomValue As String)
    uomValue = UCase$(Trim$(uomValue))
    If uomValue = "" Then Exit Sub
    If seen.Exists(uomValue) Then Exit Sub
    seen(uomValue) = True
    mCmbUom.AddItem uomValue
End Sub

Private Sub LoadEditItemOptions(Optional ByVal filterText As String = "")
    Dim item As Variant
    Dim displayText As String
    Dim rowIndex As Long
    Dim normalizedFilter As String
    Dim sku As String
    Dim itemName As String

    If mCmbEditItem Is Nothing Then Exit Sub
    mCmbEditItem.Clear
    If Not mLstEditItemResults Is Nothing Then mLstEditItemResults.Clear
    If mCatalogItems Is Nothing Then Exit Sub
    normalizedFilter = AdminNormalizeSearchText(filterText)
    For Each item In mCatalogItems
        sku = CatalogField(item, "SKU")
        itemName = CatalogField(item, "ITEM")
        If sku <> "" Then
            If normalizedFilter <> "" Then
                If Not AdminAnyTextMatchesSearch(normalizedFilter, sku, itemName, _
                                                 CatalogField(item, "UOM"), CatalogField(item, "LOCATION"), _
                                                 CatalogField(item, "DESCRIPTION"), CatalogField(item, "VENDOR(s)"), _
                                                 CatalogField(item, "VENDOR_CODE"), CatalogField(item, "CATEGORY"), _
                                                 CatalogField(item, "EXTERNAL_CODE")) Then GoTo NextItem
            End If
            displayText = itemName
            If displayText = "" Then displayText = sku
            If CatalogField(item, "UOM") <> "" Then displayText = displayText & " [" & CatalogField(item, "UOM") & "]"
            If sku <> "" Then displayText = displayText & "  " & sku
            mCmbEditItem.AddItem displayText
            rowIndex = mCmbEditItem.ListCount - 1
            mCmbEditItem.List(rowIndex, 1) = sku
        End If
NextItem:
    Next item
    mCmbEditItem.ListRows = MaxLongAdminForm(1, MinLongAdminForm(12, mCmbEditItem.ListCount))
End Sub

Private Sub LoadInventoryDimensionOptions()
    Dim selectedLocation As String
    Dim selectedCategory As String
    Dim seenLocations As Object
    Dim seenCategories As Object
    Dim item As Variant

    If mCmbLocation Is Nothing Or mCmbCategory Is Nothing Then Exit Sub
    selectedLocation = Trim$(CStr(mCmbLocation.Value))
    selectedCategory = Trim$(CStr(mCmbCategory.Value))
    Set seenLocations = CreateObject("Scripting.Dictionary")
    seenLocations.CompareMode = vbTextCompare
    Set seenCategories = CreateObject("Scripting.Dictionary")
    seenCategories.CompareMode = vbTextCompare

    mCmbLocation.Clear
    mCmbCategory.Clear
    AddDimensionOption mCmbLocation, seenLocations, mConfiguredDefaultLocation
    If Not mCatalogItems Is Nothing Then
        For Each item In mCatalogItems
            AddDimensionOption mCmbLocation, seenLocations, CatalogField(item, "LOCATION")
            AddDimensionOption mCmbCategory, seenCategories, CatalogField(item, "CATEGORY")
        Next item
    End If
    If selectedLocation <> "" Then mCmbLocation.Value = selectedLocation
    If selectedCategory <> "" Then mCmbCategory.Value = selectedCategory
End Sub

Private Sub AddDimensionOption(ByVal targetCombo As MSForms.ComboBox, _
                               ByVal seen As Object, _
                               ByVal valueText As String)
    valueText = Trim$(valueText)
    If valueText = "" Then Exit Sub
    If seen.Exists(valueText) Then Exit Sub
    seen.Add valueText, True
    targetCombo.AddItem valueText
End Sub

Private Function ComboContainsValue(ByVal targetCombo As MSForms.ComboBox, _
                                    ByVal valueText As String) As Boolean
    Dim i As Long

    If targetCombo Is Nothing Then Exit Function
    For i = 0 To targetCombo.ListCount - 1
        If StrComp(Trim$(CStr(targetCombo.List(i))), Trim$(valueText), vbTextCompare) = 0 Then
            ComboContainsValue = True
            Exit Function
        End If
    Next i
End Function

Private Sub ShowEditItemSearchResults(ByVal searchText As String)
    Dim i As Long
    Dim rowIndex As Long
    Dim resultCount As Long

    If mLstEditItemResults Is Nothing Then Exit Sub
    mLstEditItemResults.Clear
    If AdminNormalizeSearchText(searchText) = "" Or mCmbEditItem.ListCount = 0 Then
        HideEditItemSearchResults
        Exit Sub
    End If

    For i = 0 To mCmbEditItem.ListCount - 1
        mLstEditItemResults.AddItem CStr(mCmbEditItem.List(i, 0))
        rowIndex = mLstEditItemResults.ListCount - 1
        mLstEditItemResults.List(rowIndex, 1) = CStr(mCmbEditItem.List(i, 1))
    Next i

    resultCount = mLstEditItemResults.ListCount
    mLstEditItemResults.Height = 16 * MaxLongAdminForm(1, MinLongAdminForm(8, resultCount))
    mLstEditItemResults.Visible = (resultCount > 0)
    If mLstEditItemResults.Visible Then mLstEditItemResults.ZOrder 0
End Sub

Private Sub HideEditItemSearchResults()
    If mLstEditItemResults Is Nothing Then Exit Sub
    mLstEditItemResults.Visible = False
    mLstEditItemResults.Clear
End Sub

Private Sub LoadQuantityOptions()
    If mTxtQty Is Nothing Then Exit Sub
    mTxtQty.Clear
    mTxtQty.AddItem "1"
    mTxtQty.AddItem QTY_OPTION_UTILITY
    mTxtQty.AddItem QTY_OPTION_SERVICE
    mTxtQty.AddItem QTY_OPTION_NOT_COUNTED
End Sub

Private Sub RefreshGeneratedLabel()
    If mEditMode Then
        If mSelectedEditSku <> "" Then
            mLblGenerated.Caption = "Internal code: " & mSelectedEditSku & _
                                    "    Warehouse: " & mWarehouseId & " / " & mStationId
        Else
            mLblGenerated.Caption = "Choose an inventory item to edit. Warehouse: " & mWarehouseId & " / " & mStationId
        End If
    Else
        mLblGenerated.Caption = "Internal code: " & mGeneratedSku & _
                                "    Warehouse: " & mWarehouseId & " / " & mStationId
    End If
End Sub

Private Function AddLabel(ByVal controlName As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single, _
                          ByVal captionText As String) As MSForms.Label
    Set AddLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    AddLabel.Left = leftPos
    AddLabel.Top = topPos
    AddLabel.Width = widthVal
    AddLabel.Height = heightVal
    AddLabel.Caption = captionText
End Function

Private Function AddTextBox(ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBox = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    AddTextBox.Left = leftPos
    AddTextBox.Top = topPos
    AddTextBox.Width = widthVal
    AddTextBox.Height = heightVal
End Function

Private Function AddCombo(ByVal controlName As String, _
                          ByVal leftPos As Single, _
                          ByVal topPos As Single, _
                          ByVal widthVal As Single, _
                          ByVal heightVal As Single) As MSForms.ComboBox
    Set AddCombo = Me.Controls.Add("Forms.ComboBox.1", controlName, True)
    AddCombo.Left = leftPos
    AddCombo.Top = topPos
    AddCombo.Width = widthVal
    AddCombo.Height = heightVal
End Function

Private Function AddListBox(ByVal controlName As String, _
                            ByVal leftPos As Single, _
                            ByVal topPos As Single, _
                            ByVal widthVal As Single, _
                            ByVal heightVal As Single) As MSForms.ListBox
    Set AddListBox = Me.Controls.Add("Forms.ListBox.1", controlName, True)
    AddListBox.Left = leftPos
    AddListBox.Top = topPos
    AddListBox.Width = widthVal
    AddListBox.Height = heightVal
End Function

Private Function AddButton(ByVal controlName As String, _
                           ByVal leftPos As Single, _
                           ByVal topPos As Single, _
                           ByVal widthVal As Single, _
                           ByVal heightVal As Single, _
                           ByVal captionText As String) As MSForms.CommandButton
    Set AddButton = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    AddButton.Left = leftPos
    AddButton.Top = topPos
    AddButton.Width = widthVal
    AddButton.Height = heightVal
    AddButton.Caption = captionText
End Function

Private Sub mBtnAddField_Click()
    Dim fieldName As String
    Dim fieldValue As String
    Dim rowIndex As Long

    fieldName = NormalizeCustomFieldName(Trim$(CStr(mTxtCustomName.Value)))
    fieldValue = Trim$(CStr(mTxtCustomValue.Value))
    If fieldName = "" Then
        mLblStatus.Caption = "Additional field name is required."
        Exit Sub
    End If
    If IsReservedCustomField(fieldName) Then
        mLblStatus.Caption = "That field is already handled by the form."
        Exit Sub
    End If
    If fieldValue = "" Then
        mLblStatus.Caption = "Additional field value is required."
        Exit Sub
    End If

    rowIndex = FindCustomFieldRow(fieldName)
    If rowIndex < 0 Then
        mLstCustomFields.AddItem fieldName
        rowIndex = mLstCustomFields.ListCount - 1
    End If
    mLstCustomFields.List(rowIndex, 1) = fieldValue
    mTxtCustomName.Value = ""
    mTxtCustomValue.Value = ""
    mLblStatus.Caption = ""
End Sub

Private Sub mBtnRemoveField_Click()
    If mLstCustomFields.ListIndex >= 0 Then mLstCustomFields.RemoveItem mLstCustomFields.ListIndex
End Sub

Private Sub mBtnCreateInventoryTable_Click()
    Dim report As String
    Dim tableName As String

    On Error GoTo Failed
    If Not CurrentContextAllowed() Then Exit Sub
    If mOperatorWorkbook Is Nothing Then
        mLblStatus.Caption = "The captured Admin operator workbook is unavailable."
        Exit Sub
    End If
    If modAdminInventoryWorksheet.CreateInventoryWorksheetTable( _
            mOperatorWorkbook, tableName, report) Then
        mLastInventoryTableName = tableName
        mAccepted = False
    End If
    mLastInventoryWorksheetReport = report
    mLblStatus.Caption = report
    If Not modAdminInventoryWorksheet.InventoryWorksheetAutomationEnabled() Then _
        MsgBox report, IIf(tableName <> "", vbInformation, vbExclamation), "invSys Admin"
    If tableName <> "" Then Me.Hide
    Exit Sub

Failed:
    report = "Inventory worksheet creation failed: " & Err.Description
    mLastInventoryWorksheetReport = report
    mLblStatus.Caption = report
End Sub

Private Sub mBtnUploadInventoryTable_Click()
    Dim report As String
    Dim succeeded As Boolean

    On Error GoTo Failed
    If Not CurrentContextAllowed() Then Exit Sub
    If mOperatorWorkbook Is Nothing Then
        mLblStatus.Caption = "The captured Admin operator workbook is unavailable."
        Exit Sub
    End If
    succeeded = modAdminInventoryWorksheet.UploadSelectedInventoryWorksheetTable( _
        mOperatorWorkbook, mWarehouseId, mStationId, mUserId, report, mInteractiveContext)
    mLastInventoryWorksheetReport = report
    mLblStatus.Caption = report
    If Not modAdminInventoryWorksheet.InventoryWorksheetAutomationEnabled() Then _
        MsgBox report, IIf(succeeded, vbInformation, vbExclamation), "invSys Admin"
    mAccepted = False
    Me.Hide
    Exit Sub

Failed:
    report = "Inventory worksheet upload failed: " & Err.Description
    mLastInventoryWorksheetReport = report
    mLblStatus.Caption = report
End Sub

Private Sub mBtnOK_Click()
    If Not CurrentContextAllowed() Then Exit Sub
    If Not ValidateForm Then Exit Sub
    If mEditMode Then
        If Not PromptForEditReason Then Exit Sub
    End If
    If Not CurrentContextAllowed() Then Exit Sub
    mDeleteRequested = False
    mDeleteReason = ""
    mAccepted = True
    Me.Hide
End Sub

Private Function PromptForEditReason() As Boolean
    Dim answer As String

    answer = Trim$(InputBox("Why the edit?" & vbCrLf & vbCrLf & _
                            "The reason will be saved with the current user and date/time.", _
                            "invSys Admin - Edit Inventory Item", EditReason))
    If answer = "" Then
        mLblStatus.Caption = "Why the edit is required."
        Exit Function
    End If

    mTxtEditReason.Value = answer
    mLblStatus.Caption = ""
    PromptForEditReason = True
End Function

Private Sub mBtnAddMode_Click()
    If mLoading Then Exit Sub
    mEditMode = False
    ApplyModeLayout
End Sub

Private Sub mBtnEditMode_Click()
    If mLoading Then Exit Sub
    mEditMode = True
    ApplyModeLayout
End Sub

Private Sub mBtnDeleteItem_Click()
    Dim answer As VbMsgBoxResult
    Dim reason As String

    If Not CurrentContextAllowed() Then Exit Sub
    If Not mEditMode Or mSelectedEditSku = "" Then
        mLblStatus.Caption = "Choose an inventory item to delete."
        Exit Sub
    End If
    If mDeleteAutomation Then
        answer = vbYes
        reason = "Packaged deletion contract"
    Else
        answer = MsgBox("Delete '" & ItemName & "' (" & mSelectedEditSku & _
            ") from managed inventory?" & vbCrLf & vbCrLf & _
            "All active exact inventory entities for this item will be retired. " & _
            "Catalog and event history will be retained.", _
            vbQuestion Or vbYesNo Or vbDefaultButton2, "invSys Admin - Delete Item")
        If answer <> vbYes Then Exit Sub
        If Not CurrentContextAllowed() Then Exit Sub
        reason = Trim$(InputBox("Why is this item being deleted from managed inventory?", _
            "invSys Admin - Delete Item"))
    End If
    If Trim$(reason) = "" Then
        mLblStatus.Caption = "Why the deletion is required."
        Exit Sub
    End If
    If Not CurrentContextAllowed() Then Exit Sub

    mDeleteReason = Trim$(reason)
    mDeleteRequested = True
    mAccepted = True
    Me.Hide
End Sub

Private Sub mCmbEditItem_Change()
    On Error GoTo CleanFail

    Dim searchText As String

    If mLoading Then Exit Sub
    If Not mEditMode Then Exit Sub
    If mFilteringEditItems Then Exit Sub
    If CommitSelectedEditItemFromCombo() Then Exit Sub
    searchText = Trim$(CStr(mCmbEditItem.Text))
    mFilteringEditItems = True
    mLoading = True
    LoadEditItemOptions searchText
    mCmbEditItem.ListIndex = -1
    mCmbEditItem.Text = searchText
    mCmbEditItem.SelStart = Len(searchText)
    ShowEditItemSearchResults searchText
    mLoading = False
    mFilteringEditItems = False
    Exit Sub

CleanFail:
    mLoading = False
    mFilteringEditItems = False
    mLblStatus.Caption = "Inventory search failed: " & Err.Description
End Sub

Private Sub mCmbEditItem_Click()
    If mLoading Then Exit Sub
    If Not mEditMode Then Exit Sub
    If CommitSelectedEditItemFromCombo() Then Exit Sub
    ShowEditItemSearchResults CStr(mCmbEditItem.Text)
End Sub

Private Sub mCmbEditItem_AfterUpdate()
    If mLoading Then Exit Sub
    If Not mEditMode Then Exit Sub
    If mSelectedEditSku = "" Then ShowEditItemSearchResults CStr(mCmbEditItem.Text)
End Sub

Private Sub mLstEditItemResults_Click()
    Dim sku As String
    Dim displayText As String

    If mLoading Then Exit Sub
    If Not mEditMode Then Exit Sub
    If mLstEditItemResults Is Nothing Then Exit Sub
    If mLstEditItemResults.ListIndex < 0 Then Exit Sub

    displayText = CStr(mLstEditItemResults.List(mLstEditItemResults.ListIndex, 0))
    sku = CStr(mLstEditItemResults.List(mLstEditItemResults.ListIndex, 1))
    CommitEditItemSelection displayText, sku
End Sub

Private Function CommitSelectedEditItemFromCombo() As Boolean
    Dim selectedIndex As Long

    If mCmbEditItem Is Nothing Then Exit Function
    selectedIndex = mCmbEditItem.ListIndex
    If selectedIndex < 0 Then Exit Function
    CommitSelectedEditItemFromCombo = CommitEditItemSelection( _
        CStr(mCmbEditItem.List(selectedIndex, 0)), _
        CStr(mCmbEditItem.List(selectedIndex, 1)))
End Function

Private Function CommitEditItemSelection(ByVal displayText As String, _
                                         ByVal sku As String) As Boolean
    sku = Trim$(sku)
    If sku = "" Then Exit Function
    HideEditItemSearchResults

    mLoading = True
    mCmbEditItem.Text = displayText
    mCmbEditItem.ListIndex = -1
    mLoading = False

    LoadSelectedEditItemBySku sku
    If Not mBtnDeleteItem Is Nothing Then mBtnDeleteItem.Enabled = (mSelectedEditSku <> "")
    CommitEditItemSelection = (StrComp(mSelectedEditSku, sku, vbTextCompare) = 0)
End Function

Private Sub mCmbUom_Change()
    Dim newUom As String

    If mLoading Then Exit Sub
    If Trim$(CStr(mCmbUom.Value)) <> ADD_UOM_OPTION Then
        If Trim$(CStr(mCmbUom.Value)) <> "" Then mPreviousUom = Trim$(CStr(mCmbUom.Value))
        Exit Sub
    End If
    If Not mAllowUomPrompt Then
        If mPreviousUom <> "" Then
            mCmbUom.Value = mPreviousUom
        Else
            mCmbUom.Value = "EA"
        End If
        Exit Sub
    End If

    newUom = UCase$(Trim$(InputBox("New UOM:", "invSys Admin - Add UOM")))
    If Not CurrentContextAllowed() Then Exit Sub
    If newUom = "" Then
        If mPreviousUom <> "" Then
            mCmbUom.Value = mPreviousUom
        Else
            mCmbUom.Value = "EA"
        End If
        Exit Sub
    End If
    InsertUomBeforeAddOption newUom
    mCmbUom.Value = newUom
    mPreviousUom = newUom
End Sub

Private Sub mTxtQty_Change()
    If mLoading Then Exit Sub
    ApplyQuantityModeState
End Sub

Private Sub mTxtImagePath_Enter()
    If mImagePlaceholderActive Then
        mImagePlaceholderActive = False
        mTxtImagePath.Value = ""
        mTxtImagePath.ForeColor = vbWindowText
    End If
End Sub

Private Sub mTxtImagePath_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Trim$(CStr(mTxtImagePath.Value)) = "" Then ShowImagePathPlaceholder
End Sub

Private Sub ShowImagePathPlaceholder()
    If mTxtImagePath Is Nothing Then Exit Sub
    mImagePlaceholderActive = True
    mTxtImagePath.Value = IMAGE_PATH_PLACEHOLDER
    mTxtImagePath.ForeColor = RGB(128, 128, 128)
End Sub

Private Sub SetImagePathValue(ByVal imagePathValue As String)
    imagePathValue = Trim$(imagePathValue)
    If imagePathValue = "" Then
        ShowImagePathPlaceholder
    Else
        mImagePlaceholderActive = False
        mTxtImagePath.ForeColor = vbWindowText
        mTxtImagePath.Value = imagePathValue
    End If
End Sub

Private Function QuantityModeText() As String
    If mTxtQty Is Nothing Then Exit Function
    QuantityModeText = Trim$(CStr(mTxtQty.Value))
End Function

Private Function QuantityModeIsNonCounted(ByVal valueText As String) As Boolean
    Select Case UCase$(Trim$(valueText))
        Case "UTILITY", "SERVICE", "NOT COUNTED", "NON-COUNTED", "NONCOUNTED", "UNTRACKED"
            QuantityModeIsNonCounted = True
    End Select
End Function

Private Function NonCountedItemKind() As String
    Select Case UCase$(QuantityModeText())
        Case "UTILITY"
            NonCountedItemKind = "UTILITY"
        Case "SERVICE"
            NonCountedItemKind = "SERVICE"
        Case Else
            NonCountedItemKind = "NON_COUNTED"
    End Select
End Function

Private Sub ApplyQuantityModeState()
    If mTxtQty Is Nothing Or mLblQty Is Nothing Then Exit Sub
    mTxtQty.Enabled = True
    If NonCountedItem Then
        mLblQty.Caption = "Qty mode"
        If Not mCmbCategory Is Nothing Then
            If Trim$(CStr(mCmbCategory.Value)) = "" Then mCmbCategory.Value = NonCountedItemKind()
        End If
        If UCase$(QuantityModeText()) = "UTILITY" Then
            If Not mTxtVendorName Is Nothing Then
                If Trim$(CStr(mTxtVendorName.Value)) = "" Then mTxtVendorName.Value = "Utility"
            End If
        End If
    ElseIf mEditMode Then
        mLblQty.Caption = "Set qty"
    Else
        mLblQty.Caption = "Starting qty *"
    End If
End Sub

Private Sub InsertUomBeforeAddOption(ByVal uomValue As String)
    Dim i As Long

    uomValue = UCase$(Trim$(uomValue))
    If uomValue = "" Then Exit Sub
    For i = 0 To mCmbUom.ListCount - 1
        If StrComp(CStr(mCmbUom.List(i)), uomValue, vbTextCompare) = 0 Then Exit Sub
    Next i
    If mCmbUom.ListCount > 0 Then
        mCmbUom.AddItem uomValue, mCmbUom.ListCount - 1
    Else
        mCmbUom.AddItem uomValue
    End If
End Sub

Private Sub ApplyModeLayout()
    If mBtnOK Is Nothing Then Exit Sub
    mLblEditItem.Visible = mEditMode
    mCmbEditItem.Visible = mEditMode
    If Not mBtnDeleteItem Is Nothing Then
        mBtnDeleteItem.Visible = mEditMode
        mBtnDeleteItem.Enabled = (mEditMode And mSelectedEditSku <> "")
    End If
    HideEditItemSearchResults
    If Not mLblEditReason Is Nothing Then mLblEditReason.Visible = False
    If Not mTxtEditReason Is Nothing Then mTxtEditReason.Visible = False
    If mEditMode Then
        mLblTitle.Caption = "Edit inventory item"
        mLblContext.Caption = "Edit catalog fields for an existing inventory item. Qty sets the target on-hand total."
        mLblQty.Caption = "Set qty"
        mBtnOK.Caption = "Save Item"
    Else
        mLblTitle.Caption = "Add inventory item"
        mLblContext.Caption = "Fill the required fields. Internal item code is generated by invSys."
        mLblQty.Caption = "Starting qty *"
        mBtnOK.Caption = "Add Item"
    End If
    If Not mEditMode Then
        mSelectedEditSku = ""
        mDeleteRequested = False
        mDeleteReason = ""
    ElseIf mCmbEditItem.ListIndex < 0 Then
        ClearEditableFields
    End If
    ApplyQuantityModeState
    RefreshGeneratedLabel
    mLblStatus.Caption = ""
End Sub

Private Sub ClearEditableFields()
    mTxtItemName.Value = ""
    mCmbUom.Value = "EA"
    mPreviousUom = "EA"
    mTxtQty.Value = ""
    mCmbLocation.Value = ""
    mTxtDescription.Value = ""
    mCmbCategory.Value = ""
    mTxtVendorName.Value = ""
    mTxtVendorCode.Value = ""
    mTxtExternalCode.Value = ""
    If Not mTxtEditReason Is Nothing Then mTxtEditReason.Value = ""
    ApplyQuantityModeState
    ShowImagePathPlaceholder
    mTxtCustomName.Value = ""
    mTxtCustomValue.Value = ""
    mLstCustomFields.Clear
    If Not mBtnDeleteItem Is Nothing Then mBtnDeleteItem.Enabled = False
End Sub

Private Sub LoadSelectedEditItem()
    Dim sku As String

    If mCmbEditItem.ListIndex < 0 Then Exit Sub
    sku = CStr(mCmbEditItem.List(mCmbEditItem.ListIndex, 1))
    LoadSelectedEditItemBySku sku
End Sub

Private Sub LoadSelectedEditItemBySku(ByVal sku As String)
    Dim item As Object
    Dim report As String

    sku = Trim$(sku)
    If sku = "" Then Exit Sub
    Set item = FindCatalogItemBySku(sku)
    If item Is Nothing Then Exit Sub
    If mInteractiveContext <> "" Then
        If Not modInventoryCommands.SelectSku(mInteractiveContext, sku, report) Then mLblStatus.Caption = report: Exit Sub
    End If

    mSelectedEditSku = sku
    mTxtItemName.Value = CatalogField(item, "ITEM")
    If CatalogField(item, "UOM") <> "" Then
        InsertUomBeforeAddOption CatalogField(item, "UOM")
        mCmbUom.Value = CatalogField(item, "UOM")
        mPreviousUom = CatalogField(item, "UOM")
    Else
        mCmbUom.Value = "EA"
        mPreviousUom = "EA"
    End If
    mTxtQty.Value = ""
    mCmbLocation.Value = CatalogField(item, "LOCATION")
    mTxtDescription.Value = CatalogField(item, "DESCRIPTION")
    mCmbCategory.Value = CatalogField(item, "CATEGORY")
    mTxtVendorName.Value = CatalogField(item, "VENDOR(s)")
    mTxtVendorCode.Value = CatalogField(item, "VENDOR_CODE")
    mTxtExternalCode.Value = CatalogField(item, "EXTERNAL_CODE")
    If Not mTxtEditReason Is Nothing Then mTxtEditReason.Value = ""
    If CatalogItemIsNonCounted(item) Then
        Select Case UCase$(CatalogField(item, "ITEM_KIND"))
            Case "UTILITY"
                mTxtQty.Value = QTY_OPTION_UTILITY
            Case "SERVICE"
                mTxtQty.Value = QTY_OPTION_SERVICE
            Case Else
                mTxtQty.Value = QTY_OPTION_NOT_COUNTED
        End Select
    ElseIf CatalogField(item, "QTY_ON_HAND") <> "" Then
        mTxtQty.Value = CatalogField(item, "QTY_ON_HAND")
    End If
    ApplyQuantityModeState
    SetImagePathValue CatalogField(item, "IMAGE_PATH")
    mLstCustomFields.Clear
    RefreshGeneratedLabel
    mLblStatus.Caption = ""
End Sub

Private Sub mBtnCancel_Click()
    mAccepted = False
    mDeleteRequested = False
    mDeleteReason = ""
    Me.Hide
End Sub

Private Function ValidateForm() As Boolean
    If mGeneratedSku = "" Then
        mLblStatus.Caption = "Generated internal code is missing."
        Exit Function
    End If
    If mEditMode And mSelectedEditSku = "" Then
        mLblStatus.Caption = "Choose an inventory item to edit."
        Exit Function
    End If
    If ItemName = "" Then
        mLblStatus.Caption = "Item name is required."
        Exit Function
    End If
    If Uom = "" Then
        mLblStatus.Caption = "UOM is required."
        Exit Function
    End If
    If Not mEditMode Then
        If NonCountedItem Then
            ValidateForm = True
            Exit Function
        End If
        If Not IsNumeric(CStr(mTxtQty.Value)) Then
            mLblStatus.Caption = "Starting quantity must be numeric or a mode like Utility."
            Exit Function
        End If
        If StartingQty < 0 Then
            mLblStatus.Caption = "Starting quantity cannot be negative."
            Exit Function
        End If
    ElseIf Trim$(CStr(mTxtQty.Value)) <> "" And Not NonCountedItem Then
        If Not IsNumeric(CStr(mTxtQty.Value)) Then
            mLblStatus.Caption = "Set qty must be numeric."
            Exit Function
        End If
        If StartingQty < 0 Then
            mLblStatus.Caption = "Set qty cannot be negative."
            Exit Function
        End If
    End If
    ValidateForm = True
End Function

Private Function CatalogItemIsNonCounted(ByVal item As Variant) As Boolean
    Dim trackQty As String
    Dim itemKind As String

    trackQty = UCase$(CatalogField(item, "TRACK_QTY"))
    itemKind = UCase$(CatalogField(item, "ITEM_KIND"))
    CatalogItemIsNonCounted = (trackQty = "FALSE" Or trackQty = "NO" Or trackQty = "0" _
                               Or itemKind = "UTILITY" Or itemKind = "SERVICE" Or itemKind = "NON_COUNTED")
End Function

Private Function FindCatalogItemBySku(ByVal sku As String) As Object
    Dim item As Variant

    If mCatalogItems Is Nothing Then Exit Function
    For Each item In mCatalogItems
        If StrComp(CatalogField(item, "SKU"), sku, vbTextCompare) = 0 Then
            Set FindCatalogItemBySku = item
            Exit Function
        End If
    Next item
End Function

Private Function CatalogField(ByVal item As Variant, ByVal fieldName As String) As String
    On Error Resume Next
    CatalogField = Trim$(CStr(item(fieldName)))
    On Error GoTo 0
End Function

Private Function FindCustomFieldRow(ByVal fieldName As String) As Long
    Dim i As Long

    FindCustomFieldRow = -1
    For i = 0 To mLstCustomFields.ListCount - 1
        If StrComp(CStr(mLstCustomFields.List(i, 0)), fieldName, vbTextCompare) = 0 Then
            FindCustomFieldRow = i
            Exit Function
        End If
    Next i
End Function

Private Function NormalizeCustomFieldName(ByVal fieldName As String) As String
    fieldName = Replace(fieldName, vbCr, " ")
    fieldName = Replace(fieldName, vbLf, " ")
    fieldName = Replace(fieldName, vbTab, " ")
    fieldName = Replace(fieldName, "[", "(")
    fieldName = Replace(fieldName, "]", ")")
    Do While InStr(1, fieldName, "  ", vbBinaryCompare) > 0
        fieldName = Replace(fieldName, "  ", " ")
    Loop
    fieldName = Trim$(fieldName)
    If Len(fieldName) > 48 Then fieldName = Left$(fieldName, 48)
    NormalizeCustomFieldName = fieldName
End Function

Private Function IsReservedCustomField(ByVal fieldName As String) As Boolean
    Select Case UCase$(Trim$(fieldName))
        Case "SYSTEM_KEY", "SYSTEMKEY", "SYSTEM KEY", "SYSTEM-KEY", "CONDITION", "SKU", "ROW", "ITEM_CODE", "ITEM", "UOM", "LOCATION", "QTY", "TOTAL INV", _
             "QTYAVAILABLE", "DESCRIPTION", "VENDOR(S)", "VENDOR_CODE", "CATEGORY", _
             "EXTERNAL_CODE", "IMAGE_PATH", "TRACK_QTY", "ITEM_KIND", "NOTE", "IOTYPE"
            IsReservedCustomField = True
    End Select
End Function
