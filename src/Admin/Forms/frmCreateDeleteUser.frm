VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateDeleteUser
   Caption         =   "Users & Roles"
   ClientHeight    =   5850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9150
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateDeleteUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mTxtRoot As MSForms.TextBox
Private WithEvents mTxtNasUser As MSForms.TextBox
Private WithEvents mTxtNasPassword As MSForms.TextBox
Private WithEvents mBtnRootFind As MSForms.CommandButton
Private WithEvents mBtnRootScan As MSForms.CommandButton
Private WithEvents mBtnNasConnect As MSForms.CommandButton
Private WithEvents mCmbWarehouse As MSForms.ComboBox
Private WithEvents mLstWarehouses As MSForms.ListBox
Private WithEvents mLstUsers As MSForms.ListBox
Private WithEvents mTxtAuthPath As MSForms.TextBox
Private WithEvents mTxtUserId As MSForms.TextBox
Private WithEvents mTxtDisplayName As MSForms.TextBox
Private WithEvents mTxtPin As MSForms.TextBox
Private WithEvents mTxtWarehouseId As MSForms.TextBox
Private WithEvents mTxtStationId As MSForms.TextBox
Private WithEvents mChkAdmin As MSForms.CheckBox
Private WithEvents mChkReceivePost As MSForms.CheckBox
Private WithEvents mChkReceiveView As MSForms.CheckBox
Private WithEvents mChkShipPost As MSForms.CheckBox
Private WithEvents mChkProdPost As MSForms.CheckBox
Private WithEvents mChkInboxProcess As MSForms.CheckBox
Private WithEvents mBtnGeneratePin As MSForms.CommandButton
Private WithEvents mBtnCopyPin As MSForms.CommandButton
Private WithEvents mBtnRefreshUsers As MSForms.CommandButton
Private WithEvents mBtnSave As MSForms.CommandButton
Private WithEvents mBtnDeactivate As MSForms.CommandButton
Private WithEvents mBtnDelete As MSForms.CommandButton
Private WithEvents mBtnClear As MSForms.CommandButton
Private WithEvents mBtnClose As MSForms.CommandButton

Private mLblStatus As MSForms.Label
Private mWarehousePathById As Object
Private mAnchors As Object
Private mBusy As Boolean
Private mResizeInitialized As Boolean

Private Const COLOR_INFO As Long = 0
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_WARNING As Long = 192
Private Const COLOR_ERROR As Long = 255
Private Const ANCHOR_LEFT As Long = 1
Private Const ANCHOR_TOP As Long = 2
Private Const ANCHOR_RIGHT As Long = 4
Private Const ANCHOR_BOTTOM As Long = 8
Private Const NO_ERROR_WIN32 As Long = 0
Private Const ERROR_SESSION_CREDENTIAL_CONFLICT As Long = 1219
Private Const RESOURCETYPE_DISK As Long = 1
Private Const CONNECT_TEMPORARY As Long = 4
Private Const MAX_WAREHOUSE_SCAN_DEPTH As Long = 4

#If VBA7 Then
Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Private Declare PtrSafe Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
#Else
Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As String
    lpRemoteName As String
    lpComment As String
    lpProvider As String
End Type

Private Declare Function WNetAddConnection2 Lib "mpr.dll" Alias "WNetAddConnection2A" (ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUserName As String, ByVal dwFlags As Long) As Long
Private Declare Function WNetCancelConnection2 Lib "mpr.dll" Alias "WNetCancelConnection2A" (ByVal lpName As String, ByVal dwFlags As Long, ByVal fForce As Long) As Long
#End If

Private Sub UserForm_Initialize()
    Randomize
    mBusy = True
    Me.Caption = "Users & Roles"
    Me.Width = 740
    Me.Height = 585
    On Error Resume Next
    Me.ScrollBars = 0
    Me.KeepScrollBarsVisible = 0
    Me.ScrollLeft = 0
    Me.ScrollTop = 0
    On Error GoTo 0
    BuildUsersRolesLayout
    InitializeUsersRolesAnchors
    Set mWarehousePathById = CreateObject("Scripting.Dictionary")
    mWarehousePathById.CompareMode = vbTextCompare
    mTxtRoot.Value = ResolveDefaultWarehouseRootForm()
    mTxtStationId.Value = "*"
    RefreshWarehouseListForm False
    ShowStatusForm "Select a warehouse auth workbook, then create users and assign roles.", COLOR_INFO
    mBusy = False
End Sub

Private Sub UserForm_Activate()
    If Not mResizeInitialized Then
        modUserFormResizeWin.EnableResizableUserForm Me
        mResizeInitialized = True
    End If

    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Sub UserForm_Layout()
    If mAnchors Is Nothing Then Exit Sub
    mAnchors.ResizeControls
End Sub

Private Sub UserForm_Terminate()
    Set mAnchors = Nothing
End Sub

Private Sub BuildUsersRolesLayout()
    Dim topPos As Single

    AddLabelForm "lblTitle", "Users & Roles", 18, 14, 180, 18, True
    Set mLblStatus = AddLabelForm("lblStatus", "", 18, 40, 690, 30, False)

    AddLabelForm "lblRoot", "Warehouse root", 18, 78, 100, 18, False
    Set mTxtRoot = AddTextBoxForm("txtWarehouseRoot", 126, 74, 372, 22)
    Set mBtnRootFind = AddButtonForm("btnRootFind", "Find...", 506, 73, 62, 24)
    Set mBtnRootScan = AddButtonForm("btnRootScan", "Scan", 576, 73, 54, 24)

    AddLabelForm "lblNasUser", "NAS user", 18, 110, 100, 18, False
    Set mTxtNasUser = AddTextBoxForm("txtNasUser", 126, 106, 172, 22)
    AddLabelForm "lblNasPassword", "Password", 316, 110, 64, 18, False
    Set mTxtNasPassword = AddTextBoxForm("txtNasPassword", 386, 106, 112, 22)
    mTxtNasPassword.PasswordChar = "*"
    Set mBtnNasConnect = AddButtonForm("btnNasConnect", "Connect", 506, 105, 70, 24)

    AddLabelForm "lblWarehouse", "Warehouse", 18, 144, 100, 18, False
    Set mCmbWarehouse = AddComboBoxForm("cmbWarehouse", 126, 140, 220, 22)
    Set mBtnRefreshUsers = AddButtonForm("btnRefreshUsers", "Refresh Users", 356, 139, 98, 24)
    AddLabelForm "lblAuthPath", "Auth workbook", 18, 176, 100, 18, False
    Set mTxtAuthPath = AddTextBoxForm("txtAuthPath", 126, 172, 504, 22)

    AddLabelForm "lblWarehouses", "Warehouses in root", 18, 210, 114, 18, False
    Set mLstWarehouses = AddListBoxForm("lstWarehouses", 126, 204, 504, 86)
    mLstWarehouses.ColumnCount = 2
    mLstWarehouses.ColumnWidths = "150 pt;330 pt"

    AddLabelForm "lblUsers", "Users", 18, 310, 70, 18, False
    Set mLstUsers = AddListBoxForm("lstUsers", 18, 330, 270, 130)
    mLstUsers.ColumnCount = 3
    mLstUsers.ColumnWidths = "86 pt;112 pt;54 pt"

    topPos = 312
    AddLabelForm "lblUserId", "User ID", 316, topPos, 90, 18, False
    Set mTxtUserId = AddTextBoxForm("txtUserId", 420, topPos - 4, 210, 22)
    AddLabelForm "lblDisplay", "Display name", 316, topPos + 32, 90, 18, False
    Set mTxtDisplayName = AddTextBoxForm("txtDisplayName", 420, topPos + 28, 210, 22)
    AddLabelForm "lblPin", "PIN/password", 316, topPos + 64, 90, 18, False
    Set mTxtPin = AddTextBoxForm("txtPin", 420, topPos + 60, 116, 22)
    mTxtPin.ControlTipText = "This value is visible on purpose. Record it before saving; only the hash is stored."
    Set mBtnGeneratePin = AddButtonForm("btnGeneratePin", "Generate PIN", 544, topPos + 59, 86, 24)
    Set mBtnCopyPin = AddButtonForm("btnCopyPin", "Copy", 638, topPos + 59, 52, 24)
    AddLabelForm "lblWhScope", "Warehouse scope", 316, topPos + 96, 90, 18, False
    Set mTxtWarehouseId = AddTextBoxForm("txtWarehouseId", 420, topPos + 92, 116, 22)
    AddLabelForm "lblStationScope", "Station", 544, topPos + 96, 48, 18, False
    Set mTxtStationId = AddTextBoxForm("txtStationId", 594, topPos + 92, 36, 22)

    AddLabelForm "lblRoles", "Roles / capabilities", 316, 430, 140, 18, True
    Set mChkAdmin = AddCheckBoxForm("chkAdmin", "Admin maintenance", 316, 452, 146, 18)
    Set mChkReceivePost = AddCheckBoxForm("chkReceivePost", "Receiving post", 316, 474, 120, 18)
    Set mChkReceiveView = AddCheckBoxForm("chkReceiveView", "Receiving view", 316, 496, 120, 18)
    Set mChkShipPost = AddCheckBoxForm("chkShipPost", "Shipping post", 474, 452, 120, 18)
    Set mChkProdPost = AddCheckBoxForm("chkProdPost", "Production post", 474, 474, 128, 18)
    Set mChkInboxProcess = AddCheckBoxForm("chkInboxProcess", "Inbox processor", 474, 496, 128, 18)

    Set mBtnClear = AddButtonForm("btnClear", "Clear", 18, 474, 70, 26)
    Set mBtnDeactivate = AddButtonForm("btnDeactivate", "Deactivate", 96, 474, 86, 26)
    Set mBtnDelete = AddButtonForm("btnDelete", "Delete", 190, 474, 70, 26)
    Set mBtnSave = AddButtonForm("btnSave", "Create / Update", 504, 512, 126, 28)
    Set mBtnClose = AddButtonForm("btnClose", "Close", 638, 512, 70, 28)
End Sub

Private Sub InitializeUsersRolesAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me, 740, 585

    mAnchors.Add mLblStatus, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtRoot, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnRootFind, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnRootScan, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtNasUser, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mTxtNasPassword, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnNasConnect, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mCmbWarehouse, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mBtnRefreshUsers, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mTxtAuthPath, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstWarehouses, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLstUsers, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_BOTTOM
    mAnchors.Add mTxtUserId, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtDisplayName, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtPin, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnGeneratePin, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mBtnCopyPin, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtWarehouseId, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mTxtStationId, ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mChkAdmin, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mChkReceivePost, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mChkReceiveView, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mChkShipPost, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mChkProdPost, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mChkInboxProcess, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add mBtnClear, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnDeactivate, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnDelete, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnSave, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnClose, ANCHOR_RIGHT Or ANCHOR_BOTTOM
End Sub

Private Function AddLabelForm(ByVal controlName As String, ByVal captionText As String, _
                              ByVal leftPos As Single, ByVal topPos As Single, _
                              ByVal widthVal As Single, ByVal heightVal As Single, _
                              ByVal boldText As Boolean) As MSForms.Label
    Set AddLabelForm = Me.Controls.Add("Forms.Label.1", controlName, True)
    With AddLabelForm
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Font.Bold = boldText
        .WordWrap = True
    End With
End Function

Private Function AddTextBoxForm(ByVal controlName As String, ByVal leftPos As Single, _
                                ByVal topPos As Single, ByVal widthVal As Single, _
                                ByVal heightVal As Single) As MSForms.TextBox
    Set AddTextBoxForm = Me.Controls.Add("Forms.TextBox.1", controlName, True)
    With AddTextBoxForm
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddButtonForm(ByVal controlName As String, ByVal captionText As String, _
                               ByVal leftPos As Single, ByVal topPos As Single, _
                               ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.CommandButton
    Set AddButtonForm = Me.Controls.Add("Forms.CommandButton.1", controlName, True)
    With AddButtonForm
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddComboBoxForm(ByVal controlName As String, ByVal leftPos As Single, _
                                 ByVal topPos As Single, ByVal widthVal As Single, _
                                 ByVal heightVal As Single) As MSForms.ComboBox
    Set AddComboBoxForm = Me.Controls.Add("Forms.ComboBox.1", controlName, True)
    With AddComboBoxForm
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
        .Style = fmStyleDropDownList
    End With
End Function

Private Function AddListBoxForm(ByVal controlName As String, ByVal leftPos As Single, _
                                ByVal topPos As Single, ByVal widthVal As Single, _
                                ByVal heightVal As Single) As MSForms.ListBox
    Set AddListBoxForm = Me.Controls.Add("Forms.ListBox.1", controlName, True)
    With AddListBoxForm
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Function AddCheckBoxForm(ByVal controlName As String, ByVal captionText As String, _
                                 ByVal leftPos As Single, ByVal topPos As Single, _
                                 ByVal widthVal As Single, ByVal heightVal As Single) As MSForms.CheckBox
    Set AddCheckBoxForm = Me.Controls.Add("Forms.CheckBox.1", controlName, True)
    With AddCheckBoxForm
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = heightVal
    End With
End Function

Private Sub mBtnRootFind_Click()
    Dim picker As FileDialog
    Set picker = Application.FileDialog(msoFileDialogFolderPicker)
    With picker
        .Title = "Choose warehouse root"
        .AllowMultiSelect = False
        If NormalizePathForm(CStr(mTxtRoot.Value)) <> "" Then .InitialFileName = NormalizePathForm(CStr(mTxtRoot.Value))
        If .Show <> -1 Then Exit Sub
        mTxtRoot.Value = .SelectedItems(1)
    End With
    RefreshWarehouseListForm True
End Sub

Private Sub mBtnRootScan_Click()
    RefreshWarehouseListForm True
End Sub

Private Sub mBtnNasConnect_Click()
    Dim report As String
    If ConnectSelectedRootForm(report) Then
        ShowStatusForm report, COLOR_SUCCESS
        RefreshWarehouseListForm True
    Else
        ShowStatusForm report, COLOR_ERROR
    End If
End Sub

Private Sub mBtnRefreshUsers_Click()
    RefreshUsersForm
End Sub

Private Sub mCmbWarehouse_Change()
    If mBusy Then Exit Sub
    ApplySelectedWarehouseForm CStr(mCmbWarehouse.Value)
    RefreshUsersForm
End Sub

Private Sub mLstWarehouses_Click()
    If mLstWarehouses.ListIndex < 0 Then Exit Sub
    SelectWarehouseByIdForm CStr(mLstWarehouses.List(mLstWarehouses.ListIndex, 0))
End Sub

Private Sub mLstUsers_Click()
    If mLstUsers.ListIndex < 0 Then Exit Sub
    LoadSelectedUserForm CStr(mLstUsers.List(mLstUsers.ListIndex, 0))
End Sub

Private Sub mBtnGeneratePin_Click()
    mTxtPin.Value = Format$(CLng(Int((9000000# * Rnd) + 1000000#)), "0000000")
    mTxtPin.SelStart = 0
    mTxtPin.SelLength = Len(CStr(mTxtPin.Value))
    ShowStatusForm "Generated PIN is visible in the PIN/password field. Record or copy it before saving.", COLOR_WARNING
End Sub

Private Sub mBtnCopyPin_Click()
    Dim accountText As String

    accountText = BuildAccountClipboardTextForm()
    If Len(accountText) = 0 Then
        ShowStatusForm "Enter user account details before copying.", COLOR_WARNING
        Exit Sub
    End If

    If CopyTextToClipboardForm(accountText) Then
        ShowStatusForm "User account details copied. Store the PIN/password now; only the hash is saved.", COLOR_SUCCESS
    Else
        ShowStatusForm "Could not copy automatically. Account details are visible for manual copy.", COLOR_WARNING
    End If
End Sub

Private Sub mBtnClear_Click()
    ClearUserFieldsForm
End Sub

Private Sub mBtnSave_Click()
    SaveUserAndRolesForm
End Sub

Private Sub mBtnDeactivate_Click()
    SetUserStatusForm "INACTIVE"
End Sub

Private Sub mBtnDelete_Click()
    DeleteUserForm
End Sub

Private Sub mBtnClose_Click()
    Unload Me
End Sub

Private Sub RefreshWarehouseListForm(ByVal showScanResult As Boolean)
    Dim results As Collection
    Dim item As Variant
    Dim rootPath As String
    Dim rootReachable As Boolean
    Dim connectReport As String

    mBusy = True
    mCmbWarehouse.Clear
    mLstWarehouses.Clear
    Set mWarehousePathById = CreateObject("Scripting.Dictionary")
    mWarehousePathById.CompareMode = vbTextCompare
    rootPath = NormalizePathForm(CStr(mTxtRoot.Value))
    If rootPath = "" Then rootPath = ResolveDefaultWarehouseRootForm()
    rootReachable = (rootPath <> "" And FolderExistsForm(rootPath))
    If Not rootReachable And IsUncPathForm(rootPath) Then
        If Trim$(CStr(mTxtNasUser.Value)) <> "" Or CStr(mTxtNasPassword.Value) <> "" Then
            If ConnectSelectedRootForm(connectReport) Then rootReachable = FolderExistsForm(rootPath)
        End If
    End If
    If rootReachable Then RememberWarehouseRootForAdminForm rootPath
    Set results = DiscoverWarehousesForm()

    For Each item In results
        mCmbWarehouse.AddItem CStr(item)
        mLstWarehouses.AddItem CStr(item)
        mLstWarehouses.List(mLstWarehouses.ListCount - 1, 1) = CStr(mWarehousePathById(CStr(item)))
    Next item

    If mCmbWarehouse.ListCount > 0 Then
        mCmbWarehouse.ListIndex = 0
        ApplySelectedWarehouseForm CStr(mCmbWarehouse.Value)
    Else
        mTxtAuthPath.Value = ""
        mTxtWarehouseId.Value = ""
    End If
    mBusy = False

    If showScanResult Then
        If Not rootReachable Then
            If connectReport <> "" Then
                ShowStatusForm connectReport, COLOR_ERROR
            Else
                ShowStatusForm "Warehouse root is not reachable from Excel. Enter NAS credentials and click Connect, then scan again.", COLOR_ERROR
            End If
        ElseIf mCmbWarehouse.ListCount = 0 Then
            ShowStatusForm "No warehouse auth/config workbooks found under this root. Connect to the NAS if the folder requires credentials.", COLOR_WARNING
        Else
            ShowStatusForm "Warehouse root scanned. Found " & CStr(mCmbWarehouse.ListCount) & " warehouse(s).", COLOR_SUCCESS
        End If
    End If

    If mCmbWarehouse.ListCount > 0 Then
        RefreshUsersForm
    Else
        mLstUsers.Clear
    End If
End Sub

Private Function DiscoverWarehousesForm() As Collection
    Dim results As Collection
    Dim seen As Object
    Dim rootPath As String

    Set results = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare

    rootPath = NormalizePathForm(CStr(mTxtRoot.Value))
    If rootPath = "" Then rootPath = ResolveDefaultWarehouseRootForm()
    If rootPath = "" Then
        Set DiscoverWarehousesForm = results
        Exit Function
    End If

    AddWarehousesFromFolderForm results, seen, rootPath
    If Not FolderExistsForm(rootPath) Then
        Set DiscoverWarehousesForm = results
        Exit Function
    End If

    AddWarehousesFromChildFoldersForm results, seen, rootPath, 1

    Set DiscoverWarehousesForm = results
End Function

Private Sub AddWarehousesFromChildFoldersForm(ByVal results As Collection, ByVal seen As Object, _
                                             ByVal folderPath As String, ByVal currentDepth As Long)
    Dim fso As Object
    Dim rootFolder As Object
    Dim subFolder As Object

    If currentDepth > MAX_WAREHOUSE_SCAN_DEPTH Then Exit Sub

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set rootFolder = fso.GetFolder(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0
    If rootFolder Is Nothing Then Exit Sub

    On Error Resume Next
    For Each subFolder In rootFolder.SubFolders
        AddWarehousesFromFolderForm results, seen, CStr(subFolder.Path)
        AddWarehousesFromChildFoldersForm results, seen, CStr(subFolder.Path), currentDepth + 1
    Next subFolder
    On Error GoTo 0
End Sub

Private Sub AddWarehousesFromFolderForm(ByVal results As Collection, ByVal seen As Object, ByVal folderPath As String)
    AddWarehousesBySuffixForm results, seen, folderPath, ".invSys.Config.xlsb"
    AddWarehousesBySuffixForm results, seen, folderPath, ".invSys.Auth.xlsb"
End Sub

Private Sub AddWarehousesBySuffixForm(ByVal results As Collection, ByVal seen As Object, _
                                      ByVal folderPath As String, ByVal suffix As String)
    Dim candidate As String
    Dim warehouseId As String

    folderPath = NormalizePathForm(folderPath)
    If folderPath = "" Then Exit Sub

    On Error GoTo CleanExit
    candidate = Dir$(folderPath & "\*" & suffix, vbNormal)
    Do While candidate <> ""
        If Len(candidate) > Len(suffix) Then
            warehouseId = Left$(candidate, Len(candidate) - Len(suffix))
            If warehouseId <> "" And Not seen.Exists(warehouseId) Then
                seen.Add warehouseId, True
                mWarehousePathById(warehouseId) = folderPath
                results.Add warehouseId
            End If
        End If
        candidate = Dir$
    Loop

CleanExit:
End Sub

Private Sub SelectWarehouseByIdForm(ByVal warehouseId As String)
    Dim index As Long

    For index = 0 To mCmbWarehouse.ListCount - 1
        If StrComp(CStr(mCmbWarehouse.List(index)), warehouseId, vbTextCompare) = 0 Then
            mCmbWarehouse.ListIndex = index
            ApplySelectedWarehouseForm warehouseId
            RefreshUsersForm
            Exit Sub
        End If
    Next index
End Sub

Private Sub ApplySelectedWarehouseForm(ByVal warehouseId As String)
    Dim folderPath As String

    warehouseId = Trim$(warehouseId)
    mTxtWarehouseId.Value = warehouseId
    If warehouseId = "" Then
        mTxtAuthPath.Value = ""
        Exit Sub
    End If

    If Not mWarehousePathById Is Nothing Then
        If mWarehousePathById.Exists(warehouseId) Then folderPath = CStr(mWarehousePathById(warehouseId))
    End If
    If folderPath = "" Then folderPath = NormalizePathForm(CStr(mTxtRoot.Value))
    mTxtAuthPath.Value = folderPath & "\" & warehouseId & ".invSys.Auth.xlsb"
End Sub

Private Sub RefreshUsersForm()
    Dim wb As Workbook
    Dim openedTransient As Boolean
    Dim report As String
    Dim loUsers As ListObject
    Dim i As Long
    Dim rowIndex As Long

    mLstUsers.Clear
    Set wb = OpenSelectedAuthWorkbookForm(openedTransient, report)
    If wb Is Nothing Then
        If Len(report) > 0 Then ShowStatusForm report, COLOR_WARNING
        Exit Sub
    End If

    Set loUsers = FindListObjectForm(wb, "tblUsers")
    If loUsers Is Nothing Then
        ShowStatusForm "tblUsers was not found in the selected auth workbook.", COLOR_ERROR
        GoTo CleanExit
    End If

    If Not loUsers.DataBodyRange Is Nothing Then
        For i = 1 To loUsers.ListRows.Count
            rowIndex = mLstUsers.ListCount
            mLstUsers.AddItem SafeTextForm(GetCellByColumnForm(loUsers, i, "UserId"))
            mLstUsers.List(rowIndex, 1) = SafeTextForm(GetCellByColumnForm(loUsers, i, "DisplayName"))
            mLstUsers.List(rowIndex, 2) = SafeTextForm(GetCellByColumnForm(loUsers, i, "Status"))
        Next i
    End If

    ShowStatusForm "Loaded " & CStr(mLstUsers.ListCount) & " user(s) from " & CStr(mTxtWarehouseId.Value) & ".", COLOR_SUCCESS

CleanExit:
    CloseTransientWorkbookForm wb, openedTransient
End Sub

Private Sub LoadSelectedUserForm(ByVal userId As String)
    Dim wb As Workbook
    Dim openedTransient As Boolean
    Dim report As String
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim rowIndex As Long
    Dim whId As String
    Dim stId As String

    Set wb = OpenSelectedAuthWorkbookForm(openedTransient, report)
    If wb Is Nothing Then
        ShowStatusForm report, COLOR_ERROR
        Exit Sub
    End If

    Set loUsers = FindListObjectForm(wb, "tblUsers")
    Set loCaps = FindListObjectForm(wb, "tblCapabilities")
    rowIndex = FindUserRowForm(loUsers, userId)
    If rowIndex = 0 Then GoTo CleanExit

    whId = Trim$(CStr(mTxtWarehouseId.Value))
    stId = Trim$(CStr(mTxtStationId.Value))
    If stId = "" Then stId = "*"

    mTxtUserId.Value = userId
    mTxtDisplayName.Value = SafeTextForm(GetCellByColumnForm(loUsers, rowIndex, "DisplayName"))
    mTxtPin.Value = ""
    mChkAdmin.Value = HasActiveCapabilityForm(loCaps, userId, "ADMIN_MAINT", whId, stId)
    mChkReceivePost.Value = HasActiveCapabilityForm(loCaps, userId, "RECEIVE_POST", whId, stId)
    mChkReceiveView.Value = HasActiveCapabilityForm(loCaps, userId, "RECEIVE_VIEW", whId, stId)
    mChkShipPost.Value = HasActiveCapabilityForm(loCaps, userId, "SHIP_POST", whId, stId)
    mChkProdPost.Value = HasActiveCapabilityForm(loCaps, userId, "PROD_POST", whId, stId)
    mChkInboxProcess.Value = HasActiveCapabilityForm(loCaps, userId, "INBOX_PROCESS", whId, stId)

CleanExit:
    CloseTransientWorkbookForm wb, openedTransient
End Sub

Private Sub SaveUserAndRolesForm()
    Dim wb As Workbook
    Dim openedTransient As Boolean
    Dim report As String
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim userId As String
    Dim displayName As String
    Dim pinText As String
    Dim whId As String
    Dim stId As String
    Dim userRow As Long

    userId = Trim$(CStr(mTxtUserId.Value))
    displayName = Trim$(CStr(mTxtDisplayName.Value))
    pinText = CStr(mTxtPin.Value)
    whId = Trim$(CStr(mTxtWarehouseId.Value))
    stId = Trim$(CStr(mTxtStationId.Value))
    If stId = "" Then stId = "*"

    If userId = "" Then
        ShowStatusForm "User ID is required.", COLOR_ERROR
        Exit Sub
    End If
    If displayName = "" Then displayName = userId
    If whId = "" Then
        ShowStatusForm "Warehouse scope is required.", COLOR_ERROR
        Exit Sub
    End If

    Set wb = OpenSelectedAuthWorkbookForm(openedTransient, report)
    If wb Is Nothing Then
        ShowStatusForm report, COLOR_ERROR
        Exit Sub
    End If
    If wb.ReadOnly Then
        ShowStatusForm "Auth workbook is read-only; users and roles were not saved.", COLOR_ERROR
        GoTo CleanExit
    End If

    Set loUsers = FindListObjectForm(wb, "tblUsers")
    Set loCaps = FindListObjectForm(wb, "tblCapabilities")
    If loUsers Is Nothing Or loCaps Is Nothing Then
        ShowStatusForm "Auth tables were not available after schema ensure.", COLOR_ERROR
        GoTo CleanExit
    End If

    userRow = EnsureUserRowForm(loUsers, userId)
    If SafeTextForm(GetCellByColumnForm(loUsers, userRow, "PinHash")) = "" And pinText = "" Then
        ShowStatusForm "PIN/password is required for a new auth user.", COLOR_ERROR
        GoTo CleanExit
    End If

    SetCellByColumnForm loUsers, userRow, "UserId", userId
    SetCellByColumnForm loUsers, userRow, "DisplayName", displayName
    If pinText <> "" Then SetCellByColumnForm loUsers, userRow, "PinHash", modAuth.HashUserCredential(pinText)
    SetCellByColumnForm loUsers, userRow, "Status", "ACTIVE"

    SaveCapabilityChoiceForm loCaps, userId, "ADMIN_MAINT", whId, stId, CBool(mChkAdmin.Value)
    SaveCapabilityChoiceForm loCaps, userId, "RECEIVE_POST", whId, stId, CBool(mChkReceivePost.Value)
    SaveCapabilityChoiceForm loCaps, userId, "RECEIVE_VIEW", whId, stId, CBool(mChkReceiveView.Value)
    SaveCapabilityChoiceForm loCaps, userId, "SHIP_POST", whId, stId, CBool(mChkShipPost.Value)
    SaveCapabilityChoiceForm loCaps, userId, "PROD_POST", whId, stId, CBool(mChkProdPost.Value)
    SaveCapabilityChoiceForm loCaps, userId, "INBOX_PROCESS", whId, stId, CBool(mChkInboxProcess.Value)

    wb.Save
    ShowStatusForm "User and role assignments saved. Store the visible PIN/password now; it cannot be recovered later.", COLOR_SUCCESS
    RefreshUsersForm

CleanExit:
    CloseTransientWorkbookForm wb, openedTransient
End Sub

Private Sub SetUserStatusForm(ByVal statusText As String)
    Dim wb As Workbook
    Dim openedTransient As Boolean
    Dim report As String
    Dim loUsers As ListObject
    Dim rowIndex As Long
    Dim userId As String

    userId = Trim$(CStr(mTxtUserId.Value))
    If userId = "" Then
        ShowStatusForm "Select or enter a user first.", COLOR_ERROR
        Exit Sub
    End If

    Set wb = OpenSelectedAuthWorkbookForm(openedTransient, report)
    If wb Is Nothing Then
        ShowStatusForm report, COLOR_ERROR
        Exit Sub
    End If
    If wb.ReadOnly Then
        ShowStatusForm "Auth workbook is read-only; user status was not changed.", COLOR_ERROR
        GoTo CleanExit
    End If

    Set loUsers = FindListObjectForm(wb, "tblUsers")
    rowIndex = FindUserRowForm(loUsers, userId)
    If rowIndex = 0 Then
        ShowStatusForm "User was not found.", COLOR_WARNING
        GoTo CleanExit
    End If

    SetCellByColumnForm loUsers, rowIndex, "Status", UCase$(statusText)
    wb.Save
    ShowStatusForm "User status set to " & UCase$(statusText) & ".", COLOR_SUCCESS
    RefreshUsersForm

CleanExit:
    CloseTransientWorkbookForm wb, openedTransient
End Sub

Private Sub DeleteUserForm()
    Dim wb As Workbook
    Dim openedTransient As Boolean
    Dim report As String
    Dim loUsers As ListObject
    Dim loCaps As ListObject
    Dim rowIndex As Long
    Dim i As Long
    Dim userId As String

    userId = Trim$(CStr(mTxtUserId.Value))
    If userId = "" Then
        ShowStatusForm "Select or enter a user first.", COLOR_ERROR
        Exit Sub
    End If
    If MsgBox("Delete user '" & userId & "' and all assigned capabilities?", vbQuestion + vbYesNo, "Users & Roles") <> vbYes Then Exit Sub

    Set wb = OpenSelectedAuthWorkbookForm(openedTransient, report)
    If wb Is Nothing Then
        ShowStatusForm report, COLOR_ERROR
        Exit Sub
    End If
    If wb.ReadOnly Then
        ShowStatusForm "Auth workbook is read-only; user was not deleted.", COLOR_ERROR
        GoTo CleanExit
    End If

    Set loUsers = FindListObjectForm(wb, "tblUsers")
    Set loCaps = FindListObjectForm(wb, "tblCapabilities")
    rowIndex = FindUserRowForm(loUsers, userId)
    If rowIndex > 0 Then loUsers.ListRows(rowIndex).Delete

    If Not loCaps Is Nothing Then
        For i = loCaps.ListRows.Count To 1 Step -1
            If StrComp(SafeTextForm(GetCellByColumnForm(loCaps, i, "UserId")), userId, vbTextCompare) = 0 Then
                loCaps.ListRows(i).Delete
            End If
        Next i
    End If

    wb.Save
    ClearUserFieldsForm
    ShowStatusForm "User deleted.", COLOR_SUCCESS
    RefreshUsersForm

CleanExit:
    CloseTransientWorkbookForm wb, openedTransient
End Sub

Private Function OpenSelectedAuthWorkbookForm(ByRef openedTransient As Boolean, ByRef report As String) As Workbook
    Dim authPath As String
    Dim wb As Workbook
    Dim whId As String

    authPath = NormalizePathForm(CStr(mTxtAuthPath.Value))
    whId = Trim$(CStr(mTxtWarehouseId.Value))
    If authPath = "" Then
        report = "Select a warehouse first."
        Exit Function
    End If
    If Not FileExistsForm(authPath) Then
        report = "Auth workbook was not found: " & authPath
        Exit Function
    End If

    For Each wb In Application.Workbooks
        If StrComp(NormalizePathForm(wb.FullName), authPath, vbTextCompare) = 0 Then
            Set OpenSelectedAuthWorkbookForm = wb
            openedTransient = False
            Exit Function
        End If
    Next wb

    On Error GoTo FailOpen
    Set wb = Application.Workbooks.Open(Filename:=authPath, UpdateLinks:=0, ReadOnly:=False, AddToMru:=False)
    openedTransient = True
    If Not modAuth.EnsureAuthSchema(wb, whId, "svc_processor", report) Then GoTo FailSoft
    Set OpenSelectedAuthWorkbookForm = wb
    Exit Function

FailSoft:
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Set wb = Nothing
    Exit Function

FailOpen:
    report = "Unable to open auth workbook: " & Err.Description
End Function

Private Sub CloseTransientWorkbookForm(ByVal wb As Workbook, ByVal openedTransient As Boolean)
    If wb Is Nothing Or Not openedTransient Then Exit Sub
    On Error Resume Next
    wb.Close SaveChanges:=False
    On Error GoTo 0
End Sub

Private Function EnsureUserRowForm(ByVal loUsers As ListObject, ByVal userId As String) As Long
    Dim rowIndex As Long
    Dim newRow As ListRow

    rowIndex = FindUserRowForm(loUsers, userId)
    If rowIndex > 0 Then
        EnsureUserRowForm = rowIndex
        Exit Function
    End If

    Set newRow = loUsers.ListRows.Add
    EnsureUserRowForm = newRow.Index
End Function

Private Sub SaveCapabilityChoiceForm(ByVal loCaps As ListObject, ByVal userId As String, _
                                     ByVal capability As String, ByVal warehouseId As String, _
                                     ByVal stationId As String, ByVal selected As Boolean)
    Dim rowIndex As Long

    rowIndex = FindCapabilityRowForm(loCaps, userId, capability, warehouseId, stationId)
    If selected Then
        If rowIndex = 0 Then rowIndex = loCaps.ListRows.Add.Index
        SetCellByColumnForm loCaps, rowIndex, "UserId", userId
        SetCellByColumnForm loCaps, rowIndex, "Capability", capability
        SetCellByColumnForm loCaps, rowIndex, "WarehouseId", warehouseId
        SetCellByColumnForm loCaps, rowIndex, "StationId", stationId
        SetCellByColumnForm loCaps, rowIndex, "Status", "ACTIVE"
    ElseIf rowIndex > 0 Then
        SetCellByColumnForm loCaps, rowIndex, "Status", "INACTIVE"
    End If
End Sub

Private Function FindUserRowForm(ByVal loUsers As ListObject, ByVal userId As String) As Long
    Dim i As Long
    If loUsers Is Nothing Then Exit Function
    If loUsers.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To loUsers.ListRows.Count
        If StrComp(SafeTextForm(GetCellByColumnForm(loUsers, i, "UserId")), userId, vbTextCompare) = 0 Then
            FindUserRowForm = i
            Exit Function
        End If
    Next i
End Function

Private Function FindCapabilityRowForm(ByVal loCaps As ListObject, ByVal userId As String, _
                                       ByVal capability As String, ByVal warehouseId As String, _
                                       ByVal stationId As String) As Long
    Dim i As Long
    If loCaps Is Nothing Then Exit Function
    If loCaps.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To loCaps.ListRows.Count
        If StrComp(SafeTextForm(GetCellByColumnForm(loCaps, i, "UserId")), userId, vbTextCompare) <> 0 Then GoTo ContinueLoop
        If StrComp(SafeTextForm(GetCellByColumnForm(loCaps, i, "Capability")), capability, vbTextCompare) <> 0 Then GoTo ContinueLoop
        If StrComp(SafeTextForm(GetCellByColumnForm(loCaps, i, "WarehouseId")), warehouseId, vbTextCompare) <> 0 Then GoTo ContinueLoop
        If StrComp(SafeTextForm(GetCellByColumnForm(loCaps, i, "StationId")), stationId, vbTextCompare) <> 0 Then GoTo ContinueLoop
        FindCapabilityRowForm = i
        Exit Function
ContinueLoop:
    Next i
End Function

Private Function HasActiveCapabilityForm(ByVal loCaps As ListObject, ByVal userId As String, _
                                         ByVal capability As String, ByVal warehouseId As String, _
                                         ByVal stationId As String) As Boolean
    Dim i As Long
    Dim statusText As String
    If loCaps Is Nothing Then Exit Function
    If loCaps.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To loCaps.ListRows.Count
        If StrComp(SafeTextForm(GetCellByColumnForm(loCaps, i, "UserId")), userId, vbTextCompare) <> 0 Then GoTo ContinueLoop
        If StrComp(SafeTextForm(GetCellByColumnForm(loCaps, i, "Capability")), capability, vbTextCompare) <> 0 Then GoTo ContinueLoop
        If Not ScopeMatchesForm(SafeTextForm(GetCellByColumnForm(loCaps, i, "WarehouseId")), warehouseId) Then GoTo ContinueLoop
        If Not ScopeMatchesForm(SafeTextForm(GetCellByColumnForm(loCaps, i, "StationId")), stationId) Then GoTo ContinueLoop
        statusText = UCase$(SafeTextForm(GetCellByColumnForm(loCaps, i, "Status")))
        If statusText = "" Or statusText = "ACTIVE" Or statusText = "ALLOW" Then
            HasActiveCapabilityForm = True
            Exit Function
        End If
ContinueLoop:
    Next i
End Function

Private Function ScopeMatchesForm(ByVal rowScope As String, ByVal requestedScope As String) As Boolean
    rowScope = Trim$(rowScope)
    requestedScope = Trim$(requestedScope)
    ScopeMatchesForm = (rowScope = "" Or rowScope = "*" Or requestedScope = "" Or requestedScope = "*" Or _
                        StrComp(rowScope, requestedScope, vbTextCompare) = 0)
End Function

Private Function FindListObjectForm(ByVal wb As Workbook, ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Function
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                Set FindListObjectForm = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function ColumnIndexForm(ByVal lo As ListObject, ByVal columnName As String) As Long
    On Error Resume Next
    ColumnIndexForm = lo.ListColumns(columnName).Index
    On Error GoTo 0
End Function

Private Function GetCellByColumnForm(ByVal lo As ListObject, ByVal rowIndex As Long, ByVal columnName As String) As Variant
    Dim colIndex As Long
    colIndex = ColumnIndexForm(lo, columnName)
    If colIndex = 0 Then Exit Function
    GetCellByColumnForm = lo.DataBodyRange.Cells(rowIndex, colIndex).Value
End Function

Private Sub SetCellByColumnForm(ByVal lo As ListObject, ByVal rowIndex As Long, _
                                ByVal columnName As String, ByVal valueText As Variant)
    Dim colIndex As Long
    colIndex = ColumnIndexForm(lo, columnName)
    If colIndex = 0 Then Exit Sub
    lo.DataBodyRange.Cells(rowIndex, colIndex).Value = valueText
End Sub

Private Sub ClearUserFieldsForm()
    mTxtUserId.Value = ""
    mTxtDisplayName.Value = ""
    mTxtPin.Value = ""
    mChkAdmin.Value = False
    mChkReceivePost.Value = False
    mChkReceiveView.Value = False
    mChkShipPost.Value = False
    mChkProdPost.Value = False
    mChkInboxProcess.Value = False
    mLstUsers.ListIndex = -1
End Sub

Private Function ConnectSelectedRootForm(ByRef report As String) As Boolean
    Dim rootPath As String
    Dim shareRoot As String
    Dim userName As String
    Dim passwordText As String
    Dim resultCode As Long

    rootPath = NormalizePathForm(CStr(mTxtRoot.Value))
    If rootPath = "" Then
        report = "Warehouse root is required."
        Exit Function
    End If
    If Not IsUncPathForm(rootPath) Then
        report = "Warehouse root is local; NAS credentials are not needed."
        ConnectSelectedRootForm = True
        Exit Function
    End If

    shareRoot = ResolveUncShareRootForm(rootPath)
    If shareRoot = "" Then
        report = "Warehouse root must include a UNC server and share."
        Exit Function
    End If

    userName = Trim$(CStr(mTxtNasUser.Value))
    passwordText = CStr(mTxtNasPassword.Value)
    If userName = "" Or passwordText = "" Then
        report = "Enter NAS username and password, then click Connect."
        Exit Function
    End If

    resultCode = ConnectToShareWithCredentialsForm(shareRoot, userName, passwordText)
    If resultCode = ERROR_SESSION_CREDENTIAL_CONFLICT Then
        WNetCancelConnection2 shareRoot, 0, True
        resultCode = ConnectToShareWithCredentialsForm(shareRoot, userName, passwordText)
    End If

    If resultCode = NO_ERROR_WIN32 And FolderExistsForm(rootPath) Then
        mTxtNasPassword.Value = ""
        ConnectSelectedRootForm = True
        RememberWarehouseRootForAdminForm rootPath
        report = "Connected to NAS root."
    ElseIf resultCode = NO_ERROR_WIN32 Then
        report = "Connected to NAS share, but the Warehouse root folder was not found."
    ElseIf resultCode = ERROR_SESSION_CREDENTIAL_CONFLICT And FolderExistsForm(rootPath) Then
        mTxtNasPassword.Value = ""
        ConnectSelectedRootForm = True
        RememberWarehouseRootForAdminForm rootPath
        report = "Using existing Windows NAS connection."
    Else
        report = WNetConnectionErrorTextForm(resultCode)
    End If
End Function

Private Sub RememberWarehouseRootForAdminForm(ByVal rootPath As String)
    On Error Resume Next
    modAdminConsole.RememberWarehouseScanRoot rootPath
    On Error GoTo 0
End Sub

Private Function ConnectToShareWithCredentialsForm(ByVal shareRoot As String, ByVal userName As String, _
                                                   ByVal passwordText As String) As Long
    Dim resource As NETRESOURCE
    Dim qualifiedUser As String

    resource.dwType = RESOURCETYPE_DISK
    resource.lpRemoteName = shareRoot

    ConnectToShareWithCredentialsForm = WNetAddConnection2(resource, passwordText, userName, CONNECT_TEMPORARY)
    If ConnectToShareWithCredentialsForm = NO_ERROR_WIN32 Then Exit Function
    If InStr(1, userName, "\", vbTextCompare) > 0 Or InStr(1, userName, "@", vbTextCompare) > 0 Then Exit Function

    qualifiedUser = ResolveUncServerForm(shareRoot) & "\" & userName
    If qualifiedUser <> "\" & userName Then
        ConnectToShareWithCredentialsForm = WNetAddConnection2(resource, passwordText, qualifiedUser, CONNECT_TEMPORARY)
    End If
End Function

Private Function ResolveUncServerForm(ByVal pathValue As String) As String
    Dim parts() As String
    pathValue = NormalizePathForm(pathValue)
    If Not IsUncPathForm(pathValue) Then Exit Function
    parts = Split(Mid$(pathValue, 3), "\")
    If UBound(parts) >= 0 Then ResolveUncServerForm = parts(0)
End Function

Private Function ResolveUncShareRootForm(ByVal pathValue As String) As String
    Dim parts() As String
    pathValue = NormalizePathForm(pathValue)
    If Not IsUncPathForm(pathValue) Then Exit Function
    parts = Split(Mid$(pathValue, 3), "\")
    If UBound(parts) < 1 Then Exit Function
    ResolveUncShareRootForm = "\\" & parts(0) & "\" & parts(1)
End Function

Private Function WNetConnectionErrorTextForm(ByVal resultCode As Long) As String
    Select Case resultCode
        Case 5
            WNetConnectionErrorTextForm = "NAS access denied. Check username permissions for this share."
        Case 53
            WNetConnectionErrorTextForm = "NAS path was not found. Check the server, share name, and network connection."
        Case 67
            WNetConnectionErrorTextForm = "NAS network name was not found. Check the share name."
        Case 86, 1326
            WNetConnectionErrorTextForm = "NAS login failed. Check username and password."
        Case ERROR_SESSION_CREDENTIAL_CONFLICT
            WNetConnectionErrorTextForm = "Windows already has a connection to this NAS with different credentials. Disconnect that NAS session in Windows, then connect again."
        Case Else
            WNetConnectionErrorTextForm = "NAS connection failed. Windows error " & CStr(resultCode) & "."
    End Select
End Function

Private Function IsUncPathForm(ByVal pathValue As String) As Boolean
    pathValue = Trim$(pathValue)
    IsUncPathForm = (Len(pathValue) >= 3 And Left$(pathValue, 2) = "\\")
End Function

Private Function ResolveDefaultWarehouseRootForm() As String
    Dim rootPath As String
    On Error Resume Next
    rootPath = Trim$(modRuntimeWorkbooks.GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = Trim$(modRuntimeWorkbooks.ResolveCoreDataRoot("", ""))
    If rootPath = "" Then rootPath = modDeploymentPaths.DefaultRuntimeHubRootPath(False)
    On Error GoTo 0
    ResolveDefaultWarehouseRootForm = NormalizePathForm(rootPath)
End Function

Private Function NormalizePathForm(ByVal pathValue As String) As String
    pathValue = Trim$(pathValue)
    If Len(pathValue) > 3 Then
        Do While Right$(pathValue, 1) = "\" Or Right$(pathValue, 1) = "/"
            pathValue = Left$(pathValue, Len(pathValue) - 1)
        Loop
    End If
    NormalizePathForm = Replace(pathValue, "/", "\")
End Function

Private Function FolderExistsForm(ByVal folderPath As String) As Boolean
    Dim fso As Object
    Dim dirResult As String

    folderPath = NormalizePathForm(folderPath)
    If folderPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsForm = fso.FolderExists(folderPath)
    If Err.Number <> 0 Then Err.Clear
    If FolderExistsForm Then GoTo CleanExit

    dirResult = Dir$(folderPath, vbDirectory)
    If Err.Number <> 0 Then
        Err.Clear
    Else
        FolderExistsForm = (Len(dirResult) > 0)
    End If
CleanExit:
    On Error GoTo 0
End Function

Private Function FileExistsForm(ByVal filePath As String) As Boolean
    Dim fso As Object
    Dim dirResult As String

    filePath = NormalizePathForm(filePath)
    If filePath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsForm = fso.FileExists(filePath)
    If Err.Number <> 0 Then Err.Clear
    If FileExistsForm Then GoTo CleanExit

    dirResult = Dir$(filePath, vbNormal)
    If Err.Number <> 0 Then
        Err.Clear
    Else
        FileExistsForm = (Len(dirResult) > 0)
    End If
CleanExit:
    On Error GoTo 0
End Function

Private Function SafeTextForm(ByVal valueIn As Variant) As String
    If IsError(valueIn) Or IsNull(valueIn) Or IsEmpty(valueIn) Then Exit Function
    SafeTextForm = Trim$(CStr(valueIn))
End Function

Private Function BuildAccountClipboardTextForm() As String
    Dim userId As String
    Dim displayName As String
    Dim pinText As String
    Dim warehouseId As String
    Dim stationId As String
    Dim textOut As String

    userId = Trim$(CStr(mTxtUserId.Value))
    displayName = Trim$(CStr(mTxtDisplayName.Value))
    pinText = CStr(mTxtPin.Value)
    warehouseId = Trim$(CStr(mTxtWarehouseId.Value))
    stationId = Trim$(CStr(mTxtStationId.Value))

    If userId = "" And displayName = "" And pinText = "" And warehouseId = "" And stationId = "" Then Exit Function
    If stationId = "" Then stationId = "*"

    textOut = "invSys user account" & vbCrLf & _
              "User ID: " & userId & vbCrLf & _
              "Display name: " & displayName & vbCrLf & _
              "PIN/password: " & pinText & vbCrLf & _
              "Warehouse scope: " & warehouseId & vbCrLf & _
              "Station: " & stationId & vbCrLf & _
              "Roles / capabilities:" & vbCrLf & _
              "- Admin maintenance: " & YesNoTextForm(CBool(mChkAdmin.Value)) & vbCrLf & _
              "- Receiving post: " & YesNoTextForm(CBool(mChkReceivePost.Value)) & vbCrLf & _
              "- Receiving view: " & YesNoTextForm(CBool(mChkReceiveView.Value)) & vbCrLf & _
              "- Shipping post: " & YesNoTextForm(CBool(mChkShipPost.Value)) & vbCrLf & _
              "- Production post: " & YesNoTextForm(CBool(mChkProdPost.Value)) & vbCrLf & _
              "- Inbox processor: " & YesNoTextForm(CBool(mChkInboxProcess.Value))

    BuildAccountClipboardTextForm = textOut
End Function

Private Function YesNoTextForm(ByVal selected As Boolean) As String
    If selected Then
        YesNoTextForm = "Yes"
    Else
        YesNoTextForm = "No"
    End If
End Function

Private Function CopyTextToClipboardForm(ByVal valueText As String) As Boolean
    On Error GoTo FailCopy

    Dim dataObj As MSForms.DataObject
    Set dataObj = New MSForms.DataObject
    dataObj.SetText valueText
    dataObj.PutInClipboard
    CopyTextToClipboardForm = True
    Exit Function

FailCopy:
    CopyTextToClipboardForm = False
End Function

Private Sub ShowStatusForm(ByVal messageText As String, ByVal colorValue As Long)
    If mLblStatus Is Nothing Then Exit Sub
    mLblStatus.Caption = messageText
    mLblStatus.ForeColor = colorValue
End Sub
