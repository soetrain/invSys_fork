VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRetireMigrateWarehouse 
   Caption         =   "Retire / Migrate Warehouse"
   ClientHeight    =   4200
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   6400
   OleObjectBlob   =   "frmRetireMigrateWarehouse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmRetireMigrateWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mTxtWarehouseRootPath As MSForms.TextBox
Attribute mTxtWarehouseRootPath.VB_VarHelpID = -1
Private WithEvents mBtnWarehouseRootBrowse As MSForms.CommandButton
Attribute mBtnWarehouseRootBrowse.VB_VarHelpID = -1
Private WithEvents mBtnWarehouseRootRefresh As MSForms.CommandButton
Attribute mBtnWarehouseRootRefresh.VB_VarHelpID = -1
Private WithEvents mBtnArchiveDestBrowse As MSForms.CommandButton
Attribute mBtnArchiveDestBrowse.VB_VarHelpID = -1
Private WithEvents mTxtNasUser As MSForms.TextBox
Attribute mTxtNasUser.VB_VarHelpID = -1
Private WithEvents mTxtNasPassword As MSForms.TextBox
Attribute mTxtNasPassword.VB_VarHelpID = -1
Private WithEvents mBtnNasConnect As MSForms.CommandButton
Attribute mBtnNasConnect.VB_VarHelpID = -1
Private WithEvents mLstSourceWarehouses As MSForms.ListBox
Attribute mLstSourceWarehouses.VB_VarHelpID = -1
Private WithEvents mOptArchiveOnly As MSForms.OptionButton
Attribute mOptArchiveOnly.VB_VarHelpID = -1
Private WithEvents mOptArchiveMigrate As MSForms.OptionButton
Attribute mOptArchiveMigrate.VB_VarHelpID = -1
Private WithEvents mOptArchiveRetire As MSForms.OptionButton
Attribute mOptArchiveRetire.VB_VarHelpID = -1
Private WithEvents mOptArchiveRetireDelete As MSForms.OptionButton
Attribute mOptArchiveRetireDelete.VB_VarHelpID = -1
Private mLblWarehouseRoot As MSForms.Label
Private mLblWarehouseRootError As MSForms.Label
Private mLblNasUser As MSForms.Label
Private mLblNasPassword As MSForms.Label
Private mLblFoundWarehouses As MSForms.Label

Private Const PANEL_SELECTION As String = "SELECTION"
Private Const PANEL_CONFIRM As String = "CONFIRM"
Private Const PANEL_RESULT As String = "RESULT"

Private Const COLOR_ERROR As Long = 255
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_INFO As Long = 0
Private Const COLOR_WARNING As Long = 192
Private Const ANCHOR_LEFT As Long = 1
Private Const ANCHOR_TOP As Long = 2
Private Const ANCHOR_RIGHT As Long = 4
Private Const ANCHOR_BOTTOM As Long = 8
Private Const NO_ERROR_WIN32 As Long = 0
Private Const ERROR_SESSION_CREDENTIAL_CONFLICT As Long = 1219
Private Const RESOURCETYPE_DISK As Long = 1
Private Const CONNECT_TEMPORARY As Long = 4

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
#End If

Private mFormBusy As Boolean
Private mCurrentPanel As String
Private mReAuthPassed As Boolean
Private mPendingSourceWarehouseId As String
Private mPendingTargetWarehouseId As String
Private mPendingOperationMode As Long
Private mPendingAdminUser As String
Private mPendingArchiveDestPath As String
Private mPendingPublishTombstone As Boolean
Private mPendingWarehouseRootPath As String
Private mWarehousePathById As Object
Private mAnchors As Object
Private mResizeInitialized As Boolean

Private Sub UserForm_Initialize()
    mFormBusy = True
    Me.Caption = "Retire / Migrate Warehouse"
    Me.StartUpPosition = 1
    ConfigureRetireMigrateLayout
    InitializeRetireMigrateAnchors

    SetOperationModeValue modWarehouseRetire.MODE_ARCHIVE_RETIRE
    Me.chkPublishTombstone.Value = True
    Me.chkConfirmAction.Value = False
    Me.lblDeleteWarning.ForeColor = COLOR_ERROR

    ClearAllInlineErrors
    ApplyWarehouseRootDefault
    PopulateWarehouseDropdowns
    ApplyDefaultSelections
    ShowSelectionPanel
    ShowFormMessage "Select a source warehouse and operation mode, then click OK.", COLOR_INFO

    mFormBusy = False
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

Private Sub ConfigureRetireMigrateLayout()
    Dim ctl As Control
    Dim confirmLeft As Single
    Dim confirmTop As Single
    Dim resultLeft As Single
    Dim resultTop As Single

    Me.Width = 620
    Me.Height = 620
    On Error Resume Next
    Me.ScrollBars = 0
    Me.KeepScrollBarsVisible = 0
    Me.ScrollLeft = 0
    Me.ScrollTop = 0
    On Error GoTo 0

    For Each ctl In Me.Controls
        ctl.Visible = False
    Next ctl
    ConfigureWarehouseRootControls
    ConfigureNasCredentialControls
    ConfigureWarehouseListControls
    ConfigureArchiveDestinationControls
    HideDesignerOperationModeControls

    Me.btnBack.Visible = True
    Me.btnCancel.Visible = True
    Me.btnOK.Visible = True

    Me.lblTitle.Left = 18
    Me.lblTitle.Top = 18
    Me.lblTitle.Width = 564

    Me.lblSelectionIntro.Left = 18
    Me.lblSelectionIntro.Top = 44
    Me.lblSelectionIntro.Width = 584
    Me.lblSelectionIntro.Height = 30

    mLblWarehouseRoot.Left = 18
    mLblWarehouseRoot.Top = 82
    mLblWarehouseRoot.Width = 132
    mTxtWarehouseRootPath.Left = 164
    mTxtWarehouseRootPath.Top = 78
    mTxtWarehouseRootPath.Width = 320
    mBtnWarehouseRootBrowse.Left = 490
    mBtnWarehouseRootBrowse.Top = 77
    mBtnWarehouseRootBrowse.Width = 50
    mBtnWarehouseRootBrowse.Height = mTxtWarehouseRootPath.Height + 2
    mBtnWarehouseRootRefresh.Left = 546
    mBtnWarehouseRootRefresh.Top = 77
    mBtnWarehouseRootRefresh.Width = 42
    mBtnWarehouseRootRefresh.Height = mTxtWarehouseRootPath.Height + 2
    mLblWarehouseRootError.Left = 164
    mLblWarehouseRootError.Top = 100
    mLblWarehouseRootError.Width = 424

    mLblNasUser.Left = 18
    mLblNasUser.Top = 120
    mLblNasUser.Width = 132
    mTxtNasUser.Left = 164
    mTxtNasUser.Top = 116
    mTxtNasUser.Width = 150
    mLblNasPassword.Left = 324
    mLblNasPassword.Top = 120
    mLblNasPassword.Width = 58
    mTxtNasPassword.Left = 388
    mTxtNasPassword.Top = 116
    mTxtNasPassword.Width = 96
    mBtnNasConnect.Left = 490
    mBtnNasConnect.Top = 115
    mBtnNasConnect.Width = 58
    mBtnNasConnect.Height = mTxtNasPassword.Height + 2

    Me.lblSourceWarehouse.Left = 18
    Me.lblSourceWarehouse.Top = 154
    Me.lblSourceWarehouse.Width = 132
    Me.lblSourceWarehouse.Caption = "Source warehouse"

    Me.cmbSourceWarehouse.Left = 164
    Me.cmbSourceWarehouse.Top = 150
    Me.cmbSourceWarehouse.Width = 220

    Me.lblSourceWarehouseError.Left = 164
    Me.lblSourceWarehouseError.Top = 172
    Me.lblSourceWarehouseError.Width = 424

    Me.lblTargetWarehouse.Left = 18
    Me.lblTargetWarehouse.Top = 186
    Me.lblTargetWarehouse.Width = 132
    Me.lblTargetWarehouse.Caption = "Target warehouse"

    Me.cmbTargetWarehouse.Left = 164
    Me.cmbTargetWarehouse.Top = 182
    Me.cmbTargetWarehouse.Width = 220
    Me.cmbTargetWarehouse.ControlTipText = "Migration target warehouse. Create the target warehouse first, then scan the root and select it here."

    Me.lblTargetWarehouseError.Left = 164
    Me.lblTargetWarehouseError.Top = 204
    Me.lblTargetWarehouseError.Width = 424

    mLblFoundWarehouses.Left = 18
    mLblFoundWarehouses.Top = 220
    mLblFoundWarehouses.Width = 132
    mLstSourceWarehouses.Left = 164
    mLstSourceWarehouses.Top = 214
    mLstSourceWarehouses.Width = 424
    mLstSourceWarehouses.Height = 112

    Me.fraMode.Left = 18
    Me.fraMode.Top = 340
    Me.fraMode.Width = 570
    Me.fraMode.Height = 86
    ConfigureOperationModeControls

    Me.lblArchiveDestPath.Left = 18
    Me.lblArchiveDestPath.Top = 438
    Me.lblArchiveDestPath.Width = 132
    Me.lblArchiveDestPath.Caption = "Archive location"
    Me.txtArchiveDestPath.Left = 164
    Me.txtArchiveDestPath.Top = 434
    Me.txtArchiveDestPath.Width = 320
    mBtnArchiveDestBrowse.Left = 490
    mBtnArchiveDestBrowse.Top = 433
    mBtnArchiveDestBrowse.Width = 50
    mBtnArchiveDestBrowse.Height = Me.txtArchiveDestPath.Height + 2
    Me.lblArchiveDestPathError.Left = 164
    Me.lblArchiveDestPathError.Top = 456
    Me.lblArchiveDestPathError.Width = 424

    Me.chkPublishTombstone.Left = 164
    Me.chkPublishTombstone.Top = 472
    Me.chkPublishTombstone.Width = 424
    Me.lblReAuthError.Left = 164
    Me.lblReAuthError.Top = 492
    Me.lblReAuthError.Width = 424
    Me.lblDeleteWarning.Left = 164
    Me.lblDeleteWarning.Top = 510
    Me.lblDeleteWarning.Width = 424
    Me.lblDeleteWarning.Height = 20

    Me.fraConfirm.Left = 18
    Me.fraConfirm.Top = 70
    Me.fraConfirm.Width = 570
    Me.fraConfirm.Height = 430
    confirmLeft = Me.fraConfirm.Left + 12
    confirmTop = Me.fraConfirm.Top + 16
    Me.lblConfirmSummary.Left = confirmLeft
    Me.lblConfirmSummary.Top = confirmTop
    Me.lblConfirmSummary.Width = 534
    Me.lblConfirmSummary.Height = 318
    Me.chkConfirmAction.Left = confirmLeft
    Me.chkConfirmAction.Top = confirmTop + 332
    Me.chkConfirmAction.Width = 300
    Me.lblConfirmError.Left = confirmLeft
    Me.lblConfirmError.Top = confirmTop + 358
    Me.lblConfirmError.Width = 534

    Me.fraResult.Left = 18
    Me.fraResult.Top = 70
    Me.fraResult.Width = 570
    Me.fraResult.Height = 430
    resultLeft = Me.fraResult.Left + 12
    resultTop = Me.fraResult.Top + 16
    Me.lblResultSummary.Left = resultLeft
    Me.lblResultSummary.Top = resultTop
    Me.lblResultSummary.Width = 534
    Me.lblResultSummary.Height = 392

    Me.btnBack.Left = 304
    Me.btnBack.Top = 564
    Me.btnBack.Width = 88
    Me.btnCancel.Left = 400
    Me.btnCancel.Top = 564
    Me.btnCancel.Width = 88
    Me.btnOK.Left = 496
    Me.btnOK.Top = 564
    Me.btnOK.Width = 88

    ConfigureWrappedLabel Me.lblSelectionIntro
    ConfigureWrappedLabel Me.lblDeleteWarning
    ConfigureWrappedLabel Me.lblConfirmSummary
    ConfigureWrappedLabel Me.lblConfirmError
    ConfigureWrappedLabel Me.lblResultSummary
End Sub

Private Sub ConfigureWarehouseRootControls()
    If mLblWarehouseRoot Is Nothing Then
        Set mLblWarehouseRoot = Me.Controls.Add("Forms.Label.1", "lblWarehouseRootRuntime", True)
    End If
    With mLblWarehouseRoot
        .Caption = "Warehouse root"
        .Visible = True
    End With

    If mTxtWarehouseRootPath Is Nothing Then
        Set mTxtWarehouseRootPath = Me.Controls.Add("Forms.TextBox.1", "txtWarehouseRootPathRuntime", True)
    End If
    With mTxtWarehouseRootPath
        .ControlTipText = "NAS hub folder or a specific warehouse runtime folder. Examples: \\DS920\invSysWH1 or \\DS920\invSysWH1\WH1."
        .Visible = True
    End With

    If mBtnWarehouseRootBrowse Is Nothing Then
        Set mBtnWarehouseRootBrowse = Me.Controls.Add("Forms.CommandButton.1", "btnWarehouseRootBrowseRuntime", True)
    End If
    With mBtnWarehouseRootBrowse
        .Caption = "Find..."
        .ControlTipText = "Choose the NAS hub folder or local warehouse runtime folder."
        .Visible = True
        .Enabled = True
    End With

    If mBtnWarehouseRootRefresh Is Nothing Then
        Set mBtnWarehouseRootRefresh = Me.Controls.Add("Forms.CommandButton.1", "btnWarehouseRootRefreshRuntime", True)
    End If
    With mBtnWarehouseRootRefresh
        .Caption = "Scan"
        .ControlTipText = "Scan the selected root and refresh the warehouse lists."
        .Visible = True
        .Enabled = True
    End With

    If mLblWarehouseRootError Is Nothing Then
        Set mLblWarehouseRootError = Me.Controls.Add("Forms.Label.1", "lblWarehouseRootErrorRuntime", True)
    End If
    With mLblWarehouseRootError
        .Caption = ""
        .ForeColor = COLOR_ERROR
        .Visible = True
    End With
End Sub

Private Sub ConfigureNasCredentialControls()
    If mLblNasUser Is Nothing Then
        Set mLblNasUser = Me.Controls.Add("Forms.Label.1", "lblNasUserRuntime", True)
    End If
    With mLblNasUser
        .Caption = "NAS user"
        .Visible = True
    End With

    If mTxtNasUser Is Nothing Then
        Set mTxtNasUser = Me.Controls.Add("Forms.TextBox.1", "txtNasUserRuntime", True)
    End If
    With mTxtNasUser
        .ControlTipText = "NAS account for the Warehouse root. Examples: user, NAS\user, or 100.84.136.19\user."
        .Visible = True
    End With

    If mLblNasPassword Is Nothing Then
        Set mLblNasPassword = Me.Controls.Add("Forms.Label.1", "lblNasPasswordRuntime", True)
    End If
    With mLblNasPassword
        .Caption = "Password"
        .Visible = True
    End With

    If mTxtNasPassword Is Nothing Then
        Set mTxtNasPassword = Me.Controls.Add("Forms.TextBox.1", "txtNasPasswordRuntime", True)
    End If
    With mTxtNasPassword
        .PasswordChar = "*"
        .ControlTipText = "Password is used only to connect this Windows session to the NAS."
        .Visible = True
    End With

    If mBtnNasConnect Is Nothing Then
        Set mBtnNasConnect = Me.Controls.Add("Forms.CommandButton.1", "btnNasConnectRuntime", True)
    End If
    With mBtnNasConnect
        .Caption = "Connect"
        .ControlTipText = "Connect to the NAS share using the entered credentials, then scan."
        .Visible = True
        .Enabled = True
    End With
End Sub

Private Sub ConfigureWarehouseListControls()
    If mLblFoundWarehouses Is Nothing Then
        Set mLblFoundWarehouses = Me.Controls.Add("Forms.Label.1", "lblFoundWarehousesRuntime", True)
    End If
    With mLblFoundWarehouses
        .Caption = "Warehouses in root"
        .Visible = True
    End With

    If mLstSourceWarehouses Is Nothing Then
        Set mLstSourceWarehouses = Me.Controls.Add("Forms.ListBox.1", "lstSourceWarehousesRuntime", True)
    End If
    With mLstSourceWarehouses
        .ColumnCount = 2
        .ColumnWidths = "90 pt;310 pt"
        .IntegralHeight = False
        .MultiSelect = 0
        .ControlTipText = "Warehouses found under the selected NAS or local warehouse root. Columns show WarehouseId and folder path."
        .Visible = True
    End With
End Sub

Private Sub ConfigureArchiveDestinationControls()
    If mBtnArchiveDestBrowse Is Nothing Then
        Set mBtnArchiveDestBrowse = Me.Controls.Add("Forms.CommandButton.1", "btnArchiveDestBrowseRuntime", True)
    End If
    With mBtnArchiveDestBrowse
        .Caption = "Find..."
        .ControlTipText = "Choose where the retire/archive package will be written."
        .Visible = True
        .Enabled = True
    End With
End Sub

Private Sub ConfigureOperationModeControls()
    Set mOptArchiveOnly = EnsureOperationModeOption("optArchiveOnlyRuntime", "Archive package only", 12, 88)
    Set mOptArchiveRetire = EnsureOperationModeOption("optArchiveRetireRuntime", "Retire / archive warehouse", 12, 16)
    Set mOptArchiveMigrate = EnsureOperationModeOption("optArchiveMigrateRuntime", "Migrate inventory to target warehouse", 12, 34)
    Set mOptArchiveRetireDelete = EnsureOperationModeOption("optArchiveRetireDeleteRuntime", "Retire / archive and delete source files", 12, 52)
    mOptArchiveOnly.Visible = False
End Sub

Private Function EnsureOperationModeOption(ByVal controlName As String, _
                                           ByVal captionText As String, _
                                           ByVal leftPos As Single, _
                                           ByVal topPos As Single) As MSForms.OptionButton
    Dim opt As MSForms.OptionButton

    On Error Resume Next
    Set opt = Me.fraMode.Controls(controlName)
    On Error GoTo 0
    If opt Is Nothing Then
        Set opt = Me.fraMode.Controls.Add("Forms.OptionButton.1", controlName, True)
    End If

    With opt
        .Caption = captionText
        .GroupName = "RetireMigrateOperationMode"
        .Left = leftPos
        .Top = topPos
        .Width = 360
        .Visible = True
        .Enabled = True
    End With
    Set EnsureOperationModeOption = opt
End Function

Private Sub HideDesignerOperationModeControls()
    On Error Resume Next
    Me.optArchiveOnly.Visible = False
    Me.optArchiveMigrate.Visible = False
    Me.optArchiveRetire.Visible = False
    Me.optArchiveRetireDelete.Visible = False
    On Error GoTo 0
End Sub

Private Sub ConfigureWrappedLabel(ByVal lbl As MSForms.Label)
    If lbl Is Nothing Then Exit Sub

    lbl.WordWrap = True
    lbl.AutoSize = False
End Sub

Private Sub InitializeRetireMigrateAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me, 620, 620

    mAnchors.Add Me.lblSelectionIntro, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    If Not mLblWarehouseRoot Is Nothing Then mAnchors.Add mLblWarehouseRoot, ANCHOR_LEFT Or ANCHOR_TOP
    If Not mTxtWarehouseRootPath Is Nothing Then mAnchors.Add mTxtWarehouseRootPath, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    If Not mBtnWarehouseRootBrowse Is Nothing Then mAnchors.Add mBtnWarehouseRootBrowse, ANCHOR_RIGHT Or ANCHOR_TOP
    If Not mBtnWarehouseRootRefresh Is Nothing Then mAnchors.Add mBtnWarehouseRootRefresh, ANCHOR_RIGHT Or ANCHOR_TOP
    If Not mLblWarehouseRootError Is Nothing Then mAnchors.Add mLblWarehouseRootError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    If Not mLblNasUser Is Nothing Then mAnchors.Add mLblNasUser, ANCHOR_LEFT Or ANCHOR_TOP
    If Not mTxtNasUser Is Nothing Then mAnchors.Add mTxtNasUser, ANCHOR_LEFT Or ANCHOR_TOP
    If Not mLblNasPassword Is Nothing Then mAnchors.Add mLblNasPassword, ANCHOR_TOP Or ANCHOR_RIGHT
    If Not mTxtNasPassword Is Nothing Then mAnchors.Add mTxtNasPassword, ANCHOR_TOP Or ANCHOR_RIGHT
    If Not mBtnNasConnect Is Nothing Then mAnchors.Add mBtnNasConnect, ANCHOR_RIGHT Or ANCHOR_TOP
    mAnchors.Add Me.cmbSourceWarehouse, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblSourceWarehouseError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.cmbTargetWarehouse, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblTargetWarehouseError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    If Not mLblFoundWarehouses Is Nothing Then mAnchors.Add mLblFoundWarehouses, ANCHOR_LEFT Or ANCHOR_TOP
    If Not mLstSourceWarehouses Is Nothing Then mAnchors.Add mLstSourceWarehouses, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.fraMode, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblArchiveDestPath, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add Me.txtArchiveDestPath, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    If Not mBtnArchiveDestBrowse Is Nothing Then mAnchors.Add mBtnArchiveDestBrowse, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.lblArchiveDestPathError, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.chkPublishTombstone, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add Me.lblReAuthError, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.lblDeleteWarning, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.fraConfirm, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.lblConfirmSummary, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.chkConfirmAction, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add Me.lblConfirmError, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.fraResult, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.lblResultSummary, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.btnBack, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.btnCancel, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.btnOK, ANCHOR_RIGHT Or ANCHOR_BOTTOM
End Sub

Private Sub cmbSourceWarehouse_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError Me.lblSourceWarehouseError
    SelectSourceWarehouseInList Trim$(CStr(Me.cmbSourceWarehouse.Value))
    SelectDefaultTargetWarehouse
    SuggestArchiveDestination False
End Sub

Private Sub cmbTargetWarehouse_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError Me.lblTargetWarehouseError
End Sub

Private Sub txtArchiveDestPath_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError Me.lblArchiveDestPathError
End Sub

Private Sub mBtnArchiveDestBrowse_Click()
    Dim candidate As String

    candidate = modDeploymentPaths.BrowseForFolderPath(Trim$(CStr(Me.txtArchiveDestPath.Value)), "Choose Retire / Archive Destination")
    If candidate = "" Then Exit Sub

    Me.txtArchiveDestPath.Value = candidate
    ClearInlineError Me.lblArchiveDestPathError
End Sub

Private Sub mTxtWarehouseRootPath_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError mLblWarehouseRootError
End Sub

Private Sub mTxtNasUser_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError mLblWarehouseRootError
End Sub

Private Sub mTxtNasPassword_Change()
    If mFormBusy Then Exit Sub
    ClearInlineError mLblWarehouseRootError
End Sub

Private Sub mBtnNasConnect_Click()
    Dim report As String

    ClearInlineError mLblWarehouseRootError
    If ConnectSelectedWarehouseRootCredentialsForm(report) Then
        ShowFormMessage report, COLOR_SUCCESS
        RefreshWarehouseListsFromSelectedRoot "Warehouse root connected and scanned."
    Else
        SetInlineError mLblWarehouseRootError, report
        ShowFormMessage report, COLOR_ERROR
    End If
End Sub

Private Sub mBtnWarehouseRootBrowse_Click()
    Dim candidate As String

    candidate = modDeploymentPaths.BrowseForFolderPath(ResolveWarehouseRootPathForBrowse(), "Choose Warehouse Hub or Runtime Folder")
    If candidate = "" Then Exit Sub

    mFormBusy = True
    mTxtWarehouseRootPath.Value = candidate
    mFormBusy = False

    RefreshWarehouseListsFromSelectedRoot "Warehouse root selected and scanned."
End Sub

Private Sub mBtnWarehouseRootRefresh_Click()
    Dim report As String

    If ShouldConnectSelectedWarehouseRootForm() Then
        If Not ConnectSelectedWarehouseRootCredentialsForm(report) Then
            SetInlineError mLblWarehouseRootError, report
            ShowFormMessage report, COLOR_ERROR
            Exit Sub
        End If
    End If
    RefreshWarehouseListsFromSelectedRoot "Warehouse root scanned."
End Sub

Private Sub mLstSourceWarehouses_Click()
    If mFormBusy Then Exit Sub
    If mLstSourceWarehouses Is Nothing Then Exit Sub
    If mLstSourceWarehouses.ListIndex < 0 Then Exit Sub

    mFormBusy = True
    Me.cmbSourceWarehouse.Value = CStr(mLstSourceWarehouses.List(mLstSourceWarehouses.ListIndex, 0))
    mFormBusy = False
    ClearInlineError Me.lblSourceWarehouseError
    SelectDefaultTargetWarehouse
    SuggestArchiveDestination False
End Sub

Private Sub optArchiveOnly_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub optArchiveMigrate_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub optArchiveRetire_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub optArchiveRetireDelete_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub mOptArchiveOnly_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub mOptArchiveMigrate_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub mOptArchiveRetire_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub mOptArchiveRetireDelete_Click()
    If mFormBusy Then Exit Sub
    UpdateModeUi
End Sub

Private Sub chkConfirmAction_Click()
    If mCurrentPanel <> PANEL_CONFIRM Then Exit Sub
    ClearInlineError Me.lblConfirmError
    UpdateConfirmOkState
End Sub

Private Sub btnBack_Click()
    If mCurrentPanel = PANEL_CONFIRM Then
        ShowSelectionPanel
        ShowFormMessage "Adjust the selection, then click OK to continue.", COLOR_INFO
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Select Case mCurrentPanel
        Case PANEL_SELECTION
            HandleSelectionOk
        Case PANEL_CONFIRM
            HandleConfirmOk
        Case PANEL_RESULT
            Unload Me
    End Select
End Sub

Private Sub HandleSelectionOk()
    Dim sourceWarehouseId As String
    Dim targetWarehouseId As String
    Dim operationMode As Long
    Dim adminUser As String
    Dim archiveDestPath As String
    Dim publishTombstone As Boolean
    Dim warehouseRootPath As String

    ClearAllInlineErrors
    If Not BuildSpecFromSelection(sourceWarehouseId, targetWarehouseId, operationMode, adminUser, archiveDestPath, publishTombstone) Then
        ShowFormMessage "Fix the highlighted fields and try again.", COLOR_ERROR
        Exit Sub
    End If
    warehouseRootPath = ResolveSelectedWarehouseRootForm()
    If warehouseRootPath = "" Or Not FolderExistsForm(warehouseRootPath) Then
        SetInlineError mLblWarehouseRootError, "Choose the NAS hub folder or local warehouse runtime folder."
        ShowFormMessage "Warehouse root is required before continuing.", COLOR_ERROR
        Exit Sub
    End If

    If Not RequireReAuthForm("ADMIN_MAINT") Then
        SetInlineError Me.lblReAuthError, "Re-authentication required to continue"
        ShowFormMessage "Re-authentication required to continue.", COLOR_ERROR
        Exit Sub
    End If

    mPendingSourceWarehouseId = sourceWarehouseId
    mPendingTargetWarehouseId = targetWarehouseId
    mPendingOperationMode = operationMode
    mPendingAdminUser = adminUser
    mPendingArchiveDestPath = archiveDestPath
    mPendingPublishTombstone = publishTombstone
    mPendingWarehouseRootPath = warehouseRootPath
    mReAuthPassed = True
    ShowConfirmPanel
End Sub

Private Function RequireReAuthForm(ByVal requiredRole As String) As Boolean
    Dim gate As frmReAuthGate

    On Error GoTo ReAuthFail
    Set gate = New frmReAuthGate
    gate.InitializeGate ResolveRequiredRoleForm(requiredRole), ResolveCurrentUserForm()
    gate.Show vbModal
    RequireReAuthForm = gate.Authenticated
    On Error Resume Next
    Unload gate
    On Error GoTo 0
    Exit Function

ReAuthFail:
    RequireReAuthForm = False
    On Error Resume Next
    If Not gate Is Nothing Then Unload gate
    On Error GoTo 0
End Function

Private Function ResolveRequiredRoleForm(ByVal requiredRole As String) As String
    ResolveRequiredRoleForm = Trim$(requiredRole)
    If ResolveRequiredRoleForm = "" Then ResolveRequiredRoleForm = "ADMIN_MAINT"
End Function

Private Sub HandleConfirmOk()
    Dim summaryText As String
    Dim failureText As String
    Dim priorRootOverride As String

    ClearInlineError Me.lblConfirmError
    If Not mReAuthPassed Then
        SetInlineError Me.lblConfirmError, "Re-authentication required to continue"
        Exit Sub
    End If
    If Not CBool(Me.chkConfirmAction.Value) Then
        SetInlineError Me.lblConfirmError, "You must confirm this action before continuing."
        Exit Sub
    End If

    If Not modAdminConsole.ValidateRetireMigrateSpecAdmin(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingAdminUser, True, mPendingArchiveDestPath, mPendingPublishTombstone, failureText) Then
        SetInlineError Me.lblConfirmError, failureText
        Exit Sub
    End If

    priorRootOverride = modRuntimeWorkbooks.GetCoreDataRootOverride()
    If Trim$(mPendingWarehouseRootPath) <> "" Then modRuntimeWorkbooks.SetCoreDataRootOverride mPendingWarehouseRootPath
    On Error GoTo ConfirmFail

    If Not modAdminConsole.WriteArchivePackageAdmin(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingAdminUser, True, mPendingArchiveDestPath, mPendingPublishTombstone) Then
        ShowResultPanel False, "WriteArchivePackage failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
        GoTo CleanConfirm
    End If
    summaryText = "WriteArchivePackage: " & modWarehouseRetire.GetLastWarehouseRetireReport()

    If mPendingOperationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE Then
        If Not modAdminConsole.MigrateInventoryToTargetAdmin(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingAdminUser, True, mPendingArchiveDestPath, mPendingPublishTombstone) Then
            ShowResultPanel False, "MigrateInventoryToTarget failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
            GoTo CleanConfirm
        End If
        summaryText = summaryText & vbCrLf & "MigrateInventoryToTarget: " & modWarehouseRetire.GetLastWarehouseRetireReport()
    End If

    If mPendingOperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE Or _
       mPendingOperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE Then
        If Not modAdminConsole.RetireSourceWarehouseAdmin(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingAdminUser, True, mPendingArchiveDestPath, mPendingPublishTombstone) Then
            ShowResultPanel False, "RetireSourceWarehouse failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
            GoTo CleanConfirm
        End If
        summaryText = summaryText & vbCrLf & "RetireSourceWarehouse: " & modWarehouseRetire.GetLastWarehouseRetireReport()
    End If

    If mPendingOperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE Then
        If Not modAdminConsole.DeleteLocalRuntimeAdmin(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingAdminUser, True, mPendingArchiveDestPath, mPendingPublishTombstone) Then
            ShowResultPanel False, "DeleteLocalRuntime failed: " & modWarehouseRetire.GetLastWarehouseRetireReport()
            GoTo CleanConfirm
        End If
        summaryText = summaryText & vbCrLf & "DeleteLocalRuntime: " & modWarehouseRetire.GetLastWarehouseRetireReport()
    End If

    ShowResultPanel True, summaryText
    GoTo CleanConfirm

ConfirmFail:
    ShowResultPanel False, "Retire / Migrate failed: " & Err.Description

CleanConfirm:
    RestoreRootOverrideForm priorRootOverride
End Sub

Private Function BuildSpecFromSelection(ByRef sourceWarehouseId As String, _
                                        ByRef targetWarehouseId As String, _
                                        ByRef operationMode As Long, _
                                        ByRef adminUser As String, _
                                        ByRef archiveDestPath As String, _
                                        ByRef publishTombstone As Boolean) As Boolean
    Dim isValid As Boolean

    sourceWarehouseId = ResolveSelectedSourceWarehouseForm()
    targetWarehouseId = Trim$(CStr(Me.cmbTargetWarehouse.Value))
    operationMode = ResolveSelectedMode()
    adminUser = ResolveCurrentUserForm()
    archiveDestPath = Trim$(CStr(Me.txtArchiveDestPath.Value))
    publishTombstone = CBool(Me.chkPublishTombstone.Value)

    If sourceWarehouseId = "" Then
        SetInlineError Me.lblSourceWarehouseError, "Source warehouse is required."
        isValid = False
    Else
        isValid = True
    End If

    If operationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE Then
        If targetWarehouseId = "" Then
            SetInlineError Me.lblTargetWarehouseError, "Target warehouse is required for migrate mode."
            isValid = False
        ElseIf StrComp(sourceWarehouseId, targetWarehouseId, vbTextCompare) = 0 Then
            SetInlineError Me.lblTargetWarehouseError, "Target warehouse must be different from the source."
            isValid = False
        End If
    End If

    If archiveDestPath = "" Then
        SetInlineError Me.lblArchiveDestPathError, "Archive destination path is required."
        isValid = False
    End If

    If isValid Then
        If Not ValidateSelectionSpecForm(sourceWarehouseId, targetWarehouseId, operationMode, adminUser, archiveDestPath, publishTombstone) Then
            isValid = False
        End If
    End If

    BuildSpecFromSelection = isValid
End Function

Private Function ValidateSelectionSpecForm(ByVal sourceWarehouseId As String, _
                                           ByVal targetWarehouseId As String, _
                                           ByVal operationMode As Long, _
                                           ByVal adminUser As String, _
                                           ByVal archiveDestPath As String, _
                                           ByVal publishTombstone As Boolean) As Boolean
    Dim report As String

    If modAdminConsole.ValidateRetireMigrateSpecAdmin(sourceWarehouseId, targetWarehouseId, operationMode, adminUser, True, archiveDestPath, publishTombstone, report) Then
        ValidateSelectionSpecForm = True
        Exit Function
    End If

    If InStr(1, report, "SourceWarehouseId", vbTextCompare) > 0 Then
        SetInlineError Me.lblSourceWarehouseError, report
    ElseIf InStr(1, report, "TargetWarehouseId", vbTextCompare) > 0 Then
        SetInlineError Me.lblTargetWarehouseError, report
    ElseIf InStr(1, report, "ArchiveDestPath", vbTextCompare) > 0 Then
        SetInlineError Me.lblArchiveDestPathError, report
    Else
        ShowFormMessage report, COLOR_ERROR
    End If
End Function

Private Sub PopulateWarehouseDropdowns()
    Dim warehouseIds As Collection
    Dim item As Variant
    Dim rowIndex As Long

    Set mWarehousePathById = CreateObject("Scripting.Dictionary")
    mWarehousePathById.CompareMode = vbTextCompare
    Set warehouseIds = DiscoverWarehouseIdsForm()
    Me.cmbSourceWarehouse.Clear
    Me.cmbTargetWarehouse.Clear
    If Not mLstSourceWarehouses Is Nothing Then mLstSourceWarehouses.Clear

    For Each item In warehouseIds
        Me.cmbSourceWarehouse.AddItem CStr(item)
        Me.cmbTargetWarehouse.AddItem CStr(item)
        If Not mLstSourceWarehouses Is Nothing Then
            mLstSourceWarehouses.AddItem CStr(item)
            rowIndex = mLstSourceWarehouses.ListCount - 1
            mLstSourceWarehouses.List(rowIndex, 1) = WarehousePathForIdForm(CStr(item))
        End If
    Next item
End Sub

Private Sub ApplyWarehouseRootDefault()
    Dim rootPath As String

    If mTxtWarehouseRootPath Is Nothing Then Exit Sub
    rootPath = ResolveWarehouseScanRootForm()
    If rootPath = "" Then rootPath = modDeploymentPaths.DefaultRuntimeHubRootPath(False)
    mTxtWarehouseRootPath.Value = rootPath
End Sub

Private Sub RefreshWarehouseListsFromSelectedRoot(ByVal successMessage As String)
    Dim countFound As Long
    Dim rootPath As String

    ClearInlineError mLblWarehouseRootError
    rootPath = ResolveSelectedWarehouseRootForm()
    If IsUncPathForm(rootPath) And Not FolderExistsForm(rootPath) Then
        SetInlineError mLblWarehouseRootError, "Warehouse root is not reachable. Enter NAS credentials and click Connect."
        ShowFormMessage "Warehouse root is not reachable. Enter NAS credentials and click Connect.", COLOR_ERROR
        Exit Sub
    End If

    mFormBusy = True
    PopulateWarehouseDropdowns
    ApplyDefaultSelections
    countFound = Me.cmbSourceWarehouse.ListCount
    mFormBusy = False

    If countFound = 0 Then
        SetInlineError mLblWarehouseRootError, "No warehouse config files found under this root."
        ShowFormMessage "Choose the NAS hub folder or a specific warehouse runtime folder, then scan again.", COLOR_WARNING
    Else
        ShowFormMessage successMessage & " Found " & CStr(countFound) & " warehouse(s).", COLOR_SUCCESS
    End If
End Sub

Private Function ResolveWarehouseRootPathForBrowse() As String
    If Not mTxtWarehouseRootPath Is Nothing Then
        ResolveWarehouseRootPathForBrowse = NormalizePathForm(CStr(mTxtWarehouseRootPath.Value))
    End If
    If ResolveWarehouseRootPathForBrowse = "" Then ResolveWarehouseRootPathForBrowse = ResolveWarehouseScanRootForm()
    If ResolveWarehouseRootPathForBrowse = "" Then ResolveWarehouseRootPathForBrowse = modDeploymentPaths.DefaultRuntimeHubRootPath(False)
End Function

Private Function ShouldConnectSelectedWarehouseRootForm() As Boolean
    If Not IsUncPathForm(ResolveSelectedWarehouseRootForm()) Then Exit Function
    If mTxtNasUser Is Nothing Or mTxtNasPassword Is Nothing Then Exit Function
    ShouldConnectSelectedWarehouseRootForm = (Trim$(CStr(mTxtNasUser.Value)) <> "" Or CStr(mTxtNasPassword.Value) <> "")
End Function

Private Function ConnectSelectedWarehouseRootCredentialsForm(ByRef report As String) As Boolean
    Dim rootPath As String
    Dim shareRoot As String
    Dim userName As String
    Dim passwordText As String
    Dim resultCode As Long
    Dim resource As NETRESOURCE

    rootPath = ResolveSelectedWarehouseRootForm()
    If rootPath = "" Then
        report = "Warehouse root is required before connecting."
        Exit Function
    End If
    If Not IsUncPathForm(rootPath) Then
        report = "Warehouse root is local; NAS credentials are not needed."
        ConnectSelectedWarehouseRootCredentialsForm = True
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

    resource.dwType = RESOURCETYPE_DISK
    resource.lpRemoteName = shareRoot
    resultCode = WNetAddConnection2(resource, passwordText, userName, CONNECT_TEMPORARY)

    If resultCode = NO_ERROR_WIN32 Then
        mTxtNasPassword.Value = ""
        If FolderExistsForm(rootPath) Then
            report = "Connected to NAS root."
            ConnectSelectedWarehouseRootCredentialsForm = True
        Else
            report = "Connected to NAS share, but the Warehouse root folder was not found."
        End If
        Exit Function
    End If

    If resultCode = ERROR_SESSION_CREDENTIAL_CONFLICT And FolderExistsForm(rootPath) Then
        mTxtNasPassword.Value = ""
        report = "Using existing Windows NAS connection."
        ConnectSelectedWarehouseRootCredentialsForm = True
        Exit Function
    End If

    report = WNetConnectionErrorTextForm(resultCode)
End Function

Private Function DiscoverWarehouseIdsForm() As Collection
    Dim results As Collection
    Dim seen As Object
    Dim rootPath As String
    Dim scanRoots As Collection
    Dim candidate As Variant

    Set results = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    Set scanRoots = New Collection

    rootPath = ResolveSelectedWarehouseRootForm()
    If rootPath = "" Then rootPath = ResolveWarehouseScanRootForm()
    AddScanRootForm scanRoots, rootPath

    For Each candidate In scanRoots
        AddWarehousesFromRootForm results, seen, CStr(candidate)
    Next candidate

    Set DiscoverWarehouseIdsForm = results
End Function

Private Function ResolveSelectedWarehouseRootForm() As String
    If mTxtWarehouseRootPath Is Nothing Then Exit Function
    ResolveSelectedWarehouseRootForm = NormalizePathForm(CStr(mTxtWarehouseRootPath.Value))
End Function

Private Function ResolveWarehouseScanRootForm() As String
    Dim rootPath As String
    Dim parentPath As String

    rootPath = Trim$(modRuntimeWorkbooks.GetCoreDataRootOverride())
    If rootPath = "" Then rootPath = Trim$(modRuntimeWorkbooks.ResolveCoreDataRoot("", ""))
    rootPath = NormalizePathForm(rootPath)
    If rootPath = "" Then
        ResolveWarehouseScanRootForm = modDeploymentPaths.DefaultRuntimeHubRootPath(False)
        Exit Function
    End If

    If LooksLikeWarehouseRuntimeRootForm(rootPath) Then
        parentPath = GetParentFolderForm(rootPath)
        If parentPath <> "" Then
            ResolveWarehouseScanRootForm = parentPath
            Exit Function
        End If
    End If

    ResolveWarehouseScanRootForm = rootPath
End Function

Private Sub AddScanRootForm(ByVal scanRoots As Collection, ByVal rootPath As String)
    Dim item As Variant

    rootPath = NormalizePathForm(rootPath)
    If rootPath = "" Then Exit Sub

    For Each item In scanRoots
        If StrComp(CStr(item), rootPath, vbTextCompare) = 0 Then Exit Sub
    Next item

    scanRoots.Add rootPath
End Sub

Private Sub AddWarehousesFromRootForm(ByVal results As Collection, ByVal seen As Object, ByVal rootPath As String)
    Dim fso As Object
    Dim rootFolder As Object
    Dim subFolder As Object

    rootPath = NormalizePathForm(rootPath)
    If rootPath = "" Then Exit Sub
    If Not FolderExistsForm(rootPath) Then Exit Sub

    AddWarehouseIdsFromFolderForm results, seen, rootPath

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    Set rootFolder = fso.GetFolder(rootPath)
    On Error GoTo 0
    If rootFolder Is Nothing Then Exit Sub

    For Each subFolder In rootFolder.SubFolders
        AddWarehouseIdsFromFolderForm results, seen, CStr(subFolder.Path)
    Next subFolder
End Sub

Private Sub AddWarehouseIdsFromFolderForm(ByVal results As Collection, ByVal seen As Object, ByVal folderPath As String)
    Const CONFIG_SUFFIX As String = ".invSys.Config.xlsb"
    Dim candidate As String
    Dim warehouseId As String

    folderPath = NormalizePathForm(folderPath)
    If folderPath = "" Then Exit Sub

    On Error GoTo CleanExit
    candidate = Dir$(folderPath & "\*" & CONFIG_SUFFIX, vbNormal)
    Do While candidate <> ""
        If Len(candidate) > Len(CONFIG_SUFFIX) Then
            warehouseId = Left$(candidate, Len(candidate) - Len(CONFIG_SUFFIX))
            If warehouseId <> "" And Not seen.Exists(warehouseId) Then
                seen.Add warehouseId, True
                If Not mWarehousePathById Is Nothing Then mWarehousePathById(warehouseId) = folderPath
                results.Add warehouseId
            End If
        End If
        candidate = Dir$
    Loop

CleanExit:
End Sub

Private Function WarehousePathForIdForm(ByVal warehouseId As String) As String
    If mWarehousePathById Is Nothing Then Exit Function
    If mWarehousePathById.Exists(warehouseId) Then WarehousePathForIdForm = CStr(mWarehousePathById(warehouseId))
End Function

Private Sub ApplyDefaultSelections()
    If Me.cmbSourceWarehouse.ListCount > 0 Then
        Me.cmbSourceWarehouse.ListIndex = 0
        SelectSourceWarehouseInList Trim$(CStr(Me.cmbSourceWarehouse.Value))
    ElseIf Not mLstSourceWarehouses Is Nothing Then
        mLstSourceWarehouses.ListIndex = -1
    End If
    If Me.cmbTargetWarehouse.ListCount > 0 Then
        SelectDefaultTargetWarehouse
    End If
    SuggestArchiveDestination True
    UpdateModeUi
End Sub

Private Sub SelectDefaultTargetWarehouse()
    Dim sourceWarehouseId As String
    Dim index As Long

    sourceWarehouseId = ResolveSelectedSourceWarehouseForm()
    If sourceWarehouseId = "" Then Exit Sub
    If Me.cmbTargetWarehouse.ListCount = 0 Then Exit Sub
    If Trim$(CStr(Me.cmbTargetWarehouse.Value)) <> "" Then
        If StrComp(Trim$(CStr(Me.cmbTargetWarehouse.Value)), sourceWarehouseId, vbTextCompare) <> 0 Then Exit Sub
    End If

    For index = 0 To Me.cmbTargetWarehouse.ListCount - 1
        If StrComp(CStr(Me.cmbTargetWarehouse.List(index, 0)), sourceWarehouseId, vbTextCompare) <> 0 Then
            Me.cmbTargetWarehouse.ListIndex = index
            Exit Sub
        End If
    Next index
End Sub

Private Function ResolveSelectedSourceWarehouseForm() As String
    If Not mLstSourceWarehouses Is Nothing Then
        If mLstSourceWarehouses.ListIndex >= 0 Then
            ResolveSelectedSourceWarehouseForm = Trim$(CStr(mLstSourceWarehouses.List(mLstSourceWarehouses.ListIndex, 0)))
        End If
    End If
    If ResolveSelectedSourceWarehouseForm = "" Then ResolveSelectedSourceWarehouseForm = Trim$(CStr(Me.cmbSourceWarehouse.Value))
End Function

Private Sub SelectSourceWarehouseInList(ByVal warehouseId As String)
    Dim index As Long
    Dim wasBusy As Boolean

    If mLstSourceWarehouses Is Nothing Then Exit Sub
    wasBusy = mFormBusy
    mFormBusy = True
    If warehouseId = "" Then
        mLstSourceWarehouses.ListIndex = -1
        mFormBusy = wasBusy
        Exit Sub
    End If

    For index = 0 To mLstSourceWarehouses.ListCount - 1
        If StrComp(CStr(mLstSourceWarehouses.List(index, 0)), warehouseId, vbTextCompare) = 0 Then
            mLstSourceWarehouses.ListIndex = index
            mFormBusy = wasBusy
            Exit Sub
        End If
    Next index
    mFormBusy = wasBusy
End Sub

Private Sub SuggestArchiveDestination(ByVal forceApply As Boolean)
    Dim suggestedPath As String

    suggestedPath = ResolveArchiveDefaultForm(ResolveSelectedSourceWarehouseForm())
    If forceApply Or Trim$(CStr(Me.txtArchiveDestPath.Value)) = "" Then
        Me.txtArchiveDestPath.Value = suggestedPath
    End If
End Sub

Private Function ResolveArchiveDefaultForm(ByVal warehouseId As String) As String
    Dim priorRoot As String
    Dim pathValue As String
    Dim selectedRoot As String

    If warehouseId = "" Then
        ResolveArchiveDefaultForm = modDeploymentPaths.DefaultArchiveRootPath(False)
        Exit Function
    End If

    priorRoot = modRuntimeWorkbooks.GetCoreDataRootOverride()
    On Error Resume Next
    selectedRoot = ResolveSelectedWarehouseRootForm()
    If selectedRoot <> "" Then
        modRuntimeWorkbooks.SetCoreDataRootOverride selectedRoot
    Else
        modRuntimeWorkbooks.ClearCoreDataRootOverride
    End If
    If modConfig.LoadConfig(warehouseId, "") Then
        pathValue = Trim$(modConfig.GetString("PathBackupRoot", ""))
    End If
    On Error GoTo 0
    RestoreRootOverrideForm priorRoot

    pathValue = NormalizePathForm(pathValue)
    If pathValue = "" Then
        ResolveArchiveDefaultForm = modDeploymentPaths.DefaultArchiveRootPath(False)
    Else
        ResolveArchiveDefaultForm = pathValue
    End If
End Function

Private Sub UpdateModeUi()
    Dim migrateMode As Boolean
    Dim retireMode As Boolean
    Dim deleteMode As Boolean

    migrateMode = (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_MIGRATE)
    retireMode = (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE Or ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
    deleteMode = (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)

    Me.cmbTargetWarehouse.Enabled = migrateMode
    Me.lblTargetWarehouse.Enabled = migrateMode
    Me.chkPublishTombstone.Visible = retireMode
    Me.chkPublishTombstone.Enabled = retireMode
    If Not retireMode Then Me.chkPublishTombstone.Value = False

    Me.lblDeleteWarning.Visible = deleteMode
    ClearInlineError Me.lblTargetWarehouseError
    ClearInlineError Me.lblReAuthError
End Sub

Private Function ResolveSelectedMode() As Long
    If Not mOptArchiveMigrate Is Nothing Then
        If CBool(mOptArchiveMigrate.Value) Then
            ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
        ElseIf CBool(mOptArchiveRetire.Value) Then
            ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
        ElseIf CBool(mOptArchiveRetireDelete.Value) Then
            ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
        Else
            ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
        End If
        Exit Function
    End If

    If CBool(Me.optArchiveMigrate.Value) Then
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE
    ElseIf CBool(Me.optArchiveRetire.Value) Then
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE
    ElseIf CBool(Me.optArchiveRetireDelete.Value) Then
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
    Else
        ResolveSelectedMode = modWarehouseRetire.MODE_ARCHIVE_ONLY
    End If
End Function

Private Sub SetOperationModeValue(ByVal operationMode As Long)
    If Not mOptArchiveOnly Is Nothing Then
        mOptArchiveOnly.Value = (operationMode = modWarehouseRetire.MODE_ARCHIVE_ONLY)
        mOptArchiveMigrate.Value = (operationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE)
        mOptArchiveRetire.Value = (operationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE)
        mOptArchiveRetireDelete.Value = (operationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
        Exit Sub
    End If

    Me.optArchiveOnly.Value = (operationMode = modWarehouseRetire.MODE_ARCHIVE_ONLY)
    Me.optArchiveMigrate.Value = (operationMode = modWarehouseRetire.MODE_ARCHIVE_MIGRATE)
    Me.optArchiveRetire.Value = (operationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE)
    Me.optArchiveRetireDelete.Value = (operationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
End Sub

Private Sub ShowSelectionPanel()
    mCurrentPanel = PANEL_SELECTION
    SetSelectionControlsVisible True
    Me.fraConfirm.Visible = False
    Me.fraResult.Visible = False
    Me.btnBack.Visible = False
    Me.btnCancel.Caption = "Cancel"
    Me.btnOK.Caption = "OK"
    BringSelectionControlsToFront
    UpdateModeUi
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Sub ShowConfirmPanel()
    mCurrentPanel = PANEL_CONFIRM
    SetSelectionControlsVisible False
    Me.fraConfirm.Visible = True
    Me.fraResult.Visible = False
    Me.btnBack.Visible = True
    Me.btnCancel.Caption = "Cancel"
    Me.btnOK.Caption = "Run"
    Me.chkConfirmAction.Value = False
    Me.lblConfirmSummary.Caption = BuildConfirmationSummaryForm(mPendingSourceWarehouseId, mPendingTargetWarehouseId, mPendingOperationMode, mPendingArchiveDestPath, mPendingPublishTombstone)
    Me.lblDeleteWarning.Visible = (mPendingOperationMode = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
    ClearInlineError Me.lblConfirmError
    BringConfirmControlsToFront
    UpdateConfirmOkState
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Sub ShowResultPanel(ByVal wasSuccessful As Boolean, ByVal detailText As String)
    mCurrentPanel = PANEL_RESULT
    SetSelectionControlsVisible False
    Me.fraConfirm.Visible = False
    Me.fraResult.Visible = True
    Me.btnBack.Visible = False
    Me.btnCancel.Caption = "Close"
    Me.btnOK.Caption = "Close"
    Me.lblResultSummary.Caption = Trim$(detailText)
    Me.lblResultSummary.ForeColor = IIf(wasSuccessful, COLOR_SUCCESS, COLOR_ERROR)
    BringResultControlsToFront
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
End Sub

Private Function BuildConfirmationSummaryForm(ByVal sourceWarehouseId As String, _
                                              ByVal targetWarehouseId As String, _
                                              ByVal operationMode As Long, _
                                              ByVal archiveDestPath As String, _
                                              ByVal publishTombstone As Boolean) As String
    Select Case operationMode
        Case modWarehouseRetire.MODE_ARCHIVE_ONLY
            BuildConfirmationSummaryForm = _
                "Archive package only will create an archive package for " & sourceWarehouseId & "." & vbCrLf & _
                "No migration, retirement, or deletion will occur." & vbCrLf & _
                "Archive location: " & archiveDestPath
        Case modWarehouseRetire.MODE_ARCHIVE_MIGRATE
            BuildConfirmationSummaryForm = _
                "Migrate will archive " & sourceWarehouseId & " and seed current inventory into " & targetWarehouseId & "." & vbCrLf & _
                "The target remains locally authoritative. No auth, config identity, or inbox files are copied." & vbCrLf & _
                "Archive location: " & archiveDestPath
        Case modWarehouseRetire.MODE_ARCHIVE_RETIRE
            BuildConfirmationSummaryForm = _
                "Retire / archive will archive " & sourceWarehouseId & ", mark it RETIRED locally, and write a tombstone." & vbCrLf & _
                IIf(publishTombstone, "A best-effort SharePoint tombstone publish will also be attempted.", "SharePoint tombstone publish is disabled.") & vbCrLf & _
                "Archive location: " & archiveDestPath
        Case modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE
            BuildConfirmationSummaryForm = _
                "Retire / archive and delete will archive " & sourceWarehouseId & ", mark it RETIRED, write a tombstone, then delete the local runtime folder." & vbCrLf & _
                IIf(publishTombstone, "A best-effort SharePoint tombstone publish will also be attempted before deletion.", "SharePoint tombstone publish is disabled.") & vbCrLf & _
                "Archive location: " & archiveDestPath
    End Select
End Function

Private Sub UpdateConfirmOkState()
    Me.btnOK.Enabled = CBool(Me.chkConfirmAction.Value)
End Sub

Private Sub SetSelectionControlsVisible(ByVal isVisible As Boolean)
    Me.lblTitle.Visible = isVisible
    Me.lblSelectionIntro.Visible = isVisible
    If Not mLblWarehouseRoot Is Nothing Then mLblWarehouseRoot.Visible = isVisible
    If Not mTxtWarehouseRootPath Is Nothing Then mTxtWarehouseRootPath.Visible = isVisible
    If Not mBtnWarehouseRootBrowse Is Nothing Then mBtnWarehouseRootBrowse.Visible = isVisible
    If Not mBtnWarehouseRootRefresh Is Nothing Then mBtnWarehouseRootRefresh.Visible = isVisible
    If Not mLblWarehouseRootError Is Nothing Then mLblWarehouseRootError.Visible = isVisible
    If Not mLblNasUser Is Nothing Then mLblNasUser.Visible = isVisible
    If Not mTxtNasUser Is Nothing Then mTxtNasUser.Visible = isVisible
    If Not mLblNasPassword Is Nothing Then mLblNasPassword.Visible = isVisible
    If Not mTxtNasPassword Is Nothing Then mTxtNasPassword.Visible = isVisible
    If Not mBtnNasConnect Is Nothing Then mBtnNasConnect.Visible = isVisible
    Me.lblSourceWarehouse.Visible = isVisible
    Me.cmbSourceWarehouse.Visible = isVisible
    Me.lblSourceWarehouseError.Visible = isVisible
    Me.lblTargetWarehouse.Visible = isVisible
    Me.cmbTargetWarehouse.Visible = isVisible
    Me.lblTargetWarehouseError.Visible = isVisible
    If Not mLblFoundWarehouses Is Nothing Then mLblFoundWarehouses.Visible = isVisible
    If Not mLstSourceWarehouses Is Nothing Then mLstSourceWarehouses.Visible = isVisible
    Me.fraMode.Visible = isVisible
    HideDesignerOperationModeControls
    If Not mOptArchiveOnly Is Nothing Then mOptArchiveOnly.Visible = False
    If Not mOptArchiveMigrate Is Nothing Then mOptArchiveMigrate.Visible = isVisible
    If Not mOptArchiveRetire Is Nothing Then mOptArchiveRetire.Visible = isVisible
    If Not mOptArchiveRetireDelete Is Nothing Then mOptArchiveRetireDelete.Visible = isVisible
    Me.lblArchiveDestPath.Visible = isVisible
    Me.txtArchiveDestPath.Visible = isVisible
    If Not mBtnArchiveDestBrowse Is Nothing Then mBtnArchiveDestBrowse.Visible = isVisible
    Me.lblArchiveDestPathError.Visible = isVisible
    Me.chkPublishTombstone.Visible = isVisible And (ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE Or ResolveSelectedMode() = modWarehouseRetire.MODE_ARCHIVE_RETIRE_DELETE)
    Me.lblReAuthError.Visible = isVisible
End Sub

Private Sub BringSelectionControlsToFront()
    On Error Resume Next
    Me.lblTitle.ZOrder 0
    Me.lblSelectionIntro.ZOrder 0
    If Not mLblWarehouseRoot Is Nothing Then mLblWarehouseRoot.ZOrder 0
    If Not mTxtWarehouseRootPath Is Nothing Then mTxtWarehouseRootPath.ZOrder 0
    If Not mBtnWarehouseRootBrowse Is Nothing Then mBtnWarehouseRootBrowse.ZOrder 0
    If Not mBtnWarehouseRootRefresh Is Nothing Then mBtnWarehouseRootRefresh.ZOrder 0
    If Not mLblWarehouseRootError Is Nothing Then mLblWarehouseRootError.ZOrder 0
    If Not mLblNasUser Is Nothing Then mLblNasUser.ZOrder 0
    If Not mTxtNasUser Is Nothing Then mTxtNasUser.ZOrder 0
    If Not mLblNasPassword Is Nothing Then mLblNasPassword.ZOrder 0
    If Not mTxtNasPassword Is Nothing Then mTxtNasPassword.ZOrder 0
    If Not mBtnNasConnect Is Nothing Then mBtnNasConnect.ZOrder 0
    Me.lblSourceWarehouse.ZOrder 0
    Me.cmbSourceWarehouse.ZOrder 0
    Me.lblSourceWarehouseError.ZOrder 0
    Me.lblTargetWarehouse.ZOrder 0
    Me.cmbTargetWarehouse.ZOrder 0
    Me.lblTargetWarehouseError.ZOrder 0
    If Not mLblFoundWarehouses Is Nothing Then mLblFoundWarehouses.ZOrder 0
    If Not mLstSourceWarehouses Is Nothing Then mLstSourceWarehouses.ZOrder 0
    Me.fraMode.ZOrder 0
    If Not mOptArchiveOnly Is Nothing Then mOptArchiveOnly.ZOrder 0
    If Not mOptArchiveMigrate Is Nothing Then mOptArchiveMigrate.ZOrder 0
    If Not mOptArchiveRetire Is Nothing Then mOptArchiveRetire.ZOrder 0
    If Not mOptArchiveRetireDelete Is Nothing Then mOptArchiveRetireDelete.ZOrder 0
    Me.lblArchiveDestPath.ZOrder 0
    Me.txtArchiveDestPath.ZOrder 0
    If Not mBtnArchiveDestBrowse Is Nothing Then mBtnArchiveDestBrowse.ZOrder 0
    Me.lblArchiveDestPathError.ZOrder 0
    Me.chkPublishTombstone.ZOrder 0
    Me.lblReAuthError.ZOrder 0
    Me.lblDeleteWarning.ZOrder 0
    Me.btnBack.ZOrder 0
    Me.btnCancel.ZOrder 0
    Me.btnOK.ZOrder 0
    On Error GoTo 0
End Sub

Private Sub BringConfirmControlsToFront()
    On Error Resume Next
    Me.fraConfirm.ZOrder 0
    Me.lblConfirmSummary.ZOrder 0
    Me.chkConfirmAction.ZOrder 0
    Me.lblConfirmError.ZOrder 0
    Me.btnBack.ZOrder 0
    Me.btnCancel.ZOrder 0
    Me.btnOK.ZOrder 0
    On Error GoTo 0
End Sub

Private Sub BringResultControlsToFront()
    On Error Resume Next
    Me.fraResult.ZOrder 0
    Me.lblResultSummary.ZOrder 0
    Me.btnCancel.ZOrder 0
    Me.btnOK.ZOrder 0
    On Error GoTo 0
End Sub

Private Sub ClearAllInlineErrors()
    ClearInlineError mLblWarehouseRootError
    ClearInlineError Me.lblSourceWarehouseError
    ClearInlineError Me.lblTargetWarehouseError
    ClearInlineError Me.lblArchiveDestPathError
    ClearInlineError Me.lblReAuthError
    ClearInlineError Me.lblConfirmError
End Sub

Private Sub ClearInlineError(ByVal lbl As MSForms.Label)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = ""
    lbl.ForeColor = COLOR_ERROR
End Sub

Private Sub SetInlineError(ByVal lbl As MSForms.Label, ByVal messageText As String)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = Trim$(messageText)
    lbl.ForeColor = COLOR_ERROR
End Sub

Private Sub ShowFormMessage(ByVal messageText As String, ByVal foreColor As Long)
    Me.lblSelectionIntro.Caption = Trim$(messageText)
    Me.lblSelectionIntro.ForeColor = foreColor
End Sub

Private Function ResolveCurrentUserForm() As String
    ResolveCurrentUserForm = Trim$(Environ$("USERNAME"))
    If ResolveCurrentUserForm = "" Then ResolveCurrentUserForm = Trim$(Application.UserName)
End Function

Private Function NormalizePathForm(ByVal pathText As String) As String
    pathText = Trim$(Replace$(pathText, "/", "\"))
    Do While Len(pathText) > 3 And Right$(pathText, 1) = "\"
        pathText = Left$(pathText, Len(pathText) - 1)
    Loop
    NormalizePathForm = pathText
End Function

Private Function IsUncPathForm(ByVal pathText As String) As Boolean
    pathText = NormalizePathForm(pathText)
    IsUncPathForm = (Left$(pathText, 2) = "\\")
End Function

Private Function ResolveUncShareRootForm(ByVal pathText As String) As String
    Dim body As String
    Dim parts() As String

    pathText = NormalizePathForm(pathText)
    If Left$(pathText, 2) <> "\\" Then Exit Function

    body = Mid$(pathText, 3)
    If body = "" Then Exit Function
    parts = Split(body, "\")
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

Private Function GetParentFolderForm(ByVal pathText As String) As String
    GetParentFolderForm = modDeploymentPaths.GetParentFolderManaged(pathText)
End Function

Private Function GetLeafFolderNameForm(ByVal pathText As String) As String
    Dim sepPos As Long

    pathText = NormalizePathForm(pathText)
    If pathText = "" Then Exit Function
    sepPos = InStrRev(pathText, "\")
    If sepPos > 0 And sepPos < Len(pathText) Then
        GetLeafFolderNameForm = Mid$(pathText, sepPos + 1)
    Else
        GetLeafFolderNameForm = pathText
    End If
End Function

Private Function LooksLikeWarehouseRuntimeRootForm(ByVal rootPath As String) As Boolean
    Dim leafName As String

    rootPath = NormalizePathForm(rootPath)
    If rootPath = "" Then Exit Function
    leafName = Trim$(GetLeafFolderNameForm(rootPath))
    If leafName = "" Then Exit Function

    LooksLikeWarehouseRuntimeRootForm = FileExistsForm(rootPath & "\" & leafName & ".invSys.Config.xlsb")
End Function

Private Function FolderExistsForm(ByVal folderPath As String) As Boolean
    Dim fso As Object

    folderPath = NormalizePathForm(folderPath)
    If folderPath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FolderExistsForm = fso.FolderExists(folderPath)
    If Err.Number <> 0 Then
        Err.Clear
        FolderExistsForm = (Len(Dir$(folderPath, vbDirectory)) > 0)
    End If
    On Error GoTo 0
End Function

Private Function FileExistsForm(ByVal filePath As String) As Boolean
    Dim fso As Object

    filePath = NormalizePathForm(filePath)
    If filePath = "" Then Exit Function

    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso Is Nothing Then FileExistsForm = fso.FileExists(filePath)
    If Err.Number <> 0 Then
        Err.Clear
        FileExistsForm = (Len(Dir$(filePath, vbNormal)) > 0)
    End If
    On Error GoTo 0
End Function

Private Sub RestoreRootOverrideForm(ByVal priorRoot As String)
    If Trim$(priorRoot) = "" Then
        modRuntimeWorkbooks.ClearCoreDataRootOverride
    Else
        modRuntimeWorkbooks.SetCoreDataRootOverride priorRoot
    End If
End Sub
