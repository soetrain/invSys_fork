VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetupTesterStation 
   Caption         =   "Setup Tester Station"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmSetupTesterStation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSetupTesterStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const COLOR_ERROR As Long = 255
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_INFO As Long = 0
Private Const ANCHOR_LEFT As Long = 1
Private Const ANCHOR_TOP As Long = 2
Private Const ANCHOR_RIGHT As Long = 4
Private Const ANCHOR_BOTTOM As Long = 8

Private WithEvents mTxtConfirmPin As MSForms.TextBox
Attribute mTxtConfirmPin.VB_VarHelpID = -1
Private WithEvents mBtnOpen As MSForms.CommandButton
Attribute mBtnOpen.VB_VarHelpID = -1
Private WithEvents mBtnSharePointHelper As MSForms.CommandButton
Attribute mBtnSharePointHelper.VB_VarHelpID = -1

Private mLblConfirmPinError As MSForms.Label
Private mPathLocalTouched As Boolean
Private mLastSuggestedLocalPath As String
Private mSetupSucceeded As Boolean
Private mOperatorWorkbookPath As String
Private mFormBusy As Boolean
Private mAnchors As Object
Private mResizeInitialized As Boolean

Private Sub UserForm_Initialize()
    mFormBusy = True
    ConfigureShellLayout
    CreateDynamicControls
    ApplyDefaults
    ClearValidationErrors
    InitializeSetupTesterAnchors
    ShowSummary "Use the locally synced invSys root that contains Addins, Events, Snapshots, and TesterPackage, then click Setup.", COLOR_INFO
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

Private Sub txtAdminUser_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblAdminUserError
End Sub

    Private Sub txtWarehouseName_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblWarehouseNameError
End Sub

Private Sub txtWarehouseId_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblWarehouseIdError
    RefreshSuggestedLocalPath False
End Sub

Private Sub txtStationId_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblStationIdError
End Sub

Private Sub txtPathLocal_Change()
    Dim currentValue As String

    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblPathLocalError
    currentValue = Trim$(CStr(Me.txtPathLocal.Value))
    If mLastSuggestedLocalPath <> "" Then
        mPathLocalTouched = (StrComp(currentValue, mLastSuggestedLocalPath, vbTextCompare) <> 0)
    End If
End Sub

Private Sub txtPathSharePoint_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblPathSharePointError
End Sub

Private Sub mBtnSharePointHelper_Click()
    Dim candidate As String
    Dim warehouseId As String

    warehouseId = Trim$(CStr(Me.txtWarehouseId.Value))
    candidate = modTesterSetup.DetectSharePointRoot(warehouseId)
    If candidate = "" Then
        candidate = modTesterSetup.BrowseForSharePointRoot(Trim$(CStr(Me.txtPathSharePoint.Value)))
    End If

    If candidate = "" Then
        ShowSummary "SharePoint root was not detected. Pick the locally synced invSys root folder manually.", COLOR_INFO
        Exit Sub
    End If

    mFormBusy = True
    Me.txtPathSharePoint.Value = candidate
    mFormBusy = False
    ClearErrorLabel Me.lblPathSharePointError
    ShowSummary "SharePoint root detected.", COLOR_SUCCESS
End Sub

Private Sub mTxtConfirmPin_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel mLblConfirmPinError
End Sub

Private Sub btnOK_Click()
    Dim spec As modTesterSetup.TesterSetupSpec
    Dim rawPin As String
    Dim detailText As String
    Dim operatorName As String

    If mSetupSucceeded Then
        Unload Me
        Exit Sub
    End If

    ClearValidationErrors
    If Not BuildSpecFromForm(spec, rawPin) Then
        ShowSummary "Fix the highlighted fields and try again.", COLOR_ERROR
        Exit Sub
    End If

    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered(spec.PathSharePointRoot & "\Addins", detailText) Then
        ShowSummary "invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & detailText, COLOR_ERROR
        Exit Sub
    End If

    Me.btnOK.Enabled = False
    Me.btnCancel.Enabled = False
    If Not mBtnOpen Is Nothing Then mBtnOpen.Enabled = False

    modTesterSetup.SetTesterSetupProgressSink Me
    If modTesterSetup.SetupTesterStation(spec) Then
        mSetupSucceeded = True
        mOperatorWorkbookPath = modTesterSetup.GetLastTesterOperatorWorkbookPath()
        operatorName = Mid$(mOperatorWorkbookPath, InStrRev(mOperatorWorkbookPath, "\") + 1)
        If operatorName = "" Then operatorName = spec.WarehouseId & ".Receiving.Operator.xlsm"
        ShowSummary "Setup complete. Open " & operatorName & " to begin.", COLOR_SUCCESS
        Me.btnOK.Caption = "Close"
        Me.btnCancel.Caption = "Close"
        If Not mBtnOpen Is Nothing Then
            mBtnOpen.Visible = True
            mBtnOpen.Enabled = True
        End If
        If Not mAnchors Is Nothing Then mAnchors.ResizeControls
    Else
        detailText = modTesterSetup.GetLastTesterSetupReport()
        ShowSummary "Setup failed. Review the runtime path and try again." & vbCrLf & detailText, COLOR_ERROR
        Me.btnOK.Enabled = True
        Me.btnCancel.Enabled = True
    End If
    modTesterSetup.ClearTesterSetupProgressSink
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub mBtnOpen_Click()
    If mOperatorWorkbookPath = "" Then
        ShowSummary "Setup completed, but the operator workbook path was not captured.", COLOR_ERROR
        Exit Sub
    End If

    If modTesterSetup.OpenTesterReceivingWorkbook(mOperatorWorkbookPath) Then
        ShowSummary "Operator workbook opened. Use Refresh Inventory, then run your Confirm Writes test.", COLOR_SUCCESS
    Else
        ShowSummary "Setup completed, but the operator workbook could not be opened automatically.", COLOR_ERROR
    End If
End Sub

Public Sub UpdateSetupProgress(ByVal stepText As String)
    If Trim$(stepText) = "" Then Exit Sub
    ShowSummary Trim$(stepText), COLOR_INFO
    Me.Repaint
End Sub

Private Sub ConfigureShellLayout()
    Dim ctl As Control

    Me.Caption = "Setup Tester Station"
    Me.Width = 600
    Me.Height = 500
    Me.StartUpPosition = 1

    For Each ctl In Me.Controls
        ctl.Visible = False
    Next ctl

    Me.txtAdminUser.Visible = True
    Me.txtWarehouseName.Visible = True
    Me.txtWarehouseId.Visible = True
    Me.txtStationId.Visible = True
    Me.txtPathLocal.Visible = True
    Me.txtPathSharePoint.Visible = True
    Me.lblAdminUserError.Visible = True
    Me.lblWarehouseNameError.Visible = True
    Me.lblWarehouseIdError.Visible = True
    Me.lblStationIdError.Visible = True
    Me.lblPathLocalError.Visible = True
    Me.lblPathSharePointError.Visible = True
    Me.lblSummary.Visible = True
    Me.btnOK.Visible = True
    Me.btnCancel.Visible = True

    Me.txtAdminUser.Left = 170
    Me.txtAdminUser.Top = 24
    Me.txtAdminUser.Width = 240

    Me.txtWarehouseName.Left = 170
    Me.txtWarehouseName.Top = 74
    Me.txtWarehouseName.Width = 240
    Me.txtWarehouseName.PasswordChar = "*"

    Me.txtWarehouseId.Left = 170
    Me.txtWarehouseId.Top = 124
    Me.txtWarehouseId.Width = 240

    Me.txtStationId.Left = 170
    Me.txtStationId.Top = 174
    Me.txtStationId.Width = 240

    Me.txtPathLocal.Left = 170
    Me.txtPathLocal.Top = 224
    Me.txtPathLocal.Width = 360

    Me.txtPathSharePoint.Left = 170
    Me.txtPathSharePoint.Top = 274
    Me.txtPathSharePoint.Width = 258

    Me.lblAdminUserError.Left = 170
    Me.lblAdminUserError.Top = 44
    Me.lblAdminUserError.Width = 360

    Me.lblWarehouseNameError.Left = 170
    Me.lblWarehouseNameError.Top = 94
    Me.lblWarehouseNameError.Width = 360

    Me.lblWarehouseIdError.Left = 170
    Me.lblWarehouseIdError.Top = 144
    Me.lblWarehouseIdError.Width = 360

    Me.lblStationIdError.Left = 170
    Me.lblStationIdError.Top = 194
    Me.lblStationIdError.Width = 360

    Me.lblPathLocalError.Left = 170
    Me.lblPathLocalError.Top = 244
    Me.lblPathLocalError.Width = 360

    Me.lblPathSharePointError.Left = 170
    Me.lblPathSharePointError.Top = 294
    Me.lblPathSharePointError.Width = 360

    Me.lblSummary.Left = 24
    Me.lblSummary.Top = 344
    Me.lblSummary.Width = 530
    Me.lblSummary.Height = 64

    Me.btnOK.Left = 332
    Me.btnOK.Top = 420
    Me.btnOK.Width = 96
    Me.btnOK.Caption = "Setup"

    Me.btnCancel.Left = 438
    Me.btnCancel.Top = 420
    Me.btnCancel.Width = 96
    Me.btnCancel.Caption = "Close"

    If mBtnSharePointHelper Is Nothing Then
        Set mBtnSharePointHelper = Me.Controls.Add("Forms.CommandButton.1", "btnSharePointHelperRuntime", True)
    End If
    With mBtnSharePointHelper
        .Left = Me.txtPathSharePoint.Left + Me.txtPathSharePoint.Width + 8
        .Top = Me.txtPathSharePoint.Top - 1
        .Width = 72
        .Height = Me.txtPathSharePoint.Height + 2
        .Caption = "Find..."
        .ControlTipText = "Choose the locally synced invSys SharePoint root folder."
        .Visible = True
        .Enabled = True
    End With
End Sub

Private Sub CreateDynamicControls()
    CreatePromptLabel "lblUserIdPrompt", "UserId", 24, 26, 132
    CreatePromptLabel "lblPinPrompt", "PIN", 24, 76, 132
    CreatePromptLabel "lblConfirmPinPrompt", "Confirm PIN", 24, 126, 132
    CreatePromptLabel "lblWarehousePrompt", "Warehouse", 24, 176, 132
    CreatePromptLabel "lblStationPrompt", "Station", 24, 226, 132
    CreatePromptLabel "lblPathLocalPrompt", "Local Runtime Path", 24, 276, 132
    CreatePromptLabel "lblSharePointPrompt", "SharePoint Root", 24, 326, 132

    Set mTxtConfirmPin = Me.Controls.Add("Forms.TextBox.1", "txtConfirmPinRuntime", True)
    With mTxtConfirmPin
        .Left = 170
        .Top = 124
        .Width = 240
        .PasswordChar = "*"
        .Visible = True
    End With

    Set mLblConfirmPinError = Me.Controls.Add("Forms.Label.1", "lblConfirmPinErrorRuntime", True)
    With mLblConfirmPinError
        .Left = 170
        .Top = 144
        .Width = 360
        .Height = 12
        .Caption = ""
        .ForeColor = COLOR_ERROR
        .Visible = True
    End With

    Me.txtWarehouseId.Top = 174
    Me.lblWarehouseIdError.Top = 194
    Me.txtStationId.Top = 224
    Me.lblStationIdError.Top = 244
    Me.txtPathLocal.Top = 274
    Me.lblPathLocalError.Top = 294
    Me.txtPathSharePoint.Top = 324
    Me.lblPathSharePointError.Top = 344
    Me.lblSummary.Top = 370

    Me.btnOK.Top = 440
    Me.btnCancel.Top = 440

    Set mBtnOpen = Me.Controls.Add("Forms.CommandButton.1", "btnOpenReceivingRuntime", True)
    With mBtnOpen
        .Left = 24
        .Top = 440
        .Width = 138
        .Caption = "Open Workbook"
        .Visible = False
    End With
End Sub

Private Function BuildSpecFromForm(ByRef spec As modTesterSetup.TesterSetupSpec, _
                                   ByRef rawPin As String) As Boolean
    Dim confirmPin As String
    Dim isValid As Boolean

    spec.UserId = Trim$(CStr(Me.txtAdminUser.Value))
    rawPin = CStr(Me.txtWarehouseName.Value)
    confirmPin = CStr(mTxtConfirmPin.Value)
    spec.WarehouseId = Trim$(CStr(Me.txtWarehouseId.Value))
    spec.StationId = Trim$(CStr(Me.txtStationId.Value))
    spec.PathLocal = Trim$(CStr(Me.txtPathLocal.Value))
    spec.PathSharePointRoot = Trim$(CStr(Me.txtPathSharePoint.Value))

    isValid = True
    If spec.UserId = "" Then
        SetErrorLabel Me.lblAdminUserError, "UserId is required."
        isValid = False
    End If
    If rawPin = "" Then
        SetErrorLabel Me.lblWarehouseNameError, "PIN is required."
        isValid = False
    End If
    If confirmPin = "" Then
        SetErrorLabel mLblConfirmPinError, "Confirm PIN is required."
        isValid = False
    ElseIf StrComp(rawPin, confirmPin, vbBinaryCompare) <> 0 Then
        SetErrorLabel mLblConfirmPinError, "PIN entries must match."
        isValid = False
    End If
    If spec.WarehouseId = "" Then
        SetErrorLabel Me.lblWarehouseIdError, "Warehouse is required."
        isValid = False
    End If
    If spec.StationId = "" Then
        SetErrorLabel Me.lblStationIdError, "Station is required."
        isValid = False
    End If
    If spec.PathLocal = "" Then
        SetErrorLabel Me.lblPathLocalError, "Local runtime path is required."
        isValid = False
    End If
    If Not isValid Then Exit Function

    spec.PinHash = modAuth.HashUserCredential(rawPin)
    BuildSpecFromForm = True
End Function

Private Sub ApplyDefaults()
    mPathLocalTouched = False
    mLastSuggestedLocalPath = vbNullString
    mSetupSucceeded = False
    mOperatorWorkbookPath = vbNullString

    Me.txtAdminUser.Value = ResolveDefaultUserIdSetupForm()
    Me.txtWarehouseName.Value = ""
    mTxtConfirmPin.Value = ""
    Me.txtWarehouseId.Value = "WH1"
    Me.txtStationId.Value = "R1"
    Me.txtPathSharePoint.Value = modTesterSetup.DetectSharePointRoot("WH1")
    RefreshSuggestedLocalPath True
End Sub

Private Sub RefreshSuggestedLocalPath(ByVal forceApply As Boolean)
    Dim warehouseId As String
    Dim suggestedPath As String
    Dim currentValue As String

    warehouseId = Trim$(CStr(Me.txtWarehouseId.Value))
    suggestedPath = modDeploymentPaths.DefaultWarehouseRuntimeRootPath(warehouseId, False)

    currentValue = Trim$(CStr(Me.txtPathLocal.Value))
    If forceApply Or (Not mPathLocalTouched) Or currentValue = "" Or StrComp(currentValue, mLastSuggestedLocalPath, vbTextCompare) = 0 Then
        mFormBusy = True
        Me.txtPathLocal.Value = suggestedPath
        mFormBusy = False
        mPathLocalTouched = False
    End If
    mLastSuggestedLocalPath = suggestedPath
End Sub

Private Sub CreatePromptLabel(ByVal controlName As String, _
                              ByVal captionText As String, _
                              ByVal leftPos As Single, _
                              ByVal topPos As Single, _
                              ByVal widthVal As Single)
    Dim lbl As MSForms.Label

    Set lbl = Me.Controls.Add("Forms.Label.1", controlName, True)
    With lbl
        .Caption = captionText
        .Left = leftPos
        .Top = topPos
        .Width = widthVal
        .Height = 16
        .Visible = True
    End With
End Sub

Private Sub ClearValidationErrors()
    ClearErrorLabel Me.lblAdminUserError
    ClearErrorLabel Me.lblWarehouseNameError
    ClearErrorLabel mLblConfirmPinError
    ClearErrorLabel Me.lblWarehouseIdError
    ClearErrorLabel Me.lblStationIdError
    ClearErrorLabel Me.lblPathLocalError
    ClearErrorLabel Me.lblPathSharePointError
End Sub

Private Sub ClearErrorLabel(ByVal lbl As Object)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = ""
    lbl.ForeColor = COLOR_ERROR
End Sub

Private Sub SetErrorLabel(ByVal lbl As Object, ByVal messageText As String)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = Trim$(messageText)
    lbl.ForeColor = COLOR_ERROR
End Sub

Private Sub ShowSummary(ByVal messageText As String, ByVal foreColor As Long)
    Me.lblSummary.Caption = Trim$(messageText)
    Me.lblSummary.ForeColor = foreColor
    Me.Repaint
End Sub

Private Function ResolveDefaultUserIdSetupForm() As String
    ResolveDefaultUserIdSetupForm = Trim$(Environ$("USERNAME"))
    If ResolveDefaultUserIdSetupForm = "" Then ResolveDefaultUserIdSetupForm = Trim$(Application.UserName)
End Function

Private Sub InitializeSetupTesterAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me, 600, 500

    mAnchors.Add Me.txtAdminUser, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.txtWarehouseName, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mTxtConfirmPin, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.txtWarehouseId, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.txtStationId, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.txtPathLocal, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.txtPathSharePoint, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblAdminUserError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblWarehouseNameError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add mLblConfirmPinError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblWarehouseIdError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblStationIdError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblPathLocalError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblPathSharePointError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblSummary, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.btnOK, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.btnCancel, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    If Not mBtnOpen Is Nothing Then mAnchors.Add mBtnOpen, ANCHOR_LEFT Or ANCHOR_BOTTOM
    If Not mBtnSharePointHelper Is Nothing Then mAnchors.Add mBtnSharePointHelper, ANCHOR_RIGHT Or ANCHOR_TOP
End Sub
