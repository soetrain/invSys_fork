VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCreateWarehouse 
   Caption         =   "UserForm1"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   4580
   OleObjectBlob   =   "frmCreateWarehouse.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCreateWarehouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mBtnHubPathHelper As MSForms.CommandButton
Attribute mBtnHubPathHelper.VB_VarHelpID = -1
Private WithEvents mBtnSharePointHelper As MSForms.CommandButton
Attribute mBtnSharePointHelper.VB_VarHelpID = -1
Private WithEvents mBtnOpenReceiving As MSForms.CommandButton
Attribute mBtnOpenReceiving.VB_VarHelpID = -1
Private WithEvents mBtnPublishExisting As MSForms.CommandButton
Attribute mBtnPublishExisting.VB_VarHelpID = -1
Private mWriteContext As String
Private mActivity As cProvisioningActivity

Private mPathLocalTouched As Boolean
Private mLastSuggestedLocalPath As String
Private mFormBusy As Boolean
Private mLocalBootstrapComplete As Boolean
Private mCreatedWarehouseId As String
Private mCreatedWarehouseName As String
Private mCreatedStationId As String
Private mCreatedAdminUser As String
Private mCreatedPathLocal As String
Private mCreatedPathSharePoint As String
Private mCreatedOperatorWorkbookPath As String
Private mDefaultOkLeft As Single
Private mDefaultOkWidth As Single
Private mDefaultCancelLeft As Single
Private mDefaultCancelWidth As Single

Private Const COLOR_ERROR As Long = 255
Private Const COLOR_SUCCESS As Long = 32768
Private Const COLOR_INFO As Long = 0
Private Const ANCHOR_LEFT As Long = 1
Private Const ANCHOR_TOP As Long = 2
Private Const ANCHOR_RIGHT As Long = 4
Private Const ANCHOR_BOTTOM As Long = 8

Private mAnchors As Object
Private mResizeInitialized As Boolean
Private mInitialPin As MSForms.TextBox
Private mConfirmPin As MSForms.TextBox
Private WithEvents mWarehousePurpose As MSForms.ComboBox
Attribute mWarehousePurpose.VB_VarHelpID = -1

Public Sub ObserveEntry(ByVal observer As cProvisioningActivity, ByVal writeContext As String)
    Set mActivity = observer
    mWriteContext = writeContext
End Sub

Private Function BeginActivity(ByVal suffix As String) As String
    If Not mActivity Is Nothing Then BeginActivity = mActivity.BeginAction(suffix)
End Function

Private Sub FinishActivity(ByVal action As String, ByVal outcome As String)
    If Not mActivity Is Nothing Then mActivity.Finish action, outcome
End Sub

Private Sub RaiseActivityFailure(ByVal action As String)
    Dim number As Long, source As String, description As String
    number = Err.Number: source = Err.Source: description = Err.Description
    FinishActivity action, "FAILED"
    Err.Raise number, source, description
End Sub

Private Sub UserForm_Initialize()
    mWriteContext = modAuthUserCommands.CaptureContext()
    Me.Caption = "Create Warehouse"
    Me.Width = 620
    Me.Height = 470
    Me.StartUpPosition = 1
    mDefaultOkLeft = Me.btnOK.Left
    mDefaultOkWidth = Me.btnOK.Width
    mDefaultCancelLeft = Me.btnCancel.Left
    mDefaultCancelWidth = Me.btnCancel.Width
    ConfigureSummaryArea
    RestoreDefaultButtonLayout
    mFormBusy = True
    Me.txtStationId.Value = modStationIdentity.CurrentComputerStationId()
    Me.txtAdminUser.Value = ResolveDefaultAdminUserForm()
    Me.txtPathSharePoint.Value = ResolveDefaultSharePointRootForm()
    ConfigureHubPathHelperButton
    ConfigureSharePointHelperButton
    ConfigureOpenReceivingButton
    ConfigureWarehousePurpose
    Set mInitialPin = AddPrivateEntry("InitialPin", "Initial Admin PIN", 325)
    Set mConfirmPin = AddPrivateEntry("ConfirmPin", "Confirm PIN", 462)
    mInitialPin.TabIndex = mWarehousePurpose.TabIndex + 1
    mConfirmPin.TabIndex = mInitialPin.TabIndex + 1
    ConfigurePublishExistingButton
    Me.btnOK.Caption = "Create"
    Me.btnCancel.Caption = "Cancel"
    Me.chkPublishInitial.Value = True
    mPathLocalTouched = False
    mLastSuggestedLocalPath = vbNullString
    RefreshSuggestedLocalPath True
    ClearValidationErrors
    InitializeCreateWarehouseAnchors
    ShowSummary "Choose the warehouse hub folder first. For Synology, this should be the SMB warehouse path, for example \\DS920\<share>\WH1. Then choose the locally synced invSys SharePoint root that contains Addins, Events, Snapshots, and TesterPackage.", COLOR_INFO
    mFormBusy = False
End Sub

Private Sub UserForm_Activate()
    If Not mActivity Is Nothing Then mActivity.Presented Me.Visible
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

Private Sub UserForm_Terminate()
    Set mAnchors = Nothing
    Set mActivity = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim action As String
    If Cancel <> 0 Then Exit Sub
    ClearInitialCredential
    action = BeginActivity("CLOSE")
    FinishActivity action, "CLOSED"
End Sub

Private Sub txtWarehouseId_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblWarehouseIdError
    RefreshSuggestedLocalPath False
End Sub

Private Sub txtWarehouseName_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblWarehouseNameError
End Sub

Private Sub txtStationId_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblStationIdError
End Sub

Private Sub txtAdminUser_Change()
    If mFormBusy Then Exit Sub
    ClearErrorLabel Me.lblAdminUserError
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
    Dim action As String
    On Error GoTo Failed
    action = BeginActivity("FIND_SHAREPOINT")

    warehouseId = Trim$(CStr(Me.txtWarehouseId.Value))
    candidate = modTesterSetup.DetectSharePointRoot(warehouseId)
    If candidate = "" Then
        candidate = modTesterSetup.BrowseForSharePointRoot(Trim$(CStr(Me.txtPathSharePoint.Value)))
    End If

    If candidate = "" Then
        ShowSummary "SharePoint root was not detected. Pick the locally synced invSys root folder manually.", COLOR_INFO
        FinishActivity action, "UNAVAILABLE"
        Exit Sub
    End If

    mFormBusy = True
    Me.txtPathSharePoint.Value = candidate
    mFormBusy = False
    ClearErrorLabel Me.lblPathSharePointError
    ShowSummary "SharePoint root detected.", COLOR_SUCCESS
    FinishActivity action, "STAGED"
    Exit Sub
Failed:
    RaiseActivityFailure action
End Sub

Private Sub mBtnHubPathHelper_Click()
    Dim candidate As String
    Dim action As String
    On Error GoTo Failed
    action = BeginActivity("BROWSE_HUB")

    candidate = modDeploymentPaths.BrowseForFolderPath(Trim$(CStr(Me.txtPathLocal.Value)), "Choose Warehouse Hub Folder")
    If candidate = "" Then FinishActivity action, "CANCELLED": Exit Sub

    mFormBusy = True
    Me.txtPathLocal.Value = candidate
    mFormBusy = False
    mPathLocalTouched = True
    ClearErrorLabel Me.lblPathLocalError
    ShowSummary "Warehouse hub path selected. For Synology, this should be the SMB warehouse folder, for example \\DS920\<share>\WH1.", COLOR_SUCCESS
    FinishActivity action, "STAGED"
    Exit Sub
Failed:
    RaiseActivityFailure action
End Sub

Private Sub chkPublishInitial_Click()
    Dim action As String
    If mFormBusy Then Exit Sub
    action = BeginActivity("PUBLISH_OPTION")
    ClearErrorLabel Me.lblPathSharePointError
    FinishActivity action, "STAGED"
End Sub

Private Sub mWarehousePurpose_Change()
    Dim action As String
    If mFormBusy Then Exit Sub
    action = BeginActivity("PURPOSE")
    FinishActivity action, "SELECTED"
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnOK_Click()
    Dim warehouseId As String
    Dim warehouseName As String
    Dim stationId As String
    Dim adminUser As String
    Dim pathLocal As String
    Dim pathSharePoint As String
    Dim addinsFolder As String
    Dim summaryText As String
    Dim commandReport As String
    Dim action As String, outcome As String, publishOutcome As String
    On Error GoTo Failed

    If StrComp(Me.btnOK.Caption, "Close", vbTextCompare) = 0 Then
        Unload Me
        Exit Sub
    End If

    outcome = "REJECTED"
    If mLocalBootstrapComplete Then action = BeginActivity("RETRY_PUBLISH") Else action = BeginActivity("SUBMIT")
    ClearValidationErrors
    If Not BuildSpecFromForm(warehouseId, warehouseName, stationId, adminUser, pathLocal, pathSharePoint) Then
        ShowSummary "Fix the highlighted fields and try again.", COLOR_ERROR
        GoTo ActivityFinished
    End If
    If Not mLocalBootstrapComplete Then
        If Not modWarehouseCreateCommands.MatchingInitialCredential(CStr(mInitialPin.Value), CStr(mConfirmPin.Value)) Then
            ShowSummary "Enter matching initial Admin PINs before creating the warehouse.", COLOR_ERROR
            mInitialPin.SetFocus
            GoTo ActivityFinished
        End If
    End If

    If pathSharePoint <> "" Then addinsFolder = pathSharePoint & "\Addins"
    If Not modLocalAddinsRegistration.EnsureLocalInvSysAddinsRegistered(addinsFolder, summaryText) Then
        ShowSummary "invSys add-ins are not registered cleanly for this Excel session." & vbCrLf & summaryText, COLOR_ERROR
        outcome = "FAILED": GoTo ActivityFinished
    End If

    If mLocalBootstrapComplete Then
        pathLocal = mCreatedPathLocal
        mCreatedPathSharePoint = pathSharePoint
        If Not CBool(Me.chkPublishInitial.Value) Then
            ShowSummary "Warehouse hub already exists. Check the publish box to retry SharePoint publish, or click Close.", COLOR_INFO
            GoTo ActivityFinished
        End If

        If modWarehouseCreateCommands.PublishExisting(mWriteContext, mCreatedPathLocal, mCreatedPathSharePoint, commandReport, mCreatedWarehouseId, outcome) Then
            summaryText = "Warehouse hub already existed from this session." & vbCrLf & _
                          commandReport
            MarkFormComplete summaryText, True
        Else
            summaryText = "Warehouse hub was created, but SharePoint publish still failed." & vbCrLf & _
                          commandReport
            Me.btnOK.Caption = "Retry Publish"
            Me.btnCancel.Caption = "Close"
            ShowSummary summaryText, COLOR_ERROR
        End If
        GoTo ActivityFinished
    End If

    If modWarehouseBootstrap.WarehouseArtifactsExistAtPath(warehouseId, pathLocal) Then
        SetErrorCaption Me.lblWarehouseIdError, "Warehouse artifacts already exist at the selected hub path."
        ShowSummary "Choose a different WarehouseId or hub path and try again.", COLOR_ERROR
        GoTo ActivityFinished
    End If

    If Not modWarehouseCreateCommands.CreateWarehouse(mWriteContext, warehouseId, warehouseName, stationId, adminUser, pathLocal, pathSharePoint, CStr(mWarehousePurpose.Value), commandReport, mCreatedOperatorWorkbookPath, outcome, CStr(mInitialPin.Value), CStr(mConfirmPin.Value)) Then
        ClearInitialCredential
        ShowSummary commandReport, COLOR_ERROR
        GoTo ActivityFinished
    End If

    mCreatedWarehouseId = warehouseId
    mCreatedWarehouseName = warehouseName
    mCreatedStationId = stationId
    mCreatedAdminUser = adminUser
    mCreatedPathLocal = pathLocal
    mCreatedPathSharePoint = pathSharePoint
    mLocalBootstrapComplete = True
    ClearInitialCredential
    mInitialPin.Enabled = False
    mConfirmPin.Enabled = False
    mWarehousePurpose.Enabled = False
    summaryText = "Warehouse hub bootstrap complete for " & warehouseId & "."
    If mCreatedOperatorWorkbookPath <> "" Then
        summaryText = summaryText & vbCrLf & "Receiving workbook created: " & GetFileNameForm(mCreatedOperatorWorkbookPath)
        If Not mBtnOpenReceiving Is Nothing Then
            mBtnOpenReceiving.Visible = True
            mBtnOpenReceiving.Enabled = True
        End If
    End If

    If CBool(Me.chkPublishInitial.Value) Then
        If modWarehouseCreateCommands.PublishExisting(mWriteContext, pathLocal, pathSharePoint, commandReport, mCreatedWarehouseId, publishOutcome) Then
            outcome = "CREATED_PUBLISHED"
            summaryText = summaryText & vbCrLf & commandReport
            MarkFormComplete summaryText, True
        Else
            outcome = "CREATED_PUBLISH_FAILED"
            If publishOutcome = "DENIED" Then outcome = "CREATED_PUBLISH_DENIED"
            mCreatedPathSharePoint = pathSharePoint
            summaryText = summaryText & vbCrLf & commandReport & vbCrLf & _
                          "Use Retry Publish, or reopen and choose Publish Existing Warehouse."
            Me.btnOK.Caption = "Retry Publish"
            Me.btnCancel.Caption = "Close"
            ShowSummary summaryText, COLOR_ERROR
        End If
    Else
        summaryText = summaryText & vbCrLf & "Initial SharePoint publish skipped."
        MarkFormComplete summaryText, True
    End If
ActivityFinished:
    FinishActivity action, outcome
    Exit Sub
Failed:
    ClearInitialCredential
    FinishActivity action, "FAILED"
    ShowSummary "Warehouse creation could not finish. Inspect the destination before retrying.", COLOR_ERROR
End Sub

Private Sub mBtnPublishExisting_Click()
    Dim report As String, action As String, outcome As String
    On Error GoTo Failed
    action = BeginActivity("PUBLISH_EXISTING")
    If modWarehouseCreateCommands.PublishExisting(mWriteContext, CStr(Me.txtPathLocal.Value), CStr(Me.txtPathSharePoint.Value), report, "", outcome) Then
        MarkFormComplete report, True
    Else
        ShowSummary report, COLOR_ERROR
    End If
    FinishActivity action, outcome
    Exit Sub
Failed:
    RaiseActivityFailure action
End Sub

Private Sub ConfigurePublishExistingButton()
    Set mBtnPublishExisting = Me.Controls.Add("Forms.CommandButton.1", "btnPublishExisting", True)
    With mBtnPublishExisting
        .Caption = "Publish Existing Warehouse"
        .ControlTipText = "Publish the saved warehouse in the hub folder without creating it again."
        .Left = 150: .Top = Me.btnOK.Top: .Width = 186: .Height = Me.btnOK.Height
    End With
    Me.btnOK.Left = Me.InsideWidth - 216: Me.btnOK.Width = 96
    Me.btnCancel.Left = Me.InsideWidth - 108: Me.btnCancel.Width = 96
End Sub

Private Sub mBtnOpenReceiving_Click()
    Dim wb As Workbook
    Dim action As String

    action = BeginActivity("OPEN_WORKBOOK")

    If mCreatedOperatorWorkbookPath = "" Then
        ShowSummary "Receiving workbook path was not captured.", COLOR_ERROR
        FinishActivity action, "REJECTED"
        Exit Sub
    End If
    If Len(Dir$(mCreatedOperatorWorkbookPath, vbNormal)) = 0 Then
        ShowSummary "Receiving workbook was not found: " & mCreatedOperatorWorkbookPath, COLOR_ERROR
        FinishActivity action, "REJECTED"
        Exit Sub
    End If

    On Error GoTo FailOpen
    Set wb = Application.Workbooks.Open(Filename:=mCreatedOperatorWorkbookPath, UpdateLinks:=0, ReadOnly:=False, IgnoreReadOnlyRecommended:=True, Notify:=False, AddToMru:=False)
    If wb Is Nothing Then GoTo FailOpen
    wb.Activate
    ShowSummary "Receiving workbook opened. Use Confirm Writes to test the new warehouse.", COLOR_SUCCESS
    FinishActivity action, "OPENED"
    Exit Sub

FailOpen:
    ShowSummary "Receiving workbook could not be opened: " & Err.Description, COLOR_ERROR
    FinishActivity action, "FAILED"
End Sub

Private Sub MarkFormComplete(ByVal summaryText As String, ByVal includeCloseHint As Boolean)
    ClearInitialCredential
    If includeCloseHint Then summaryText = summaryText & vbCrLf & "Click Close to finish."
    Me.Tag = "COMPLETE"
    Me.btnOK.Caption = "Close"
    ConfigureCompleteButtonLayout
    If Not mAnchors Is Nothing Then mAnchors.ResizeControls
    ShowSummary summaryText, COLOR_SUCCESS
End Sub

Private Function AddPrivateEntry(ByVal controlName As String, ByVal prompt As String, ByVal leftPos As Single) As MSForms.TextBox
    Dim fieldLabel As MSForms.Label, entry As MSForms.TextBox
    Set fieldLabel = Me.Controls.Add("Forms.Label.1", "lbl" & controlName, True)
    fieldLabel.Caption = prompt
    fieldLabel.Move leftPos, mWarehousePurpose.Top - 3, 124, 13
    Set entry = Me.Controls.Add("Forms.TextBox.1", "txt" & controlName, True)
    entry.PasswordChar = "*"
    entry.Move leftPos, fieldLabel.Top + fieldLabel.Height, 124, 18
    Set AddPrivateEntry = entry
End Function

Private Sub ClearInitialCredential()
    If Not mInitialPin Is Nothing Then mInitialPin.Value = ""
    If Not mConfirmPin Is Nothing Then mConfirmPin.Value = ""
End Sub

Private Function BuildSpecFromForm(ByRef warehouseId As String, _
                                   ByRef warehouseName As String, _
                                   ByRef stationId As String, _
                                   ByRef adminUser As String, _
                                   ByRef pathLocal As String, _
                                   ByRef pathSharePoint As String) As Boolean
    Dim report As String
    Dim isValid As Boolean

    warehouseId = Trim$(CStr(Me.txtWarehouseId.Value))
    warehouseName = Trim$(CStr(Me.txtWarehouseName.Value))
    stationId = Trim$(CStr(Me.txtStationId.Value))
    adminUser = Trim$(CStr(Me.txtAdminUser.Value))
    pathLocal = Trim$(CStr(Me.txtPathLocal.Value))
    pathSharePoint = Trim$(CStr(Me.txtPathSharePoint.Value))

    isValid = modAdminConsole.ValidateWarehouseSpecAdmin(warehouseId, warehouseName, stationId, adminUser, pathLocal, pathSharePoint, report, CStr(mWarehousePurpose.Value))
    If Not isValid Then SetErrorCaption Me.lblWarehouseIdError, report

    If stationId = "" Then
        SetErrorCaption Me.lblStationIdError, "StationId is required."
        isValid = False
    End If
    If adminUser = "" Then
        SetErrorCaption Me.lblAdminUserError, "AdminUser is required."
        isValid = False
    End If
    If pathLocal = "" Then
        SetErrorCaption Me.lblPathLocalError, "Warehouse hub path is required."
        isValid = False
    End If
    If CBool(Me.chkPublishInitial.Value) And pathSharePoint = "" Then
        SetErrorCaption Me.lblPathSharePointError, "SharePoint path is required when initial publish is enabled."
        isValid = False
    End If

    BuildSpecFromForm = isValid
End Function

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

Private Sub ClearValidationErrors()
    ClearErrorLabel Me.lblWarehouseIdError
    ClearErrorLabel Me.lblWarehouseNameError
    ClearErrorLabel Me.lblStationIdError
    ClearErrorLabel Me.lblAdminUserError
    ClearErrorLabel Me.lblPathLocalError
    ClearErrorLabel Me.lblPathSharePointError
End Sub

Private Sub ClearErrorLabel(ByVal lbl As MSForms.Label)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = ""
    lbl.foreColor = COLOR_ERROR
End Sub

Private Sub SetErrorCaption(ByVal lbl As MSForms.Label, ByVal messageText As String)
    If lbl Is Nothing Then Exit Sub
    lbl.Caption = Trim$(messageText)
    lbl.foreColor = COLOR_ERROR
End Sub

Private Sub ShowSummary(ByVal messageText As String, ByVal foreColor As Long)
    Me.lblSummary.Caption = modCreateWarehousePresentation.FormatSummaryMessage(messageText)
    Me.lblSummary.foreColor = foreColor
    Me.Repaint
End Sub

Private Sub ConfigureSummaryArea()
    With Me.lblSummary
        .WordWrap = True
        .AutoSize = False
        .Left = 18
        .Top = 282
        .Width = Me.InsideWidth - 36
        .Height = 108
    End With
End Sub

Private Sub ConfigureCompleteButtonLayout()
    Me.btnCancel.Visible = False
    Me.btnCancel.Enabled = False
    Me.btnOK.Width = 96
    Me.btnOK.Left = Me.InsideWidth - Me.btnOK.Width - 24
End Sub

Private Sub RestoreDefaultButtonLayout()
    Me.btnCancel.Visible = True
    Me.btnCancel.Enabled = True
    Me.btnCancel.Left = mDefaultCancelLeft
    Me.btnCancel.Width = mDefaultCancelWidth
    Me.btnOK.Left = mDefaultOkLeft
    Me.btnOK.Width = mDefaultOkWidth
End Sub

Private Sub InitializeCreateWarehouseAnchors()
    Set mAnchors = modDynamicForms.CreateFormAnchorManager()
    mAnchors.Initialize Me, 620, 610
    mAnchors.Add mWarehousePurpose, ANCHOR_LEFT Or ANCHOR_TOP
    mAnchors.Add Me.Controls("lblWarehousePurpose"), ANCHOR_LEFT Or ANCHOR_TOP

    mAnchors.Add Me.txtWarehouseId, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.txtWarehouseName, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.txtStationId, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.txtAdminUser, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.txtPathLocal, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.txtPathSharePoint, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblWarehouseIdError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblWarehouseNameError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblStationIdError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblAdminUserError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblPathLocalError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.lblPathSharePointError, ANCHOR_LEFT Or ANCHOR_TOP Or ANCHOR_RIGHT
    mAnchors.Add Me.chkPublishInitial, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add Me.lblSummary, ANCHOR_LEFT Or ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.btnCancel, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    mAnchors.Add Me.btnOK, ANCHOR_RIGHT Or ANCHOR_BOTTOM
    If Not mBtnHubPathHelper Is Nothing Then mAnchors.Add mBtnHubPathHelper, ANCHOR_RIGHT Or ANCHOR_TOP
    If Not mBtnSharePointHelper Is Nothing Then mAnchors.Add mBtnSharePointHelper, ANCHOR_RIGHT Or ANCHOR_TOP
    If Not mBtnOpenReceiving Is Nothing Then mAnchors.Add mBtnOpenReceiving, ANCHOR_LEFT Or ANCHOR_BOTTOM
    mAnchors.Add mBtnPublishExisting, ANCHOR_LEFT Or ANCHOR_BOTTOM
End Sub

Private Sub ConfigureWarehousePurpose()
    Dim ctl As Object, purposeLabel As MSForms.Label, rowTop As Single
    rowTop = Me.chkPublishInitial.Top
    For Each ctl In Me.Controls
        If ctl.Top >= rowTop Then ctl.Top = ctl.Top + 36
    Next ctl
    Me.lblSummary.Top = Me.chkPublishInitial.Top + Me.chkPublishInitial.Height + 12
    Me.btnOK.Top = Me.lblSummary.Top + Me.lblSummary.Height + 18
    Me.btnCancel.Top = Me.btnOK.Top
    mBtnOpenReceiving.Top = Me.btnOK.Top
    Me.Height = 610
    Set purposeLabel = Me.Controls.Add("Forms.Label.1", "lblWarehousePurpose", True)
    With purposeLabel
        .Caption = "Warehouse purpose"
        .Left = 18
        .Top = rowTop + 3
        .Width = 118
        .Height = 18
    End With
    Set mWarehousePurpose = Me.Controls.Add("Forms.ComboBox.1", "cboWarehousePurpose", True)
    With mWarehousePurpose
        .Left = 140
        .Top = rowTop
        .Width = 180
        .Height = 22
        .Style = fmStyleDropDownList
        .AddItem "Operational"
        .AddItem "Training"
        .ListIndex = 0
        .ControlTipText = "Automatic How-To replay requires a separately generated Training warehouse."
    End With
End Sub

Private Function ResolveDefaultAdminUserForm() As String
    ResolveDefaultAdminUserForm = Trim$(Environ$("USERNAME"))
End Function

Private Function ResolveDefaultSharePointRootForm() As String
    On Error Resume Next
    If modConfig.IsLoaded() Then
        ResolveDefaultSharePointRootForm = Trim$(CStr(modConfig.GetString("PathSharePointRoot", "")))
    End If
    On Error GoTo 0

    If ResolveDefaultSharePointRootForm = "" Then
        ResolveDefaultSharePointRootForm = modTesterSetup.DetectSharePointRoot(Trim$(CStr(Me.txtWarehouseId.Value)))
    End If
End Function

Private Function GetFileNameForm(ByVal fullPath As String) As String
    Dim slashPos As Long

    fullPath = Trim$(Replace$(fullPath, "/", "\"))
    slashPos = InStrRev(fullPath, "\")
    If slashPos > 0 Then
        GetFileNameForm = Mid$(fullPath, slashPos + 1)
    Else
        GetFileNameForm = fullPath
    End If
End Function

Private Sub ConfigureHubPathHelperButton()
    If mBtnHubPathHelper Is Nothing Then
        Set mBtnHubPathHelper = Me.Controls.Add("Forms.CommandButton.1", "btnHubPathHelperRuntime", True)
    End If

    Me.txtPathLocal.Width = 258
    With mBtnHubPathHelper
        .Caption = "Browse..."
        .Left = Me.txtPathLocal.Left + Me.txtPathLocal.Width + 8
        .Top = Me.txtPathLocal.Top - 1
        .Width = 72
        .Height = Me.txtPathLocal.Height + 2
        .ControlTipText = "Choose the warehouse hub folder. For Synology, pick the SMB warehouse folder such as \\DS920\\<share>\\WH1."
        .Visible = True
        .Enabled = True
    End With
End Sub

Private Sub ConfigureOpenReceivingButton()
    If mBtnOpenReceiving Is Nothing Then
        Set mBtnOpenReceiving = Me.Controls.Add("Forms.CommandButton.1", "btnOpenReceivingRuntime", True)
    End If

    With mBtnOpenReceiving
        .Caption = "Open Workbook"
        .Left = 18
        .Top = Me.btnOK.Top
        .Width = 120
        .Height = Me.btnOK.Height
        .ControlTipText = "Open the generated receiving operator workbook for this warehouse."
        .Visible = False
        .Enabled = False
    End With
End Sub

Private Sub ConfigureSharePointHelperButton()
    If mBtnSharePointHelper Is Nothing Then
        Set mBtnSharePointHelper = Me.Controls.Add("Forms.CommandButton.1", "btnSharePointHelperRuntime", True)
    End If

    Me.txtPathSharePoint.Width = 258
    With mBtnSharePointHelper
        .Caption = "Find..."
        .Left = Me.txtPathSharePoint.Left + Me.txtPathSharePoint.Width + 8
        .Top = Me.txtPathSharePoint.Top - 1
        .Width = 72
        .Height = Me.txtPathSharePoint.Height + 2
        .ControlTipText = "Choose the locally synced invSys SharePoint root folder."
        .Visible = True
        .Enabled = True
    End With
End Sub
