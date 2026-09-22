VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAdminEventDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mFamily As MSForms.ComboBox
Private WithEvents mRows As MSForms.ListBox
Private WithEvents mEnabled As MSForms.CheckBox
Private WithEvents mUp As MSForms.CommandButton
Private WithEvents mDown As MSForms.CommandButton
Private WithEvents mSave As MSForms.CommandButton
Private WithEvents mReset As MSForms.CommandButton
Private WithEvents mReload As MSForms.CommandButton
Private mPreview As MSForms.TextBox
Private mStatus As MSForms.Label
Private mVersionLabel As MSForms.Label
Private mContext As String
Private mRequest As String
Private mVersion As Long
Private mLoading As Boolean
Private mCanEdit As Boolean

Public Sub Initialize(ByVal page As Object, ByVal context As String)
    Dim family As Variant
    mContext = context
    modAdminSettingsControls.AddControl page, "Label", "lblDetailFamily", "Event family", 12, 46, 68, 18
    Set mFamily = modAdminSettingsControls.AddControl(page, "ComboBox", "cmbDetailFamily", "", 84, 42, 248, 22)
    mFamily.Style = fmStyleDropDownList
    For Each family In Split(modEventDetailSettings.FamilyChoices(), vbLf)
        mFamily.AddItem family
    Next family
    mLoading = True: mFamily.ListIndex = 0: mLoading = False
    Set mVersionLabel = modAdminSettingsControls.AddControl(page, "Label", "lblDetailVersion", "", 350, 46, 332, 22)
    modAdminSettingsControls.AddControl page, "Label", "lblDetailRequired", _
        "Required identity, outcome, source, time and safety context remain visible.", 12, 74, 320, 38
    modAdminSettingsControls.AddControl page, "Label", "lblDetailColumns", "Field                               Show       Order       Required", 12, 112, 320, 16
    Set mRows = modAdminSettingsControls.AddControl(page, "ListBox", "lstDetailFields", "", 12, 130, 320, 222)
    mRows.ColumnCount = 5
    mRows.ColumnWidths = "0 pt;175 pt;35 pt;40 pt;55 pt"
    Set mEnabled = modAdminSettingsControls.AddControl(page, "CheckBox", "chkDetailEnabled", "Show selected field", 12, 362, 300, 22)
    Set mUp = modAdminSettingsControls.AddControl(page, "CommandButton", "btnDetailMoveUp", "Move Up", 12, 392, 86, 26)
    Set mDown = modAdminSettingsControls.AddControl(page, "CommandButton", "btnDetailMoveDown", "Move Down", 110, 392, 86, 26)
    modAdminSettingsControls.AddControl page, "Label", "lblDetailPreview", "Synthetic preview", 350, 86, 332, 18
    Set mPreview = modAdminSettingsControls.AddControl(page, "TextBox", "txtDetailPreview", "", 350, 112, 332, 290)
    mPreview.MultiLine = True: mPreview.Locked = True: mPreview.ScrollBars = fmScrollBarsVertical
    Set mSave = modAdminSettingsControls.AddControl(page, "CommandButton", "btnSaveDetailProfile", "Save Detail Profile", 12, 426, 160, 28)
    Set mReset = modAdminSettingsControls.AddControl(page, "CommandButton", "btnResetDetailProfile", "Reset to Default", 184, 426, 130, 28)
    Set mReload = modAdminSettingsControls.AddControl(page, "CommandButton", "btnReloadDetailProfile", "Reload", 326, 426, 86, 28)
    Set mStatus = modAdminSettingsControls.AddControl(page, "Label", "lblDetailProfileStatus", "", 12, 468, 670, 40)
    ReloadProfile
End Sub

Public Sub ReloadProfile(Optional ByRef outcome As String = "")
    Dim report As String
    mCanEdit = modEventDetailSettings.ReadEditor(mContext, mVersion, mRequest, report)
    outcome = IIf(mCanEdit, "REFRESHED", "FAILED")
    mSave.Enabled = mCanEdit: mReset.Enabled = mCanEdit
    If mRequest <> "" Then
        Render
    Else
        mLoading = True
        mRows.Clear: mPreview.Value = ""
        mLoading = False
    End If
    ShowSelection
    mStatus.Caption = report
    mVersionLabel.Caption = "Profile version " & CStr(mVersion) & IIf(mVersion = 0, " (built-in defaults)", "")
End Sub

Public Sub ResetProfile(Optional ByRef outcome As String = "")
    outcome = "REJECTED"
    If Not mCanEdit Then Exit Sub
    mRequest = modEventDetailSettings.DefaultRequest()
    Render
    mStatus.Caption = "Profile defaults staged. Save Detail Profile to apply."
    outcome = "STAGED"
End Sub

Public Function SaveProfile(Optional ByRef outcome As String = "") As Boolean
    Dim report As String
    SaveProfile = modEventDetailSettings.SaveProfile(mContext, mVersion, mRequest, report, outcome)
    If SaveProfile Then ReloadProfile
    mStatus.Caption = report
End Function

Private Sub Render()
    Dim selectedId As String, data As String, line As Variant, values As Variant, index As Long
    If mRows.ListIndex >= 0 Then selectedId = mRows.List(mRows.ListIndex, 0)
    mLoading = True
    mRows.Clear
    data = modEventDetailSettings.FieldRows(mRequest, CStr(mFamily.Value))
    If data <> "" Then
        For Each line In Split(data, vbLf)
            values = Split(CStr(line), vbTab)
            mRows.AddItem values(0)
            index = mRows.ListCount - 1
            mRows.List(index, 1) = values(1)
            mRows.List(index, 2) = IIf(values(2) = "True", "Yes", "No")
            mRows.List(index, 3) = values(3)
            mRows.List(index, 4) = IIf(values(4) = "True", "Required", "Optional")
            If values(0) = selectedId Then mRows.ListIndex = index
        Next line
    End If
    If mRows.ListIndex < 0 And mRows.ListCount > 0 Then mRows.ListIndex = 0
    mPreview.Value = modEventDetailSettings.Preview(mRequest, CStr(mFamily.Value))
    mLoading = False
    ShowSelection
End Sub

Private Sub ShowSelection()
    Dim selected As Boolean
    If mLoading Then Exit Sub
    selected = (mRows.ListIndex >= 0)
    mLoading = True
    mEnabled.Enabled = False: mUp.Enabled = False: mDown.Enabled = False
    If selected Then
        mEnabled.Value = (mRows.List(mRows.ListIndex, 2) = "Yes")
        mEnabled.Enabled = mCanEdit And mRows.List(mRows.ListIndex, 4) <> "Required"
        mUp.Enabled = mCanEdit And mRows.ListIndex > 0
        mDown.Enabled = mCanEdit And mRows.ListIndex < mRows.ListCount - 1
    End If
    mLoading = False
End Sub

Private Sub StageSelected(ByVal move As Long, Optional ByRef outcome As String = "")
    Dim staged As String
    outcome = "REJECTED"
    If mLoading Or Not mCanEdit Or mRows.ListIndex < 0 Then Exit Sub
    staged = modEventDetailSettings.StageField(mRequest, CStr(mFamily.Value), _
        CStr(mRows.List(mRows.ListIndex, 0)), CBool(mEnabled.Value), move)
    If staged = "" Then Exit Sub
    mRequest = staged
    Render
    mStatus.Caption = "Profile changes staged. Save Detail Profile to apply."
    outcome = "STAGED"
End Sub

Private Sub mFamily_Change()
    PerformSettings "ADMIN_DETAIL_SELECT_FAMILY"
End Sub
Private Sub mRows_Change()
    PerformSettings "ADMIN_DETAIL_SELECT_FIELD"
End Sub
Private Sub mEnabled_Click()
    PerformSettings "ADMIN_DETAIL_SHOW_FIELD"
End Sub
Private Sub mUp_Click()
    PerformSettings "ADMIN_DETAIL_MOVE_UP"
End Sub
Private Sub mDown_Click()
    PerformSettings "ADMIN_DETAIL_MOVE_DOWN"
End Sub
Private Sub mSave_Click()
    PerformSettings "ADMIN_DETAIL_SAVE"
End Sub
Private Sub mReset_Click()
    PerformSettings "ADMIN_DETAIL_RESET"
End Sub
Private Sub mReload_Click()
    PerformSettings "ADMIN_DETAIL_RELOAD"
End Sub

Private Sub PerformSettings(ByVal controlId As String)
    Dim activityId As String, notice As String, outcome As String
    If mLoading Or mStatus Is Nothing Then Exit Sub
    On Error GoTo Failed
    If mContext = "" Or mContext <> modActivity.CaptureContext() Then
        mStatus.Caption = "Session or warehouse changed. Reopen Settings."
        Exit Sub
    End If
    activityId = modActivity.BeginAction(controlId, mContext, notice)
    outcome = "REJECTED"
    Select Case controlId
        Case "ADMIN_DETAIL_SELECT_FAMILY"
            If mRequest <> "" Then Render: outcome = "SELECTED"
        Case "ADMIN_DETAIL_SELECT_FIELD"
            If mRows.ListIndex >= 0 Then ShowSelection: outcome = "SELECTED"
        Case "ADMIN_DETAIL_SHOW_FIELD": StageSelected 0, outcome
        Case "ADMIN_DETAIL_MOVE_UP": StageSelected -1, outcome
        Case "ADMIN_DETAIL_MOVE_DOWN": StageSelected 1, outcome
        Case "ADMIN_DETAIL_SAVE": SaveProfile outcome
        Case "ADMIN_DETAIL_RESET": ResetProfile outcome
        Case "ADMIN_DETAIL_RELOAD": ReloadProfile outcome
    End Select
Done:
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    If notice <> "" Then mStatus.Caption = mStatus.Caption & " " & notice
    Exit Sub
Failed:
    mLoading = False
    If outcome <> "COMPLETED" Then outcome = "FAILED"
    mStatus.Caption = "Settings action or display update failed. Reopen Settings to verify."
    Resume Done
End Sub
