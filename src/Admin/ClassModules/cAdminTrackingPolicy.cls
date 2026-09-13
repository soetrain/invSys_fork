VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAdminTrackingPolicy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mRows As MSForms.ListBox
Private WithEvents mCapture As MSForms.CheckBox
Private WithEvents mAdminVisible As MSForms.CheckBox
Private WithEvents mView As MSForms.ComboBox
Private WithEvents mCollect As MSForms.CheckBox
Private WithEvents mVisible As MSForms.CheckBox
Private WithEvents mSequence As MSForms.CheckBox
Private WithEvents mSave As MSForms.CommandButton
Private WithEvents mReset As MSForms.CommandButton
Private WithEvents mReload As MSForms.CommandButton
Private mStatus As MSForms.Label
Private mContext As String
Private mRequest As String
Private mVersion As Long
Private mSavedCatalog As Long
Private mLoading As Boolean

Public Sub Initialize(ByVal page As Object, ByVal context As String)
    mContext = context
    Set mCapture = AddControl(page, "CheckBox", "chkCaptureControls", "Capture recorded controls", 12, 40, 170, 22)
    Set mAdminVisible = AddControl(page, "CheckBox", "chkAdminEventsVisible", "Show optional Admin events", 190, 40, 190, 22)
    AddControl page, "Label", "lblWarehouseDefaultView", "Warehouse default view", 396, 20, 200, 18
    Set mView = AddControl(page, "ComboBox", "cmbWarehouseDefaultView", "", 396, 40, 220, 22)
    mView.Style = fmStyleDropDownList
    mView.AddItem "How-To": mView.AddItem "Diagnostic": mView.AddItem "Compare both"
    AddControl page, "Label", "lblRequiredTracking", "Canonical business events and audit: Required (always collected).", 12, 70, 670, 18
    AddControl page, "Label", "lblPolicyFamily", "Family / control", 12, 96, 340, 18
    AddControl page, "Label", "lblPolicyFlags", "Collect       Visible       Sequence       Availability", 366, 96, 316, 18
    Set mRows = AddControl(page, "ListBox", "lstTrackingControls", "", 12, 114, 670, 134)
    mRows.ColumnCount = 7
    mRows.ColumnWidths = "0 pt;140 pt;210 pt;45 pt;45 pt;55 pt;160 pt"
    Set mCollect = AddControl(page, "CheckBox", "chkControlCollect", "Collect selected control", 12, 258, 200, 22)
    Set mVisible = AddControl(page, "CheckBox", "chkControlVisible", "Visible in Viewer", 224, 258, 170, 22)
    Set mSequence = AddControl(page, "CheckBox", "chkControlSequence", "Eligible for recorded sequence", 408, 258, 270, 22)
    Set mSave = AddControl(page, "CommandButton", "btnSaveTrackingPolicy", "Save Tracking Policy", 12, 294, 166, 28)
    Set mReset = AddControl(page, "CommandButton", "btnResetTrackingPolicy", "Reset to Default", 190, 294, 138, 28)
    Set mReload = AddControl(page, "CommandButton", "btnReloadTrackingPolicy", "Reload", 340, 294, 86, 28)
    Set mStatus = AddControl(page, "Label", "lblTrackingPolicyStatus", "", 12, 330, 670, 48)
    ReloadPolicy
End Sub

Public Sub ReloadPolicy()
    Dim report As String, loaded As Boolean
    loaded = modTrackingPolicySettings.ReadEditor(mContext, mVersion, mRequest, report, mSavedCatalog)
    mSave.Enabled = loaded: mReset.Enabled = loaded
    mRows.Enabled = loaded: mCapture.Enabled = loaded: mAdminVisible.Enabled = loaded: mView.Enabled = loaded
    mCollect.Enabled = loaded: mVisible.Enabled = loaded: mSequence.Enabled = loaded
    If loaded Then Render Else mRows.Clear
    mStatus.Caption = report
End Sub

Public Sub ResetPolicy()
    If mRequest = "" Then Exit Sub
    mRequest = modTrackingPolicySettings.DefaultRequest()
    mSavedCatalog = 0
    Render
    mStatus.Caption = "Defaults staged. Save Tracking Policy to apply."
End Sub

Public Function SavePolicy() As Boolean
    Dim report As String
    SavePolicy = modTrackingPolicySettings.SavePolicy(mContext, mVersion, mRequest, report)
    If SavePolicy Then
        ReloadPolicy
        If mRequest = "" Then report = report & " Reload unavailable. Reopen Settings to verify."
    End If
    mStatus.Caption = report
End Function

Private Sub Render()
    Dim rows As Variant, row As Variant, values As Variant, column As Long, selected As Long
    selected = mRows.ListIndex
    mLoading = True
    mCapture.Value = CBool(modTrackingPolicySettings.FieldValue(mRequest, "ViewerActionPathCaptureEnabled"))
    mAdminVisible.Value = CBool(modTrackingPolicySettings.FieldValue(mRequest, "AdminViewerEventLoggingEnabled"))
    mView.Value = modTrackingPolicySettings.FieldValue(mRequest, "DefaultView")
    mRows.Clear
    rows = Split(modTrackingPolicySettings.ControlRows(mRequest, mSavedCatalog), vbLf)
    For Each row In rows
        values = Split(CStr(row), vbTab)
        mRows.AddItem values(0)
        For column = 1 To 6
            mRows.List(mRows.ListCount - 1, column) = values(column)
        Next column
    Next row
    If selected < 0 Or selected >= mRows.ListCount Then selected = 0
    mRows.ListIndex = selected
    mLoading = False
    ShowSelection
End Sub

Private Sub ShowSelection()
    If mRows.ListIndex < 0 Then Exit Sub
    mLoading = True
    mCollect.Value = CBool(mRows.List(mRows.ListIndex, 3))
    mVisible.Value = CBool(mRows.List(mRows.ListIndex, 4))
    mSequence.Value = CBool(mRows.List(mRows.ListIndex, 5))
    mLoading = False
End Sub

Private Sub StageTop(ByVal field As String, ByVal value As String)
    Dim staged As String
    If mLoading Or mRequest = "" Then Exit Sub
    staged = modTrackingPolicySettings.StageValue(mRequest, field, value)
    If staged = "" Then Exit Sub
    mRequest = staged
    mStatus.Caption = "Policy changes staged. Save Tracking Policy to apply."
End Sub

Private Sub StageControl(ByVal field As String, ByVal value As Boolean)
    Dim staged As String
    If mLoading Or mRows.ListIndex < 0 Or mRequest = "" Then Exit Sub
    staged = modTrackingPolicySettings.StageValue(mRequest, field, CStr(value), CStr(mRows.List(mRows.ListIndex, 0)))
    If staged = "" Then Exit Sub
    mRequest = staged
    Render
    mStatus.Caption = "Policy changes staged. Save Tracking Policy to apply."
End Sub

Private Sub mRows_Click()
    If Not mLoading Then ShowSelection
End Sub
Private Sub mCapture_Click()
    StageTop "ViewerActionPathCaptureEnabled", CStr(CBool(mCapture.Value))
End Sub
Private Sub mAdminVisible_Click()
    StageTop "AdminViewerEventLoggingEnabled", CStr(CBool(mAdminVisible.Value))
End Sub
Private Sub mView_Change()
    StageTop "DefaultView", CStr(mView.Value)
End Sub
Private Sub mCollect_Click()
    StageControl "Collect", CBool(mCollect.Value)
End Sub
Private Sub mVisible_Click()
    StageControl "Visible", CBool(mVisible.Value)
End Sub
Private Sub mSequence_Click()
    StageControl "SequenceEligible", CBool(mSequence.Value)
End Sub
Private Sub mSave_Click()
    SavePolicy
End Sub
Private Sub mReset_Click()
    ResetPolicy
End Sub
Private Sub mReload_Click()
    ReloadPolicy
End Sub

Private Function AddControl(ByVal page As Object, ByVal kind As String, ByVal name As String, _
                            ByVal caption As String, ByVal x As Single, ByVal y As Single, _
                            ByVal width As Single, ByVal height As Single) As Object
    Set AddControl = modAdminSettingsControls.AddControl(page, kind, name, caption, x, y, width, height)
End Function
