VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEventTrackingSettings
   Caption         =   "Event Tracking Settings"
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9000
   StartUpPosition =   1
End
Attribute VB_Name = "frmEventTrackingSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@RuntimeStubUserFormCode
Option Explicit

Private WithEvents mChoice As MSForms.ComboBox
Private WithEvents mSave As MSForms.CommandButton
Private WithEvents mReset As MSForms.CommandButton
Private WithEvents mReload As MSForms.CommandButton
Private WithEvents mClose As MSForms.CommandButton
Private mPolicyRows As MSForms.ListBox
Private mPolicyStatus As MSForms.Label
Private mEffective As MSForms.Label
Private mEvidence As MSForms.Label
Private mStatus As MSForms.Label
Private mScope As MSForms.Label
Private mLayout As cOperationsAnchorManager
Private mContext As String
Private mLoading As Boolean
Private mResizeReady As Boolean

Private Sub UserForm_Initialize()
    Dim pages As MSForms.MultiPage, page As Object, choice As Variant
    Me.Caption = "Event Tracking Settings"
    Me.Width = 744: Me.Height = 640
    Set mLayout = modOperationsLayout.OperationsAnchorManager()
    mLayout.ConfigureForForm Me, 744, 640
    Set pages = CreateControl(Me, "MultiPage", "mpOperationsSettings", "", 12, 8, 712, 550, 15)
    pages.Pages(0).Caption = "Event Tracking"
    pages.Pages.Remove 1
    Set page = pages.Pages(0)
    Set mScope = CreateControl(page, "Label", "lblPreferenceScope", "", 12, 12, 670, 24, 7)
    Set mPolicyStatus = CreateControl(page, "Label", "lblReadOnlyPolicy", "Warehouse tracking policy (read only)", 12, 42, 670, 24, 7)
    Set mPolicyRows = CreateControl(page, "ListBox", "lstReadOnlyTrackingPolicy", "", 12, 96, 670, 142, 15)
    mPolicyRows.ColumnCount = 7
    mPolicyRows.ColumnWidths = "0 pt;140 pt;200 pt;45 pt;45 pt;55 pt;160 pt"
    mPolicyRows.Locked = True
    CreatePolicyHeaders page
    CreateControl page, "Label", "lblRequiredCollection", "Required canonical/audit collection stays enabled. Navigation also requires capture and explicit recording.", 12, 248, 670, 30, 13
    CreateControl page, "Label", "lblPreferredView", "Preferred Action Path view", 12, 290, 210, 20, 9
    Set mChoice = CreateControl(page, "ComboBox", "cmbPreferredActionPathView", "", 232, 286, 320, 24, 9)
    mChoice.Style = fmStyleDropDownList
    For Each choice In Split(modActionPathPreference.Choices(), vbLf)
        mChoice.AddItem choice
    Next choice
    Set mEffective = CreateControl(page, "Label", "lblEffectiveView", "", 12, 322, 670, 30, 13)
    Set mEvidence = CreateControl(page, "Label", "lblDiagnosticEvidence", "", 12, 360, 670, 32, 13)
    Set mSave = CreateControl(page, "CommandButton", "btnSaveMyPreference", "Save My Preference", 12, 408, 160, 28, 9)
    Set mReset = CreateControl(page, "CommandButton", "btnResetMyPreference", "Reset to Default", 184, 408, 130, 28, 9)
    Set mReload = CreateControl(page, "CommandButton", "btnReloadMyPreference", "Reload", 326, 408, 86, 28, 9)
    Set mStatus = CreateControl(page, "Label", "lblPreferenceStatus", "", 12, 450, 670, 48, 13)
    Set mClose = CreateControl(Me, "CommandButton", "btnClose", "Close", 654, 574, 66, 28, 12)
End Sub

Private Sub CreatePolicyHeaders(ByVal page As Object)
    Dim widths As Variant, captions As Variant, index As Long, x As Single
    widths = Split(mPolicyRows.ColumnWidths, ";")
    captions = Array("Operation", "Control", "Collect", "Viewer", "Sequence", "Availability")
    x = mPolicyRows.Left + 3
    For index = 1 To 6
        CreateControl page, "Label", "lblPolicyColumn" & CStr(index), CStr(captions(index - 1)), x, 74, CSng(Val(widths(index))), 18, 3
        x = x + CSng(Val(widths(index)))
    Next index
End Sub

Public Function BindContext(ByVal context As String) As Boolean
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    If mContext <> "" And mContext <> context Then Exit Function
    mContext = context
    mScope.Caption = "Your personal preferences for warehouse " & modNasConnection.GetCurrentTargetWarehouseId()
    ReloadSettings
    BindContext = True
End Function

Public Function HasContext(ByVal context As String) As Boolean
    HasContext = (mContext <> "" And mContext = context)
End Function

Public Sub ReloadSettings(Optional ByRef outcome As String = "")
    Dim choice As String, effective As String, evidence As String, report As String, request As String
    Dim version As Long, catalogVersion As Long, readable As Boolean, rows As String, line As Variant, fields As Variant, index As Long, column As Long
    readable = modActionPathPreference.ReadPreference(mContext, choice, effective, evidence, report, version, request, catalogVersion)
    outcome = IIf(readable, "REFRESHED", "FAILED")
    mLoading = True
    mChoice.ListIndex = -1
    If choice <> "" Then mChoice.Value = choice
    mChoice.Enabled = readable: mSave.Enabled = readable: mReset.Enabled = readable
    mLoading = False
    mEffective.Caption = "Effective view: " & effective
    mEvidence.Caption = evidence
    mStatus.Caption = report
    mPolicyRows.Clear
    mPolicyStatus.Caption = "Tracking policy unavailable (read only)."
    If request = "" Then Exit Sub
    mPolicyStatus.Caption = IIf(version = 0, "Built-in tracking defaults", "Tracking policy version " & CStr(version)) & " (read only)."
    rows = modTrackingPolicySettings.ControlRows(request, catalogVersion)
    If rows = "" Then Exit Sub
    For Each line In Split(rows, vbLf)
        fields = Split(CStr(line), vbTab)
        mPolicyRows.AddItem fields(0)
        index = mPolicyRows.ListCount - 1
        For column = 1 To 6
            If column >= 3 And column <= 5 Then
                mPolicyRows.List(index, column) = IIf(fields(column) = "True", "Yes", "No")
            ElseIf column = 6 And version = 0 Then
                mPolicyRows.List(index, column) = "Built-in defaults"
            Else
                mPolicyRows.List(index, column) = fields(column)
            End If
        Next column
    Next line
End Sub

Public Function SaveMyPreference(Optional ByRef outcome As String = "") As Boolean
    Dim report As String
    SaveMyPreference = modActionPathPreference.SavePreference(mContext, CStr(mChoice.Value), report, outcome)
    If SaveMyPreference Then ReloadSettings
    mStatus.Caption = report
End Function

Private Sub mChoice_Change()
    PerformSettings "VIEWER_PATH_PREFERENCE_SELECT"
End Sub
Private Sub mSave_Click()
    PerformSettings "VIEWER_PATH_PREFERENCE_SAVE"
End Sub
Private Sub mReset_Click()
    PerformSettings "VIEWER_PATH_PREFERENCE_RESET"
End Sub
Private Sub mReload_Click()
    PerformSettings "VIEWER_PATH_PREFERENCE_RELOAD"
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
        Case "VIEWER_PATH_PREFERENCE_SELECT"
            If mChoice.Enabled And mChoice.ListIndex >= 0 Then
                mStatus.Caption = "Personal choice staged. Save My Preference to apply."
                outcome = "STAGED"
            End If
        Case "VIEWER_PATH_PREFERENCE_SAVE": SaveMyPreference outcome
        Case "VIEWER_PATH_PREFERENCE_RESET"
            If mChoice.Enabled Then
                mLoading = True: mChoice.ListIndex = 0: mLoading = False
                mStatus.Caption = "Warehouse default staged. Save My Preference to apply."
                outcome = "STAGED"
            End If
        Case "VIEWER_PATH_PREFERENCE_RELOAD": ReloadSettings outcome
    End Select
Done:
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    If notice <> "" Then mStatus.Caption = mStatus.Caption & " " & notice
    Exit Sub
Failed:
    mLoading = False
    If outcome <> "COMPLETED" And outcome <> "UNCHANGED" Then outcome = "FAILED"
    mStatus.Caption = "Settings action or display update failed. Reopen Settings to verify."
    Resume Done
End Sub
Private Sub mClose_Click()
    Unload Me
End Sub
Private Sub UserForm_Terminate()
    modOperationsTrackingSettings.ReleaseSettings Me
    Set mLayout = Nothing
End Sub
Private Sub UserForm_Activate()
    If Not mResizeReady Then
        modUserFormResizeWin.EnableResizableUserForm Me, True, True
        mResizeReady = True
    End If
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
End Sub
Private Sub UserForm_Layout()
    If Not mLayout Is Nothing Then mLayout.ApplyAnchoredLayout
End Sub

Private Function CreateControl(ByVal parent As Object, ByVal kind As String, ByVal name As String, ByVal caption As String, _
                               ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal anchors As Long) As Object
    Dim control As Object
    Set control = parent.Controls.Add("Forms." & kind & ".1", name, True)
    control.Move x, y, width, height
    If kind = "Label" Or kind = "CommandButton" Then control.Caption = caption
    mLayout.RegisterControl control, anchors
    Set CreateControl = control
End Function
