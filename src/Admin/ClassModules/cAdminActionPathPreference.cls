VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAdminActionPathPreference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mChoice As MSForms.ComboBox
Private WithEvents mSave As MSForms.CommandButton
Private WithEvents mReset As MSForms.CommandButton
Private WithEvents mReload As MSForms.CommandButton
Private mEffective As MSForms.Label
Private mEvidence As MSForms.Label
Private mStatus As MSForms.Label
Private mContext As String
Private mLoading As Boolean

Public Sub Initialize(ByVal page As Object, ByVal context As String)
    Dim choice As Variant
    mContext = context
    modAdminSettingsControls.AddControl page, "Label", "lblActionPathPreference", "Preferred Action Path view", 12, 48, 220, 20
    Set mChoice = modAdminSettingsControls.AddControl(page, "ComboBox", "cmbPreferredActionPathView", "", 12, 74, 300, 24)
    mChoice.Style = fmStyleDropDownList
    For Each choice In Split(modActionPathPreference.Choices(), vbLf)
        mChoice.AddItem choice
    Next choice
    Set mEffective = modAdminSettingsControls.AddControl(page, "Label", "lblActionPathEffectiveView", "", 12, 116, 650, 40)
    Set mEvidence = modAdminSettingsControls.AddControl(page, "Label", "lblActionPathEvidence", "", 12, 168, 650, 48)
    modAdminSettingsControls.AddControl page, "Label", "lblActionPathPreferenceScope", _
        "Saved for your Windows account, invSys sign-in and this warehouse. This choice does not change tracking policy.", 12, 228, 650, 48
    Set mSave = modAdminSettingsControls.AddControl(page, "CommandButton", "btnSaveActionPathPreference", "Save My Preference", 12, 294, 160, 28)
    Set mReset = modAdminSettingsControls.AddControl(page, "CommandButton", "btnResetActionPathPreference", "Reset to Default", 184, 294, 130, 28)
    Set mReload = modAdminSettingsControls.AddControl(page, "CommandButton", "btnReloadActionPathPreference", "Reload", 326, 294, 86, 28)
    Set mStatus = modAdminSettingsControls.AddControl(page, "Label", "lblActionPathPreferenceStatus", "", 12, 346, 650, 60)
    ReloadPreference
End Sub

Public Sub ReloadPreference(Optional ByRef outcome As String = "")
    Dim choice As String, effective As String, evidence As String, report As String, readable As Boolean
    readable = modActionPathPreference.ReadPreference(mContext, choice, effective, evidence, report)
    outcome = IIf(readable, "REFRESHED", "FAILED")
    mLoading = True
    mChoice.ListIndex = -1
    If choice <> "" Then mChoice.Value = choice
    mLoading = False
    mChoice.Enabled = readable: mSave.Enabled = readable: mReset.Enabled = readable
    mEffective.Caption = "Effective view: " & effective
    mEvidence.Caption = evidence
    mStatus.Caption = report
End Sub

Public Sub ResetPreference(Optional ByRef outcome As String = "")
    outcome = "REJECTED"
    If Not mChoice.Enabled Then Exit Sub
    mLoading = True
    mChoice.ListIndex = 0
    mLoading = False
    mStatus.Caption = "Warehouse default staged. Save My Preference to apply."
    outcome = "STAGED"
End Sub

Public Function SavePreference(Optional ByRef outcome As String = "") As Boolean
    Dim report As String
    SavePreference = modActionPathPreference.SavePreference(mContext, CStr(mChoice.Value), report, outcome)
    If SavePreference Then ReloadPreference
    mStatus.Caption = report
End Function

Private Sub mChoice_Change()
    PerformSettings "ADMIN_PATH_PREFERENCE_SELECT"
End Sub
Private Sub mSave_Click()
    PerformSettings "ADMIN_PATH_PREFERENCE_SAVE"
End Sub
Private Sub mReset_Click()
    PerformSettings "ADMIN_PATH_PREFERENCE_RESET"
End Sub
Private Sub mReload_Click()
    PerformSettings "ADMIN_PATH_PREFERENCE_RELOAD"
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
        Case "ADMIN_PATH_PREFERENCE_SELECT"
            If mChoice.Enabled And mChoice.ListIndex >= 0 Then
                mStatus.Caption = "Personal choice staged. Save My Preference to apply."
                outcome = "STAGED"
            End If
        Case "ADMIN_PATH_PREFERENCE_SAVE": SavePreference outcome
        Case "ADMIN_PATH_PREFERENCE_RESET": ResetPreference outcome
        Case "ADMIN_PATH_PREFERENCE_RELOAD": ReloadPreference outcome
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
