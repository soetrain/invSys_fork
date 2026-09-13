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

Public Sub ReloadPreference()
    Dim choice As String, effective As String, evidence As String, report As String, readable As Boolean
    readable = modActionPathPreference.ReadPreference(mContext, choice, effective, evidence, report)
    mLoading = True
    mChoice.ListIndex = -1
    If choice <> "" Then mChoice.Value = choice
    mLoading = False
    mChoice.Enabled = readable: mSave.Enabled = readable: mReset.Enabled = readable
    mEffective.Caption = "Effective view: " & effective
    mEvidence.Caption = evidence
    mStatus.Caption = report
End Sub

Public Sub ResetPreference()
    If Not mChoice.Enabled Then Exit Sub
    mChoice.ListIndex = 0
    mStatus.Caption = "Warehouse default staged. Save My Preference to apply."
End Sub

Public Function SavePreference() As Boolean
    Dim report As String
    SavePreference = modActionPathPreference.SavePreference(mContext, CStr(mChoice.Value), report)
    If SavePreference Then ReloadPreference
    mStatus.Caption = report
End Function

Private Sub mChoice_Change()
    If Not mLoading Then mStatus.Caption = "Personal choice staged. Save My Preference to apply."
End Sub
Private Sub mSave_Click()
    SavePreference
End Sub
Private Sub mReset_Click()
    ResetPreference
End Sub
Private Sub mReload_Click()
    ReloadPreference
End Sub
