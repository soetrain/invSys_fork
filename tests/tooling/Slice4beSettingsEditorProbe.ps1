# D13 Settings probes: call each actual private handler exactly once.
# Install into disposable packages before compilation and before forms exist.
function Install-SettingsActivityProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    . (Join-Path $PSScriptRoot 'Slice4beTrackingPolicy.ps1')
    Install-Slice4beTrackingPolicyProbe $admin.VBComponents.Item('TestD5Commands') $admin.VBComponents.Item('frmAdminSettings').CodeModule
    $admin.VBComponents.Item('cAdminTrackingPolicy').CodeModule.AddFromString(@'
Public Function SettingsActivityAction(ByVal action As String, ByVal value As String) As String
    Dim index As Long
    mLoading = True
    Select Case action
        Case "SelectControl"
            mRows.ListIndex = -1
            For index = 0 To mRows.ListCount - 1
                If CStr(mRows.List(index, 0)) = value Then mRows.ListIndex = index: Exit For
            Next index
        Case "Capture": mCapture.Value = CBool(value)
        Case "AdminVisible": mAdminVisible.Value = CBool(value)
        Case "DefaultView": mView.Value = value
        Case "Collect": mCollect.Value = CBool(value)
        Case "Visible": mVisible.Value = CBool(value)
        Case "Sequence": mSequence.Value = CBool(value)
    End Select
    mLoading = False
    Select Case action
        Case "SelectControl": mRows_Click
        Case "Capture": mCapture_Click
        Case "AdminVisible": mAdminVisible_Click
        Case "DefaultView": mView_Change
        Case "Collect": mCollect_Click
        Case "Visible": mVisible_Click
        Case "Sequence": mSequence_Click
        Case "Save": mSave_Click
        Case "Reset": mReset_Click
        Case "Reload": mReload_Click
        Case "Request": SettingsActivityAction = mRequest: Exit Function
        Case "Version": SettingsActivityAction = CStr(mVersion): Exit Function
        Case "StaleVersionFixture": mVersion = mVersion - 1: Exit Function
        Case "Selected": SettingsActivityAction = CStr(mRows.List(mRows.ListIndex, 0)): Exit Function
        Case "RenderOnly": Render
        Case Else: Err.Raise 5, , "Unknown Settings fixture action"
    End Select
    SettingsActivityAction = mStatus.Caption
End Function
'@)
    $admin.VBComponents.Item('cAdminEventDetail').CodeModule.AddFromString(@'
Public Function SettingsActivityAction(ByVal action As String, ByVal value As String) As String
    Dim index As Long
    mLoading = True
    Select Case action
        Case "SelectFamily": mFamily.Value = value
        Case "SelectField"
            mRows.ListIndex = -1
            For index = 0 To mRows.ListCount - 1
                If CStr(mRows.List(index, 0)) = value Then mRows.ListIndex = index: Exit For
            Next index
        Case "ShowField": mEnabled.Value = CBool(value)
    End Select
    mLoading = False
    Select Case action
        Case "SelectFamily": mFamily_Change
        Case "SelectField": mRows_Change
        Case "ShowField": mEnabled_Click
        Case "MoveUp": mUp_Click
        Case "MoveDown": mDown_Click
        Case "Save": mSave_Click
        Case "Reset": mReset_Click
        Case "Reload": mReload_Click
        Case "Request": SettingsActivityAction = mRequest: Exit Function
        Case "Version": SettingsActivityAction = CStr(mVersion): Exit Function
        Case "StaleVersionFixture": mVersion = mVersion - 1: Exit Function
        Case "Selected": SettingsActivityAction = CStr(mRows.List(mRows.ListIndex, 0)): Exit Function
        Case "Family": SettingsActivityAction = CStr(mFamily.Value): Exit Function
        Case "RenderOnly": Render
        Case Else: Err.Raise 5, , "Unknown Settings fixture action"
    End Select
    SettingsActivityAction = mStatus.Caption
End Function
'@)
    $preference=@'
Public Function SettingsActivityAction(ByVal action As String, ByVal value As String) As String
    If action = "Select" Then
        mLoading = True: mChoice.Value = value: mLoading = False
    End If
    Select Case action
        Case "Select": mChoice_Change
        Case "Save": mSave_Click
        Case "Reset": mReset_Click
        Case "Reload": mReload_Click
        Case "Choice": SettingsActivityAction = CStr(mChoice.Value): Exit Function
        Case Else: Err.Raise 5, , "Unknown Settings fixture action"
    End Select
    SettingsActivityAction = mStatus.Caption
End Function
'@
    $admin.VBComponents.Item('cAdminActionPathPreference').CodeModule.AddFromString($preference)
    $admin.VBComponents.Item('frmAdminSettings').CodeModule.AddFromString(@'
Public Function SettingsActivityAction(ByVal section As String, ByVal action As String, ByVal value As String) As String
    Select Case section
        Case "Tracking": SettingsActivityAction = mTracking.SettingsActivityAction(action, value)
        Case "Detail": SettingsActivityAction = mDetail.SettingsActivityAction(action, value)
        Case "Preference": SettingsActivityAction = mPreference.SettingsActivityAction(action, value)
        Case Else: Err.Raise 5, , "Unknown Settings fixture section"
    End Select
End Function
Public Sub SettingsActivitySection(ByVal section As String)
    mPages.Value = 1
    Select Case section
        Case "Tracking": mTrackingSections.Value = 0
        Case "Detail": mTrackingSections.Value = 1
        Case "Preference": mTrackingSections.Value = 2
        Case Else: Err.Raise 5, , "Unknown Settings fixture section"
    End Select
    Me.Show vbModeless
    Me.Repaint
End Sub
'@)
    $admin.VBComponents.Item('TestD5Commands').CodeModule.AddFromString(@'
Public Function SettingsActivityAction(ByVal section As String, ByVal action As String, ByVal value As String) As String
    SettingsActivityAction = mForm.SettingsActivityAction(section, action, value)
End Function
Public Sub SettingsActivitySection(ByVal section As String)
    mForm.SettingsActivitySection section
End Sub
'@)
    $ops=$packages['invSys.Operations.xlam'].VBProject
    $ops.VBComponents.Item('frmEventTrackingSettings').CodeModule.AddFromString($preference)
    $ops.VBComponents.Item('modOperationsTrackingSettings').CodeModule.AddFromString(@'
Public Function SettingsActivityAction(ByVal action As String, ByVal value As String) As String
    If mSettings Is Nothing Then Err.Raise 5, , "Actual Operations Settings is unavailable"
    SettingsActivityAction = mSettings.SettingsActivityAction(action, value)
End Function
'@)
    $ops.VBComponents.Item('modInventoryViewer').CodeModule.AddFromString(@'
Public Function SettingsActivityOpen() As Boolean
    Dim control As Object
    If mInventoryViewer Is Nothing Then Exit Function
    For Each control In mInventoryViewer.Controls
        If control.Name = "btnSettings" Then
            If Not control.Visible Or Not control.Enabled Then Exit Function
            control.Value = True
            SettingsActivityOpen = True
            Exit Function
        End If
    Next control
End Function
'@)
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $core.Name='TestSettingsOwnerState'
    $core.CodeModule.AddFromString(@'
Public Function CatalogIds(ByVal version As Long) As String
    Dim ids As Variant
    ids = modActivityCatalog.ControlIds(version)
    If IsArray(ids) Then CatalogIds = Join(ids, vbLf)
End Function
Public Function State() As String
    Dim context As String, version As Long, request As String, report As String
    Dim choice As String, effective As String, evidence As String, result As Object
    context = modActivity.CaptureContext()
    Set result = CreateObject("Scripting.Dictionary")
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Err.Raise 5, , "Policy fixture read unavailable"
    result.Add "PolicyVersion", version: result.Add "PolicyRequest", request
    If Not modEventDetailSettings.ReadEditor(context, version, request, report) Then Err.Raise 5, , "Profile fixture read unavailable"
    result.Add "ProfileVersion", version: result.Add "ProfileRequest", request
    If Not modActionPathPreference.ReadPreference(context, choice, effective, evidence, report) Then Err.Raise 5, , "Preference fixture read unavailable"
    result.Add "Preference", choice
    State = modTrainingJson.EncodeObject(result)
End Function
Public Function SetupRecordingPolicy() As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrackingPolicySettings.StageValue(modTrackingPolicySettings.DefaultRequest(), "ViewerActionPathCaptureEnabled", "True")
    SetupRecordingPolicy = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function SetupPreference(ByVal choice As String) As Boolean
    Dim report As String
    SetupPreference = modActionPathPreference.SavePreference(modActivity.CaptureContext(), choice, report)
End Function
Public Function PersonalChoice() As String
    Dim choice As String, effective As String, evidence As String, report As String
    If Not modActionPathPreference.ReadPreference(modActivity.CaptureContext(), choice, effective, evidence, report) Then Err.Raise 5, , "Personal fixture read unavailable"
    PersonalChoice = choice
End Function
Public Function ProfileReadable() As Boolean
    Dim version As Long, request As String, report As String
    ProfileReadable = modEventDetailSettings.ReadEditor(modActivity.CaptureContext(), version, request, report)
End Function
Public Function SetupProfile() As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modEventDetailSettings.ReadEditor(context, version, request, report) Then Exit Function
    SetupProfile = modEventDetailSettings.SaveProfile(context, version, request, report)
End Function
'@)
}
