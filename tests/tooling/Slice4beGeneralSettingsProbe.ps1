# Invoke existing General Settings callbacks in disposable compiled packages.
function Install-GeneralSettingsProbe {
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $admin.VBComponents.Item('frmAdminSettings').CodeModule.AddFromString(@'
Public Function GeneralSettingsActionForTest(ByVal action As String, ByVal value As String) As String
    Dim rows As MSForms.ListBox, i As Long
    Select Case action
        Case "SelectConfig": Set rows = mLstConfig
        Case "SelectCarrier", "CarrierRemove": Set rows = mLstCarriers
        Case "SelectUom": Set rows = mLstUoms
    End Select
    If Not rows Is Nothing Then
        mLoading = True: rows.ListIndex = -1
        For i = 0 To rows.ListCount - 1
            If CStr(rows.List(i, 0)) = value Then rows.ListIndex = i: Exit For
        Next i
        mLoading = False
    End If
    Select Case action
        Case "SelectConfig": mLstConfig_Click
        Case "SelectCarrier": mLstCarriers_Click
        Case "SelectUom": mLstUoms_Click
        Case "Reload": mBtnReloadConfig_Click
        Case "CarrierAdd": mTxtCarrier.Value = value: mBtnAdd_Click
        Case "CarrierRemove": mBtnRemove_Click
        Case "CarrierReset": mBtnReset_Click
        Case "ConnectionChoice": mChkManualServerCredentials.Value = CBool(value)
        Case "ConnectionSave": mBtnSaveConnectionPolicy_Click
        Case "Render": LoadConfigRows: LoadConnectionPolicy: LoadCarriers: LoadUoms
        Case "ConfigKey": GeneralSettingsActionForTest = CStr(mTxtConfigKey.Value): Exit Function
        Case "CarrierDraft": GeneralSettingsActionForTest = CStr(mTxtCarrier.Value): Exit Function
        Case "UomDraft": GeneralSettingsActionForTest = CStr(mTxtUom.Value): Exit Function
        Case "ConnectionDraft": GeneralSettingsActionForTest = CStr(CBool(mChkManualServerCredentials.Value)): Exit Function
        Case "Rows": GeneralSettingsActionForTest = CStr(mLstConfig.ListCount): Exit Function
        Case Else: Err.Raise 5, , "Unknown General Settings fixture action."
    End Select
    GeneralSettingsActionForTest = mLblStatus.Caption
End Function
'@)
    $admin.VBComponents.Item('TestD5Commands').CodeModule.AddFromString(@'
Public Function GeneralSettingsAction(ByVal action As String, ByVal value As String) As String
    GeneralSettingsAction = mForm.GeneralSettingsActionForTest(action, value)
End Function
'@)
    $core=$packages['invSys.Core.xlam'].VBProject.VBComponents.Add(1)
    $core.Name='TestGeneralSettings'
    $core.CodeModule.AddFromString(@'
Option Explicit
Public Function Definitions(ByVal version As Long) As String
    Dim ids As Variant, id As Variant, definition As Object
    ids = modActivityCatalog.ControlIds(version)
    If Not IsArray(ids) Then Exit Function
    For Each id In ids
        Set definition = modActivityCatalog.Control(CStr(id), version)
        If definition Is Nothing Then Err.Raise 5, , "Registered definition missing."
        Definitions = Definitions & modTrainingJson.EncodeObject(definition) & vbLf
    Next id
End Function
'@)
}
