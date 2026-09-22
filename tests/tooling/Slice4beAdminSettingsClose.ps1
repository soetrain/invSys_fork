# D18 Close-discard through the real modal Admin launcher/default instance.
function Install-AdminSettingsDefaultCloseProbe($TestModule,$FormCode) {
    $FormCode.AddFromString(@'
Public Function DefaultCloseStateForTest(ByVal phase As Long) As String
    If phase = 1 Then
        PreferenceTestChoose "Compare both"
        PolicyTestCapture True
        DefaultCloseStateForTest = CStr(PreferenceTestChoice() = "Compare both") & "|" & _
            CStr(modTrackingPolicySettings.FieldValue(PolicyTestRequest(), "ViewerActionPathCaptureEnabled") = "True")
    Else
        DefaultCloseStateForTest = CStr(PreferenceTestChoice() = "Use warehouse default") & "|" & _
            CStr(modTrackingPolicySettings.FieldValue(PolicyTestRequest(), "ViewerActionPathCaptureEnabled") = "False")
    End If
End Function
'@)
    $TestModule.CodeModule.AddFromString(@'
Private mDefaultClosePhase As Long
Private mDefaultCloseResult As String
Private mDefaultCloseEntries As Long
Public Sub ArmDefaultClose(ByVal phase As Long)
    mDefaultClosePhase = phase
    mDefaultCloseResult = ""
End Sub
Public Function DriveDefaultClose(ByVal instance As frmAdminSettings) As Boolean
    Dim phase As Long
    phase = mDefaultClosePhase
    If phase = 0 Then Exit Function
    mDefaultClosePhase = 0
    mDefaultCloseEntries = mDefaultCloseEntries + 1
    mDefaultCloseResult = instance.DefaultCloseStateForTest(phase)
    instance.PolicyTestCloseAction
    DriveDefaultClose = True
End Function
Public Function DefaultCloseResult() As String
    DefaultCloseResult = mDefaultCloseResult
End Function
Public Function DefaultCloseEntries() As Long
    DefaultCloseEntries = mDefaultCloseEntries
End Function
Public Function RemainingSettingsInstances() As Long
    Dim instance As Object
    For Each instance In VBA.UserForms
        If TypeName(instance) = "frmAdminSettings" Then RemainingSettingsInstances = RemainingSettingsInstances + 1
    Next instance
End Function
Public Sub CleanupDefaultSettings()
    Unload frmAdminSettings
End Sub
'@)
    $line = $FormCode.ProcBodyLine('UserForm_Activate',0)
    $FormCode.InsertLines($line+1,'    If TestD5Commands.DriveDefaultClose(Me) Then Exit Sub')
}

function Test-AdminSettingsDefaultClose($Fixture) {
    $operator = $excel.Workbooks.Add()
    try {
        $operator.SaveAs((Join-Path $runRoot 'settings-operator.xlsm'),52)
        $operator.Activate()
        $before = (Get-FileHash -LiteralPath $Fixture.Config).Hash
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ArmDefaultClose' @(1))
        [void](Run 'invSys.Admin.xlam' 'modAdmin.Open_Settings')
        Check 'AdminClose.RealLauncherStagesAndCloses' ([string](Run 'invSys.Admin.xlam' 'TestD5Commands.DefaultCloseResult') -ceq 'True|True')
        Check 'AdminClose.NoHiddenDefaultInstance' ([int](Run 'invSys.Admin.xlam' 'TestD5Commands.RemainingSettingsInstances') -eq 0)
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.ArmDefaultClose' @(2))
        [void](Run 'invSys.Admin.xlam' 'modAdmin.Open_Settings')
        $flags = ([string](Run 'invSys.Admin.xlam' 'TestD5Commands.DefaultCloseResult')).Split('|')
        if($flags.Count -ne 2) { throw 'The real Settings activation was not observed.' }
        Check 'AdminClose.RealReopenDiscardsPersonalStaging' ($flags[0] -ceq 'True')
        Check 'AdminClose.RealReopenDiscardsPolicyStaging' ($flags[1] -ceq 'True')
        Check 'AdminClose.BothRealActivationsObserved' ([int](Run 'invSys.Admin.xlam' 'TestD5Commands.DefaultCloseEntries') -eq 2)
        Check 'AdminClose.ConfigBytesUnchanged' ($before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    } finally {
        [void](Run 'invSys.Admin.xlam' 'TestD5Commands.CleanupDefaultSettings')
        $operator.Close($false)
    }
}
