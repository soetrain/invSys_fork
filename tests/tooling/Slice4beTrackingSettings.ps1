# D18/4be.2 surface discovery through the existing packaged form's Initialize
# and Show path. Inspect the live controls, never VBA source or invented seams.
# Only fixed booleans leave Excel; no Config values, paths or actors are emitted.
function Install-Slice4beTrackingSettingsProbe($TestModule) {
    $TestModule.CodeModule.AddFromString(@'
Public Function TrackingSettingsSurface() As String
    Dim general As Object, tracking As Object, capture As Object, views As Object
    mForm.Show vbModeless
    mForm.Repaint
    Set general = TrackingPage(mForm, "General")
    Set tracking = TrackingPage(mForm, "Event Tracking")
    TrackingSettingsSurface = CStr(mForm.Visible) & "|" & _
        CStr(Not TrackingControl(mForm, "ListBox", "", "lstConfig") Is Nothing And _
             Not TrackingControl(mForm, "ListBox", "", "lstCarriers") Is Nothing And _
             Not TrackingControl(mForm, "ListBox", "", "lstUoms") Is Nothing) & "|" & _
        CStr(Not general Is Nothing And Not tracking Is Nothing)
    TrackingSettingsSurface = TrackingSettingsSurface & "|" & _
        CStr(TrackingHasCaption(tracking, "Tracking") And _
             TrackingHasCaption(tracking, "Event Detail") And _
             TrackingHasCaption(tracking, "Action Paths")) & "|" & _
        CStr(TrackingButton(tracking, "Save Tracking Policy") And _
             TrackingButton(tracking, "Save Detail Profile") And _
             TrackingButton(tracking, "Save My Preference") And _
             TrackingButton(tracking, "Reload") And _
             TrackingButton(tracking, "Reset to Default"))
    Set capture = TrackingControl(tracking, "CheckBox", "Capture recorded controls")
    TrackingSettingsSurface = TrackingSettingsSurface & "|" & CStr(TrackingUnchecked(capture))
    Set views = TrackingControl(tracking, "ComboBox", "", "cmbPreferredActionPathView")
    TrackingSettingsSurface = TrackingSettingsSurface & "|" & CStr(TrackingViewChoices(views))
End Function
Private Function TrackingPage(ByVal parent As Object, ByVal caption As String) As Object
    Dim item As Object, page As Object, found As Object
    If parent Is Nothing Then Exit Function
    For Each item In parent.Controls
        If TypeName(item) = "MultiPage" Then
            For Each page In item.Pages
                If page.Caption = caption Then Set TrackingPage = page: Exit Function
            Next page
        ElseIf TypeName(item) = "Frame" Then
            Set found = TrackingPage(item, caption)
            If Not found Is Nothing Then Set TrackingPage = found: Exit Function
        End If
    Next item
End Function
Private Function TrackingControl(ByVal parent As Object, ByVal kind As String, _
                                 Optional ByVal caption As String = "", _
                                 Optional ByVal name As String = "") As Object
    Dim item As Object, page As Object, found As Object, matches As Boolean
    If parent Is Nothing Then Exit Function
    For Each item In parent.Controls
        If TypeName(item) = kind Then
            matches = (name = "" Or item.Name = name)
            If caption <> "" Then matches = matches And (item.Caption = caption)
            If matches Then Set TrackingControl = item: Exit Function
        End If
        If TypeName(item) = "MultiPage" Then
            For Each page In item.Pages
                Set found = TrackingControl(page, kind, caption, name)
                If Not found Is Nothing Then Set TrackingControl = found: Exit Function
            Next page
        ElseIf TypeName(item) = "Frame" Then
            Set found = TrackingControl(item, kind, caption, name)
            If Not found Is Nothing Then Set TrackingControl = found: Exit Function
        End If
    Next item
End Function
Private Function TrackingHasCaption(ByVal parent As Object, ByVal caption As String) As Boolean
    TrackingHasCaption = Not TrackingControl(parent, "Label", caption) Is Nothing Or _
                         Not TrackingControl(parent, "Frame", caption) Is Nothing
End Function
Private Function TrackingButton(ByVal parent As Object, ByVal caption As String) As Boolean
    TrackingButton = Not TrackingControl(parent, "CommandButton", caption) Is Nothing
End Function
Private Function TrackingUnchecked(ByVal control As Object) As Boolean
    If control Is Nothing Then Exit Function
    If IsNull(control.Value) Then Exit Function
    TrackingUnchecked = (control.Value = False)
End Function
Private Function TrackingViewChoices(ByVal control As Object) As Boolean
    Dim index As Long, choices As String
    If control Is Nothing Then Exit Function
    For index = 0 To control.ListCount - 1
        choices = choices & "|" & CStr(control.List(index, 0))
    Next index
    TrackingViewChoices = (choices = "|Use warehouse default|How-To|Diagnostic|Compare both")
End Function
'@)
}
function Test-Slice4beTrackingSettingsSurface($Fixture) {
    $before=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    $observed=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.TrackingSettingsSurface')
    $flags=$observed.Split('|')
    if ($flags.Count -ne 7 -or @($flags | Where-Object { $_ -cnotin @('True','False') }).Count) {
        throw 'Settings surface observation did not return seven Boolean facts.'
    }
    Check 'TrackingSettings.FormShown' ($flags[0] -ceq 'True')
    Check 'TrackingSettings.ExistingGeneralEditorsPresent' ($flags[1] -ceq 'True')
    Check 'TrackingSettings.GeneralAndEventTrackingTabs' ($flags[2] -ceq 'True')
    Check 'TrackingSettings.TrackingDetailAndActionPathSections' ($flags[3] -ceq 'True')
    Check 'TrackingSettings.SeparateSaveReloadResetActions' ($flags[4] -ceq 'True')
    Check 'TrackingSettings.CaptureDefaultsOff' ($flags[5] -ceq 'True')
    Check 'TrackingSettings.PersonalViewChoices' ($flags[6] -ceq 'True')
    Check 'TrackingSettings.OpenDoesNotWriteConfig' ($before -ceq (Get-FileHash -LiteralPath $Fixture.Config).Hash)
    if ($CaptureEvidence) { CaptureFormEvidence 'invSys Settings' 'tracking-settings-open.png' }
}
