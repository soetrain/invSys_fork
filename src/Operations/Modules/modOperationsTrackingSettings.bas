Attribute VB_Name = "modOperationsTrackingSettings"
Option Explicit

Private mSettings As frmEventTrackingSettings

Public Function OpenSettings(ByVal context As String, ByRef report As String) As Boolean
    report = "Session or warehouse changed. Reopen Viewer."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    If mSettings Is Nothing Then
        Set mSettings = New frmEventTrackingSettings
        If Not mSettings.BindContext(context) Then CloseSettings: Exit Function
    ElseIf Not mSettings.HasContext(context) Then
        report = "Close Event Tracking Settings, then reopen it for the current session."
        Exit Function
    End If
    If Not mSettings.Visible Then mSettings.Show vbModeless
    report = "Event Tracking Settings opened."
    OpenSettings = True
End Function

Public Sub CloseSettings()
    If Not mSettings Is Nothing Then Unload mSettings
    Set mSettings = Nothing
End Sub

Public Sub ReleaseSettings(ByVal instance As frmEventTrackingSettings)
    If Not mSettings Is Nothing Then
        If mSettings Is instance Then Set mSettings = Nothing
    End If
End Sub
