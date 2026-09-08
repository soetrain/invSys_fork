Attribute VB_Name = "modAdminSettingsAction"
Option Explicit

Public Function SaveValue(ByVal key As String, ByVal value As Variant, ByVal warehouse As String, _
                          ByVal station As String, ByVal context As String, ByRef report As String) As Boolean
    Dim activityId As String, notice As String, outcome As String
    On Error GoTo Failed
    report = "Session or warehouse changed. Reopen Settings before saving."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    activityId = modActivity.BeginAction("ADMIN_SETTINGS_SAVE_VALUE", context, notice)
    outcome = "REJECTED"
    If key = "" Then
        report = "Select a config key first."
    ElseIf Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report) Then
        outcome = "DENIED"
    Else
        SaveValue = modConfigCommands.UpdateConfigValue(key, value, report, warehouse, station, outcome)
    End If
Done:
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    If notice <> "" Then report = report & " " & notice
    Exit Function
Failed:
    SaveValue = False
    outcome = "FAILED"
    report = "Configuration save failed. Reopen Settings to verify the current values."
    Resume Done
End Function
