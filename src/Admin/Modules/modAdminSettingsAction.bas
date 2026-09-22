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

' Only deliberate form commands enter this wrapper; Core services stay unobserved.
Public Function ChangeUom(ByVal controlId As String, ByVal rawValue As Variant, _
                          ByVal hasSelection As Boolean, ByVal context As String, ByRef report As String) As Boolean
    Dim activityId As String, notice As String, outcome As String
    On Error GoTo Failed
    report = "Session or warehouse changed. Reopen Settings before saving."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    Select Case controlId
        Case "ADMIN_UOM_ADD", "ADMIN_UOM_REMOVE", "ADMIN_UOM_RESET"
        Case Else: report = "UOM action is unavailable.": Exit Function
    End Select
    activityId = modActivity.BeginAction(controlId, context, notice)
    outcome = "REJECTED": report = ""
    If controlId = "ADMIN_UOM_REMOVE" And Not hasSelection Then
        report = "Select a UOM."
        GoTo Done
    End If
    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report) Then
        outcome = "DENIED"
        GoTo Done
    End If
    Select Case controlId
        Case "ADMIN_UOM_ADD"
            ChangeUom = modUomSettings.AddConfiguredUom(CStr(rawValue), report, outcome)
        Case "ADMIN_UOM_REMOVE"
            ChangeUom = modUomSettings.RemoveConfiguredUom(CStr(rawValue), report, outcome)
        Case "ADMIN_UOM_RESET"
            If MsgBox("Reset the warehouse UOM catalog to defaults?", vbQuestion + vbYesNo, "invSys Settings") <> vbYes Then
                outcome = "CANCELLED": GoTo Done
            End If
            If context <> modActivity.CaptureContext() Then
                report = "Session or warehouse changed. Reopen Settings before saving."
                GoTo Done
            End If
            ChangeUom = modUomSettings.ResetConfiguredUoms(report, outcome)
    End Select
Done:
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    If notice <> "" Then report = report & " " & notice
    Exit Function
Failed:
    ChangeUom = False: outcome = "FAILED"
    report = "UOM catalog change failed. Reopen Settings to verify the current values."
    Resume Done
End Function
