Attribute VB_Name = "modAdminUsersAction"
Option Explicit
Option Private Module

' Admin-local transient input only. Never serialized or passed across projects.
Public Type UserWriteRequest
    Command As String
    WriteContext As String
    ActivityContext As String
    AuthPath As String
    UserId As String
    DisplayName As String
    Credential As String
    Warehouse As String
    Station As String
    CapabilityMask As Long
    GuideChoice As Long
End Type

Private Const TRACKING_NOTICE As String = "Tracking unavailable: the Users & Roles action could not be fully recorded."

Public Function Execute(ByRef request As UserWriteRequest, ByRef report As String, ByRef notice As String) As Boolean
    Dim activity As String, outcome As String, surface As Object, saved As Boolean
    On Error GoTo Failed
    outcome = "FAILED": report = "": notice = ""
    Set surface = modAdminActionObservation.CurrentSurface()
    activity = BeginObservation(request.Command, request.ActivityContext, notice)
    modAdminActionObservation.RestoreSurface surface
    If request.Command <> "SAVE" And Trim$(request.UserId) = "" Then
        outcome = "REJECTED": report = "Select or enter a user first.": GoTo CleanExit
    End If
    Select Case request.Command
        Case "SAVE"
            saved = modAuthUserCommands.SaveUser(request.WriteContext, request.AuthPath, request.UserId, _
                request.DisplayName, request.Credential, request.Warehouse, request.Station, request.CapabilityMask, report, outcome, request.GuideChoice)
            If saved Then report = "User and role assignments saved. Store the visible PIN/password now; it cannot be recovered later."
        Case "DEACTIVATE"
            saved = modAuthUserCommands.DeactivateUser(request.WriteContext, request.AuthPath, request.UserId, report, outcome)
            If saved Then report = "User status set to INACTIVE."
        Case "DELETE"
            If MsgBox("Delete user '" & Trim$(request.UserId) & "' and all assigned capabilities?", vbQuestion + vbYesNo, "Users & Roles") <> vbYes Then
                outcome = "CANCELLED": GoTo CleanExit
            End If
            saved = modAuthUserCommands.DeleteUser(request.WriteContext, request.AuthPath, request.UserId, report, outcome)
            If saved Then report = "User deleted."
        Case Else
            outcome = "REJECTED": report = "Unknown Users & Roles command."
    End Select
    Execute = saved
CleanExit:
    Set surface = modAdminActionObservation.CurrentSurface()
    modAdminActionObservation.FinishObservation activity, outcome, notice, TRACKING_NOTICE
    modAdminActionObservation.RestoreSurface surface
    If notice <> "" Then notice = TRACKING_NOTICE
    Exit Function
Failed:
    outcome = "FAILED": Execute = False
    report = "Users & Roles could not finish. Inspect the current Auth state before retrying."
    Resume CleanExit
End Function

Private Function BeginObservation(ByVal command As String, ByVal context As String, ByRef notice As String) As String
    Dim current As String
    On Error GoTo Unavailable
    BeginObservation = modActivity.BeginAction("ADMIN_USER_" & command, context, notice)
    current = modActivity.CaptureContext()
    If context <> "" And context <> current Then notice = TRACKING_NOTICE
    Exit Function
Unavailable:
    notice = TRACKING_NOTICE
End Function

Public Sub ShowNotice(ByVal notice As String)
    On Error Resume Next
    If notice <> "" Then MsgBox TRACKING_NOTICE, vbExclamation, "Users & Roles"
End Sub
