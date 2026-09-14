Attribute VB_Name = "modActionRecording"
Option Explicit

' D18 public cross-project boundary: primitives only; Core owns no form.
Public Function Control(ByVal context As String, ByVal command As String, ByRef report As String) As Boolean
    Control = modRecordingSession.ExecuteCommand(context, command, report)
End Function

Public Function Status(ByVal context As String, ByRef active As Boolean, ByRef canStart As Boolean) As String
    Status = modRecordingSession.ReadStatus(context, active, canStart)
End Function

Public Sub ContextClosed(ByVal context As String)
    modRecordingSession.InterruptContext context, "VIEWER_CLOSED"
End Sub
