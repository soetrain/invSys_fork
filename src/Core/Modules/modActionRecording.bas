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

' EXPECTATION1 projection: header marker/draft/sequence/terminal StepId/kind;
' subsequent rows are StepId/control/outcome/retry/caption, tab/CRLF separated.
' These drafts are authored intent only, never activity or workflow commands.
Public Function OpenExpectation(ByVal context As String, ByRef projection As String, ByRef notice As String) As Boolean
    OpenExpectation = modExpectationDraft.OpenRecording(context, projection, notice)
End Function

Public Function ExpectationChoices(ByVal context As String, ByVal draftId As String, ByVal controlId As String) As String
    ExpectationChoices = modExpectationDraft.Choices(context, draftId, controlId)
End Function

Public Function EditExpectation(ByVal context As String, ByVal draftId As String, ByVal command As String, _
                                 ByVal stepId As String, ByVal controlId As String, ByVal outcome As String, _
                                 ByVal retry As Boolean, ByRef projection As String, ByRef notice As String) As Boolean
    EditExpectation = modExpectationDraft.Edit(context, draftId, command, stepId, controlId, outcome, retry, projection, notice)
End Function

Public Function UseExpectation(ByVal context As String, ByVal draftId As String, ByVal terminalStepId As String, _
                                ByVal terminalKind As String, ByRef notice As String) As Boolean
    UseExpectation = modExpectationDraft.UseRecording(context, draftId, terminalStepId, terminalKind, notice)
End Function

Public Sub CloseExpectation(ByVal context As String, ByVal draftId As String)
    modExpectationDraft.CloseContext context, draftId
End Sub
