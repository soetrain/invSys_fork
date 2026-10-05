Attribute VB_Name = "modProductionCompleteCodes"
Option Explicit
Option Private Module

Public Function Control(ByVal id As String) As Object
    If id <> "PRODUCTION_RUN_COMPLETE" Then Exit Function
    Set Control = modProductionControlCatalog.Command(id, "PRODUCTION_RUN_COMPLETION", _
        "Complete Run", "Operations > Production > Production Run - List")
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim severity As String, effect As String, message As String, nextStep As String, record As Object
    If id <> "PRODUCTION_RUN_COMPLETE" Then Exit Function
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = "Production completion requested."
        Case "CONFIRMED"
            message = "The selected Production completion command and local refresh finished."
            nextStep = "Use exact source evidence to verify application; other Processes may remain."
        Case "PENDING"
            severity = "Warning": message = "Production events were submitted; processing or refresh remains unfinished."
            nextStep = "Inspect the submitted events before taking further action."
        Case "REJECTED"
            severity = "Warning": effect = "Unchanged"
            message = "Selection or validation prevented Production submission."
        Case "DENIED"
            severity = "Blocked": effect = "Unchanged"
            message = "Production permission is required; completion owner work was not started."
        Case "FAILED"
            severity = "Error": message = "Production completion failed or was interrupted; partial effects may remain."
            nextStep = "Inspect exact source evidence before retrying; no rollback is implied."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
