Attribute VB_Name = "modProductionCloseCodes"
Option Explicit
Option Private Module

' Catalog 21 observes dismissal only, never saved work or Domain application.
Public Function Control(ByVal id As String) As Object
    Dim record As Object
    If id <> "PRODUCTION_CLOSE" Then Exit Function
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "PRODUCTION_WORKFLOW"
    record.Add "Class", "Command": record.Add "Role", "Production"
    record.Add "Caption", "Close": record.Add "Surface", "Operations > Production"
    record.Add "Capability", "PROD_POST": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal code As String) As Object
    Dim record As Object, severity As String, effect As String, message As String, nextStep As String
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = "Production form dismissal requested."
        Case "CLOSED"
            effect = "Unchanged"
            message = "Production form dismissed; no save, posting or Domain application is asserted."
        Case "FAILED"
            severity = "Error": message = "Production form dismissal failed; inspect the form before retrying."
            nextStep = "Inspect whether the Production form is still open before retrying Close."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", "PRODUCTION_CLOSE_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
