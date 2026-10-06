Attribute VB_Name = "modProductionPrintCodes"
Option Explicit
Option Private Module

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim severity As String, effect As String, message As String, nextStep As String, record As Object
    If id <> "PRODUCTION_RUN_PRINT" Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED"
            effect = "Unknown": message = "Print Recall requested."
        Case "PREVIEW_RETURNED"
            message = "Report prepared and print preview closed."
            nextStep = "Preview return does not confirm physical printing or report freshness."
        Case "REJECTED"
            severity = "Warning": message = "Validation prevented report preparation."
            nextStep = "Review the message in Production before trying again."
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; no new report was prepared."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "Report preparation or preview did not finish; earlier local report changes may remain."
            nextStep = "Inspect the report before trying again; no rollback is implied."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
