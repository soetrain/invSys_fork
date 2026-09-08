Attribute VB_Name = "modReceivingActivityCodes"
Option Explicit
Option Private Module

' Registered facts only; the Receiving owner selects the outcome.
Public Function Outcome(ByVal code As String) As Object
    Dim record As Object, severity As String, effect As String, message As String, nextStep As String
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = "Action requested."
        Case "CONFIRMED"
            message = "Receiving command completed; Domain application not asserted."
            nextStep = "Inspect published outcomes for every related event."
        Case "PENDING"
            severity = "Warning"
            message = "Receiving events were submitted; processing or refresh did not finish."
            nextStep = "Inspect the owning workflow and published outcomes before retrying."
        Case "DENIED"
            severity = "Blocked": effect = "Unchanged"
            message = "The action was not authorized."
            nextStep = "Review the owning form's prerequisites and permissions."
        Case "REJECTED"
            severity = "Warning"
            message = "Receiving validation stopped the command."
            nextStep = "Review the staged entries and the owning form's prerequisites."
        Case "FAILED"
            severity = "Error"
            message = "The action failed. Its final data state requires verification."
            nextStep = "Inspect the owning workflow before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", "RECEIVE_CONFIRM_" & code
    record.Add "OutcomeCode", code
    record.Add "Severity", severity
    record.Add "DataEffect", effect
    record.Add "UserMessage", message
    record.Add "NextStep", nextStep
    Set Outcome = record
End Function
