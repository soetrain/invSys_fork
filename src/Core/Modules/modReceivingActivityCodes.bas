Attribute VB_Name = "modReceivingActivityCodes"
Option Explicit
Option Private Module

' Registered facts only; the Receiving owner selects the outcome.
Public Function Outcome(ByVal code As String, Optional ByVal disposition As Boolean = False) As Object
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
    If disposition Then
        record("EventCode") = "DISPOSITION_CONFIRM_" & code
        record("UserMessage") = Replace(message, "Receiving", "Inventory disposition")
    End If
    Set Outcome = record
End Function

Public Function LocalOutcome(ByVal code As String, ByVal clearing As Boolean) As Object
    Dim record As Object, severity As String, effect As String, message As String, nextStep As String, prefix As String
    severity = "Info": effect = "Unknown": prefix = "RECEIVE_REFRESH_"
    If clearing Then prefix = "RECEIVE_CLEAR_"
    Select Case code
        Case "REQUESTED"
            message = "Workbook-local refresh requested."
            If clearing Then message = "Local staging clear requested."
        Case "REFRESHED"
            If clearing Then Exit Function
            effect = "Changed"
            message = "Workbook-local Receiving projections refreshed; no inventory event submitted."
            nextStep = "Review displayed inventory and staged entries."
        Case "CLEARED", "EMPTY"
            If Not clearing Then Exit Function
            effect = "Changed"
            message = "Workbook-local Receiving staging cleared; no inventory event submitted."
            If code = "EMPTY" Then
                effect = "Unchanged": message = "Workbook-local Receiving staging was already empty."
            End If
            nextStep = "Add entries to stage further work."
        Case "FAILED"
            severity = "Error"
            message = "Local refresh failed; its final state requires verification."
            If clearing Then message = "Local staging clear failed; its final state requires verification."
            nextStep = "Inspect the captured Receiving workbook before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", prefix & code
    record.Add "OutcomeCode", code
    record.Add "Severity", severity
    record.Add "DataEffect", effect
    record.Add "UserMessage", message
    record.Add "NextStep", nextStep
    Set LocalOutcome = record
End Function

Public Function StagingOutcome(ByVal code As String, ByVal prefix As String) As Object
    Dim record As Object, severity As String, effect As String, message As String, nextStep As String
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = "Local staging action requested."
        Case "STAGED"
            effect = "Changed"
            message = "Workbook-local staging changed; no inventory event submitted."
            nextStep = "Review staged entries before confirming them."
        Case "REJECTED"
            severity = "Warning": effect = "Unchanged"
            message = "Form validation prevented local staging."
            nextStep = "Review the owning form's required inputs."
        Case "FAILED"
            severity = "Error"
            message = "Local staging failed; its final state requires verification."
            nextStep = "Inspect staged entries before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", prefix & code
    record.Add "OutcomeCode", code
    record.Add "Severity", severity
    record.Add "DataEffect", effect
    record.Add "UserMessage", message
    record.Add "NextStep", nextStep
    Set StagingOutcome = record
End Function
