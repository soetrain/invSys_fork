Attribute VB_Name = "modReceivingNavigationCodes"
Option Explicit
Option Private Module

' Catalog 6: fixed control labels only. Selected values never enter activity.
Public Function Control(ByVal controlId As String) As Object
    Dim caption As String, surface As String, record As Object
    surface = "Operations > Receiving"
    Select Case controlId
        Case "RECEIVING_PAGE_RECEIPTS": caption = "Receiving"
        Case "RECEIVING_PAGE_RETURNS": caption = "Returns"
        Case "RECEIVING_PAGE_PURCHASING": caption = "Purchasing"
        Case "RECEIVING_SELECT_ITEM": caption = "Receive Item Results"
        Case "DISPOSITION_SELECT_ITEM": caption = "Return Item Results"
        Case "RECEIVING_SELECT_AGGREGATE": caption = "Aggregate Received"
        Case "DISPOSITION_SELECT_AGGREGATE": caption = "Aggregate Returns"
        Case "RECEIVING_SELECT_HISTORY": caption = "Receiving Entries History"
        Case "DISPOSITION_SELECT_HISTORY": caption = "Return Entries History"
        Case "RECEIVING_SELECT_STAGED": caption = "Received Tally"
        Case "DISPOSITION_SELECT_STAGED": caption = "Return Tally"
        Case "RECEIVING_SELECT_CONDITION": caption = "Condition *"
        Case "DISPOSITION_SELECT_KIND": caption = "Disposition *"
        Case Else: Exit Function
    End Select
    If Left$(controlId, 12) = "DISPOSITION_" Then surface = surface & " > Returns"
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", controlId
    record.Add "OwnerId", "RECEIVING_NAVIGATION"
    record.Add "Class", "Navigation"
    record.Add "Role", "Receiving"
    record.Add "Caption", caption
    record.Add "Surface", surface
    record.Add "Capability", "RECEIVE_POST"
    record.Add "CodePrefix", controlId & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal controlId As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, message As String, severity As String, effect As String, nextStep As String
    Set definition = Control(controlId)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = "Receiving UI selection requested."
        Case "SELECTED"
            effect = "Unchanged"
            message = "Receiving UI selection completed; no staged work or inventory was changed."
        Case "FAILED"
            severity = "Error"
            message = "Receiving UI selection failed; its final display state requires verification."
            nextStep = "Inspect the owning form before continuing."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", controlId & "_" & code
    record.Add "OutcomeCode", code
    record.Add "Severity", severity
    record.Add "DataEffect", effect
    record.Add "UserMessage", message
    record.Add "NextStep", nextStep
    Set Outcome = record
End Function
