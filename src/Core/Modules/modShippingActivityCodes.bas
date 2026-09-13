Attribute VB_Name = "modShippingActivityCodes"
Option Explicit
Option Private Module

' Definitions describe owner facts; they never execute a Shipping command.
Public Function Control(ByVal controlId As String) As Object
    Dim caption As String, record As Object
    Select Case controlId
        Case "SHIPPING_ADD": caption = "Add"
        Case "SHIPPING_UPDATE": caption = "Update Row"
        Case "SHIPPING_REMOVE": caption = "Remove"
        Case "SHIPPING_HOLD": caption = "Send Hold"
        Case "SHIPPING_RETURN": caption = "Return"
        Case "SHIPPING_STAGE": caption = "To Shipments"
        Case "SHIPPING_SEND": caption = "Shipments Sent"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", controlId
    record.Add "OwnerId", "SHIPPING_WORKFLOW"
    record.Add "Class", "Command"
    record.Add "Role", "Shipping"
    record.Add "Caption", caption
    record.Add "Surface", "Operations > Shipping"
    record.Add "Capability", "SHIP_POST"
    record.Add "CodePrefix", controlId & "_"
    Set Control = record
End Function

Public Function HasInventorySources(ByVal controlId As String) As Boolean
    Select Case controlId
        Case "SHIPPING_ADD", "SHIPPING_UPDATE", "SHIPPING_REMOVE", "SHIPPING_STAGE", "SHIPPING_SEND"
            HasInventorySources = True
    End Select
End Function

Public Function Outcome(ByVal controlId As String, ByVal code As String) As Object
    Dim definition As Object, record As Object
    Dim severity As String, effect As String, message As String, nextStep As String
    Set definition = Control(controlId)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED"
            message = "Shipping action requested."
        Case "DENIED"
            severity = "Blocked": effect = "Unchanged"
            message = "Shipping authorization was not established; the action was stopped."
            nextStep = "Review Shipping access before continuing."
        Case "REJECTED"
            severity = "Warning": effect = "Unchanged"
            message = "Shipping validation stopped the action before staging or submission."
            nextStep = "Review the Shipping selection and prerequisites."
        Case "STAGED"
            If controlId = "SHIPPING_SEND" Then Exit Function
            effect = "Changed"
            message = "Local Shipping staging changed; inventory application and restart persistence are unconfirmed."
            nextStep = "Review the captured Shipping staging."
        Case "PENDING"
            If Not HasInventorySources(controlId) Then Exit Function
            severity = "Notice"
            message = "Shipping submission accepted; inventory application is unconfirmed."
            nextStep = "Inspect published outcomes for every related event before retrying."
        Case "CONFIRMED"
            If controlId <> "SHIPPING_SEND" Then Exit Function
            message = "Shipping processing and refresh finished; individual inventory application is unconfirmed."
            nextStep = "Inspect published outcomes for every related event."
        Case "FAILED"
            severity = "Error"
            message = "Shipping did not finish cleanly; its final data state requires verification."
            nextStep = "Inspect the Shipping workflow before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", definition("CodePrefix") & code
    record.Add "OutcomeCode", code
    record.Add "Severity", severity
    record.Add "DataEffect", effect
    record.Add "UserMessage", message
    record.Add "NextStep", nextStep
    Set Outcome = record
End Function
