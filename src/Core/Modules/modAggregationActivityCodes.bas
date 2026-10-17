Attribute VB_Name = "modAggregationActivityCodes"
Option Explicit
Option Private Module

Public Function Control(ByVal id As String) As Object
    Dim caption As String, kind As String, record As Object
    kind = "Command"
    Select Case id
        Case "ADMIN_AGGREGATION_OPEN": caption = "Aggregate Global Snapshot": kind = "Navigation"
        Case "ADMIN_AGGREGATION_REFRESH": caption = "Refresh Connected"
        Case "ADMIN_AGGREGATION_ADD_SERVER": caption = "Add Server": kind = "Navigation"
        Case "ADMIN_AGGREGATION_CONNECT": caption = "Connect + Discover"
        Case "ADMIN_AGGREGATION_SELECT_SOURCE": caption = "Select Sources": kind = "Navigation"
        Case "ADMIN_AGGREGATION_AGGREGATE": caption = "Aggregate Selected Sources"
        Case "ADMIN_AGGREGATION_CLOSE": caption = "Close": kind = "Navigation"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "ADMIN_HQ_AGGREGATION"
    record.Add "Class", kind: record.Add "Role", "Admin": record.Add "Caption", caption
    record.Add "Surface", "Admin > Aggregation": record.Add "Capability", "ADMIN_MAINT"
    record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function PositiveOutcome(ByVal id As String, ByVal code As String) As Boolean
    Select Case id
        Case "ADMIN_AGGREGATION_OPEN": PositiveOutcome = (code = "OPENED")
        Case "ADMIN_AGGREGATION_REFRESH": PositiveOutcome = (code = "DISCOVERED")
        Case "ADMIN_AGGREGATION_ADD_SERVER": PositiveOutcome = (code = "STAGED")
        Case "ADMIN_AGGREGATION_CONNECT": PositiveOutcome = (code = "CONNECTED")
        Case "ADMIN_AGGREGATION_SELECT_SOURCE": PositiveOutcome = (code = "SELECTED")
        Case "ADMIN_AGGREGATION_AGGREGATE": PositiveOutcome = (code = "COMPLETED")
        Case "ADMIN_AGGREGATION_CLOSE": PositiveOutcome = (code = "CLOSED")
    End Select
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, message As String, severity As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info"
    Select Case code
        Case "REQUESTED": message = definition("Caption") & " requested."
        Case "REJECTED": severity = "Warning": message = "The request was refused. Read the form result."
        Case "DENIED": severity = "Warning": message = "Source context changed. Reopen Aggregation. No rollback is asserted."
        Case "FAILED": severity = "Error": message = "The action did not finish. Read the form result; no rollback is asserted."
        Case "CANCELLED", "INDETERMINATE", "COMPLETED_STOPPED"
            If id <> "ADMIN_AGGREGATION_AGGREGATE" Then Exit Function
            Select Case code
                Case "CANCELLED": message = "The retained-attempt review was cancelled. Its existing state is retained."
                Case "INDETERMINATE": severity = "Warning": message = "Output persistence needs review. Its bytes and retry evidence are retained."
                Case "COMPLETED_STOPPED": severity = "Warning": message = "Complete output was saved. Further work stopped; inspect the output and source context before retrying."
            End Select
        Case Else
            If Not PositiveOutcome(id, code) Then Exit Function
            Select Case code
                Case "OPENED": message = "Aggregation opened."
                Case "DISCOVERED": message = "Source discovery returned. Review READY and rejected sources."
                Case "STAGED": message = "Connection entry was selected. No Windows connection is asserted."
                Case "CONNECTED": message = "The Windows connection returned success. Review the discovered sources."
                Case "SELECTED": message = "The session-only source selection changed. No aggregation is asserted."
                Case "COMPLETED": message = "The owning command confirmed complete advisory output. Warehouse inventory remains authoritative."
                Case "CLOSED": message = "Aggregation closed. No aggregation or rollback is asserted."
            End Select
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", "Unknown"
    record.Add "UserMessage", message: record.Add "NextStep", ""
    Set Outcome = record
End Function
