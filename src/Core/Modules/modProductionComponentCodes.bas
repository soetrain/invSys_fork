Attribute VB_Name = "modProductionComponentCodes"
Option Explicit
Option Private Module

' Catalog 16: local component edits never assert saved or Domain effects.
Public Function ControlIds() As Variant
    ControlIds = Split("PRODUCTION_PROCESS_REQUIREMENT_ADD|PRODUCTION_PROCESS_REQUIREMENT_UPDATE|" & _
                       "PRODUCTION_PROCESS_REQUIREMENT_REMOVE|PRODUCTION_PROCESS_REQUIREMENT_UP|" & _
                       "PRODUCTION_PROCESS_REQUIREMENT_DOWN|PRODUCTION_PROCESS_OUTPUT_ADD|" & _
                       "PRODUCTION_PROCESS_OUTPUT_UPDATE|PRODUCTION_PROCESS_OUTPUT_REMOVE|" & _
                       "PRODUCTION_PROCESS_OUTPUT_UP|PRODUCTION_PROCESS_OUTPUT_DOWN", "|")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, surface As String, action As String, record As Object
    If Left$(id, 31) = "PRODUCTION_PROCESS_REQUIREMENT_" Then
        surface = "Requirements": action = Mid$(id, 32)
    ElseIf Left$(id, 26) = "PRODUCTION_PROCESS_OUTPUT_" Then
        surface = "Outputs": action = Mid$(id, 27)
    Else
        Exit Function
    End If
    Select Case action
        Case "ADD": caption = "Add"
        Case "UPDATE": caption = "Update"
        Case "REMOVE": caption = "Remove"
        Case "UP": caption = "Up"
        Case "DOWN": caption = "Down"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "PRODUCTION_DESIGNER"
    record.Add "Class", "Command": record.Add "Role", "Production"
    record.Add "Caption", caption
    record.Add "Surface", "Operations > Production > Process Designer > " & surface
    record.Add "Capability", "PROD_POST": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, severity As String, effect As String
    Dim message As String, nextStep As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED": effect = "Unknown": message = "Process component edit requested."
        Case "STAGED"
            message = "Process components changed locally; saved definitions were not changed."
            nextStep = "Validate the process draft before using the owning Save command."
        Case "REJECTED"
            severity = "Warning": message = "The component edit requires valid input or selection; saved definitions were not changed."
            nextStep = "Review the current editor and selection before retrying."
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; the draft was not changed."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "The component edit failed; inspect the current draft before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
