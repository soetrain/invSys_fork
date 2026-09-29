Attribute VB_Name = "modProductionInstructionCodes"
Option Explicit
Option Private Module

' Catalog 14: local instruction edits, with no saved-definition or source effect.
Public Function ControlIds() As Variant
    ControlIds = Split("PRODUCTION_PROCESS_INSTRUCTION_ADD|PRODUCTION_PROCESS_INSTRUCTION_UPDATE|" & _
                       "PRODUCTION_PROCESS_INSTRUCTION_REMOVE|PRODUCTION_PROCESS_INSTRUCTION_UP|" & _
                       "PRODUCTION_PROCESS_INSTRUCTION_DOWN", "|")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, record As Object
    Select Case id
        Case "PRODUCTION_PROCESS_INSTRUCTION_ADD": caption = "Add"
        Case "PRODUCTION_PROCESS_INSTRUCTION_UPDATE": caption = "Update"
        Case "PRODUCTION_PROCESS_INSTRUCTION_REMOVE": caption = "Remove"
        Case "PRODUCTION_PROCESS_INSTRUCTION_UP": caption = "Up"
        Case "PRODUCTION_PROCESS_INSTRUCTION_DOWN": caption = "Down"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "PRODUCTION_DESIGNER"
    record.Add "Class", "Command": record.Add "Role", "Production"
    record.Add "Caption", caption
    record.Add "Surface", "Operations > Production > Process Designer > Instructions"
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
        Case "REQUESTED": effect = "Unknown": message = "Process instruction edit requested."
        Case "STAGED"
            message = "Process instructions changed locally; saved definitions were not changed."
            nextStep = "Validate the process draft before using the owning Save command."
        Case "REJECTED"
            severity = "Warning": message = "The instruction edit requires valid input or selection; the draft was not changed."
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; the draft was not changed."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "The instruction edit failed; verify the current draft before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
