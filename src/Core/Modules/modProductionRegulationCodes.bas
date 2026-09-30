Attribute VB_Name = "modProductionRegulationCodes"
Option Explicit
Option Private Module

Public Function ControlIds() As Variant
    ControlIds = Array("PRODUCTION_OUTPUT_REGULATION_APPLY", "PRODUCTION_OUTPUT_REGULATION_CLEAR")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, record As Object
    Select Case id
        Case "PRODUCTION_OUTPUT_REGULATION_APPLY": caption = "Apply Regulation"
        Case "PRODUCTION_OUTPUT_REGULATION_CLEAR": caption = "Clear Override"
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "PRODUCTION_DESIGNER"
    record.Add "Class", "Command": record.Add "Role", "Production"
    record.Add "Caption", caption: record.Add "Surface", "Operations > Production > Production Settings"
    record.Add "Capability", "PROD_POST": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object
    Dim severity As String, effect As String, message As String, nextStep As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED"
            effect = "Unknown": message = "Output regulation action requested."
        Case "STAGED"
            message = "Output regulation staged locally; saved definitions were not changed."
            nextStep = "Inspect and validate the local draft before using the owning Save command."
        Case "REJECTED"
            severity = "Warning": message = "Output regulation requires valid input or selection; saved definitions were not changed."
            nextStep = "Review the current editor and selection before retrying."
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; the draft was not changed."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "The output regulation action failed; inspect the current local draft before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
