Attribute VB_Name = "modProductionDraftCodes"
Option Explicit
Option Private Module

' D18 catalog 12: fixed observations of local drafts, never saved design effects.
Public Function ControlIds() As Variant
    ControlIds = Split("PRODUCTION_PROCESS_NEW|PRODUCTION_PROCESS_CLEAR|PRODUCTION_PROCESS_VALIDATE|" & _
                       "PRODUCTION_RECIPE_NEW|PRODUCTION_RECIPE_CLEAR|PRODUCTION_RECIPE_VALIDATE", "|")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, surface As String, record As Object
    Select Case id
        Case "PRODUCTION_PROCESS_NEW": caption = "New Process"
        Case "PRODUCTION_RECIPE_NEW": caption = "New Recipe"
        Case "PRODUCTION_PROCESS_CLEAR", "PRODUCTION_RECIPE_CLEAR": caption = "Clear"
        Case "PRODUCTION_PROCESS_VALIDATE": caption = "Validate"
        Case "PRODUCTION_RECIPE_VALIDATE": caption = "Validate Recipe"
        Case Else: Exit Function
    End Select
    surface = "Operations > Production > Process Designer"
    If Left$(id, 18) = "PRODUCTION_RECIPE_" Then surface = "Operations > Production > Recipe Designer"
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "PRODUCTION_DESIGNER"
    record.Add "Class", "Command": record.Add "Role", "Production"
    record.Add "Caption", caption: record.Add "Surface", surface
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
        Case "REQUESTED": effect = "Unknown": message = "Designer draft action requested."
        Case "STAGED"
            If Right$(id, 9) = "_VALIDATE" Then Exit Function
            message = "Designer draft reset locally; saved definitions were not changed."
            nextStep = "Enter the draft and validate it before using the owning Save command."
        Case "VALIDATED"
            If Right$(id, 9) <> "_VALIDATE" Then Exit Function
            message = "Local draft validation passed; no definition was saved or released."
        Case "REJECTED"
            If Right$(id, 9) <> "_VALIDATE" Then Exit Function
            severity = "Warning": message = "Local draft validation found issues; saved definitions were not changed."
            nextStep = "Review the owning designer's validation results."
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; the draft was not changed."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "The designer draft action failed; verify its current state."
            nextStep = "Inspect the captured Production workbook and designer before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
