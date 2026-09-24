Attribute VB_Name = "modProductionLifecycleCodes"
Option Explicit
Option Private Module

' D18 catalog 13. Command observations never assert Designs application.
Public Function ControlIds() As Variant
    ControlIds = Split("PRODUCTION_PROCESS_SAVE|PRODUCTION_PROCESS_RELEASE|PRODUCTION_PROCESS_OBSOLETE|" & _
                       "PRODUCTION_RECIPE_SAVE|PRODUCTION_RECIPE_RELEASE|PRODUCTION_RECIPE_OBSOLETE", "|")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, surface As String, record As Object
    Select Case id
        Case "PRODUCTION_PROCESS_SAVE", "PRODUCTION_RECIPE_SAVE": caption = "Save Draft"
        Case "PRODUCTION_PROCESS_RELEASE", "PRODUCTION_RECIPE_RELEASE": caption = "Release"
        Case "PRODUCTION_PROCESS_OBSOLETE", "PRODUCTION_RECIPE_OBSOLETE": caption = "Obsolete"
        Case Else: Exit Function
    End Select
    surface = "Operations > Production > Process Designer"
    If Left$(id, 18) = "PRODUCTION_RECIPE_" Then surface = "Operations > Production > Recipe Designer"
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "PRODUCTION_DESIGN_LIFECYCLE"
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
    severity = "Info": effect = "Unknown"
    Select Case code
        Case "REQUESTED": message = "Design lifecycle action requested."
        Case "REJECTED"
            severity = "Warning": effect = "Unchanged"
            message = "Local validation prevented submission."
            nextStep = "Review the owning designer's validation results."
        Case "DENIED"
            severity = "Blocked": effect = "Unchanged"
            message = "Production permission is required; no event was submitted."
        Case "CANCELLED"
            If Right$(id, 5) = "_SAVE" Then Exit Function
            severity = "Notice": effect = "Unchanged"
            message = "Design lifecycle confirmation declined; no event was submitted."
        Case "PENDING"
            severity = "Notice": message = "Design event submitted; the owning command did not finish successfully."
            nextStep = "Inspect the exact event in published Designs evidence before retrying."
        Case "CONFIRMED"
            message = "Design event submitted and the owning command finished; application requires Designs evidence."
        Case "FAILED"
            severity = "Error": message = "Design lifecycle action failed; its final data state requires verification."
            nextStep = "Inspect the captured designer and any referenced Designs event before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
