Attribute VB_Name = "modProductionDesignReadCodes"
Option Explicit
Option Private Module

Public Function ControlIds() As Variant
    ControlIds = Array("PRODUCTION_PROCESS_REFRESH", "PRODUCTION_PROCESS_LOAD", "PRODUCTION_PROCESS_REUSE", _
                       "PRODUCTION_RECIPE_REFRESH", "PRODUCTION_RECIPE_LOAD")
End Function

Public Function Control(ByVal id As String) As Object
    Dim caption As String, record As Object, designer As String
    Select Case id
        Case "PRODUCTION_PROCESS_REFRESH", "PRODUCTION_RECIPE_REFRESH": caption = "Refresh"
        Case "PRODUCTION_PROCESS_LOAD": caption = "View Process"
        Case "PRODUCTION_PROCESS_REUSE": caption = "Edit as New Version"
        Case "PRODUCTION_RECIPE_LOAD": caption = "Load"
        Case Else: Exit Function
    End Select
    designer = "Process Designer"
    If Left$(id, 18) = "PRODUCTION_RECIPE_" Then designer = "Recipe Designer"
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "ControlId", id: record.Add "OwnerId", "PRODUCTION_DESIGNER"
    record.Add "Class", "Command": record.Add "Role", "Production"
    record.Add "Caption", caption: record.Add "Surface", "Operations > Production > " & designer
    record.Add "Capability", "PROD_POST": record.Add "CodePrefix", id & "_"
    Set Control = record
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, positive As String
    Dim severity As String, effect As String, message As String, nextStep As String
    Set definition = Control(id)
    If definition Is Nothing Then Exit Function
    positive = "PRESENTED"
    If Right$(id, 7) = "REFRESH" Then positive = "REFRESHED"
    If id = "PRODUCTION_PROCESS_REUSE" Then positive = "STAGED"
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED"
            effect = "Unknown": message = "Designer action requested."
        Case positive
            If code = "REFRESHED" Then
                message = "Design list refresh finished; saved definitions were not changed."
                nextStep = "Check the displayed lists before selecting a saved design; source availability is not established by this observation."
            ElseIf code = "PRESENTED" Then
                message = "The designer view was updated locally; saved definitions were not changed."
                nextStep = "Inspect the displayed design; this observation does not establish its validity or release status."
            Else
                message = "A process draft was prepared locally for editing; saved definitions were not changed."
                nextStep = "Validate before saving; the proposed version has not been reserved or saved."
            End If
        Case "REJECTED"
            severity = "Warning": message = "Select a saved design before loading it; saved definitions were not changed."
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; the draft was not changed."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "The designer action failed; inspect the current designer before retrying."
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
