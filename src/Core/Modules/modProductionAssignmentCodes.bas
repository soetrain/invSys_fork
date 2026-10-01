Attribute VB_Name = "modProductionAssignmentCodes"
Option Explicit
Option Private Module

' D18 catalog23: local Assignment gestures and one owning Designs save.
Public Function ControlIds() As Variant
    ControlIds = Array("PRODUCTION_ASSIGNMENT_REFRESH", "PRODUCTION_ASSIGNMENT_PROCESS", _
        "PRODUCTION_ASSIGNMENT_REQUIREMENT", "PRODUCTION_ASSIGNMENT_ADD", "PRODUCTION_ASSIGNMENT_REMOVE", _
        "PRODUCTION_ASSIGNMENT_CLEAR", "PRODUCTION_ASSIGNMENT_SAVE", _
        "PRODUCTION_ASSIGNMENT_PROCESS_SELECT", "PRODUCTION_ASSIGNMENT_REQUIREMENT_SELECT")
End Function

Public Function Control(ByVal id As String) As Object
    Dim ids As Variant, captions As Variant, index As Long, record As Object
    ids = ControlIds()
    captions = Array("Refresh", "Select Process", "Select Requirement", "Add Acceptable", "Remove Row", _
                     "Clear", "Save Alternatives", "Processes", "Ingredient Requirements")
    For index = LBound(ids) To UBound(ids)
        If id = ids(index) Then
            Set record = modProductionControlCatalog.Command(id, "PRODUCTION_ASSIGNMENT", CStr(captions(index)), _
                                                            "Operations > Production > Ingredients Assignment")
            If id = "PRODUCTION_ASSIGNMENT_PROCESS_SELECT" Or id = "PRODUCTION_ASSIGNMENT_REQUIREMENT_SELECT" Then record("Class") = "Navigation"
            Set Control = record
            Exit Function
        End If
    Next index
End Function

Public Function PositiveOutcome(ByVal id As String) As String
    Select Case id
        Case "PRODUCTION_ASSIGNMENT_REFRESH": PositiveOutcome = "REFRESHED"
        Case "PRODUCTION_ASSIGNMENT_PROCESS", "PRODUCTION_ASSIGNMENT_PROCESS_SELECT": PositiveOutcome = "PRESENTED"
        Case "PRODUCTION_ASSIGNMENT_REQUIREMENT", "PRODUCTION_ASSIGNMENT_REQUIREMENT_SELECT": PositiveOutcome = "SELECTED"
        Case "PRODUCTION_ASSIGNMENT_ADD", "PRODUCTION_ASSIGNMENT_REMOVE", "PRODUCTION_ASSIGNMENT_CLEAR": PositiveOutcome = "STAGED"
        Case "PRODUCTION_ASSIGNMENT_SAVE": PositiveOutcome = "CONFIRMED"
    End Select
End Function

Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim record As Object, positive As String, severity As String, effect As String
    Dim message As String, nextStep As String
    positive = PositiveOutcome(id)
    If positive = "" Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED"
            effect = "Unknown": message = "Ingredients Assignment action requested."
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; no Assignment read, draft edit or Designs submission was made."
        Case "REJECTED"
            severity = "Warning": message = "Selection or validation prevented the Assignment action; saved definitions were not changed."
            nextStep = "Review the form's selection and validation results; local preparation before Save may remain."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            message = "The Assignment action failed or partly completed; its final data state requires verification."
            nextStep = "Inspect the local draft and any exact referenced Designs events before retrying; prior work may remain."
        Case "PENDING"
            If id <> "PRODUCTION_ASSIGNMENT_SAVE" Then Exit Function
            severity = "Notice": effect = "Unknown"
            message = "The Designs save was submitted, but its processing, projection or refresh command did not finish."
            nextStep = "Check the exact referenced Designs event before retrying; submission does not prove application."
        Case "REFRESHED", "PRESENTED", "SELECTED", "STAGED", "CONFIRMED"
            If code <> positive Then Exit Function
            Select Case code
                Case "REFRESHED": message = "Assignment lists were refreshed locally; source availability and completeness were not established."
                Case "PRESENTED": message = "Selected Process data was presented locally; no complete or released Process is certified."
                Case "SELECTED": message = "The selected requirement's acceptable-item view was presented locally."
                Case "STAGED": message = "Acceptable-item assignments changed locally; saved definitions were not changed."
                Case "CONFIRMED"
                    effect = "Unknown"
                    message = "The owning Designs save command finished; application requires exact Designs evidence."
            End Select
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
