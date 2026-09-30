Attribute VB_Name = "modProductionLocalEditCodes"
Option Explicit
Option Private Module

' Shared facts for local component and recipe edits. Owner-specific
' wording remains exact; no successful local edit asserts saved authority.
Public Function Outcome(ByVal id As String, ByVal code As String) As Object
    Dim definition As Object, record As Object, recipeOrder As Boolean, recipeStructure As Boolean
    Dim severity As String, effect As String, message As String, nextStep As String
    Set definition = modProductionRecipeOrderCodes.Control(id)
    recipeOrder = Not definition Is Nothing
    If definition Is Nothing Then Set definition = modRecipeStructureCodes.Control(id)
    recipeStructure = (Not definition Is Nothing) And Not recipeOrder
    If definition Is Nothing Then Set definition = modProductionComponentCodes.Control(id)
    If definition Is Nothing Then Exit Function
    severity = "Info": effect = "Unchanged"
    Select Case code
        Case "REQUESTED"
            effect = "Unknown"
            If recipeOrder Then
                message = "Recipe ordering requested."
            ElseIf recipeStructure Then
                message = "Recipe structure edit requested."
            Else
                message = "Process component edit requested."
            End If
        Case "STAGED"
            If recipeOrder Then
                message = "Recipe execution order staged locally; saved definitions were not changed."
                nextStep = "Validate the recipe draft before using the owning Save command."
            ElseIf recipeStructure Then
                message = "Recipe structure changed locally; saved definitions were not changed."
                nextStep = "Validate the recipe draft before using the owning Save command."
            Else
                message = "Process components changed locally; saved definitions were not changed."
                nextStep = "Validate the process draft before using the owning Save command."
            End If
        Case "REJECTED"
            severity = "Warning"
            If recipeOrder Then
                message = "Recipe ordering was rejected; saved definitions were not changed."
                nextStep = "Review the selection and recipe connections; local ordering changes may remain."
            ElseIf recipeStructure Then
                message = "The recipe structure edit requires valid input or selection; saved definitions were not changed."
                nextStep = "Review the current editor and selection before retrying."
            Else
                message = "The component edit requires valid input or selection; saved definitions were not changed."
                nextStep = "Review the current editor and selection before retrying."
            End If
        Case "DENIED"
            severity = "Blocked": message = "Production permission is required; the draft was not changed."
        Case "FAILED"
            severity = "Error": effect = "Unknown"
            If recipeOrder Then
                message = "The recipe ordering action failed; inspect the current draft before retrying."
            ElseIf recipeStructure Then
                message = "The recipe structure edit failed; inspect the current draft before retrying."
            Else
                message = "The component edit failed; inspect the current draft before retrying."
            End If
        Case Else: Exit Function
    End Select
    Set record = CreateObject("Scripting.Dictionary")
    record.Add "EventCode", id & "_" & code: record.Add "OutcomeCode", code
    record.Add "Severity", severity: record.Add "DataEffect", effect
    record.Add "UserMessage", message: record.Add "NextStep", nextStep
    Set Outcome = record
End Function
