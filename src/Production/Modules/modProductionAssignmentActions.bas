Attribute VB_Name = "modProductionAssignmentActions"
Option Explicit
Option Private Module

Public Enum ProductionAssignmentCommand
    AssignmentRefresh = 1
    AssignmentProcess
    AssignmentRequirement
    AssignmentAdd
    AssignmentRemove
    AssignmentClear
    AssignmentSave
    AssignmentProcessSelect
    AssignmentRequirementSelect
End Enum

Public Function Execute(ByVal owner As frmProduction, ByVal command As ProductionAssignmentCommand, _
                        ByVal context As String, ByVal operatorBook As Workbook, _
                        ByRef loading As Boolean, ByRef busy As Boolean) As String
    Dim action As cProductionWorksheetAction, facts As cProductionLifecycleFacts
    Dim suffix As String, report As String, priorLoading As Boolean
    Dim errorNumber As Long, errorSource As String, errorText As String
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    Select Case command
        Case AssignmentRefresh: suffix = "REFRESH"
        Case AssignmentProcess: suffix = "PROCESS"
        Case AssignmentRequirement: suffix = "REQUIREMENT"
        Case AssignmentAdd: suffix = "ADD"
        Case AssignmentRemove: suffix = "REMOVE"
        Case AssignmentClear: suffix = "CLEAR"
        Case AssignmentSave: suffix = "SAVE"
        Case AssignmentProcessSelect: suffix = "PROCESS_SELECT"
        Case AssignmentRequirementSelect: suffix = "REQUIREMENT_SELECT"
        Case Else: Exit Function
    End Select
    priorLoading = loading: busy = True
    ' The existing Operations observer has no worksheet mutation authority.
    Set action = New cProductionWorksheetAction
    If Not action.Begin("PRODUCTION_ASSIGNMENT_" & suffix, context, operatorBook, report) Then GoTo Done
    Set facts = New cProductionLifecycleFacts: facts.OutcomeCode = "FAILED"
    facts.BindContinuation action
    action.OutcomeCode = owner.ApplyAssignmentCommand(command, action, facts)
    report = owner.TestStatusText()
    If command = AssignmentSave Then
        action.ObserveSubmission facts
        action.OutcomeCode = facts.OutcomeCode
    End If
Done:
    loading = priorLoading
    If Not action Is Nothing Then action.Finish report
    busy = False
    Execute = report
    If errorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise errorNumber, errorSource, errorText
    End If
    Exit Function
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    If Not action Is Nothing Then
        action.OutcomeCode = "FAILED"
        If Not facts Is Nothing Then
            action.ObserveSubmission facts
            If facts.EventId <> "" Then action.OutcomeCode = facts.OutcomeCode
        End If
    End If
    Resume Done
End Function
