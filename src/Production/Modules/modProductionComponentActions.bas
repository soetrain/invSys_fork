Attribute VB_Name = "modProductionComponentActions"
Option Explicit
Option Private Module

Public Enum ProductionComponentKind
    ComponentRequirement = 1
    ComponentOutput
End Enum

Public Enum ProductionComponentAction
    ComponentAdd = 1
    ComponentUpdate
    ComponentRemove
    ComponentUp
    ComponentDown
End Enum

Public Function EditComponent(ByVal owner As frmProduction, ByVal kind As ProductionComponentKind, _
                              ByVal action As ProductionComponentAction, ByVal capturedContext As String, _
                              ByVal operatorWorkbook As Workbook, ByRef loading As Boolean, ByRef busy As Boolean) As String
    Dim id As String, suffix As String, activityId As String, notice As String, report As String
    Dim outcome As String, permitted As Boolean, wasLoading As Boolean
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    Select Case kind
        Case ComponentRequirement: id = "PRODUCTION_PROCESS_REQUIREMENT_"
        Case ComponentOutput: id = "PRODUCTION_PROCESS_OUTPUT_"
        Case Else: Exit Function
    End Select
    Select Case action
        Case ComponentAdd: suffix = "ADD"
        Case ComponentUpdate: suffix = "UPDATE"
        Case ComponentRemove: suffix = "REMOVE"
        Case ComponentUp: suffix = "UP"
        Case ComponentDown: suffix = "DOWN"
        Case Else: Exit Function
    End Select
    report = "Session, warehouse, or captured workbook changed. Reopen Production before editing the draft."
    If Not modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then EditComponent = report: Exit Function
    wasLoading = loading: busy = True
    activityId = modActivity.BeginAction(id & suffix, capturedContext, notice)
    If Not modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then GoTo Done
    permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("PROD_POST", report)
    If Not permitted Then permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report)
    If Not permitted Then
        outcome = "DENIED": report = "Production permission is required; the draft was not changed."
    ElseIf owner.ApplyComponentEdit(kind, action, report) Then
        outcome = "STAGED": report = "Process components changed locally; saved definitions were not changed."
    Else
        outcome = "REJECTED"
    End If
Done:
    loading = wasLoading
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    busy = False
    If notice <> "" Then report = report & " " & notice
    EditComponent = report
    Exit Function
Failed:
    outcome = "FAILED": report = "The component edit failed; inspect the current draft before retrying."
    Resume Done
End Function
