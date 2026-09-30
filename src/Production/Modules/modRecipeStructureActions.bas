Attribute VB_Name = "modRecipeStructureActions"
Option Explicit
Option Private Module

Public Enum ProductionRecipeStructureAction
    RecipeAddProcess = 1
    RecipeRemoveProcess
    RecipeConnect
    RecipeUpdateConnection
    RecipeDisconnect
End Enum

Public Function EditStructure(ByVal owner As frmProduction, ByVal action As ProductionRecipeStructureAction, _
                              ByVal capturedContext As String, ByVal operatorWorkbook As Workbook, _
                              ByRef loading As Boolean, ByRef busy As Boolean) As String
    Dim id As String, activityId As String, notice As String, report As String
    Dim outcome As String, permitted As Boolean, wasLoading As Boolean
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    Select Case action
        Case RecipeAddProcess: id = "PRODUCTION_RECIPE_ADD_PROCESS"
        Case RecipeRemoveProcess: id = "PRODUCTION_RECIPE_REMOVE_PROCESS"
        Case RecipeConnect: id = "PRODUCTION_RECIPE_CONNECT"
        Case RecipeUpdateConnection: id = "PRODUCTION_RECIPE_UPDATE_CONNECTION"
        Case RecipeDisconnect: id = "PRODUCTION_RECIPE_DISCONNECT"
        Case Else: Exit Function
    End Select
    report = "Session, warehouse, or captured workbook changed. Reopen Production before editing the draft."
    If Not modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then EditStructure = report: Exit Function
    wasLoading = loading: busy = True
    activityId = modActivity.BeginAction(id, capturedContext, notice)
    If Not modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then GoTo Done
    permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("PROD_POST", report)
    If Not permitted Then permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report)
    If Not permitted Then
        outcome = "DENIED": report = "Production permission is required; the draft was not changed."
    Else
        report = vbNullString
        If owner.ApplyRecipeStructureAction(action) Then outcome = "STAGED" Else outcome = "REJECTED"
    End If
Done:
    loading = wasLoading
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    busy = False
    If notice <> "" Then report = report & " " & notice
    EditStructure = report
    Exit Function
Failed:
    outcome = "FAILED": report = "The recipe structure edit failed; inspect the current draft before retrying."
    Resume Done
End Function
