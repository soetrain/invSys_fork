Attribute VB_Name = "modProductionRecipeOrderActions"
Option Explicit
Option Private Module

Public Enum ProductionRecipeOrderAction
    RecipeOrderUp = 1
    RecipeOrderDown
    RecipeOrderAuto
End Enum

Public Function EditOrder(ByVal owner As frmProduction, ByVal action As ProductionRecipeOrderAction, _
                          ByVal capturedContext As String, ByVal operatorWorkbook As Workbook, _
                          ByRef loading As Boolean, ByRef busy As Boolean) As String
    Dim id As String, activityId As String, notice As String, report As String
    Dim outcome As String, permitted As Boolean, wasLoading As Boolean
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    Select Case action
        Case RecipeOrderUp: id = "PRODUCTION_RECIPE_MOVE_UP"
        Case RecipeOrderDown: id = "PRODUCTION_RECIPE_MOVE_DOWN"
        Case RecipeOrderAuto: id = "PRODUCTION_RECIPE_AUTO_ORDER"
        Case Else: Exit Function
    End Select
    report = "Session, warehouse, or captured workbook changed. Reopen Production before editing the draft."
    If Not modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then EditOrder = report: Exit Function
    wasLoading = loading: busy = True
    activityId = modActivity.BeginAction(id, capturedContext, notice)
    If Not modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then GoTo Done
    permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("PROD_POST", report)
    If Not permitted Then permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report)
    If Not permitted Then
        outcome = "DENIED": report = "Production permission is required; the draft was not changed."
    Else
        report = vbNullString
        If owner.ApplyRecipeOrderAction(action, report) Then outcome = "STAGED" Else outcome = "REJECTED"
    End If
Done:
    loading = wasLoading
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    busy = False
    If notice <> "" Then report = report & " " & notice
    EditOrder = report
    Exit Function
Failed:
    outcome = "FAILED": report = "The recipe ordering action failed; inspect the current draft before retrying."
    Resume Done
End Function

' Preserve the existing bounded local ordering algorithm, including partial
' movement before a rejected graph. The form owns its existing acyclic check.
Public Sub OrderNodes(ByVal nodes As MSForms.ListBox, ByVal connections As MSForms.ListBox, _
                      ByVal instructions As MSForms.ListBox)
    Dim pass As Long, i As Long, fromIndex As Long, toIndex As Long, changed As Boolean
    For pass = 1 To nodes.ListCount * nodes.ListCount
        changed = False
        For i = 0 To connections.ListCount - 1
            fromIndex = modProductionRecipeLists.NodeIndex(nodes, modProductionRecipeLists.ListText(connections.List(i, 0)))
            toIndex = modProductionRecipeLists.NodeIndex(nodes, modProductionRecipeLists.ListText(connections.List(i, 2)))
            If fromIndex >= toIndex And fromIndex >= 0 And toIndex >= 0 Then
                nodes.ListIndex = fromIndex
                Call modProductionComponentLists.MoveRow(nodes, -1, instructions)
                changed = True
            End If
        Next i
        If Not changed Then Exit For
    Next pass
End Sub
