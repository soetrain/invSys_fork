Attribute VB_Name = "modProductionDesignReadActions"
Option Explicit
Option Private Module

Public Enum ProductionDesignReadAction
    DesignReadProcessRefresh = 1
    DesignReadProcessLoad
    DesignReadProcessReuse
    DesignReadRecipeRefresh
    DesignReadRecipeLoad
End Enum

Public Function ReadDesigner(ByVal owner As frmProduction, ByVal action As ProductionDesignReadAction, _
                             ByVal processes As MSForms.ListBox, ByVal recipes As MSForms.ListBox, _
                             ByVal capturedContext As String, ByVal operatorWorkbook As Workbook, _
                             ByRef loading As Boolean, ByRef busy As Boolean) As String
    Dim id As String, activityId As String, notice As String, report As String
    Dim outcome As String, permitted As Boolean, wasLoading As Boolean
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    Select Case action
        Case DesignReadProcessRefresh: id = "PRODUCTION_PROCESS_REFRESH"
        Case DesignReadProcessLoad: id = "PRODUCTION_PROCESS_LOAD"
        Case DesignReadProcessReuse: id = "PRODUCTION_PROCESS_REUSE"
        Case DesignReadRecipeRefresh: id = "PRODUCTION_RECIPE_REFRESH"
        Case DesignReadRecipeLoad: id = "PRODUCTION_RECIPE_LOAD"
        Case Else: Exit Function
    End Select
    report = "Session, warehouse, or captured workbook changed. Reopen Production before using the designer."
    If Not modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then ReadDesigner = report: Exit Function
    wasLoading = loading: busy = True
    activityId = modActivity.BeginAction(id, capturedContext, notice)
    If Not modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then GoTo Done
    permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("PROD_POST", report)
    If Not permitted Then permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report)
    If Not permitted Then
        outcome = "DENIED": report = "Production permission is required; the draft was not changed."
    Else
        outcome = ApplyLocalRead(owner, action, processes, recipes, report)
    End If
Done:
    loading = wasLoading
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    busy = False
    If notice <> "" Then report = report & " " & notice
    ReadDesigner = report
    Exit Function
Failed:
    outcome = "FAILED": report = "The designer action failed; inspect the current designer before retrying."
    Resume Done
End Function

Private Function ApplyLocalRead(ByVal owner As frmProduction, ByVal action As ProductionDesignReadAction, _
                                ByVal processes As MSForms.ListBox, ByVal recipes As MSForms.ListBox, _
                                ByRef report As String) As String
    Dim selected As MSForms.ListBox, index As Long, loaded As Boolean
    If action = DesignReadProcessRefresh Or action = DesignReadRecipeRefresh Then
        owner.RefreshReusableDesignLists
        report = "Process Designer refreshed."
        If action = DesignReadRecipeRefresh Then report = "Recipe Designer refreshed."
        ApplyLocalRead = "REFRESHED"
        Exit Function
    End If
    Set selected = processes
    If action = DesignReadRecipeLoad Then Set selected = recipes
    index = selected.ListIndex
    If index < 0 Then
        report = "Select a saved Process first."
        If action = DesignReadRecipeLoad Then report = "Select a saved Recipe first."
        ApplyLocalRead = "REJECTED"
        Exit Function
    End If
    If action = DesignReadRecipeLoad Then
        loaded = owner.LoadRecipeDefinitionIntoDesigner(modProductionRecipeLists.ListText(selected.List(index, 0)), _
                                                       modProductionRecipeLists.ListText(selected.List(index, 1)), report)
    Else
        loaded = owner.LoadProcessDefinitionIntoDesigner(modProductionRecipeLists.ListText(selected.List(index, 0)), _
                                                        modProductionRecipeLists.ListText(selected.List(index, 1)), _
                                                        action = DesignReadProcessReuse, report)
    End If
    ApplyLocalRead = "FAILED"
    If loaded Then
        ApplyLocalRead = "PRESENTED"
        If action = DesignReadProcessReuse Then ApplyLocalRead = "STAGED"
    End If
End Function
