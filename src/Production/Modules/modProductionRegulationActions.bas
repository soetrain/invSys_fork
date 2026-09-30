Attribute VB_Name = "modProductionRegulationActions"
Option Explicit
Option Private Module

Public Function Stage(ByVal owner As frmProduction, ByVal applyRegulation As Boolean, _
                      ByVal capturedContext As String, ByVal operatorWorkbook As Workbook, _
                      ByRef loading As Boolean, ByRef busy As Boolean) As String
    Dim id As String, activityId As String, notice As String, report As String
    Dim outcome As String, permitted As Boolean, wasLoading As Boolean, staged As Boolean
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    id = "PRODUCTION_OUTPUT_REGULATION_CLEAR"
    If applyRegulation Then id = "PRODUCTION_OUTPUT_REGULATION_APPLY"
    report = "Session, warehouse, or captured workbook changed. Reopen Production before using the designer."
    If Not modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then Stage = report: Exit Function
    wasLoading = loading: busy = True
    activityId = modActivity.BeginAction(id, capturedContext, notice)
    If Not modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then GoTo Done
    permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("PROD_POST", report)
    If Not permitted Then permitted = modRoleUiAccess.CanCurrentUserPerformCapabilityCached("ADMIN_MAINT", report)
    If Not permitted Then
        outcome = "DENIED": report = "Production permission is required; the draft was not changed."
    Else
        report = ""
        If applyRegulation Then
            staged = owner.ApplyOutputRegulation(report)
        Else
            staged = owner.ClearOutputRegulation(report)
        End If
        outcome = "REJECTED"
        If staged Then outcome = "STAGED"
    End If
Done:
    loading = wasLoading
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    busy = False
    If notice <> "" Then report = Trim$(report & " " & notice)
    Stage = report
    Exit Function
Failed:
    outcome = "FAILED": report = "The output regulation action failed; inspect the current local draft before retrying."
    Resume Done
End Function

Public Function FindRecipe(ByVal regulations As Collection, ByVal nodeId As String, ByVal outputId As String) As Object
    Dim rawRecord As Variant, record As Object
    If regulations Is Nothing Then Exit Function
    For Each rawRecord In regulations
        Set record = rawRecord
        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(record, "OutputId"), outputId, vbTextCompare) = 0 Then
            Set FindRecipe = record
            Exit Function
        End If
    Next rawRecord
End Function

Public Sub SetProcess(ByRef regulations As Object, ByVal outputId As String, ByVal enabled As Boolean, _
                                       ByVal floorText As String, ByVal ceilingText As String)
    If regulations Is Nothing Then
        Set regulations = CreateObject("Scripting.Dictionary")
        regulations.CompareMode = vbTextCompare
    End If
    regulations(outputId) = Array(CStr(enabled), Trim$(floorText), Trim$(ceilingText))
End Sub

Public Sub RemoveRecipe(ByVal regulations As Collection, ByVal nodeId As String, ByVal outputId As String)
    Dim i As Long, record As Object
    If regulations Is Nothing Then Exit Sub
    For i = regulations.Count To 1 Step -1
        Set record = regulations(i)
        If StrComp(modProductionReusableDesigns.ReusableRecordText(record, "ProcessNodeId"), nodeId, vbTextCompare) = 0 _
           And StrComp(modProductionReusableDesigns.ReusableRecordText(record, "OutputId"), outputId, vbTextCompare) = 0 Then
            regulations.Remove i
        End If
    Next i
End Sub
