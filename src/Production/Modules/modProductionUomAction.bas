Attribute VB_Name = "modProductionUomAction"
Option Explicit

Public Function OpenWorkbench(ByVal workbook As Workbook, ByVal context As String, _
                              ByVal loading As Boolean, ByRef busy As Boolean) As String
    Dim activityId As String, notice As String, report As String, outcome As String
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    report = "Session, warehouse, or captured workbook changed. Reopen Production before opening the UOM workbench."
    If Not modProductionDesignerActions.ContextIsCurrent(context, workbook) Then OpenWorkbench = report: Exit Function
    busy = True
    activityId = modActivity.BeginAction("PRODUCTION_UOM_EDIT", context, notice)
    If Not modProductionDesignerActions.ContextIsCurrent(context, workbook) Then GoTo Done
    If Not modRoleUiAccess.CanCurrentUserPerformCapabilityCached("PROD_POST", report) Then
        outcome = "DENIED": report = "Production permission is required; staging was not changed."
    Else
        If Not modProductionUomCatalog.SendUomCatalogToWorksheet(workbook, report, outcome) Then
            report = "UOM Catalog export failed: " & report
        End If
    End If
Done:
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    busy = False
    If notice <> "" Then report = report & " " & notice
    OpenWorkbench = report
    Exit Function
Failed:
    outcome = "FAILED": report = "The UOM workbench could not be opened. Verify the staging worksheet before retrying."
    Resume Done
End Function

' Called by the form's deliberate Retrieve action, not by staging/service reads.
Public Function Retrieve(ByVal workbook As Workbook, ByVal context As String, ByRef report As String) As Boolean
    Dim activityId As String, notice As String, outcome As String
    On Error GoTo Failed
    report = "Session or warehouse changed. Reopen Production before retrieving the catalog."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    activityId = modActivity.BeginAction("PRODUCTION_UOM_RETRIEVE", context, notice)
    Retrieve = modProductionUomCatalog.RetrieveUomCatalogFromWorksheet(workbook, report, outcome)
    If Not Retrieve Then report = "UOM Catalog retrieval failed: " & report
Done:
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    If notice <> "" Then report = report & " " & notice
    Exit Function
Failed:
    Retrieve = False
    outcome = "FAILED"
    report = "UOM Catalog retrieval failed. Verify the current catalog before retrying."
    Resume Done
End Function
