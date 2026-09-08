Attribute VB_Name = "modProductionUomAction"
Option Explicit

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
