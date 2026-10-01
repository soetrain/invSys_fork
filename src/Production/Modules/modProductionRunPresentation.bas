Attribute VB_Name = "modProductionRunPresentation"
Option Explicit
Option Private Module

' Observe the existing local tree owner; no inventory or Designs submission.
Public Function Execute(ByVal owner As frmProduction, ByVal collapsed As Boolean, _
                        ByVal context As String, ByVal operatorBook As Workbook, _
                        ByRef loading As Boolean, ByRef busy As Boolean) As String
    Dim action As cProductionWorksheetAction, report As String, suffix As String
    Dim priorLoading As Boolean, number As Long, source As String, description As String
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    priorLoading = loading: busy = True
    suffix = IIf(collapsed, "TREE_COLLAPSE", "TREE_EXPAND")
    Set action = New cProductionWorksheetAction
    If Not action.Begin("PRODUCTION_RUN_" & suffix, context, operatorBook, report) Then GoTo Done
    If Not action.CanContinue(report) Then GoTo Done
    owner.SetAllRunTreeGroupsCollapsed collapsed
    report = IIf(collapsed, "All ingredient choices hidden.", "All ingredient choices shown.")
    action.OutcomeCode = "PRESENTED"
Done:
    loading = priorLoading
    If Not action Is Nothing Then action.Finish report
    busy = False
    Execute = report
    If number <> 0 Then
        On Error GoTo 0
        Err.Raise number, source, description
    End If
    Exit Function
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    If Not action Is Nothing Then action.OutcomeCode = "FAILED"
    Resume Done
End Function
