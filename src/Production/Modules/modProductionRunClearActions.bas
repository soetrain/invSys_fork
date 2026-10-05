Attribute VB_Name = "modProductionRunClearActions"
Option Explicit
Option Private Module

' Observe the existing local reset/cleanup owners without submitting Domain work.
Public Function Execute(ByVal owner As frmProduction, ByVal context As String, _
                        ByVal operatorBook As Workbook, ByRef loading As Boolean, ByRef busy As Boolean, _
                        ByVal lines As MSForms.ListBox, ByVal palette As MSForms.ListBox, _
                        ByVal checks As MSForms.ListBox, ByVal outputs As MSForms.ListBox, _
                        ByVal instructions As MSForms.ListBox) As String
    Dim action As cProductionWorksheetAction, report As String, priorLoading As Boolean
    Dim number As Long, source As String, description As String
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    priorLoading = loading: busy = True
    Set action = New cProductionWorksheetAction
    If Not action.Begin("PRODUCTION_RUN_CLEAR", context, operatorBook, report) Then GoTo Done
    If Not action.CanContinue(report) Then GoTo Done
    Set owner.RunActionContinuation = action
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        modProductionReusableRun.ClearReusableRun
        If Not action.CanContinue(report) Then GoTo Done
        lines.Clear
        If Not action.CanContinue(report) Then GoTo Done
        palette.Clear
        If Not action.CanContinue(report) Then GoTo Done
        checks.Clear
        If Not action.CanContinue(report) Then GoTo Done
        outputs.Clear
        If Not action.CanContinue(report) Then GoTo Done
        instructions.Clear
        If Not action.CanContinue(report) Then GoTo Done
        report = "Reusable Production Run cleared."
    Else
        If Not modProductionRunBinding.BindWorksheetOwner(operatorBook) Then GoTo Done
        mProduction.BtnClearRecipeChooser
        If Not action.CanContinue(report) Then GoTo Done
        owner.RefreshLoaderState
        If Not action.CanContinue(report) Then GoTo Done
        owner.RefreshManagerState
        If Not action.CanContinue(report) Then GoTo Done
        report = "Production Run cleared."
    End If
    action.OutcomeCode = "STAGED"
Done:
    Set owner.RunActionContinuation = Nothing
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

Public Function ContinueRefresh(ByVal owner As frmProduction, ByVal action As cProductionWorksheetAction) As Boolean
    Dim report As String
    If Not modOperationsFormLifetime.IsLoaded(owner) Then Exit Function
    ContinueRefresh = True
    If action Is Nothing Then Exit Function
    ContinueRefresh = action.CanContinue(report)
    If Not ContinueRefresh Then owner.ShowStatus report
End Function
