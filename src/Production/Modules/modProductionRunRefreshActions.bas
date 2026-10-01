Attribute VB_Name = "modProductionRunRefreshActions"
Option Explicit
Option Private Module

' Retain the distinct Loader/Manager sequences and the local owner's result.
Public Function Execute(ByVal owner As frmProduction, ByVal context As String, _
                        ByVal operatorBook As Workbook, ByRef loading As Boolean, ByRef busy As Boolean, _
                        ByVal isLoader As Boolean) As String
    Dim action As cProductionWorksheetAction, controlId As String
    Dim report As String, refreshReport As String, refreshed As Boolean, priorLoading As Boolean
    Dim number As Long, source As String, description As String
    If loading Or busy Then Exit Function
    On Error GoTo Failed
    priorLoading = loading: busy = True
    If isLoader Then controlId = "PRODUCTION_RUN_LOADER_REFRESH" Else controlId = "PRODUCTION_RUN_MANAGER_REFRESH"
    Set action = New cProductionWorksheetAction
    If Not action.Begin(controlId, context, operatorBook, report) Then GoTo Done
    If Not action.CanContinue(report) Then GoTo Done
    Set owner.RunActionContinuation = action
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        If isLoader Then
            owner.ResetInventoryCache
            owner.RefreshReusableDesignLists
            If Not action.CanContinue(report) Then GoTo Done
        End If
        owner.RefreshReusableRunControls True, action
        If Not action.CanContinue(report) Then GoTo Done
        report = "Reusable Production Run inventory refreshed from the exact entity projection."
    Else
        refreshed = owner.RefreshProductionInventoryReadModel(refreshReport)
        If Not action.CanContinue(report) Then GoTo Done
        owner.ResetInventoryCache
        If isLoader Then
            owner.RefreshReusableDesignLists
            If Not action.CanContinue(report) Then GoTo Done
        End If
        owner.RefreshLoaderState
        If Not action.CanContinue(report) Then GoTo Done
        owner.RefreshManagerState
        If Not action.CanContinue(report) Then GoTo Done
        If Not refreshed Then
            owner.ShowStatus refreshReport
            MsgBox refreshReport, vbExclamation, "Production Inventory Refresh"
            If Not action.CanContinue(report) Then GoTo Done
            report = refreshReport
            GoTo Done
        End If
        report = "Production Run inventory refreshed. " & refreshReport
    End If
    action.OutcomeCode = "REFRESHED"
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
