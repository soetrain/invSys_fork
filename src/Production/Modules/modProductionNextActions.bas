Attribute VB_Name = "modProductionNextActions"
Option Explicit
Option Private Module

' Observe the actual owner acknowledgment, then the captured local refresh.
Public Sub Execute(ByVal owner As frmProduction, ByVal context As String, _
                   ByVal operatorBook As Workbook, ByRef loading As Boolean, ByRef busy As Boolean)
    Dim priorLoading As Boolean, priorBusy As Boolean, completed As Boolean
    Dim action As cProductionWorksheetAction, report As String
    Dim number As Long, source As String, description As String, helpFile As String, helpContext As Long
    If loading Or busy Then Exit Sub
    priorLoading = loading: priorBusy = busy
    On Error GoTo Failed
    busy = True
    Set action = New cProductionWorksheetAction
    If Not action.Begin("PRODUCTION_RUN_NEXT_BATCH", context, operatorBook, report) Then
        If action.OutcomeCode = "DENIED" Then report = "Production permission changed. Reopen Production before continuing."
        GoTo Done
    End If
    If Not action.CanContinue(report) Then GoTo Done
    Set owner.RunActionContinuation = action
    If modProductionReusableRun.ReusableRunIsLoaded() Then
        action.OutcomeCode = "REJECTED"
        If Not modProductionReusableRun.BeginNextReusableBatch(report) Then GoTo Done
        action.OutcomeCode = "FAILED"
        If Not action.CanContinue(report) Then GoTo Done
        owner.RefreshReusableRunControls True, action
    Else
        If Not modProductionRunBinding.BindWorksheetOwner(operatorBook) Then
            report = "Production worksheet is unavailable.": GoTo Done
        End If
        mProduction.BtnNextBatch completed, report
        If Not completed Then GoTo Done
        If Not action.CanContinue(report) Then GoTo Done
        owner.ResetInventoryCache
        owner.RefreshLoaderState
        If Not action.CanContinue(report) Then GoTo Done
        owner.RefreshManagerState
        report = "Next Batch completed."
    End If
    If Not action.CanContinue(report) Then GoTo Done
    action.OutcomeCode = "STAGED"
Done:
    ' Workbook closure may already have unloaded this exact form.
    If modOperationsFormLifetime.IsLoaded(owner) Then Set owner.RunActionContinuation = Nothing
    loading = priorLoading
    If Not action Is Nothing Then action.Finish report
    busy = priorBusy
    If report <> "" And modOperationsFormLifetime.IsLoaded(owner) Then owner.ShowStatus report
    If number <> 0 Then
        On Error GoTo 0
        Err.Raise number, source, description, helpFile, helpContext
    End If
    Exit Sub
Failed:
    number = Err.Number: source = Err.Source: description = Err.Description
    helpFile = Err.HelpFile: helpContext = Err.HelpContext
    If Not action Is Nothing Then action.OutcomeCode = "FAILED"
    Resume Done
End Sub
