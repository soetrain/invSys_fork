Attribute VB_Name = "modReceivingActivityAction"
Option Explicit

' Called only by the real form action. Direct posting is not a user-control event.
Public Function ConfirmWrites(ByVal operatorWb As Workbook, ByVal context As String, _
                              ByVal trackReceipts As Boolean, ByRef report As String) As Boolean
    Dim activityId As String, notice As String, outcome As String
    Dim eventIds As String, submissionState As String, references As String
    On Error GoTo Failed
    report = "Session or warehouse changed. Reopen Receiving before confirming."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    If trackReceipts Then
        activityId = modActivity.BeginAction("RECEIVING_CONFIRM_WRITES", context, notice)
    Else
        activityId = modActivity.BeginAction("DISPOSITION_CONFIRM", context, notice)
    End If
    ConfirmWrites = modReceivingPostingService.ExecuteConfirmWrites( _
        operatorWb, report, outcome, eventIds, submissionState)
Done:
    On Error GoTo ObservationFailed
    If activityId <> "" Then
        references = modActivity.InventorySourceReferences(activityId, eventIds, submissionState)
        modActivity.FinishAction activityId, outcome, notice, references
    End If
    If notice <> "" Then report = report & " " & notice
    Exit Function
ObservationFailed:
    report = report & " Tracking unavailable: the result could not be recorded."
    Exit Function
Failed:
    ConfirmWrites = False: outcome = "FAILED"
    report = "Receiving confirmation failed. Verify current staging before retrying."
    Resume Done
End Function

' Shared Receiving/Returns buttons; only existing local owners perform writes.
Public Function LocalAction(ByVal operatorWb As Workbook, ByVal context As String, _
                            ByVal clearStaging As Boolean, ByRef report As String) As Boolean
    Dim activityId As String, notice As String, outcome As String, controlId As String, changed As Boolean
    On Error GoTo Failed
    report = "Session or warehouse changed. Reopen Receiving before continuing."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    report = "Receiving workbook is no longer open. Reopen Receiving before continuing."
    If operatorWb Is Nothing Then Exit Function
    controlId = "RECEIVING_REFRESH"
    If clearStaging Then controlId = "RECEIVING_CLEAR"
    activityId = modActivity.BeginAction(controlId, context, notice)
    outcome = "FAILED"
    If clearStaging Then
        modTS_Received.ClearReceivingFormStagingForWorkbook operatorWb, changed
        outcome = "EMPTY"
        If changed Then outcome = "CLEARED"
        report = "Receiving form staging cleared."
        LocalAction = True
    Else
        report = ""
        LocalAction = modTS_Received.RefreshReceivingUiForWorkbook(operatorWb, "LOCAL", report)
        If LocalAction Then
            outcome = "REFRESHED"
            report = "Receiving history, managed items, and staging refreshed."
        ElseIf report = "" Then
            report = "Receiving refresh did not complete. Verify the captured workbook before retrying."
        End If
    End If
Done:
    modReceivingAddInput.Finish activityId, outcome, notice, report
    Exit Function
Failed:
    LocalAction = False: outcome = "FAILED"
    If clearStaging Then report = "Clear failed: " & Err.Description Else report = "Refresh failed: " & Err.Description
    Resume Done
End Function
