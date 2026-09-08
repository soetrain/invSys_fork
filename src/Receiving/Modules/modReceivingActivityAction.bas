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
    If trackReceipts Then activityId = modActivity.BeginAction("RECEIVING_CONFIRM_WRITES", context, notice)
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
