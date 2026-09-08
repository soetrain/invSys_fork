Attribute VB_Name = "modReceivingActivityAction"
Option Explicit

' The real Ribbon entry observes the attempt before its existing Core guard.
Public Function BeginOpen(ByVal userControlAction As Boolean, ByRef activityId As String, _
                          ByRef notice As String) As Boolean
    BeginOpen = True
    If Not userControlAction Then Exit Function
    activityId = modActivity.BeginAction("RECEIVING_OPEN", modActivity.CaptureContext(), notice)
    If modRoleUiAccess.RequireCurrentUserCapabilityCached( _
        "RECEIVE_POST", "Current user does not have RECEIVE_POST for this warehouse/station.") Then Exit Function
    BeginOpen = False
    FinishLifecycle activityId, "DENIED", notice
End Function

Public Sub ShowMessage(ByVal messageText As String, ByVal style As VbMsgBoxStyle)
    If modUiQuiet.QuietUiIsActive() Then
        Debug.Print "invSys Receiving: " & messageText
    Else
        MsgBox messageText, style, "invSys Receiving"
    End If
End Sub

' Completion follows the owner's dismissal; internal Unload is never attributed.
Public Sub CloseForm(ByVal form As frmReceiving, ByVal operatorWb As Workbook, ByVal context As String, _
                     Optional ByVal nativeClosing As Boolean = False)
    Dim activityId As String, notice As String
    activityId = BeginClose(operatorWb, context, notice)
    ' Native QueryClose must return to Windows without destroying its live VBA frame.
    ' Hide commits UI dismissal now; the uncancelled native message completes teardown.
    If nativeClosing Then form.Hide Else Unload form
    FinishLifecycle activityId, "CLOSED", notice
End Sub

Public Function BeginClose(ByVal operatorWb As Workbook, ByVal context As String, ByRef notice As String) As String
    If Not operatorWb Is Nothing And context <> "" And context = modActivity.CaptureContext() Then _
        BeginClose = modActivity.BeginAction("RECEIVING_CLOSE", context, notice)
End Function

Public Sub FinishLifecycle(ByVal activityId As String, ByVal outcome As String, ByVal notice As String)
    Dim report As String
    On Error Resume Next
    modReceivingAddInput.Finish activityId, outcome, notice, report
    If report <> "" Then modTS_Received.ShowReceivingMessage Trim$(report), vbExclamation
    On Error GoTo 0
End Sub

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
    Dim refreshState As String
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
        LocalAction = modTS_Received.RefreshReceivingUiForWorkbook(operatorWb, "LOCAL", report, refreshState)
        If LocalAction Then
            Select Case refreshState
                Case "REFRESHED"
                    outcome = refreshState
                    report = "Receiving history, managed items, and staging refreshed."
                Case "STALE"
                    outcome = refreshState
                Case Else
                    report = report & " Receiving refresh freshness could not be verified. Review the captured workbook."
            End Select
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
