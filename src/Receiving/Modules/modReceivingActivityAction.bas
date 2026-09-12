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

' The existing form handlers select their own registered control identity.
Public Function ConfirmWrites(ByVal operatorWb As Workbook, ByVal context As String, _
                              ByVal trackReceipts As Boolean, ByRef report As String) As Boolean
    Dim controlId As String, notice As String
    controlId = "DISPOSITION_CONFIRM"
    If trackReceipts Then controlId = "RECEIVING_CONFIRM_WRITES"
    ConfirmWrites = ConfirmAction(operatorWb, context, controlId, report, notice)
End Function

' Capture comes from the native entry; never redirect to another active workbook.
Public Function ConfirmWorksheetEntry(ByVal operatorWb As Workbook, ByVal callerSheet As Object, _
                                      ByVal caller As Variant, ByRef report As String, ByRef notice As String) As Boolean
    Dim nativeAction As Boolean, context As String, resolvedName As String
    notice = ""
    report = "Activate a Receiving operator workbook before confirming writes."
    If operatorWb Is Nothing Then Exit Function
    If VarType(caller) = vbString Then
        If CStr(caller) = "btnConfirmWrites" And TypeOf callerSheet Is Worksheet Then
            If callerSheet.Name = "ReceivedTally" Then nativeAction = (callerSheet.Parent Is operatorWb)
        End If
    End If
    If Not nativeAction Then
        ConfirmWorksheetEntry = modReceivingPostingService.ExecuteConfirmWrites(operatorWb, report)
        Exit Function
    End If
    context = modActivity.CaptureContext()
    report = "Session or warehouse changed. Sign in and confirm from the intended Receiving worksheet."
    If context = "" Then Exit Function
    If Not modOperationsPrimitiveBridge.ResolveEligibleRoleOperatorWorkbookName( _
        operatorWb.Name, "RECEIVING", resolvedName, report) Then Exit Function
    If resolvedName <> operatorWb.Name Then
        report = "The captured worksheet is not an eligible Receiving operator workbook."
        Exit Function
    End If
    ConfirmWorksheetEntry = ConfirmAction(operatorWb, context, "RECEIVING_WORKSHEET_CONFIRM", report, notice)
End Function

Private Function ConfirmAction(ByVal operatorWb As Workbook, ByVal context As String, _
                               ByVal controlId As String, ByRef report As String, ByRef notice As String) As Boolean
    Dim activityId As String, outcome As String
    Dim eventIds As String, submissionState As String, references As String
    On Error GoTo Failed
    report = "Session or warehouse changed. Reopen Receiving before confirming."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    activityId = modActivity.BeginAction(controlId, context, notice)
    If controlId = "RECEIVING_WORKSHEET_CONFIRM" And context <> modActivity.CaptureContext() Then
        outcome = "REJECTED"
        GoTo Done
    End If
    ConfirmAction = modReceivingPostingService.ExecuteConfirmWrites( _
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
    notice = "Tracking unavailable: the result could not be recorded."
    report = report & " " & notice
    Exit Function
Failed:
    ConfirmAction = False: outcome = "FAILED"
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
