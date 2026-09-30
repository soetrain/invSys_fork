Attribute VB_Name = "modProductionCloseActions"
Option Explicit
Option Private Module

' Internal unload and workbook shutdown never call this user-action boundary.
Public Sub CloseForm(ByVal form As frmProduction, ByVal operatorWorkbook As Workbook, _
                     ByVal capturedContext As String, Optional ByVal nativeClosing As Boolean = False)
    Dim activityId As String, notice As String, outcome As String
    ' Optional observation must never prevent dismissal or redirect attribution.
    On Error Resume Next
    If modProductionDesignerActions.ContextIsCurrent(capturedContext, operatorWorkbook) Then _
        activityId = modActivity.BeginAction("PRODUCTION_CLOSE", capturedContext, notice)
    On Error GoTo DismissalFailed
    ' Hide commits native dismissal before QueryClose returns to normal teardown.
    If nativeClosing Then form.Hide Else Unload form
    outcome = "CLOSED"
    GoTo Finished
DismissalFailed:
    outcome = "FAILED"
Finished:
    On Error Resume Next
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    If outcome = "FAILED" Then notice = "Production could not be dismissed. Inspect the form before retrying. " & notice
    If notice <> "" Then MsgBox Trim$(notice), vbExclamation, "Production"
End Sub
