Attribute VB_Name = "modReceivingAddInput"
Option Explicit
Option Private Module

' Existing form validation, separate from staging and from activity publication.
Public Function Valid(ByVal form As frmReceiving, ByRef receiptType As String, _
                      ByRef report As String) As Boolean
    report = "Select a managed item to receive first."
    If form.Controls("lstReceiveItems").ListIndex < 0 Then Exit Function
    report = "Ref number is required."
    If Trim$(CStr(form.Controls("txtRef").Value)) = "" Then Exit Function
    report = "Quantity must be greater than zero."
    If CDbl(Val(CStr(form.Controls("txtQty").Value))) <= 0 Then Exit Function
    report = "Receive location is required."
    If Trim$(CStr(form.Controls("txtReceiveLocation").Value)) = "" Then Exit Function
    report = "Choose the condition of the received goods."
    If form.Controls("cboCondition").ListIndex < 0 Then Exit Function
    receiptType = "RECEIPT"
    If form.Controls("tabsReceiving").Value = 1 Then
        report = "Choose RETURN or DUMP."
        If form.Controls("cboDisposition").ListIndex < 0 Then Exit Function
        receiptType = UCase$(Trim$(CStr(form.Controls("cboDisposition").Value)))
    End If
    report = "Disposition reason is required."
    If receiptType <> "RECEIPT" And Trim$(CStr(form.Controls("txtReturnReason").Value)) = "" Then Exit Function
    report = ""
    Valid = True
End Function

' Optional evidence cannot hide the original staging status or retry the service.
Public Sub Finish(ByVal activityId As String, ByVal outcome As String, _
                  ByVal notice As String, ByRef report As String)
    On Error GoTo Failed
    If activityId <> "" Then modActivity.FinishAction activityId, outcome, notice
    If notice <> "" Then report = report & " " & notice
    Exit Sub
Failed:
    report = report & " Tracking unavailable: the result could not be recorded."
End Sub
