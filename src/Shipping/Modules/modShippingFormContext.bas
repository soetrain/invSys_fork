Attribute VB_Name = "modShippingFormContext"
Option Explicit

' This local UI binding is independent of optional collection or its storage.
Public Function IsCurrent(ByVal operatorWb As Workbook, ByVal context As String, _
                          Optional ByRef report As String = "") As Boolean
    Dim openBook As Workbook
    On Error GoTo Unavailable
    report = "Session or warehouse changed. Reopen Shipping before continuing."
    If context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    report = "The captured Shipping operator workbook is no longer available. Reopen Shipping."
    If operatorWb Is Nothing Then Exit Function
    For Each openBook In Application.Workbooks
        If openBook Is operatorWb Then
            If openBook.IsAddin Then Exit Function
            IsCurrent = True
            report = ""
            Exit Function
        End If
    Next openBook
Unavailable:
End Function

Public Function CanAct(ByVal operatorWb As Workbook, ByVal context As String, _
                       ByRef report As String, Optional ByVal requireCapability As Boolean = True) As Boolean
    Dim allowed As Boolean
    On Error GoTo Denied
    If Not IsCurrent(operatorWb, context, report) Then Exit Function
    If requireCapability Then
        allowed = modRoleUiAccess.CanCurrentUserPerformCapability("SHIP_POST")
        If Not IsCurrent(operatorWb, context, report) Then Exit Function
        If Not allowed Then GoTo Denied
    End If
    CanAct = True
    Exit Function
Denied:
    If Not IsCurrent(operatorWb, context, report) Then Exit Function
    report = "Shipping permission could not be verified. Review Shipping access before continuing."
End Function

Public Function CanReuse(ByVal form As frmShipmentsTally) As Boolean
    If form Is Nothing Then Exit Function
    CanReuse = form.HasCurrentContext()
End Function

' Explicit relaunch creates a new binding; it never renews an old form's session.
Public Function CreateBoundForm(ByVal operatorWb As Workbook, _
                                ByVal previous As frmShipmentsTally, _
                                ByVal preserveActiveRows As Boolean) As frmShipmentsTally
    Dim fresh As frmShipmentsTally
    If Not previous Is Nothing Then Unload previous
    Set fresh = New frmShipmentsTally
    fresh.SetOperatorWorkbook operatorWb
    fresh.InitializeFromShipping preserveActiveRows
    Set CreateBoundForm = fresh
End Function
