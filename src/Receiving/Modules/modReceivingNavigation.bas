Attribute VB_Name = "modReceivingNavigation"
Option Explicit

Public Function CreateInputs(ByVal form As frmReceiving) As Collection
    Dim inputs As New Collection, name As Variant, inputState As cReceivingSelectionInput
    For Each name In Array("tabsReceiving", "lstReceiveItems", "lstAggregate", "lstInventory", "lstStaged", "cboCondition", "cboDisposition")
        Set inputState = New cReceivingSelectionInput
        inputState.Attach form.Controls(CStr(name))
        inputs.Add inputState, CStr(name)
    Next name
    Set CreateInputs = inputs
End Function

Public Sub Detach(ByRef inputs As Collection)
    Dim inputState As cReceivingSelectionInput
    If inputs Is Nothing Then Exit Sub
    For Each inputState In inputs
        inputState.Detach
    Next inputState
    Set inputs = Nothing
End Sub

Public Function BeginSelection(ByVal form As frmReceiving, ByVal inputs As Collection, ByVal name As String, _
                               ByVal wb As Workbook, ByVal context As String, ByRef activityId As String, _
                               ByRef notice As String, ByRef report As String) As Boolean
    Dim inputState As cReceivingSelectionInput, controlId As String
    BeginSelection = True
    If inputs Is Nothing Then Exit Function
    Set inputState = inputs(name)
    If Not inputState.Consume() Then Exit Function
    BeginSelection = False
    report = "Session or warehouse changed. Reopen Receiving before selecting."
    If wb Is Nothing Or context = "" Or context <> modActivity.CaptureContext() Then Exit Function
    report = ""
    controlId = ControlIdFor(name, CLng(form.Controls("tabsReceiving").Value))
    If controlId = "" Then Exit Function
    activityId = modActivity.BeginAction(controlId, context, notice)
    BeginSelection = True
End Function

Private Function ControlIdFor(ByVal name As String, ByVal pageIndex As Long) As String
    Dim prefix As String
    If name = "tabsReceiving" Then
        Select Case pageIndex
            Case 0: ControlIdFor = "RECEIVING_PAGE_RECEIPTS"
            Case 1: ControlIdFor = "RECEIVING_PAGE_RETURNS"
            Case 2: ControlIdFor = "RECEIVING_PAGE_PURCHASING"
        End Select
        Exit Function
    End If
    If pageIndex <> 0 And pageIndex <> 1 Then Exit Function
    prefix = "RECEIVING_SELECT_"
    If pageIndex = 1 Then prefix = "DISPOSITION_SELECT_"
    Select Case name
        Case "lstReceiveItems": ControlIdFor = prefix & "ITEM"
        Case "lstAggregate": ControlIdFor = prefix & "AGGREGATE"
        Case "lstInventory": ControlIdFor = prefix & "HISTORY"
        Case "lstStaged": ControlIdFor = prefix & "STAGED"
        Case "cboCondition": If pageIndex = 0 Then ControlIdFor = prefix & "CONDITION"
        Case "cboDisposition": If pageIndex = 1 Then ControlIdFor = prefix & "KIND"
    End Select
End Function

' Existing tab presentation remains local to the Receiving form.
Public Function ApplyTab(ByVal form As frmReceiving, ByVal pageIndex As Long) As String
    Dim control As Object, name As Variant, showReceiving As Boolean, showReturns As Boolean, showOperational As Boolean
    showReceiving = (pageIndex = 0): showReturns = (pageIndex = 1)
    showOperational = showReceiving Or showReturns
    For Each control In form.Controls
        Select Case CStr(control.Name)
            Case "tabsReceiving", "btnClose", "txtStatus": control.Visible = True
            Case "lblPurchasingStub": control.Visible = Not showOperational
            Case "lblReturnReason", "txtReturnReason", "lblDisposition", "cboDisposition": control.Visible = showReturns
            Case Else: control.Visible = showOperational
        End Select
    Next control
    If showReceiving Or showReturns Then
        SetLabels form, showReturns
        For Each name In Array("txtReceiveLocation", "txtLotNumber", "cboCondition")
            Set control = form.Controls(CStr(name))
            control.Locked = showReturns
            If showReturns Then control.BackColor = &HEFEFEF Else control.BackColor = &HFFFFFF
        Next name
        If showReturns Then
            ApplyTab = "Outbound inventory disposition is ready. Choose RETURN or DUMP."
        Else
            ApplyTab = "Receiving is ready."
        End If
    Else
        ApplyTab = "Purchasing is not yet operational."
    End If
End Function

Private Sub SetLabels(ByVal form As frmReceiving, ByVal returns As Boolean)
    Dim names As Variant, captions As Variant, index As Long
    names = Array("lblRef", "lblItemSearch", "lblReceiveItemsTitle", "btnAdd", "btnConfirm", _
                  "lblInventoryTitle", "lblStagedTitle", "lblAggregateTitle", "lblReceiveLocation", "lblReturnReason")
    If returns Then
        captions = Array("Disposition Ref", "Return item search", "Return Item Results", "Add Disposition", "Confirm Dispositions", _
                         "Return Entries History", "Return Tally", "Aggregate Returns", "Source location", "Disposition reason *")
    Else
        captions = Array("PO/BOL Ref", "Receive item search", "Receive Item Results", "Add Selected", "Confirm Writes", _
                         "Receiving Entries History", "Received Tally", "Aggregate Received", "Receive location *", "Return reason *")
    End If
    For index = LBound(names) To UBound(names)
        form.Controls(names(index)).Caption = captions(index)
    Next index
End Sub
