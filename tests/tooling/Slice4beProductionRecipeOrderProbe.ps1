# Disposable adapters invoke the existing Recipe ordering Click handlers.
function Install-ProductionRecipeOrderProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines($form.CountOfDeclarationLines+1,('Private mOrderNestedForTest As Boolean'+[Environment]::NewLine+'Private mOrderNestedEnteredForTest As Boolean'+[Environment]::NewLine+'Private mOrderFailAfterWriteForTest As Boolean'+[Environment]::NewLine+'Private mOrderRestoredForTest As Boolean'))
    $start=$form.ProcStartLine('RenumberRecipeExecutionOrder',0)
    $body=[string]$form.Lines($start,$form.ProcCountLines('RenumberRecipeExecutionOrder',0))
    $lines=$body -split '\r?\n'
    $hits=@(for($i=0;$i -lt $lines.Count;$i++){if($lines[$i].Trim() -ieq 'mLstRecipeNodes.List(i, 4) = CStr(i + 1)'){$i}})
    if($hits.Count -ne 1){throw 'Ordering partial-write anchor missing; not product RED.'}
    $form.InsertLines($start+$hits[0]+1,'        If mOrderFailAfterWriteForTest Then mOrderFailAfterWriteForTest = False: Err.Raise 5432, , "Synthetic ordering interruption"')
    $line=$form.ProcBodyLine('RenumberRecipeExecutionOrder',0)
    $form.InsertLines($line+1,@'
    If mOrderNestedForTest Then
        mOrderNestedForTest = False: mOrderNestedEnteredForTest = True
        Call OrderActForTest("DOWN")
    End If
'@)
    $form.AddFromString(@'
Public Function OrderActForTest(ByVal action As String) As String
    On Error GoTo Failed
    Select Case action
        Case "UP": mBtnRecipeMoveUp_Click
        Case "DOWN": mBtnRecipeMoveDown_Click
        Case "AUTO": mBtnRecipeAutoOrder_Click
        Case Else: Err.Raise 5
    End Select
    OrderActForTest = mTxtStatus.Text
    Exit Function
Failed:
    OrderActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Public Sub OrderStageForTest(ByVal canary As String, ByVal mode As String)
    Dim prior As Boolean, i As Long, selected As Long
    prior = mLoading: mLoading = True: mPages.Value = 1
    mTxtReusableRecipeId.Text = "R1": mTxtReusableRecipeVersion.Text = "1"
    mTxtReusableRecipeName.Text = canary: mTxtReusableRecipeDescription.Text = canary
    mLstRecipeNodes.Clear: mLstRecipeConnections.Clear: mLstRecipeConnectionDisplay.Clear
    mLstProcessInstructions.Clear
    If mode <> "Empty" Then
        For i = 0 To 2
            mLstRecipeNodes.AddItem "N" & CStr(i + 1)
            mLstRecipeNodes.List(i, 1) = "P" & CStr(i + 1)
            mLstRecipeNodes.List(i, 2) = "1"
            mLstRecipeNodes.List(i, 3) = canary & CStr(i + 1)
            mLstRecipeNodes.List(i, 4) = CStr(i + 9)
            mLstProcessInstructions.AddItem CStr(i + 9)
            mLstProcessInstructions.List(i, 1) = canary & "STEP" & CStr(i + 1)
        Next i
        mLstRecipeConnections.AddItem "N3"
        mLstRecipeConnections.List(0, 1) = "OUT"
        mLstRecipeConnections.List(0, 2) = "N1"
        mLstRecipeConnections.List(0, 3) = "REQ"
        mLstRecipeConnections.List(0, 4) = "1"
        mLstRecipeConnections.List(0, 5) = "100"
        mLstRecipeConnections.List(0, 6) = "EA"
        Select Case mode
            Case "Ordered"
                mLstRecipeConnections.List(0, 0) = "N1": mLstRecipeConnections.List(0, 2) = "N3"
            Case "Cycle"
                mLstRecipeConnections.AddItem "N1"
                mLstRecipeConnections.List(1, 1) = "OUT2": mLstRecipeConnections.List(1, 2) = "N3"
                mLstRecipeConnections.List(1, 3) = "REQ2": mLstRecipeConnections.List(1, 4) = "2"
                mLstRecipeConnections.List(1, 5) = "50": mLstRecipeConnections.List(1, 6) = "EA"
            Case "Self"
                mLstRecipeConnections.List(0, 2) = "N3"
            Case "MissingSource": mLstRecipeConnections.List(0, 0) = "UNKNOWN"
            Case "MissingTarget": mLstRecipeConnections.List(0, 2) = "UNKNOWN"
            Case "CaseInsensitive": mLstRecipeConnections.List(0, 0) = "n3"
        End Select
        selected = 1
        If mode = "NoSelection" Then selected = -1
        If mode = "Top" Then selected = 0
        If mode = "Bottom" Then selected = 2
        mLstRecipeNodes.ListIndex = selected
    End If
    mLoading = prior
End Sub
Private Function OrderListTextForTest(ByVal rows As MSForms.ListBox, ByVal fields As Long) As String
    Dim i As Long, j As Long
    For i = 0 To rows.ListCount - 1
        If i > 0 Then OrderListTextForTest = OrderListTextForTest & vbLf
        For j = 0 To fields - 1
            If j > 0 Then OrderListTextForTest = OrderListTextForTest & vbTab
            OrderListTextForTest = OrderListTextForTest & NzStr(rows.List(i, j))
        Next j
    Next i
End Function
Public Function OrderRowsForTest(ByVal kind As String) As String
    Select Case kind
        Case "Nodes": OrderRowsForTest = OrderListTextForTest(mLstRecipeNodes, 5)
        Case "Connections": OrderRowsForTest = OrderListTextForTest(mLstRecipeConnections, 7)
        Case "Instructions": OrderRowsForTest = OrderListTextForTest(mLstProcessInstructions, 2)
    End Select
End Function
Public Function OrderAuxForTest() As String
    OrderAuxForTest = CStr(mLstRecipeNodes.ListIndex) & "|" & CStr(mCmbConnectionFromNode.ListCount)
End Function
Public Function OrderStateForTest() As String
    OrderStateForTest = OrderRowsForTest("Nodes") & vbCr & OrderRowsForTest("Connections") & vbCr & _
        OrderRowsForTest("Instructions") & vbCr & OrderAuxForTest() & vbCr & _
        mTxtReusableRecipeId.Text & "|" & mTxtReusableRecipeVersion.Text & "|" & mTxtReusableRecipeName.Text
End Function
Public Function OrderGuardForTest(ByVal action As String, ByVal guard As String) As String
    Dim prior As Boolean, original As MSForms.ListBox
    Select Case guard
        Case "Loading"
            prior = mLoading: mLoading = True
            OrderGuardForTest = OrderActForTest(action): mLoading = prior
        Case "Busy"
            prior = mDesignerActionInProgress: mDesignerActionInProgress = True
            OrderGuardForTest = OrderActForTest(action): mDesignerActionInProgress = prior
        Case "Failure"
            Set original = mLstRecipeNodes: Set mLstRecipeNodes = Nothing
            OrderGuardForTest = OrderActForTest(action): Set mLstRecipeNodes = original
        Case "PartialFailure"
            mOrderFailAfterWriteForTest = True
            OrderGuardForTest = OrderActForTest(action)
            mOrderRestoredForTest = Not mLoading And Not mDesignerActionInProgress
            mOrderFailAfterWriteForTest = False
        Case "Nested"
            mOrderNestedEnteredForTest = False: mOrderNestedForTest = True
            Call OrderActForTest(action): mOrderNestedForTest = False
            OrderGuardForTest = CStr(mOrderNestedEnteredForTest)
    End Select
End Function
Public Function OrderRestoredForTest() As Boolean
    OrderRestoredForTest = mOrderRestoredForTest
End Function
Public Function OrderCaptionForTest(ByVal action As String) As String
    Select Case action
        Case "UP": OrderCaptionForTest = mBtnRecipeMoveUp.Caption
        Case "DOWN": OrderCaptionForTest = mBtnRecipeMoveDown.Caption
        Case "AUTO": OrderCaptionForTest = mBtnRecipeAutoOrder.Caption
    End Select
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function OrderAct(ByVal action As String) As String
    OrderAct = mForm.OrderActForTest(action)
End Function
Public Sub OrderStage(ByVal canary As String, ByVal mode As String)
    mForm.OrderStageForTest canary, mode
End Sub
Public Function OrderRows(ByVal kind As String) As String
    OrderRows = mForm.OrderRowsForTest(kind)
End Function
Public Function OrderState() As String
    OrderState = mForm.OrderStateForTest()
End Function
Public Function OrderAux() As String
    OrderAux = mForm.OrderAuxForTest()
End Function
Public Function OrderGuard(ByVal action As String, ByVal guard As String) As String
    OrderGuard = mForm.OrderGuardForTest(action, guard)
End Function
Public Function OrderRestored() As Boolean
    OrderRestored = mForm.OrderRestoredForTest()
End Function
Public Function OrderCaption(ByVal action As String) As String
    OrderCaption = mForm.OrderCaptionForTest(action)
End Function
Public Sub OrderShow()
    mForm.Show vbModeless
End Sub
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function OrderPolicyForTest(ByVal enabled As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(enabled))
    OrderPolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function OrderTerminalForTest(ByVal recordJson As String) As Boolean
    Dim record As Object
    Set record = modTrainingJson.DecodeObject(recordJson)
    OrderTerminalForTest = modEvaluationMatches.CommandCompleted(record)
End Function
'@)
}
