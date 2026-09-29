# Disposable adapters call the real five form handlers, without replacing them.
function Install-ProductionInstructionProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines($form.CountOfDeclarationLines+1,"Private mInstructionNestedForTest As Boolean`r`nPrivate mInstructionNestedEnteredForTest As Boolean")
    $line=$form.ProcBodyLine('mLstProcessInstructions_Click',0)
    $form.InsertLines($line+1,@'
    If mInstructionNestedForTest Then
        mInstructionNestedForTest = False: mInstructionNestedEnteredForTest = True
        Call InstructionActForTest("ADD")
    End If
'@)
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function InstructionActForTest(ByVal action As String) As String
    On Error GoTo Failed
    Select Case action
        Case "ADD": mBtnProcessInstructionAdd_Click
        Case "UPDATE": mBtnProcessInstructionUpdate_Click
        Case "REMOVE": mBtnProcessInstructionRemove_Click
        Case "UP": mBtnProcessInstructionUp_Click
        Case "DOWN": mBtnProcessInstructionDown_Click
        Case Else: Err.Raise 5
    End Select
    InstructionActForTest = mTxtStatus.Text
    Exit Function
Failed:
    InstructionActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Public Sub InstructionStageForTest(ByVal canary As String, ByVal selected As Long, ByVal value As String)
    Dim prior As Boolean, i As Long
    prior = mLoading: mLoading = True
    mLstProcessInstructions.Clear
    For i = 1 To 3
        mLstProcessInstructions.AddItem CStr(i)
        mLstProcessInstructions.List(i - 1, 1) = canary & CStr(i)
    Next i
    mLstProcessInstructions.ListIndex = selected
    mTxtProcessInstruction.Text = value
    mLoading = prior
End Sub
Public Function InstructionRowsForTest() As String
    Dim i As Long
    For i = 0 To mLstProcessInstructions.ListCount - 1
        If i > 0 Then InstructionRowsForTest = InstructionRowsForTest & vbLf
        InstructionRowsForTest = InstructionRowsForTest & CStr(mLstProcessInstructions.List(i, 0)) & vbTab & CStr(mLstProcessInstructions.List(i, 1))
    Next i
End Function
Public Function InstructionEditorForTest() As String
    InstructionEditorForTest = mTxtProcessInstruction.Text
End Function
Public Function InstructionSelectionForTest() As Long
    InstructionSelectionForTest = mLstProcessInstructions.ListIndex
End Function
Public Function InstructionCaptionForTest(ByVal action As String) As String
    Select Case action
        Case "ADD": InstructionCaptionForTest = mBtnProcessInstructionAdd.Caption
        Case "UPDATE": InstructionCaptionForTest = mBtnProcessInstructionUpdate.Caption
        Case "REMOVE": InstructionCaptionForTest = mBtnProcessInstructionRemove.Caption
        Case "UP": InstructionCaptionForTest = mBtnProcessInstructionUp.Caption
        Case "DOWN": InstructionCaptionForTest = mBtnProcessInstructionDown.Caption
    End Select
End Function
Public Function InstructionFailureForTest(ByVal action As String) As String
    Dim original As MSForms.ListBox
    Set original = mLstProcessInstructions
    Set mLstProcessInstructions = Nothing
    InstructionFailureForTest = InstructionActForTest(action)
    Set mLstProcessInstructions = original
End Function
Public Function InstructionBusyForTest(ByVal action As String) As String
    Dim prior As Boolean
    prior = mDesignerActionInProgress: mDesignerActionInProgress = True
    InstructionBusyForTest = InstructionActForTest(action)
    mDesignerActionInProgress = prior
End Function
Public Function InstructionLoadingForTest(ByVal action As String) As String
    Dim prior As Boolean
    prior = mLoading: mLoading = True
    InstructionLoadingForTest = InstructionActForTest(action)
    mLoading = prior
End Function
Public Function InstructionNestedForTest(ByVal action As String) As Boolean
    mInstructionNestedEnteredForTest = False: mInstructionNestedForTest = True
    Call InstructionActForTest(action)
    mInstructionNestedForTest = False
    InstructionNestedForTest = mInstructionNestedEnteredForTest
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function InstructionAct(ByVal action As String) As String
    InstructionAct = mForm.InstructionActForTest(action)
End Function
Public Sub InstructionStage(ByVal canary As String, ByVal selected As Long, ByVal value As String)
    mForm.InstructionStageForTest canary, selected, value
End Sub
Public Function InstructionRows() As String
    InstructionRows = mForm.InstructionRowsForTest()
End Function
Public Function InstructionEditor() As String
    InstructionEditor = mForm.InstructionEditorForTest()
End Function
Public Function InstructionSelection() As Long
    InstructionSelection = mForm.InstructionSelectionForTest()
End Function
Public Function InstructionCaption(ByVal action As String) As String
    InstructionCaption = mForm.InstructionCaptionForTest(action)
End Function
Public Function InstructionFailure(ByVal action As String) As String
    InstructionFailure = mForm.InstructionFailureForTest(action)
End Function
Public Function InstructionBusy(ByVal action As String) As String
    InstructionBusy = mForm.InstructionBusyForTest(action)
End Function
Public Function InstructionLoading(ByVal action As String) As String
    InstructionLoading = mForm.InstructionLoadingForTest(action)
End Function
Public Function InstructionNested(ByVal action As String) As Boolean
    InstructionNested = mForm.InstructionNestedForTest(action)
End Function
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function InstructionPolicyForTest(ByVal enabled As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(enabled))
    InstructionPolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function InstructionTerminalForTest(ByVal recordJson As String) As Boolean
    Dim record As Object
    Set record = modTrainingJson.DecodeObject(recordJson)
    InstructionTerminalForTest = modEvaluationMatches.CommandCompleted(record)
End Function
'@)
}
