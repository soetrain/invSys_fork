# Disposable adapters enter the existing Recipe structure Click handlers.
function Install-ProductionRecipeStructureProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines($form.CountOfDeclarationLines+1,@'
Private mStructureNestedActionForTest As String
Private mStructureNestedEnteredForTest As Boolean
Private mStructureFailAfterWriteForTest As Boolean
Private mStructureRestoredForTest As Boolean
Private mStructureActionActiveForTest As Boolean
Private mStructureConnectionClicksForTest As Long
'@)
    $line=$form.ProcBodyLine('mLstRecipeConnections_Click',0)
    $form.InsertLines($line+1,'    If mStructureActionActiveForTest Then mStructureConnectionClicksForTest = mStructureConnectionClicksForTest + 1')
    foreach($boundary in @('RefreshConnectionNodeCombos','RefreshRecipeConnectionDisplay')){
        $line=$form.ProcBodyLine($boundary,0)
        $form.InsertLines($line+1,'    StructureBoundaryForTest')
    }
    $form.AddFromString(@'
Private Sub StructureBoundaryForTest()
    Dim action As String
    If mStructureFailAfterWriteForTest Then
        mStructureFailAfterWriteForTest = False
        Err.Raise 5432, , "Synthetic structure interruption"
    End If
    If mStructureNestedActionForTest <> "" Then
        action = mStructureNestedActionForTest: mStructureNestedActionForTest = ""
        mStructureNestedEnteredForTest = True
        Call StructureActForTest(action)
    End If
End Sub
Public Function StructureActForTest(ByVal action As String) As String
    On Error GoTo Failed
    mStructureConnectionClicksForTest = 0: mStructureActionActiveForTest = True
    Select Case action
        Case "ADD_PROCESS": mBtnRecipeAddProcess_Click
        Case "REMOVE_PROCESS": mBtnRecipeRemoveProcess_Click
        Case "CONNECT": mBtnRecipeConnect_Click
        Case "UPDATE_CONNECTION": mBtnRecipeUpdateConnection_Click
        Case "DISCONNECT": mBtnRecipeDisconnect_Click
        Case Else: Err.Raise 5
    End Select
    StructureActForTest = mTxtStatus.Text
    mStructureActionActiveForTest = False
    Exit Function
Failed:
    mStructureActionActiveForTest = False
    StructureActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Public Function StructureEditorForTest() As String
    StructureEditorForTest = ComboText(mCmbConnectionFromNode) & vbTab & ConnectionOutputId() & vbTab & _
        ConnectionTargetNodeId() & vbTab & ConnectionRequirementId() & vbTab & Trim$(mTxtConnectionQty.Text) & vbTab & _
        Trim$(mTxtConnectionPercent.Text) & vbTab & ComboText(mCmbConnectionUom)
End Function
Public Function StructureConnectionClicksForTest() As Long
    StructureConnectionClicksForTest = mStructureConnectionClicksForTest
End Function
Public Sub StructureStageForTest(ByVal canary As String, ByVal action As String, ByVal mode As String)
    Dim prior As Boolean, i As Long, target As String, requirement As String
    prior = mLoading: mLoading = True: mPages.Value = 1
    mTxtReusableRecipeId.Text = "R1": mTxtReusableRecipeVersion.Text = "1"
    mTxtReusableRecipeName.Text = canary: mTxtReusableRecipeDescription.Text = canary
    mLstRecipeNodes.Clear: mLstRecipeConnections.Clear: mLstRecipeConnectionDisplay.Clear
    mLstProcessInstructions.Clear: mLstReleasedProcesses.Clear
    For i = 0 To 2
        mLstRecipeNodes.AddItem "N" & CStr(i + 1)
        mLstRecipeNodes.List(i, 1) = "P" & CStr(i + 1)
        mLstRecipeNodes.List(i, 2) = "1"
        mLstRecipeNodes.List(i, 3) = canary & CStr(i + 1)
        mLstRecipeNodes.List(i, 4) = CStr(i + 9)
        mLstProcessInstructions.AddItem CStr(i + 9)
        mLstProcessInstructions.List(i, 1) = canary & "STEP" & CStr(i + 1)
    Next i
    If mode = "Collision" Then mLstRecipeNodes.List(2, 0) = "n4"
    If mode = "CaseInsensitive" Then mLstRecipeNodes.List(1, 0) = "n2"
    mLstRecipeNodes.ListIndex = 1
    mLstReleasedProcesses.AddItem "P4"
    mLstReleasedProcesses.List(0, 1) = "2": mLstReleasedProcesses.List(0, 2) = canary & "RELEASED"
    mLstReleasedProcesses.ListIndex = 0
    For i = 0 To 2
        mLstRecipeConnections.AddItem "N1"
        mLstRecipeConnections.List(i, 1) = "O" & CStr(i + 1)
        mLstRecipeConnections.List(i, 2) = "N3"
        mLstRecipeConnections.List(i, 3) = "R" & CStr(i + 1)
        mLstRecipeConnections.List(i, 4) = CStr(i + 1)
        mLstRecipeConnections.List(i, 5) = "100"
        mLstRecipeConnections.List(i, 6) = "EA"
    Next i
    mLstRecipeConnections.List(0, 2) = "N2"
    mLstRecipeConnections.List(1, 0) = "N2"
    mLstRecipeConnections.ListIndex = 0
    If action = "UPDATE_CONNECTION" And (mode = "Duplicate" Or mode = "DuplicateCase") Then mLstRecipeConnections.ListIndex = 2
    ' The hidden-list Click refreshes editors and clears mLoading; restore fixture suppression.
    mLoading = True
    mLstRecipeConnectionDisplay.AddItem "Fixture connection"
    mLstRecipeConnectionDisplay.List(0, 7) = "1"
    mLstRecipeConnectionDisplay.ListIndex = 0
    mCmbConnectionFromNode.Clear: mCmbConnectionOutput.Clear
    mCmbConnectionToNode.Clear: mCmbConnectionRequirement.Clear: mCmbConnectionUom.Clear
    mCmbConnectionFromNode.AddItem "N1": mCmbConnectionFromNode.List(0, 1) = canary & "SOURCE"
    mCmbConnectionFromNode.ListIndex = 0
    mCmbConnectionOutput.AddItem "OEDIT": mCmbConnectionOutput.List(0, 1) = canary & "OUTPUT"
    mCmbConnectionOutput.ListIndex = 0
    target = "N3": requirement = "RNEW"
    If action = "UPDATE_CONNECTION" Or mode = "Duplicate" Then target = "N2": requirement = "R1"
    If mode = "Self" Then target = "n1"
    If mode = "DuplicateCase" Then target = "n2": requirement = "r1"
    mCmbConnectionToNode.AddItem target
    mCmbConnectionToNode.List(0, 1) = requirement: mCmbConnectionToNode.List(0, 2) = canary & "TARGET"
    mCmbConnectionToNode.ListIndex = 0
    mCmbConnectionUom.AddItem "EA": mCmbConnectionUom.ListIndex = 0
    mTxtConnectionQty.Text = " 4 ": mTxtConnectionPercent.Text = " 50 "
    Select Case mode
        Case "NoSelection"
            If action = "ADD_PROCESS" Then mLstReleasedProcesses.ListIndex = -1
            If action = "REMOVE_PROCESS" Then mLstRecipeNodes.ListIndex = -1
            If action = "DISCONNECT" Then mLstRecipeConnectionDisplay.ListIndex = -1: mLstRecipeConnections.ListIndex = -1
        Case "NoSource": mCmbConnectionFromNode.ListIndex = -1
        Case "NoOutput": mCmbConnectionOutput.ListIndex = -1
        Case "NoTarget": mCmbConnectionToNode.ListIndex = -1
        Case "NoRequirement": mCmbConnectionToNode.List(0, 1) = ""
        Case "NonPositive": mTxtConnectionQty.Text = "0": mTxtConnectionPercent.Text = "0"
        Case "NoUom": mCmbConnectionUom.ListIndex = -1
        Case "FractionalEA": mTxtConnectionQty.Text = "1.5"
        Case "PercentOnly": mTxtConnectionQty.Text = ""
        Case "OtherText": mTxtConnectionQty.Text = "not numeric"
        Case "NegativeOther": mTxtConnectionQty.Text = "-2"
        Case "AppendFallback": mLstRecipeConnections.ListIndex = -1: mCmbConnectionToNode.List(0, 1) = "RNEW"
        Case "Unchanged"
            mCmbConnectionOutput.List(0, 0) = "O1"
            mTxtConnectionQty.Text = "1": mTxtConnectionPercent.Text = "100"
        Case "HiddenFallback": mLstRecipeConnectionDisplay.List(0, 7) = "-1"
        Case "NoDisplaySelection": mLstRecipeConnectionDisplay.ListIndex = -1
        Case "InvalidDisplay": mLstRecipeConnectionDisplay.List(0, 7) = "99"
    End Select
    mTxtStatus.Text = "Prior local status"
    mLoading = prior
End Sub
Private Function StructureListTextForTest(ByVal rows As MSForms.ListBox, ByVal fields As Long) As String
    Dim i As Long, j As Long
    For i = 0 To rows.ListCount - 1
        If i > 0 Then StructureListTextForTest = StructureListTextForTest & vbLf
        For j = 0 To fields - 1
            If j > 0 Then StructureListTextForTest = StructureListTextForTest & vbTab
            StructureListTextForTest = StructureListTextForTest & NzStr(rows.List(i, j))
        Next j
    Next i
End Function
Public Function StructureRowsForTest(ByVal kind As String) As String
    Select Case kind
        Case "Nodes": StructureRowsForTest = StructureListTextForTest(mLstRecipeNodes, 5)
        Case "Connections": StructureRowsForTest = StructureListTextForTest(mLstRecipeConnections, 7)
        Case "Instructions": StructureRowsForTest = StructureListTextForTest(mLstProcessInstructions, 2)
    End Select
End Function
Public Function StructureAuxForTest() As String
    StructureAuxForTest = CStr(mLstRecipeNodes.ListIndex) & "|" & CStr(mLstRecipeConnections.ListIndex) & "|" & _
        CStr(mCmbConnectionFromNode.ListCount) & "|" & mTxtConnectionQty.Text & "|" & mTxtConnectionPercent.Text
End Function
Public Function StructureStateForTest() As String
    StructureStateForTest = StructureRowsForTest("Nodes") & vbCr & StructureRowsForTest("Connections") & vbCr & _
        StructureRowsForTest("Instructions") & vbCr & StructureAuxForTest() & vbCr & _
        mTxtReusableRecipeId.Text & "|" & mTxtReusableRecipeVersion.Text & "|" & mTxtReusableRecipeName.Text
End Function
Public Function StructureGuardForTest(ByVal action As String, ByVal guard As String) As String
    Dim prior As Boolean, original As MSForms.ListBox
    Select Case guard
        Case "Loading"
            prior = mLoading: mLoading = True
            StructureGuardForTest = StructureActForTest(action): mLoading = prior
        Case "Busy"
            prior = mDesignerActionInProgress: mDesignerActionInProgress = True
            StructureGuardForTest = StructureActForTest(action): mDesignerActionInProgress = prior
        Case "Failure"
            If action = "ADD_PROCESS" Or action = "REMOVE_PROCESS" Then
                Set original = mLstRecipeNodes: Set mLstRecipeNodes = Nothing
                StructureGuardForTest = StructureActForTest(action): Set mLstRecipeNodes = original
            Else
                Set original = mLstRecipeConnections: Set mLstRecipeConnections = Nothing
                StructureGuardForTest = StructureActForTest(action): Set mLstRecipeConnections = original
            End If
        Case "PartialFailure"
            mStructureFailAfterWriteForTest = True
            StructureGuardForTest = StructureActForTest(action)
            mStructureRestoredForTest = Not mLoading And Not mDesignerActionInProgress
            mStructureFailAfterWriteForTest = False
        Case "Nested"
            mStructureNestedEnteredForTest = False: mStructureNestedActionForTest = action
            Call StructureActForTest(action): mStructureNestedActionForTest = ""
            StructureGuardForTest = CStr(mStructureNestedEnteredForTest)
    End Select
End Function
Public Function StructureRestoredForTest() As Boolean
    StructureRestoredForTest = mStructureRestoredForTest
End Function
Public Function StructureCaptionForTest(ByVal action As String) As String
    Select Case action
        Case "ADD_PROCESS": StructureCaptionForTest = mBtnRecipeAddProcess.Caption
        Case "REMOVE_PROCESS": StructureCaptionForTest = mBtnRecipeRemoveProcess.Caption
        Case "CONNECT": StructureCaptionForTest = mBtnRecipeConnect.Caption
        Case "UPDATE_CONNECTION": StructureCaptionForTest = mBtnRecipeUpdateConnection.Caption
        Case "DISCONNECT": StructureCaptionForTest = mBtnRecipeDisconnect.Caption
    End Select
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function StructureAct(ByVal action As String) As String
    StructureAct = mForm.StructureActForTest(action)
End Function
Public Function StructureEditor() As String
    StructureEditor = mForm.StructureEditorForTest()
End Function
Public Function StructureConnectionClicks() As Long
    StructureConnectionClicks = mForm.StructureConnectionClicksForTest()
End Function
Public Sub StructureStage(ByVal canary As String, ByVal action As String, ByVal mode As String)
    mForm.StructureStageForTest canary, action, mode
End Sub
Public Function StructureRows(ByVal kind As String) As String
    StructureRows = mForm.StructureRowsForTest(kind)
End Function
Public Function StructureState() As String
    StructureState = mForm.StructureStateForTest()
End Function
Public Function StructureAux() As String
    StructureAux = mForm.StructureAuxForTest()
End Function
Public Function StructureGuard(ByVal action As String, ByVal guard As String) As String
    StructureGuard = mForm.StructureGuardForTest(action, guard)
End Function
Public Function StructureRestored() As Boolean
    StructureRestored = mForm.StructureRestoredForTest()
End Function
Public Function StructureCaption(ByVal action As String) As String
    StructureCaption = mForm.StructureCaptionForTest(action)
End Function
Public Sub StructureShow()
    mForm.Show vbModeless
End Sub
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function StructurePolicyForTest(ByVal enabled As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(enabled))
    StructurePolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function StructureTerminalForTest(ByVal recordJson As String) As Boolean
    Dim record As Object
    Set record = modTrainingJson.DecodeObject(recordJson)
    StructureTerminalForTest = modEvaluationMatches.CommandCompleted(record)
End Function
'@)
}
