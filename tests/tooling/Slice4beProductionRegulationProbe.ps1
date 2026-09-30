# Disposable, unsaved adapters call the two operator Click handlers unchanged.
function Install-ProductionRegulationProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines($form.CountOfDeclarationLines+1,@'
Private mRegPartialForTest As Boolean
Private mRegNestedForTest As Boolean
Private mRegNestedEnteredForTest As Boolean
Private mRegUnavailableForTest As Boolean
Private mRegReadCountForTest As Long
Private mRegRestoredForTest As Boolean
'@)
    $start=$form.ProcStartLine('RefreshOutputRegulationSettings',0)
    $lines=$form.Lines($start,$form.ProcCountLines('RefreshOutputRegulationSettings',0)) -split '\r?\n'
    $reads=@(for($i=0;$i -lt $lines.Count;$i++){if($lines[$i].Contains('modOperationsPrimitiveBridge.GetProcessVersion(processId, processVersion)')){$start+$i}})
    $boundaries=@(for($i=0;$i -lt $lines.Count;$i++){if($lines[$i].Trim() -ceq 'If mLstOutputRegulations Is Nothing Then Exit Sub'){$start+$i}})
    if($reads.Count -ne 1 -or $boundaries.Count -ne 1){throw 'Regulation fixture boundary missing/ambiguous; not product RED.'}
    $form.ReplaceLine($reads[0],$form.Lines($reads[0],1).Replace('modOperationsPrimitiveBridge.GetProcessVersion(processId, processVersion)','RegulationReadForTest(processId, processVersion)'))
    $form.InsertLines($boundaries[0],'    RegulationBoundaryForTest')
    $form.AddFromString(@'
Private Function RegulationReadForTest(ByVal identity As String, ByVal version As String) As String
    mRegReadCountForTest = mRegReadCountForTest + 1
    If Not mRegUnavailableForTest Then RegulationReadForTest = modOperationsPrimitiveBridge.GetProcessVersion(identity, version)
End Function
Private Sub RegulationBoundaryForTest()
    If mRegPartialForTest Then Err.Raise 5432, , "Synthetic regulation interruption"
    If mRegNestedForTest Then
        mRegNestedForTest = False: mRegNestedEnteredForTest = True
        mLoading = True
        mLstOutputRegulations.ListIndex = 0
        mLoading = False
        Call RegulationActForTest("CLEAR")
    End If
End Sub
Public Function RegulationActForTest(ByVal action As String) As String
    On Error GoTo Failed
    If action = "APPLY" Then
        mBtnOutputRegulationApply_Click
    ElseIf action = "CLEAR" Then
        mBtnOutputRegulationClear_Click
    Else
        Err.Raise 5
    End If
    RegulationActForTest = mTxtStatus.Text
    Exit Function
Failed:
    RegulationActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Private Sub RegulationRecordForTest(ByVal node As String, ByVal output As String, ByVal identity As String, ByVal version As String, ByVal floor As String, ByVal ceiling As String)
    Dim record As Object
    Set record = NewReusableRecord("OUTPUT_REGULATION")
    record("ProcessNodeId") = node: record("ProcessId") = identity
    record("ProcessVersion") = version: record("OutputId") = output
    record("OutputRegulationEnabled") = True
    AddNumericReusableField record, "OutputFloorQty", floor
    AddNumericReusableField record, "OutputCeilingQty", ceiling
    mRecipeOutputRegulations.Add record
End Sub
Public Sub RegulationStageForTest(ByVal scope As String, ByVal mode As String, ByVal canary As String, ByVal identity As String, ByVal version As String)
    Dim prior As Boolean, i As Long
    prior = mLoading: mLoading = True
    mRegPartialForTest = False: mRegNestedForTest = False: mRegUnavailableForTest = False
    mRegNestedEnteredForTest = False: mRegReadCountForTest = 0
    mPages.Value = 5
    mTxtProcessId.Text = identity: mTxtProcessVersion.Text = version
    mTxtProcessName.Text = canary: mTxtProcessDescription.Text = canary
    mTxtReusableRecipeId.Text = "LOCAL-R": mTxtReusableRecipeVersion.Text = "99"
    mTxtReusableRecipeName.Text = canary: mTxtReusableRecipeDescription.Text = canary
    mCmbOutputRegulationScope.ListIndex = IIf(scope = "Recipe", 1, 0)
    mLstProcessOutputs.Clear: mLstRecipeNodes.Clear: mLstOutputRegulations.Clear
    Set mProcessOutputRegulations = CreateObject("Scripting.Dictionary")
    mProcessOutputRegulations.CompareMode = vbTextCompare
    Set mRecipeOutputRegulations = New Collection
    For i = 0 To 1
        mLstProcessOutputs.AddItem IIf(i = 0, "A01", "B02")
        mLstProcessOutputs.List(i, 1) = canary
        mLstProcessOutputs.List(i, 2) = "RECIPE-FIXTURE"
        mLstProcessOutputs.List(i, 3) = "LOCAL-DESIGN"
        mLstProcessOutputs.List(i, 4) = "1"
        mLstProcessOutputs.List(i, 5) = "1"
        mLstProcessOutputs.List(i, 8) = "EA"
        mLstProcessOutputs.List(i, 9) = "FIXED"
        mLstRecipeNodes.AddItem "NODE" & CStr(i + 1)
        mLstRecipeNodes.List(i, 1) = identity: mLstRecipeNodes.List(i, 2) = version
        mLstRecipeNodes.List(i, 3) = canary: mLstRecipeNodes.List(i, 4) = CStr(i + 1)
    Next i
    SetProcessOutputRegulation "A01", True, "1", "4"
    SetProcessOutputRegulation "B02", True, "5", "9"
    RegulationRecordForTest "NODE1", "A01", identity, version, "1", "4"
    RegulationRecordForTest "NODE2", "A01", identity, version, "5", "9"
    RegulationRecordForTest "NODE1", "B02", identity, version, "6", "10"
    If mode = "NoExisting" Then
        mProcessOutputRegulations.Remove "A01"
        RemoveRecipeOutputRegulation "NODE1", "A01"
    End If
    mLstOutputRegulations.AddItem "A01"
    mLstOutputRegulations.List(0, 1) = canary: mLstOutputRegulations.List(0, 2) = "EA"
    mLstOutputRegulations.List(0, 3) = "True": mLstOutputRegulations.List(0, 4) = "1"
    mLstOutputRegulations.List(0, 5) = "4"
    mLstOutputRegulations.List(0, 6) = IIf(scope = "Recipe", "NODE1", "0")
    mLstOutputRegulations.List(0, 7) = identity: mLstOutputRegulations.List(0, 8) = version
    mCmbOutputRegulationNode.Clear: mCmbOutputRegulationNode.AddItem "NODE1"
    mCmbOutputRegulationNode.List(0, 1) = identity: mCmbOutputRegulationNode.List(0, 2) = version
    mCmbOutputRegulationNode.List(0, 3) = canary: mCmbOutputRegulationNode.ListIndex = 0
    mLstOutputRegulations.ListIndex = 0
    mChkOutputRegulated.Value = True
    mTxtOutputRegulationFloor.Text = " 2 ": mTxtOutputRegulationCeiling.Text = " 8 "
    Select Case mode
        Case "NoSelection": mLstOutputRegulations.ListIndex = -1
        Case "Negative": mTxtOutputRegulationFloor.Text = "-1"
        Case "Zero": mTxtOutputRegulationFloor.Text = "0"
        Case "Reversed": mTxtOutputRegulationFloor.Text = "9"
        Case "Equal": mTxtOutputRegulationFloor.Text = "8"
        Case "FractionEA": mTxtOutputRegulationFloor.Text = "1.5"
        Case "LowercaseEA": mLstOutputRegulations.List(0, 2) = "ea": mTxtOutputRegulationCeiling.Text = "2.5"
        Case "FractionLB": mLstOutputRegulations.List(0, 2) = "LB": mTxtOutputRegulationFloor.Text = "1.5": mTxtOutputRegulationCeiling.Text = "2.5"
        Case "Blank": mTxtOutputRegulationFloor.Text = ""
        Case "NonNumeric": mTxtOutputRegulationFloor.Text = canary
        Case "DisabledBlank": mChkOutputRegulated.Value = False: mTxtOutputRegulationFloor.Text = "": mTxtOutputRegulationCeiling.Text = ""
        Case "DisabledText": mChkOutputRegulated.Value = False: mTxtOutputRegulationFloor.Text = canary: mTxtOutputRegulationCeiling.Text = canary
        Case "DisabledNegative": mChkOutputRegulated.Value = False: mTxtOutputRegulationFloor.Text = "-2": mTxtOutputRegulationCeiling.Text = "-1"
        Case "ReadUnavailable": mRegUnavailableForTest = True
    End Select
    mTxtStatus.Text = "Fixture prepared"
    mLoading = prior
End Sub
Public Function RegulationValueForTest(ByVal scope As String, ByVal node As String, ByVal output As String) As String
    Dim record As Object
    If scope = "Process" Then
        RegulationValueForTest = ProcessOutputRegulationValue(output, 0) & "|" & ProcessOutputRegulationValue(output, 1) & "|" & ProcessOutputRegulationValue(output, 2)
    Else
        Set record = FindRecipeOutputRegulation(node, output)
        If record Is Nothing Then RegulationValueForTest = "NONE": Exit Function
        RegulationValueForTest = modProductionReusableDesigns.ReusableRecordText(record, "OutputRegulationEnabled") & "|" & _
            modProductionReusableDesigns.ReusableRecordText(record, "OutputFloorQty") & "|" & _
            modProductionReusableDesigns.ReusableRecordText(record, "OutputCeilingQty")
    End If
End Function
Public Function RegulationBindingForTest() As String
    Dim record As Object
    Set record = FindRecipeOutputRegulation("NODE1", "A01")
    If record Is Nothing Then Exit Function
    RegulationBindingForTest = modProductionReusableDesigns.ReusableRecordText(record, "ProcessId") & "|" & _
        modProductionReusableDesigns.ReusableRecordText(record, "ProcessVersion") & "|" & _
        modProductionReusableDesigns.ReusableRecordText(record, "ProcessNodeId") & "|" & _
        modProductionReusableDesigns.ReusableRecordText(record, "OutputId")
End Function
Public Function RegulationStateForTest() As String
    RegulationStateForTest = RegulationValueForTest("Process", "", "A01") & vbCr & _
        RegulationValueForTest("Process", "", "B02") & vbCr & RegulationValueForTest("Recipe", "NODE1", "A01") & vbCr & _
        RegulationValueForTest("Recipe", "NODE2", "A01") & vbCr & RegulationValueForTest("Recipe", "NODE1", "B02")
End Function
Public Function RegulationAuxForTest() As String
    RegulationAuxForTest = CStr(mLstOutputRegulations.ListCount) & "|" & CStr(mRegReadCountForTest) & "|" & _
        CStr(mChkOutputRegulated.Value) & "|" & mTxtOutputRegulationFloor.Text & "|" & mTxtOutputRegulationCeiling.Text
End Function
Public Function RegulationGuardForTest(ByVal action As String, ByVal guard As String) As String
    Dim prior As Boolean, original As MSForms.ListBox
    Select Case guard
        Case "Loading"
            prior = mLoading: mLoading = True
            RegulationGuardForTest = RegulationActForTest(action): mLoading = prior
        Case "Busy"
            prior = mDesignerActionInProgress: mDesignerActionInProgress = True
            RegulationGuardForTest = RegulationActForTest(action): mDesignerActionInProgress = prior
        Case "Failure"
            Set original = mLstOutputRegulations: Set mLstOutputRegulations = Nothing
            RegulationGuardForTest = RegulationActForTest(action): Set mLstOutputRegulations = original
        Case "PartialFailure"
            mRegPartialForTest = True
            RegulationGuardForTest = RegulationActForTest(action)
            mRegRestoredForTest = Not mLoading And Not mDesignerActionInProgress
            mRegPartialForTest = False
        Case "Nested"
            mRegNestedEnteredForTest = False: mRegNestedForTest = True
            Call RegulationActForTest(action): mRegNestedForTest = False
            RegulationGuardForTest = CStr(mRegNestedEnteredForTest)
    End Select
End Function
Public Function RegulationRestoredForTest() As Boolean
    RegulationRestoredForTest = mRegRestoredForTest
End Function
Public Function RegulationCaptionForTest(ByVal action As String) As String
    If action = "APPLY" Then RegulationCaptionForTest = mBtnOutputRegulationApply.Caption Else RegulationCaptionForTest = mBtnOutputRegulationClear.Caption
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function RegulationAct(ByVal action As String) As String
    RegulationAct = mForm.RegulationActForTest(action)
End Function
Public Sub RegulationStage(ByVal scope As String, ByVal mode As String, ByVal canary As String, ByVal identity As String, ByVal version As String)
    mForm.RegulationStageForTest scope, mode, canary, identity, version
End Sub
Public Function RegulationValue(ByVal scope As String, ByVal node As String, ByVal output As String) As String
    RegulationValue = mForm.RegulationValueForTest(scope, node, output)
End Function
Public Function RegulationBinding() As String
    RegulationBinding = mForm.RegulationBindingForTest()
End Function
Public Function RegulationState() As String
    RegulationState = mForm.RegulationStateForTest()
End Function
Public Function RegulationAux() As String
    RegulationAux = mForm.RegulationAuxForTest()
End Function
Public Function RegulationGuard(ByVal action As String, ByVal guard As String) As String
    RegulationGuard = mForm.RegulationGuardForTest(action, guard)
End Function
Public Function RegulationRestored() As Boolean
    RegulationRestored = mForm.RegulationRestoredForTest()
End Function
Public Function RegulationCaption(ByVal action As String) As String
    RegulationCaption = mForm.RegulationCaptionForTest(action)
End Function
Public Sub RegulationShow()
    mForm.Show vbModeless
End Sub
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function RegulationPolicyForTest(ByVal enabled As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(enabled))
    RegulationPolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function RegulationTerminalForTest(ByVal recordJson As String) As Boolean
    Dim record As Object
    Set record = modTrainingJson.DecodeObject(recordJson)
    RegulationTerminalForTest = modEvaluationMatches.CommandCompleted(record)
End Function
'@)
}
