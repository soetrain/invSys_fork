# Disposable test adapters invoke the ten actual packaged Click handlers.
function Install-ProductionComponentProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines($form.CountOfDeclarationLines+1,"Private mComponentNestedForTest As Boolean`r`nPrivate mComponentNestedEnteredForTest As Boolean`r`nPrivate mComponentFailAfterWriteForTest As Boolean`r`nPrivate mComponentFailureRestoredForTest As Boolean")
    foreach($pair in @(@('WriteRequirementEditorToList','mTxtRequirementId'),@('WriteOutputEditorToList','mTxtProcessOutputId'))){
        $start=$form.ProcStartLine($pair[0],0);$body=[string]$form.Lines($start,$form.ProcCountLines($pair[0],0))
        $lines=$body -split '\r?\n';$needle='.List(idx, 0) = Trim$('+$pair[1]+'.Text)'
        # VBE normalizes identifier casing (Text becomes text); VBA is case-insensitive.
        $hits=@(for($i=0;$i -lt $lines.Count;$i++){if($lines[$i].Trim() -ieq $needle){$i}})
        if($hits.Count -ne 1){throw 'Partial-write probe anchor unavailable; not product RED.'}
        $form.InsertLines($start+$hits[0]+1,'        If mComponentFailAfterWriteForTest Then mComponentFailAfterWriteForTest = False: Err.Raise 5432, , "Synthetic component write interruption"')
    }
    foreach($kind in @('Requirements','Outputs')){
        $line=$form.ProcBodyLine(('mLstProcess'+$kind+'_Click'),0)
        $singular=if($kind -ceq 'Requirements'){'REQUIREMENT'}else{'OUTPUT'}
        $form.InsertLines($line+1,(@'
    If mComponentNestedForTest Then
        mComponentNestedForTest = False: mComponentNestedEnteredForTest = True
        Call ComponentActForTest("__KIND__", "ADD")
    End If
'@).Replace('__KIND__',$singular))
    }
    $form.AddFromString(@'
Public Function ComponentActForTest(ByVal kind As String, ByVal action As String) As String
    On Error GoTo Failed
    Select Case kind & "_" & action
        Case "REQUIREMENT_ADD": mBtnProcessRequirementAdd_Click
        Case "REQUIREMENT_UPDATE": mBtnProcessRequirementUpdate_Click
        Case "REQUIREMENT_REMOVE": mBtnProcessRequirementRemove_Click
        Case "REQUIREMENT_UP": mBtnProcessRequirementUp_Click
        Case "REQUIREMENT_DOWN": mBtnProcessRequirementDown_Click
        Case "OUTPUT_ADD": mBtnProcessOutputAdd_Click
        Case "OUTPUT_UPDATE": mBtnProcessOutputUpdate_Click
        Case "OUTPUT_REMOVE": mBtnProcessOutputRemove_Click
        Case "OUTPUT_UP": mBtnProcessOutputUp_Click
        Case "OUTPUT_DOWN": mBtnProcessOutputDown_Click
        Case Else: Err.Raise 5
    End Select
    ComponentActForTest = mTxtStatus.Text
    Exit Function
Failed:
    ComponentActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Private Function ComponentListForTest(ByVal kind As String) As MSForms.ListBox
    If kind = "REQUIREMENT" Then Set ComponentListForTest = mLstProcessRequirements Else Set ComponentListForTest = mLstProcessOutputs
End Function
Public Sub ComponentStageForTest(ByVal kind As String, ByVal canary As String, ByVal mode As String)
    Dim prior As Boolean, i As Long, selected As Long
    prior = mLoading: mLoading = True
    mTxtProcessId.Text = "P1": mTxtProcessVersion.Text = "1"
    mLstProcessRequirements.Clear: mLstProcessOutputs.Clear: mLstProcessInstructions.Clear
    Set mProcessOutputRegulations = CreateObject("Scripting.Dictionary")
    mProcessOutputRegulations.CompareMode = vbTextCompare
    For i = 0 To 2
        mLstProcessRequirements.AddItem "A" & CStr(i + 1)
        mLstProcessRequirements.List(i, 1) = canary & CStr(i + 1)
        mLstProcessRequirements.List(i, 2) = "2"
        mLstProcessRequirements.List(i, 3) = "100"
        mLstProcessRequirements.List(i, 4) = "2"
        mLstProcessRequirements.List(i, 5) = "EA"
        mLstProcessRequirements.List(i, 6) = "FIXED"
        mLstProcessOutputs.AddItem "B" & CStr(i + 1)
        mLstProcessOutputs.List(i, 1) = canary & CStr(i + 1)
        mLstProcessOutputs.List(i, 2) = canary & "ITEM" & CStr(i + 1)
        mLstProcessOutputs.List(i, 3) = canary & "DESIGN" & CStr(i + 1)
        mLstProcessOutputs.List(i, 4) = "1"
        mLstProcessOutputs.List(i, 5) = "2"
        mLstProcessOutputs.List(i, 6) = "100"
        mLstProcessOutputs.List(i, 7) = "2"
        mLstProcessOutputs.List(i, 8) = "EA"
        mLstProcessOutputs.List(i, 9) = "FIXED"
        SetProcessOutputRegulation "B" & CStr(i + 1), True, "1", "9"
        mLstProcessInstructions.AddItem CStr(i + 9)
        mLstProcessInstructions.List(i, 1) = canary & "STEP" & CStr(i + 1)
    Next i
    selected = 1
    If mode = "NoSelection" Or mode = "Append" Then selected = -1
    If mode = "Top" Then selected = 0
    If mode = "Bottom" Then selected = 2
    mLstProcessRequirements.ListIndex = selected: mSelectedProcessRequirementIndex = selected
    mLstProcessOutputs.ListIndex = selected: mSelectedProcessOutputIndex = selected
    mTxtRequirementId.Text = "A2": mTxtRequirementName.Text = " " & canary & "EDIT "
    mTxtRequirementQty.Text = "4": mTxtRequirementPercent.Text = "100": mTxtRequirementYieldBasis.Text = "4"
    mTxtRequirementUom.Text = "EA"
    SelectComboText mCmbRequirementQtyMode, "Enter a number"
    mTxtProcessOutputId.Text = "B2": mTxtProcessOutputName.Text = " " & canary & "EDIT "
    mTxtProcessOutputItemCode.Text = canary & "ITEM": mTxtProcessOutputDesignId.Text = canary & "DESIGN"
    mTxtProcessOutputDesignVersion.Text = "1"
    mTxtProcessOutputQty.Text = "4": mTxtProcessOutputPercent.Text = "100": mTxtProcessOutputYieldBasis.Text = "4"
    mCmbProcessOutputUom.Clear: mCmbProcessOutputUom.AddItem "EA": mCmbProcessOutputUom.ListIndex = 0
    SelectComboText mCmbProcessOutputQtyMode, "Enter a number"
    Select Case mode
        Case "Invalid": mTxtRequirementName.Text = "": mTxtProcessOutputName.Text = ""
        Case "NoSelection", "Append": mTxtRequirementId.Text = "Z1": mTxtProcessOutputId.Text = "Z2"
        Case "Fallback"
            mLstProcessRequirements.ListIndex = -1: mTxtRequirementId.Text = "Z1"
            mLstProcessOutputs.ListIndex = -1: mTxtProcessOutputId.Text = "Z2"
        Case "IdentityLookup": mLstProcessRequirements.ListIndex = 0: mLstProcessOutputs.ListIndex = 0
        Case "Fraction": mTxtRequirementQty.Text = "1.5": mTxtProcessOutputQty.Text = "1.5"
        Case "YieldDefault": mTxtProcessOutputPercent.Text = "": mTxtProcessOutputYieldBasis.Text = ""
        Case "YieldChange": mTxtProcessOutputYieldBasis.Text = "2"
        Case "Actual"
            SelectComboText mCmbRequirementQtyMode, "Variable -- determined at Check In"
            SelectComboText mCmbProcessOutputQtyMode, "Variable -- determined by Actual Output"
            Call ApplyRequirementQtyMode
            Call ApplyOutputQtyMode
    End Select
    mLoading = prior
End Sub
Public Function ComponentRowsForTest(ByVal kind As String) As String
    Dim rows As MSForms.ListBox, i As Long, j As Long, fields As Long
    On Error GoTo Failed
    Set rows = ComponentListForTest(kind)
    fields = IIf(kind = "REQUIREMENT", 7, 10)
    For i = 0 To rows.ListCount - 1
        If i > 0 Then ComponentRowsForTest = ComponentRowsForTest & vbLf
        For j = 0 To fields - 1
            If j > 0 Then ComponentRowsForTest = ComponentRowsForTest & vbTab
            ComponentRowsForTest = ComponentRowsForTest & NzStr(rows.List(i, j))
        Next j
    Next i
    Exit Function
Failed:
    ComponentRowsForTest = "PROBE_ERROR|" & CStr(Err.Number)
End Function
Public Function ComponentEditorForTest(ByVal kind As String) As String
    If kind = "REQUIREMENT" Then
        ComponentEditorForTest = mTxtRequirementId.Text & vbTab & mTxtRequirementName.Text & vbTab & mTxtRequirementQty.Text & vbTab & mTxtRequirementPercent.Text & vbTab & mTxtRequirementYieldBasis.Text & vbTab & mTxtRequirementUom.Text & vbTab & RequirementQtyMode()
    Else
        ComponentEditorForTest = mTxtProcessOutputId.Text & vbTab & mTxtProcessOutputName.Text & vbTab & mTxtProcessOutputQty.Text & vbTab & mTxtProcessOutputPercent.Text & vbTab & mTxtProcessOutputYieldBasis.Text & vbTab & ComboText(mCmbProcessOutputUom) & vbTab & OutputQtyMode()
    End If
End Function
Public Function ComponentAuxForTest(ByVal kind As String) As String
    Dim i As Long
    ComponentAuxForTest = CStr(ComponentListForTest(kind).ListIndex) & "|"
    If kind = "REQUIREMENT" Then ComponentAuxForTest = ComponentAuxForTest & CStr(mSelectedProcessRequirementIndex) Else ComponentAuxForTest = ComponentAuxForTest & CStr(mSelectedProcessOutputIndex)
    ComponentAuxForTest = ComponentAuxForTest & "|" & Join(mProcessOutputRegulations.Keys, ",") & "|"
    For i = 0 To mLstProcessInstructions.ListCount - 1
        ComponentAuxForTest = ComponentAuxForTest & NzStr(mLstProcessInstructions.List(i, 0)) & ","
    Next i
End Function
Public Function ComponentSnapshotForTest() As String
    On Error GoTo Failed
    ComponentSnapshotForTest = ComponentRowsForTest("REQUIREMENT") & vbCr & ComponentRowsForTest("OUTPUT") & vbCr & ComponentEditorForTest("REQUIREMENT") & vbCr & ComponentEditorForTest("OUTPUT") & vbCr & ComponentAuxForTest("REQUIREMENT") & vbCr & ComponentAuxForTest("OUTPUT")
    Exit Function
Failed:
    ComponentSnapshotForTest = "PROBE_ERROR|" & CStr(Err.Number)
End Function
Public Function ComponentGuardForTest(ByVal kind As String, ByVal action As String, ByVal guard As String) As String
    Dim prior As Boolean, original As MSForms.ListBox
    Select Case guard
        Case "Busy"
            prior = mDesignerActionInProgress: mDesignerActionInProgress = True
            ComponentGuardForTest = ComponentActForTest(kind, action)
            mDesignerActionInProgress = prior
        Case "Loading"
            prior = mLoading: mLoading = True
            ComponentGuardForTest = ComponentActForTest(kind, action)
            mLoading = prior
        Case "Failure"
            Set original = ComponentListForTest(kind)
            If kind = "REQUIREMENT" Then Set mLstProcessRequirements = Nothing Else Set mLstProcessOutputs = Nothing
            ComponentGuardForTest = ComponentActForTest(kind, action)
            If kind = "REQUIREMENT" Then Set mLstProcessRequirements = original Else Set mLstProcessOutputs = original
        Case "Nested"
            mComponentNestedEnteredForTest = False: mComponentNestedForTest = True
            ComponentGuardForTest = ComponentActForTest(kind, action)
            mComponentNestedForTest = False
            ComponentGuardForTest = CStr(mComponentNestedEnteredForTest)
        Case "PartialFailure"
            prior = mLoading: mComponentFailAfterWriteForTest = True
            ComponentGuardForTest = ComponentActForTest(kind, action)
            mComponentFailureRestoredForTest = Not mLoading And Not mDesignerActionInProgress
            mLoading = prior: mComponentFailAfterWriteForTest = False
    End Select
End Function
Public Function ComponentFailureRestoredForTest() As Boolean
    ComponentFailureRestoredForTest = mComponentFailureRestoredForTest
End Function
Public Sub ComponentSelectAddedForTest(ByVal kind As String, ByVal canary As String)
    Dim rows As MSForms.ListBox, prior As Boolean
    Set rows = ComponentListForTest(kind)
    prior = mLoading: mLoading = True
    rows.ListIndex = rows.ListCount - 1
    mLoading = prior
    If kind = "REQUIREMENT" Then
        Call mLstProcessRequirements_Click
        mTxtRequirementName.Text = canary & "UPDATED"
    Else
        Call mLstProcessOutputs_Click
        mTxtProcessOutputName.Text = canary & "UPDATED"
    End If
End Sub
Public Function ComponentCaptionForTest(ByVal kind As String, ByVal action As String) As String
    Select Case kind & "_" & action
        Case "REQUIREMENT_ADD": ComponentCaptionForTest = mBtnProcessRequirementAdd.Caption
        Case "REQUIREMENT_UPDATE": ComponentCaptionForTest = mBtnProcessRequirementUpdate.Caption
        Case "REQUIREMENT_REMOVE": ComponentCaptionForTest = mBtnProcessRequirementRemove.Caption
        Case "REQUIREMENT_UP": ComponentCaptionForTest = mBtnProcessRequirementUp.Caption
        Case "REQUIREMENT_DOWN": ComponentCaptionForTest = mBtnProcessRequirementDown.Caption
        Case "OUTPUT_ADD": ComponentCaptionForTest = mBtnProcessOutputAdd.Caption
        Case "OUTPUT_UPDATE": ComponentCaptionForTest = mBtnProcessOutputUpdate.Caption
        Case "OUTPUT_REMOVE": ComponentCaptionForTest = mBtnProcessOutputRemove.Caption
        Case "OUTPUT_UP": ComponentCaptionForTest = mBtnProcessOutputUp.Caption
        Case "OUTPUT_DOWN": ComponentCaptionForTest = mBtnProcessOutputDown.Caption
    End Select
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function ComponentAct(ByVal kind As String, ByVal action As String) As String
    ComponentAct = mForm.ComponentActForTest(kind, action)
End Function
Public Sub ComponentShow()
    mForm.Show vbModeless
End Sub
Public Sub ComponentStage(ByVal kind As String, ByVal canary As String, ByVal mode As String)
    mForm.ComponentStageForTest kind, canary, mode
End Sub
Public Function ComponentRows(ByVal kind As String) As String
    ComponentRows = mForm.ComponentRowsForTest(kind)
End Function
Public Function ComponentEditor(ByVal kind As String) As String
    ComponentEditor = mForm.ComponentEditorForTest(kind)
End Function
Public Function ComponentAux(ByVal kind As String) As String
    ComponentAux = mForm.ComponentAuxForTest(kind)
End Function
Public Function ComponentSnapshot() As String
    ComponentSnapshot = mForm.ComponentSnapshotForTest()
End Function
Public Function ComponentGuard(ByVal kind As String, ByVal action As String, ByVal guard As String) As String
    ComponentGuard = mForm.ComponentGuardForTest(kind, action, guard)
End Function
Public Function ComponentFailureRestored() As Boolean
    ComponentFailureRestored = mForm.ComponentFailureRestoredForTest()
End Function
Public Sub ComponentSelectAdded(ByVal kind As String, ByVal canary As String)
    mForm.ComponentSelectAddedForTest kind, canary
End Sub
Public Function ComponentCaption(ByVal kind As String, ByVal action As String) As String
    ComponentCaption = mForm.ComponentCaptionForTest(kind, action)
End Function
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function ComponentPolicyForTest(ByVal enabled As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(enabled))
    ComponentPolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function ComponentTerminalForTest(ByVal recordJson As String) As Boolean
    Dim record As Object
    Set record = modTrainingJson.DecodeObject(recordJson)
    ComponentTerminalForTest = modEvaluationMatches.CommandCompleted(record)
End Function
'@)
}
