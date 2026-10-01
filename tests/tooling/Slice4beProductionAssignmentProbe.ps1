# Unsaved fixture adapters call the nine real handlers; no replacement handlers.
# The shared read adapter supplies controlled boundary reads and a released fixture.
function Install-ProductionAssignmentProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $start=$form.ProcStartLine('DesignerReleasedProcessForTest',0)
    $body=$form.Lines($start,$form.ProcCountLines('DesignerReleasedProcessForTest',0)) -split '\r?\n'
    $anchor=@(for($i=0;$i -lt $body.Count;$i++){if($body[$i].Trim() -ieq 'stage = "SAVE"'){$start+$i}})
    if($anchor.Count -ne 1){throw 'Released Assignment fixture anchor changed; not product RED.'}
    $form.InsertLines($anchor[0],@'
    mLstProcessRequirements.AddItem "A02"
    mLstProcessRequirements.List(0, 1) = fixtureName
    mLstProcessRequirements.List(0, 2) = "1"
    mLstProcessRequirements.List(0, 5) = "EA"
    mLstProcessRequirements.List(0, 6) = "FIXED"
'@)
    $start=$form.ProcStartLine('SelectReusableAssignmentProcess',0)
    $body=$form.Lines($start,$form.ProcCountLines('SelectReusableAssignmentProcess',0)) -split '\r?\n'
    $anchor=@(for($i=0;$i -lt $body.Count;$i++){if($body[$i].Trim() -ieq 'jsonText = modOperationsPrimitiveBridge.GetProcessVersion( _'){$start+$i}})
    if($anchor.Count -ne 1){throw 'Assignment read boundary changed; not product RED.'}
    $form.ReplaceLine($anchor[0],'    jsonText = DesignReadPayloadForTest(True, _')
    $form.AddFromString(@'
Public Function AssignmentPrepareForTest(ByVal canary As String) As Boolean
    Dim ready As String, parts As Variant
    mReadCanaryForTest = canary
    ready = DesignerReleasedProcessForTest(canary)
    parts = Split(ready, "|")
    mReadSetupStageForTest = "ReleasedProcess"
    If Left$(ready, 6) <> "READY|" Then
        If UBound(parts) >= 1 Then mReadSetupStageForTest = CStr(parts(1))
        Exit Function
    End If
    mReadProcessIdForTest = CStr(parts(1)): mReadProcessVersionForTest = CStr(parts(2))
    mReadSetupStageForTest = "Ready"
    AssignmentPrepareForTest = True
End Function
Public Sub AssignmentStageForTest(ByVal mode As String)
    Dim prior As Boolean, alternative As Object
    prior = mLoading: mLoading = True: mReadModeForTest = ""
    DesignerStageForTest "Process", mReadCanaryForTest
    mLstAssignRecipes.Clear: mLstAssignRecipes.AddItem mReadProcessIdForTest
    mLstAssignRecipes.List(0, 1) = mReadProcessVersionForTest
    mLstAssignRecipes.List(0, 2) = mReadCanaryForTest
    mLstAssignIngredients.Clear: mLstAssignIngredients.AddItem "A02"
    mLstAssignIngredients.List(0, 1) = mReadCanaryForTest
    mLstAssignIngredients.List(0, 2) = "EA"
    mLstAssignIngredients.List(0, 3) = mReadCanaryForTest
    mLstAssignIngredients.List(0, 4) = "REQUIREMENT"
    mLstAssignIngredients.List(0, 5) = "1"
    Set mProcessAlternatives = New Collection
    Set alternative = NewReusableRecord("ALTERNATIVE")
    alternative("RequirementId") = "A02": alternative("ITEM_CODE") = "ASSIGN-OLD"
    mProcessAlternatives.Add alternative
    mLstAssignInventory.Clear: mLstAssignInventory.AddItem "ASSIGN-EXACT-KEY"
    mLstAssignInventory.List(0, 1) = mReadCanaryForTest
    mLstAssignInventory.List(0, 2) = "EA": mLstAssignInventory.List(0, 6) = "ASSIGN-NEW"
    mLstAssignRecipes.ListIndex = 0: mLoading = True
    mLstAssignIngredients.ListIndex = 0: mLoading = True
    mLstAssignInventory.ListIndex = 0
    RefreshReusableAllowedItems
    mLstAssignAllowed.ListIndex = 0
    If mode = "NoSelection" Then
        mLstAssignRecipes.ListIndex = -1: mLoading = True
        mLstAssignIngredients.ListIndex = -1: mLoading = True
        mLstAssignInventory.ListIndex = -1: mLstAssignAllowed.ListIndex = -1
    End If
    If mode = "Duplicate" Then mLstAssignInventory.List(0, 6) = "assign-old"
    If mode = "NoMatch" Then mLstAssignAllowed.List(0, 6) = "ABSENT"
    If mode = "Empty" Then
        mLstAssignIngredients.Clear: mLoading = True
        mLstAssignAllowed.Clear: Set mProcessAlternatives = New Collection
    End If
    mReadModeForTest = mode: mReadCallsForTest = 0: mLoading = prior
End Sub
Public Function AssignmentActForTest(ByVal action As String) As String
    On Error GoTo Failed
    Select Case action
        Case "REFRESH": mBtnAssignRefresh_Click
        Case "PROCESS": mBtnAssignRecipe_Click
        Case "REQUIREMENT": mBtnAssignIngredient_Click
        Case "ADD": mBtnAssignAdd_Click
        Case "REMOVE": mBtnAssignRemove_Click
        Case "CLEAR": mBtnAssignClear_Click
        Case "SAVE": mBtnAssignSave_Click
        Case "PROCESS_SELECT": mLstAssignRecipes_Click
        Case "REQUIREMENT_SELECT": mLstAssignIngredients_Click
        Case Else: Err.Raise 5
    End Select
    AssignmentActForTest = mTxtStatus.Text
    Exit Function
Failed:
    AssignmentActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Public Function AssignmentStateForTest() As String
    Dim record As Variant, result As String
    result = DesignerStateForTest("Process") & "|" & CStr(mLstAssignIngredients.ListCount) & "|" & _
        CStr(mLstAssignAllowed.ListCount) & "|" & CStr(mProcessAlternatives.Count)
    For Each record In mProcessAlternatives
        result = result & "|" & modProductionReusableDesigns.ReusableRecordText(record, "RequirementId") & _
            "|" & modProductionReusableDesigns.ReusableRecordText(record, "ITEM_CODE")
    Next record
    AssignmentStateForTest = result
End Function
Public Function AssignmentPreservedForTest(ByVal action As String, ByVal mode As String) As Boolean
    Dim records As Collection, record As Variant, report As String, hasDraft As Boolean, hasAlternative As Boolean
    Select Case action
        Case "REFRESH"
            AssignmentPreservedForTest = (mLstAssignIngredients.ListCount = 0 And mLstAssignAllowed.ListCount = 0 And mProcessAlternatives.Count = 1)
        Case "PROCESS", "PROCESS_SELECT"
            If mode = "EmptyArray" Then
                AssignmentPreservedForTest = (mLstAssignIngredients.ListCount = 0 And mProcessAlternatives.Count = 0)
            Else
                AssignmentPreservedForTest = (mLstAssignIngredients.ListCount = 1 And mProcessAlternatives.Count = 0)
                If AssignmentPreservedForTest Then AssignmentPreservedForTest = (NzStr(mLstAssignIngredients.List(0, 0)) = "A02")
            End If
        Case "REQUIREMENT", "REQUIREMENT_SELECT"
            AssignmentPreservedForTest = (mLstAssignAllowed.ListCount = 1 And mProcessAlternatives.Count = 1)
        Case "ADD"
            AssignmentPreservedForTest = (mProcessAlternatives.Count = 2 And mLstAssignAllowed.ListCount = 2)
            If AssignmentPreservedForTest Then AssignmentPreservedForTest = (modProductionReusableDesigns.ReusableRecordText(mProcessAlternatives(2), "ITEM_CODE") = "ASSIGN-NEW")
        Case "REMOVE", "CLEAR"
            AssignmentPreservedForTest = (mProcessAlternatives.Count = 0 And mLstAssignAllowed.ListCount = 0)
        Case "SAVE"
            Set records = modProductionReusableDesigns.ParseReusableDefinitionRecords( _
                modOperationsPrimitiveBridge.GetProcessVersion(mTxtProcessId.Text, mTxtProcessVersion.Text), report)
            If records Is Nothing Then Exit Function
            For Each record In records
                If modProductionReusableDesigns.ReusableRecordText(record, "RecordType") = "PROCESS" Then _
                    hasDraft = (modProductionReusableDesigns.ReusableRecordText(record, "Status") = "DRAFT")
                If modProductionReusableDesigns.ReusableRecordText(record, "RecordType") = "ALTERNATIVE" Then _
                    hasAlternative = (modProductionReusableDesigns.ReusableRecordText(record, "RequirementId") = "A02" And _
                        modProductionReusableDesigns.ReusableRecordText(record, "ITEM_CODE") = "ASSIGN-OLD")
            Next record
            AssignmentPreservedForTest = hasDraft And hasAlternative And mTxtProcessId.Text = mReadProcessIdForTest And _
                mTxtProcessVersion.Text <> mReadProcessVersionForTest
    End Select
End Function
Public Function AssignmentGuardForTest(ByVal action As String, ByVal guard As String) As String
    If guard = "Loading" Then mLoading = True
    If guard = "Busy" Then mDesignerActionInProgress = True
    AssignmentGuardForTest = AssignmentActForTest(action)
    mLoading = False: mDesignerActionInProgress = False
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function AssignmentPrepare(ByVal canary As String) As Boolean
    AssignmentPrepare = mForm.AssignmentPrepareForTest(canary)
End Function
Public Sub AssignmentStage(ByVal mode As String)
    mForm.AssignmentStageForTest mode
End Sub
Public Function AssignmentAct(ByVal action As String) As String
    AssignmentAct = mForm.AssignmentActForTest(action)
End Function
Public Function AssignmentState() As String
    AssignmentState = mForm.AssignmentStateForTest()
End Function
Public Function AssignmentPreserved(ByVal action As String, ByVal mode As String) As Boolean
    AssignmentPreserved = mForm.AssignmentPreservedForTest(action, mode)
End Function
Public Function AssignmentGuard(ByVal action As String, ByVal guard As String) As String
    AssignmentGuard = mForm.AssignmentGuardForTest(action, guard)
End Function
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function AssignmentPolicyForTest(ByVal enabled As Boolean, ByVal navigation As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String, model As Object, row As Variant
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    Set model = modTrackingPolicyModel.Defaults(enabled)
    For Each row In model("Controls")
        If Right$(CStr(row("ControlId")), 7) = "_SELECT" And Left$(CStr(row("ControlId")), 22) = "PRODUCTION_ASSIGNMENT_" Then row("Collect") = enabled And navigation
    Next row
    request = modTrainingJson.EncodeObject(model)
    AssignmentPolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
'@)
}
