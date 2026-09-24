# Disposable adapters invoke real packaged handlers; no runtime handler replacement.
function Install-ProductionDesignerProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function DesignerActionForTest(ByVal designer As String, ByVal action As String) As String
    On Error GoTo ObservedError
    If designer = "Process" Then
        Select Case action
            Case "New": mBtnProcessNew_Click
            Case "Clear": mBtnProcessClear_Click
            Case "Validate": mBtnProcessValidate_Click
            Case Else: Err.Raise 5
        End Select
    ElseIf designer = "Recipe" Then
        Select Case action
            Case "New": mBtnRecipeNew_Click
            Case "Clear": mBtnRecipeClear_Click
            Case "Validate": mBtnRecipeValidate_Click
            Case Else: Err.Raise 5
        End Select
    Else
        Err.Raise 5
    End If
    DesignerActionForTest = mTxtStatus.Text
    Exit Function
ObservedError:
    DesignerActionForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Public Sub DesignerStageForTest(ByVal designer As String, ByVal value As String)
    If designer = "Process" Then
        mTxtProcessName.Text = value
        mTxtProcessId.Text = "P1": mTxtProcessVersion.Text = "1"
        mTxtProcessDescription.Text = value
        mLstProcessOutputs.Clear
    Else
        mTxtReusableRecipeName.Text = value
        mTxtReusableRecipeId.Text = "R1": mTxtReusableRecipeVersion.Text = "1"
        mTxtReusableRecipeDescription.Text = value
        mLstRecipeNodes.Clear: mLstRecipeConnections.Clear
    End If
End Sub
Public Function DesignerStateForTest(ByVal designer As String) As String
    If designer = "Process" Then
        DesignerStateForTest = mTxtProcessName.Text & "|" & mTxtProcessId.Text & "|" & mTxtProcessVersion.Text & "|" & mTxtProcessDescription.Text & "|" & CStr(mLstProcessOutputs.ListCount)
    Else
        DesignerStateForTest = mTxtReusableRecipeName.Text & "|" & mTxtReusableRecipeId.Text & "|" & mTxtReusableRecipeVersion.Text & "|" & mTxtReusableRecipeDescription.Text & "|" & CStr(mLstRecipeNodes.ListCount)
    End If
End Function
Public Sub DesignerValidProcessForTest()
    DesignerStageForTest "Process", "Local validation fixture"
    mLstProcessOutputs.AddItem "A01"
    mLstProcessOutputs.List(0, 1) = "Fixture output"
    mLstProcessOutputs.List(0, 2) = "FIXTURE"
    mLstProcessOutputs.List(0, 5) = "1"
    mLstProcessOutputs.List(0, 8) = "EA"
    mLstProcessOutputs.List(0, 9) = "FIXED"
End Sub
Public Function DesignerCaptionForTest(ByVal designer As String, ByVal action As String) As String
    If designer = "Process" Then
        Select Case action
            Case "New": DesignerCaptionForTest = mBtnProcessNew.Caption
            Case "Clear": DesignerCaptionForTest = mBtnProcessClear.Caption
            Case "Validate": DesignerCaptionForTest = mBtnProcessValidate.Caption
        End Select
    Else
        Select Case action
            Case "New": DesignerCaptionForTest = mBtnRecipeNew.Caption
            Case "Clear": DesignerCaptionForTest = mBtnRecipeClear.Caption
            Case "Validate": DesignerCaptionForTest = mBtnRecipeValidate.Caption
        End Select
    End If
End Function
Public Function DesignerReleasedProcessForTest(ByVal fixtureName As String) As String
    Dim identity As String, version As String, stage As String, row As Long
    On Error GoTo Failed
    stage = "NEW"
    mBtnProcessNew_Click
    identity = mTxtProcessId.Text: version = mTxtProcessVersion.Text
    If identity = "" Or FindIdentityListRow(mLstProcesses, identity, version) >= 0 Then GoTo Failed
    mTxtProcessName.Text = fixtureName
    mTxtProcessOutputId.Text = "A01"
    mTxtProcessOutputName.Text = fixtureName
    mTxtProcessOutputItemCode.Text = "RECIPE-FIXTURE"
    mTxtProcessOutputQty.Text = "1"
    RefreshProcessOutputUomCatalog "EA"
    mBtnProcessOutputAdd_Click
    If mLstProcessOutputs.ListCount <> 1 Then GoTo Failed
    stage = "SAVE"
    mBtnProcessSave_Click
    row = FindIdentityListRow(mLstProcesses, identity, version)
    If row < 0 Then GoTo Failed
    If CStr(mLstProcesses.List(row, 4)) <> "DRAFT" Then GoTo Failed
    stage = "RELEASE"
    mReusableActionTestInProgress = True
    mBtnProcessRelease_Click
    mReusableActionTestInProgress = False
    row = FindIdentityListRow(mLstReleasedProcesses, identity, version)
    If row < 0 Then GoTo Failed
    If CStr(mLstReleasedProcesses.List(row, 4)) <> "RELEASED" Then GoTo Failed
    DesignerReleasedProcessForTest = "READY|" & identity & "|" & version
    Exit Function
Failed:
    mReusableActionTestInProgress = False
    DesignerReleasedProcessForTest = "FIXTURE_FAILED|" & stage & "|" & CStr(Err.Number)
End Function
Public Sub DesignerRefreshForTest()
    mBtnRecipeRefresh_Click
End Sub
Public Function DesignerReleasedRecipeForTest(ByVal identity As String, ByVal version As String, ByVal fixtureName As String) As Boolean
    Dim row As Long, records As Collection, report As String, record As Object, hasOutput As Boolean, released As Boolean
    ClearRecipeDraft True
    mTxtReusableRecipeName.Text = fixtureName
    row = FindIdentityListRow(mLstReleasedProcesses, identity, version)
    ShowStatus "FIXTURE_RECIPE_ROW"
    If row < 0 Then Exit Function
    mLstReleasedProcesses.ListIndex = row
    mBtnRecipeAddProcess_Click
    ShowStatus "FIXTURE_RECIPE_NODE"
    If mLstRecipeNodes.ListCount <> 1 Then Exit Function
    Set records = ProcessRecordsForRecipeNode(0, report)
    ShowStatus "FIXTURE_RECIPE_RECORDS"
    If records Is Nothing Then Exit Function
    For Each record In records
        If modProductionReusableDesigns.ReusableRecordText(record, "RecordType") = "OUTPUT" Then hasOutput = True
        If modProductionReusableDesigns.ReusableRecordText(record, "RecordType") = "PROCESS" Then
            released = (modProductionReusableDesigns.ReusableRecordText(record, "Status") = "RELEASED")
        End If
    Next record
    ShowStatus "FIXTURE_RECIPE_OUTPUT_STATUS"
    DesignerReleasedRecipeForTest = hasOutput And released
End Function
'@)
    $module=$project.VBComponents.Add(1);$module.Name='TestProductionDesigner'
    $module.CodeModule.AddFromString(@'
Private mForm As frmProduction
Public Sub OpenDesigner(ByVal workbookName As String)
    CloseDesigner
    Set mForm = New frmProduction
    mForm.SetOperatorWorkbook Application.Workbooks(workbookName)
End Sub
Public Sub CloseDesigner()
    If mForm Is Nothing Then Exit Sub
    Unload mForm: Set mForm = Nothing
End Sub
Public Function Act(ByVal designer As String, ByVal action As String) As String
    Act = mForm.DesignerActionForTest(designer, action)
End Function
Public Sub Stage(ByVal designer As String, ByVal value As String)
    mForm.DesignerStageForTest designer, value
End Sub
Public Function State(ByVal designer As String) As String
    State = mForm.DesignerStateForTest(designer)
End Function
Public Sub ValidProcess()
    mForm.DesignerValidProcessForTest
End Sub
Public Function Caption(ByVal designer As String, ByVal action As String) As String
    Caption = mForm.DesignerCaptionForTest(designer, action)
End Function
Public Function ReleasedProcess(ByVal fixtureName As String) As String
    ReleasedProcess = mForm.DesignerReleasedProcessForTest(fixtureName)
End Function
Public Function ReleasedRecipe(ByVal identity As String, ByVal version As String, ByVal fixtureName As String) As Boolean
    ReleasedRecipe = mForm.DesignerReleasedRecipeForTest(identity, version, fixtureName)
End Function
Public Sub RefreshDesigners()
    mForm.DesignerRefreshForTest
End Sub
Public Function RecipeFixtureStatus() As String
    Dim status As String
    status = mForm.TestStatusText()
    If Left$(status, 15) = "FIXTURE_RECIPE_" Then RecipeFixtureStatus = status Else RecipeFixtureStatus = "UNAVAILABLE"
End Function
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function DesignerDisableTrackingForTest() As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(False))
    DesignerDisableTrackingForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
'@)
}
