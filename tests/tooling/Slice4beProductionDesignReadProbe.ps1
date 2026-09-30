# Unsaved adapters exercise the existing five Click handlers. Fixture values stay
# in memory; controlled read responses preserve the existing parser/population path.
function Install-ProductionDesignReadProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines($form.CountOfDeclarationLines+1,@'
Private mReadModeForTest As String
Private mReadCallsForTest As Long
Private mReadProcessIdForTest As String
Private mReadProcessVersionForTest As String
Private mReadRecipeIdForTest As String
Private mReadRecipeVersionForTest As String
Private mReadCanaryForTest As String
Private mReadSetupStageForTest As String
Private mReadNestedActionForTest As String
Private mReadNestedEnteredForTest As Boolean
'@)
    $replacements=@{
        '    processes = modOperationsPrimitiveBridge.ListProcesses("")'='    processes = DesignReadListForTest(True, "")'
        '    releasedProcesses = modOperationsPrimitiveBridge.ListProcesses("RELEASED")'='    releasedProcesses = DesignReadListForTest(True, "RELEASED")'
        '    recipes = modOperationsPrimitiveBridge.ListRecipes("")'='    recipes = DesignReadListForTest(False, "")'
        '    releasedRecipes = modOperationsPrimitiveBridge.ListRecipes("RELEASED")'='    releasedRecipes = DesignReadListForTest(False, "RELEASED")'
        '    jsonText = modOperationsPrimitiveBridge.GetProcessVersion(processId, processVersion)'='    jsonText = DesignReadPayloadForTest(True, processId, processVersion)'
        '    jsonText = modOperationsPrimitiveBridge.GetRecipeGraph(recipeId, recipeVersion)'='    jsonText = DesignReadPayloadForTest(False, recipeId, recipeVersion)'
    }
    $lines=$form.Lines(1,$form.CountOfLines) -split '\r?\n'
    foreach($old in $replacements.Keys){
        $matches=@(for($line=0;$line -lt $lines.Count;$line++){if($lines[$line] -ceq $old){$line+1}})
        if($matches.Count -ne 1){throw 'Design read fixture anchor missing/ambiguous; not product RED.'}
        $form.ReplaceLine($matches[0],$replacements[$old])
    }
    foreach($entry in @(
        @{Procedure='RefreshReusableDesignLists';Anchor='    FillListFromArray mLstProcesses, processes'},
        @{Procedure='LoadProcessDefinitionIntoDesigner';Anchor='    ClearProcessDraft False'},
        @{Procedure='LoadRecipeDefinitionIntoDesigner';Anchor='    ClearRecipeDraft False'}
    )){
        $start=$form.ProcStartLine($entry.Procedure,0);$count=$form.ProcCountLines($entry.Procedure,0)
        $body=$form.Lines($start,$count) -split '\r?\n'
        $matches=@(for($i=0;$i -lt $body.Count;$i++){if($body[$i] -ceq $entry.Anchor){$start+$i}})
        if($matches.Count -ne 1){throw 'Design read partial-effect boundary changed; not product RED.'}
        $form.InsertLines($matches[0]+1,'    DesignReadBoundaryForTest')
    }
    $form.AddFromString(@'
Private Sub DesignReadBoundaryForTest()
    Dim action As String
    If mReadModeForTest = "PartialFailure" Then Err.Raise 5432, , "Synthetic local read interruption"
    If mReadNestedActionForTest <> "" Then
        action = mReadNestedActionForTest: mReadNestedActionForTest = ""
        mReadNestedEnteredForTest = True
        Call DesignReadActForTest(action)
    End If
End Sub
Private Function DesignReadListForTest(ByVal process As Boolean, ByVal status As String) As Variant
    mReadCallsForTest = mReadCallsForTest + 1
    If mReadModeForTest = "ReadError" Then Err.Raise 5432, , "Synthetic read interruption"
    If mReadModeForTest = "EmptyLists" Then Exit Function
    If process Then
        DesignReadListForTest = modOperationsPrimitiveBridge.ListProcesses(status)
    Else
        DesignReadListForTest = modOperationsPrimitiveBridge.ListRecipes(status)
    End If
End Function
Private Function DesignReadPayloadForTest(ByVal process As Boolean, ByVal identity As String, ByVal version As String) As String
    mReadCallsForTest = mReadCallsForTest + 1
    Select Case mReadModeForTest
        Case "ReadError": Err.Raise 5432, , "Synthetic read interruption"
        Case "Malformed": DesignReadPayloadForTest = "[{": Exit Function
        Case "EmptyArray": DesignReadPayloadForTest = "[]": Exit Function
        Case "Unavailable": Exit Function
    End Select
    If process Then
        DesignReadPayloadForTest = modOperationsPrimitiveBridge.GetProcessVersion(identity, version)
    Else
        DesignReadPayloadForTest = modOperationsPrimitiveBridge.GetRecipeGraph(identity, version)
    End If
End Function
Public Function DesignReadPrepareForTest(ByVal canary As String) As Boolean
    Dim ready As String, row As Long
    On Error GoTo Failed
    mReadCanaryForTest = canary
    mReadSetupStageForTest = "ReleasedProcess"
    ready = DesignerReleasedProcessForTest(canary)
    If Left$(ready, 6) <> "READY|" Then Exit Function
    mReadProcessIdForTest = mTxtProcessId.Text
    mReadProcessVersionForTest = mTxtProcessVersion.Text
    mReadSetupStageForTest = "RecipeDraft"
    If Not DesignerReleasedRecipeForTest(mReadProcessIdForTest, mReadProcessVersionForTest, canary) Then Exit Function
    mReadRecipeIdForTest = mTxtReusableRecipeId.Text
    mReadRecipeVersionForTest = mTxtReusableRecipeVersion.Text
    mReadSetupStageForTest = "RecipeSave"
    mBtnRecipeSave_Click
    mReadSetupStageForTest = "RecipeRelease"
    mReusableActionTestInProgress = True
    mBtnRecipeRelease_Click
    mReusableActionTestInProgress = False
    mReadSetupStageForTest = "ReleasedStatus"
    row = FindIdentityListRow(mLstLoaderRecipes, mReadRecipeIdForTest, mReadRecipeVersionForTest)
    If row < 0 Then Exit Function
    ' The Run picker has three columns; status belongs to the six-column list.
    row = FindIdentityListRow(mLstRecipes, mReadRecipeIdForTest, mReadRecipeVersionForTest)
    If row < 0 Then Exit Function
    DesignReadPrepareForTest = (NzStr(mLstRecipes.List(row, 4)) = "RELEASED")
    If DesignReadPrepareForTest Then mReadSetupStageForTest = "Ready"
    Exit Function
Failed:
    mReusableActionTestInProgress = False
End Function
Public Function DesignReadSetupStageForTest() As String
    DesignReadSetupStageForTest = mReadSetupStageForTest
End Function
Public Function DesignReadFixtureForTest() As Variant
    DesignReadFixtureForTest = Array(mReadProcessIdForTest, mReadProcessVersionForTest, _
        mReadRecipeIdForTest, mReadRecipeVersionForTest, mReadCanaryForTest)
End Function
Public Sub DesignReadRestoreFixtureForTest(ByVal values As Variant)
    mReadProcessIdForTest = CStr(values(0)): mReadProcessVersionForTest = CStr(values(1))
    mReadRecipeIdForTest = CStr(values(2)): mReadRecipeVersionForTest = CStr(values(3))
    mReadCanaryForTest = CStr(values(4))
End Sub
Public Sub DesignReadStageForTest(ByVal mode As String)
    Dim prior As Boolean
    prior = mLoading: mLoading = True: mReadModeForTest = ""
    ClearProcessDraft False: ClearRecipeDraft False
    mTxtProcessId.Text = "LOCAL-P": mTxtProcessVersion.Text = "99"
    mTxtProcessName.Text = mReadCanaryForTest: mTxtProcessDescription.Text = mReadCanaryForTest
    mTxtReusableRecipeId.Text = "LOCAL-R": mTxtReusableRecipeVersion.Text = "99"
    mTxtReusableRecipeName.Text = mReadCanaryForTest: mTxtReusableRecipeDescription.Text = mReadCanaryForTest
    mLstProcesses.Clear: mLstProcesses.AddItem mReadProcessIdForTest
    mLstProcesses.List(0, 1) = mReadProcessVersionForTest
    mLstRecipes.Clear: mLstRecipes.AddItem mReadRecipeIdForTest
    mLstRecipes.List(0, 1) = mReadRecipeVersionForTest
    mLstReleasedProcesses.Clear: mLstReleasedProcesses.AddItem "STALE"
    mLstAssignRecipes.Clear: mLstAssignRecipes.AddItem "STALE"
    mLstLoaderRecipes.Clear: mLstLoaderRecipes.AddItem "STALE"
    If mode <> "NoSelection" Then
        mLstProcesses.ListIndex = 0: mLstRecipes.ListIndex = 0
    End If
    mReadModeForTest = mode: mReadCallsForTest = 0: mLoading = prior
    mReadNestedActionForTest = "": mReadNestedEnteredForTest = False
End Sub
Public Function DesignReadActForTest(ByVal action As String) As String
    On Error GoTo Failed
    Select Case action
        Case "PROCESS_REFRESH": mBtnProcessRefresh_Click
        Case "PROCESS_LOAD": mBtnProcessLoad_Click
        Case "PROCESS_REUSE": mBtnProcessReuse_Click
        Case "RECIPE_REFRESH": mBtnRecipeRefresh_Click
        Case "RECIPE_LOAD": mBtnRecipeLoad_Click
        Case Else: Err.Raise 5
    End Select
    DesignReadActForTest = mTxtStatus.Text
    Exit Function
Failed:
    DesignReadActForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Public Function DesignReadStateForTest() As String
    DesignReadStateForTest = DesignerStateForTest("Process") & vbLf & DesignerStateForTest("Recipe") & vbLf & _
        CStr(mLstProcesses.ListCount) & "|" & CStr(mLstProcesses.ListIndex) & "|" & _
        CStr(mLstReleasedProcesses.ListCount) & "|" & CStr(mLstRecipes.ListCount) & "|" & CStr(mLstRecipes.ListIndex) & "|" & _
        CStr(mLstAssignRecipes.ListCount) & "|" & CStr(mLstLoaderRecipes.ListCount)
End Function
Public Function DesignReadCallsForTest() As Long
    DesignReadCallsForTest = mReadCallsForTest
End Function
Public Function DesignReadGuardForTest(ByVal action As String, ByVal guard As String) As String
    On Error GoTo Done
    If guard = "Loading" Then mLoading = True
    If guard = "Busy" Then mDesignerActionInProgress = True
    If guard = "Nested" Then mReadNestedActionForTest = action
    DesignReadGuardForTest = DesignReadActForTest(action)
Done:
    mLoading = False: mDesignerActionInProgress = False
End Function
Public Function DesignReadNestedEnteredForTest() As Boolean
    DesignReadNestedEnteredForTest = mReadNestedEnteredForTest
End Function
Public Function DesignReadGuardsRestoredForTest() As Boolean
    DesignReadGuardsRestoredForTest = Not mLoading And Not mDesignerActionInProgress
End Function
Public Function DesignReadPreservedForTest(ByVal action As String, ByVal mode As String) As Boolean
    If Right$(action, 7) = "REFRESH" Then
        If mode = "EmptyLists" Then
            DesignReadPreservedForTest = (mLstProcesses.ListCount = 0 And mLstReleasedProcesses.ListCount = 0 And _
                mLstRecipes.ListCount = 0 And mLstAssignRecipes.ListCount = 0 And mLstLoaderRecipes.ListCount = 0)
        Else
            DesignReadPreservedForTest = (FindIdentityListRow(mLstProcesses, mReadProcessIdForTest, mReadProcessVersionForTest) >= 0 And _
                FindIdentityListRow(mLstReleasedProcesses, mReadProcessIdForTest, mReadProcessVersionForTest) >= 0 And _
                FindIdentityListRow(mLstRecipes, mReadRecipeIdForTest, mReadRecipeVersionForTest) >= 0 And _
                FindIdentityListRow(mLstAssignRecipes, mReadProcessIdForTest, mReadProcessVersionForTest) >= 0 And _
                FindIdentityListRow(mLstLoaderRecipes, mReadRecipeIdForTest, mReadRecipeVersionForTest) >= 0)
        End If
        DesignReadPreservedForTest = DesignReadPreservedForTest And mTxtProcessId.Text = "LOCAL-P" And mTxtReusableRecipeId.Text = "LOCAL-R"
    ElseIf mode = "EmptyArray" Then
        If action = "RECIPE_LOAD" Then
            DesignReadPreservedForTest = (mTxtReusableRecipeId.Text = "" And mLstRecipeNodes.ListCount = 0)
        Else
            DesignReadPreservedForTest = (mTxtProcessId.Text = "" And mLstProcessOutputs.ListCount = 0)
        End If
    ElseIf action = "RECIPE_LOAD" Then
        DesignReadPreservedForTest = (mTxtReusableRecipeId.Text = mReadRecipeIdForTest And _
            mTxtReusableRecipeVersion.Text = mReadRecipeVersionForTest And mLstRecipeNodes.ListCount = 1)
    Else
        DesignReadPreservedForTest = (mTxtProcessId.Text = mReadProcessIdForTest And mLstProcessOutputs.ListCount = 1)
        If Not DesignReadPreservedForTest Then Exit Function
        If action = "PROCESS_LOAD" Then
            DesignReadPreservedForTest = (mTxtProcessVersion.Text = mReadProcessVersionForTest)
        Else
            DesignReadPreservedForTest = (mTxtProcessVersion.Text = "2" And NzStr(mLstProcessOutputs.List(0, 4)) = "2")
        End If
    End If
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function ReadPrepare(ByVal canary As String) As Boolean
    ReadPrepare = mForm.DesignReadPrepareForTest(canary)
End Function
Public Sub ReadReopen(ByVal workbookName As String)
    Dim values As Variant
    values = mForm.DesignReadFixtureForTest()
    OpenDesigner workbookName
    mForm.DesignReadRestoreFixtureForTest values
End Sub
Public Function ReadSetupStage() As String
    ReadSetupStage = mForm.DesignReadSetupStageForTest()
End Function
Public Sub ReadStage(ByVal mode As String)
    mForm.DesignReadStageForTest mode
End Sub
Public Function ReadAct(ByVal action As String) As String
    ReadAct = mForm.DesignReadActForTest(action)
End Function
Public Function ReadState() As String
    ReadState = mForm.DesignReadStateForTest()
End Function
Public Function ReadCalls() As Long
    ReadCalls = mForm.DesignReadCallsForTest()
End Function
Public Function ReadGuard(ByVal action As String, ByVal guard As String) As String
    ReadGuard = mForm.DesignReadGuardForTest(action, guard)
End Function
Public Function ReadNestedEntered() As Boolean
    ReadNestedEntered = mForm.DesignReadNestedEnteredForTest()
End Function
Public Function ReadGuardsRestored() As Boolean
    ReadGuardsRestored = mForm.DesignReadGuardsRestoredForTest()
End Function
Public Function ReadPreserved(ByVal action As String, ByVal mode As String) As Boolean
    ReadPreserved = mForm.DesignReadPreservedForTest(action, mode)
End Function
'@)
    $packages['invSys.Core.xlam'].VBProject.VBComponents.Item('TestShippingCatalog').CodeModule.AddFromString(@'
Public Function ReadPolicyForTest(ByVal enabled As Boolean) As Boolean
    Dim context As String, version As Long, request As String, report As String
    context = modActivity.CaptureContext()
    If Not modTrackingPolicySettings.ReadEditor(context, version, request, report) Then Exit Function
    request = modTrainingJson.EncodeObject(modTrackingPolicyModel.Defaults(enabled))
    ReadPolicyForTest = modTrackingPolicySettings.SavePolicy(context, version, request, report)
End Function
Public Function ReadTerminalForTest(ByVal recordJson As String) As Boolean
    Dim record As Object
    Set record = modTrainingJson.DecodeObject(recordJson)
    ReadTerminalForTest = modEvaluationMatches.CommandCompleted(record)
End Function
'@)
}
