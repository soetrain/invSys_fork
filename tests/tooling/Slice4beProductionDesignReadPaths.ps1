# Selection-only adapters; the recorded commands still use the five Click handlers.
function Install-ProductionDesignReadPathProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Sub DesignReadPathInputForTest(ByVal action As String)
    If action = "PROCESS_LOAD" Or action = "PROCESS_REUSE" Then
        mLstProcesses.ListIndex = FindIdentityListRow(mLstProcesses, mReadProcessIdForTest, mReadProcessVersionForTest)
    ElseIf action = "RECIPE_LOAD" Then
        mLstRecipes.ListIndex = FindIdentityListRow(mLstRecipes, mReadRecipeIdForTest, mReadRecipeVersionForTest)
    End If
End Sub
Public Function DesignReadPathResultForTest() As Boolean
    DesignReadPathResultForTest = DesignReadPreservedForTest("PROCESS_REUSE", "Normal") And _
        DesignReadPreservedForTest("RECIPE_LOAD", "Normal")
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub ReadPathInput(ByVal action As String)
    mForm.DesignReadPathInputForTest action
End Sub
Public Function ReadPathResult() As Boolean
    ReadPathResult = mForm.DesignReadPathResultForTest()
End Function
'@)
}
