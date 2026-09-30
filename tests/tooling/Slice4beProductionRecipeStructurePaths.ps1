function Install-ProductionRecipeStructurePathProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $form=$project.VBComponents.Item('frmProduction').CodeModule
    $form.InsertLines($form.CountOfDeclarationLines+1,@'
Private mStructurePathNodesForTest As Variant
Private mStructurePathConnectionsForTest As Variant
'@)
    $form.AddFromString(@'
Public Sub StructurePathRememberForTest()
    mStructurePathNodesForTest = mLstRecipeNodes.List
    mStructurePathConnectionsForTest = mLstRecipeConnections.List
End Sub
Public Sub StructurePathResetForTest()
    mLoading = True
    mLstRecipeNodes.List = mStructurePathNodesForTest
    mLstRecipeConnections.List = mStructurePathConnectionsForTest
    mLstRecipeConnections.ListIndex = 0
    mLoading = False
    RefreshConnectionNodeCombos
    RefreshRecipeConnectionDisplay 0
    LoadConnectionEditorFromIndex 0
End Sub
Public Sub StructurePathInputForTest(ByVal action As String)
    Select Case action
        Case "CONNECT"
            SelectComboText mCmbConnectionFromNode, NzStr(mLstRecipeNodes.List(0, 0))
            mCmbConnectionFromNode_Change
            SelectComboText mCmbConnectionOutput, "001"
            mCmbConnectionOutput_Change
            SelectCompatibleTarget NzStr(mLstRecipeNodes.List(1, 0)), "001"
        Case "UPDATE_CONNECTION"
            mTxtConnectionQty.Text = "4": mTxtConnectionPercent.Text = "75"
        Case "ADD_PROCESS": mLstReleasedProcesses.ListIndex = 0
    End Select
End Sub
Public Function StructurePathResultForTest() As Boolean
    If mLstRecipeNodes.ListCount <> 2 Or mLstRecipeConnections.ListCount <> 1 Then Exit Function
    StructurePathResultForTest = (NzStr(mLstRecipeConnections.List(0, 4)) = "4") And _
        (NzStr(mLstRecipeConnections.List(0, 5)) = "75") And _
        (NzStr(mLstRecipeConnections.List(0, 0)) = NzStr(mLstRecipeNodes.List(0, 0))) And _
        (NzStr(mLstRecipeConnections.List(0, 2)) = NzStr(mLstRecipeNodes.List(1, 0)))
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub StructurePathRemember()
    mForm.StructurePathRememberForTest
End Sub
Public Sub StructurePathReset()
    mForm.StructurePathResetForTest
End Sub
Public Sub StructurePathInput(ByVal action As String)
    mForm.StructurePathInputForTest action
End Sub
Public Function StructurePathResult() As Boolean
    StructurePathResult = mForm.StructurePathResultForTest()
End Function
'@)
}
