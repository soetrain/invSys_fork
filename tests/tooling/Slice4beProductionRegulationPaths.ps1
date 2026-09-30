# Only input selection is adapted; the recorded actions use the original Click handlers.
function Install-ProductionRegulationPathProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Sub RegulationPathInputForTest(ByVal action As String)
    mCmbOutputRegulationScope.ListIndex = IIf(Left$(action, 6) = "RECIPE", 1, 0)
    mLstOutputRegulations.ListIndex = 0
    If Right$(action, 5) = "APPLY" Then
        mChkOutputRegulated.Value = True
        mTxtOutputRegulationFloor.Text = "2": mTxtOutputRegulationCeiling.Text = "8"
    End If
End Sub
Public Function RegulationPathResultForTest() As Boolean
    RegulationPathResultForTest = RegulationValueForTest("Process", "NODE1", "A01") = "False||" And _
        RegulationValueForTest("Recipe", "NODE1", "A01") = "NONE" And _
        RegulationValueForTest("Process", "NODE1", "B02") = "True|5|9" And _
        RegulationValueForTest("Recipe", "NODE2", "A01") = "True|5|9"
End Function
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Sub RegulationPathInput(ByVal action As String)
    mForm.RegulationPathInputForTest action
End Sub
Public Function RegulationPathResult() As Boolean
    RegulationPathResult = mForm.RegulationPathResultForTest()
End Function
'@)
}
