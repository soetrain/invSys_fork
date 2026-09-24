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
