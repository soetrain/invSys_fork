# Reuse established UI adapters; all are installed before constructing any forms.
function Install-ProductionPathsProbe {
    . (Join-Path $PSScriptRoot 'Slice4beViewerPublishedRead.ps1')
    Test-Slice4beViewerPublishedRead $null $null $true
    . (Join-Path $PSScriptRoot 'Slice4beActionRecording.ps1')
    Test-Slice4beActionRecording $null $true
    . (Join-Path $PSScriptRoot 'Slice4beRecordingReader.ps1')
    Test-Slice4beRecordingReader $null $null $true
    . (Join-Path $PSScriptRoot 'Slice4beRecordingEvaluation.ps1')
    Install-RecordingEvaluationProbe
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function DesignerPresentForTest(ByVal pageIndex As Long) As Double
    mPages.Value = pageIndex
    If Not Me.Visible Then Me.Show vbModeless
    Me.Repaint: DoEvents
    DesignerPresentForTest = CDbl(modUserFormResizeWin.GetUserFormWindowHandle(Me))
End Function
'@)
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function Present(ByVal pageIndex As Long) As Double
    Present = mForm.DesignerPresentForTest(pageIndex)
End Function
'@)
    $packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('modInventoryViewer').CodeModule.AddFromString(@'
Public Function PublishedSelectedFieldForTest(ByVal caption As String, ByVal expected As String) As Boolean
    Dim instance As Object, fields As MSForms.ListBox, index As Long
    For Each instance In VBA.UserForms
        If TypeName(instance) = "frmEventDetail" Then
            If Not instance.Visible Then Exit Function
            Set fields = instance.Controls("lstEventFields")
            For index = 0 To fields.ListCount - 1
                If CStr(fields.List(index, 0)) = caption Then
                    PublishedSelectedFieldForTest = (CStr(fields.List(index, 1)) = expected)
                    Exit Function
                End If
            Next index
            Exit Function
        End If
    Next instance
End Function
'@)
    $loaded=[long](Run 'invSys.Admin.xlam' 'TestD5Commands.LoadedFormsForTest')
    Check 'Harness.ProductionPathsProbesInstalledBeforeForms' ($loaded -eq 0)
    if($loaded -ne 0){throw 'Path adapters must precede all forms.'}
}
