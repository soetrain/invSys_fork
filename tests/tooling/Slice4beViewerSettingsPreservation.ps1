# Supplemental D18 state preservation against the published Viewer fixture.
# Row values are compared only inside VBA; reports contain Boolean facts.
function Initialize-PopulatedViewerSettings($OperationsBook) {
    $form = $OperationsBook.VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
    if($form.Lines(1,$form.CountOfLines) -notmatch 'Private Sub mBtnSettings_Click\(') { return }
    if($form.Lines(1,$form.CountOfLines) -notmatch 'Public Function PopulatedSettingsForTest\(') {
        $form.AddFromString(@'
Private Function PopulatedViewerStateForTest() As String
    Dim rowIndex As Long, columnIndex As Long, value As String
    PopulatedViewerStateForTest = CStr(mTabs.Value) & "|" & CStr(mColumnCount) & "|" & _
        CStr(mLstInventory.ListIndex) & "|" & CStr(mLstInventory.ListCount) & "|" & CStr(mTxtSearch.Value) & "|" & mLblStatus.Caption
    For rowIndex = LBound(mRows, 1) To UBound(mRows, 1)
        For columnIndex = LBound(mRows, 2) To UBound(mRows, 2)
            value = CStr(mRows(rowIndex, columnIndex))
            PopulatedViewerStateForTest = PopulatedViewerStateForTest & "|" & CStr(Len(value)) & ":" & value
        Next columnIndex
    Next rowIndex
End Function
Public Function PopulatedSettingsForTest(ByVal expectedTab As Long) As Boolean
    Dim before As String, instance As Object, settingsCount As Long
    If expectedTab = 2 Then mTabs.Value = 2
    If mTabs.Value <> expectedTab Or IsEmpty(mRows) Or mLstInventory.ListCount = 0 Then Exit Function
    mLstInventory.ListIndex = mLstInventory.ListCount - 1
    before = PopulatedViewerStateForTest()
    mBtnSettings_Click
    mBtnSettings_Click
    For Each instance In VBA.UserForms
        If TypeName(instance) = "frmEventTrackingSettings" Then
            If instance.Visible Then settingsCount = settingsCount + 1
        End If
    Next instance
    PopulatedSettingsForTest = (settingsCount = 1 And before = PopulatedViewerStateForTest())
End Function
'@)
        $module = $OperationsBook.VBProject.VBComponents.Item('modInventoryViewer').CodeModule
        $module.AddFromString(@'
Public Function PopulatedSettingsForTest(ByVal expectedTab As Long) As Boolean
    If Not mInventoryViewer Is Nothing Then PopulatedSettingsForTest = mInventoryViewer.PopulatedSettingsForTest(expectedTab)
End Function
'@)
    }
}
function Test-PopulatedViewerSettings($Excel,$OperationsBook,[int]$ExpectedTab) {
    $form = $OperationsBook.VBProject.VBComponents.Item('frmInventoryViewer').CodeModule
    if($form.Lines(1,$form.CountOfLines) -notmatch 'Public Function PopulatedSettingsForTest\(') { return $false }
    return [bool](Run-WorkbookMacro -Excel $Excel -WorkbookName $OperationsBook.Name -MacroName 'modInventoryViewer.PopulatedSettingsForTest' -Arguments @($ExpectedTab))
}
