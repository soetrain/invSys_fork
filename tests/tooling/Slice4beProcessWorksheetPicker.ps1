# D4/D14/D15: unsaved adapters observe the real shared picker CommitSelection.
function Install-ProcessWorksheetPickerProbe {
    $core=$packages['invSys.Core.xlam'].VBProject
    $core.VBComponents.Item('cDynItemSearch').CodeModule.AddFromString(@'
Public Function WorksheetPickerExpectedForTest(ByVal field As String) As String
    If lst Is Nothing Then Exit Function
    If lst.ListCount = 0 Then Exit Function
    If field = "Name" Then WorksheetPickerExpectedForTest = CStr(lst.List(0, 2))
    If field = "Sku" Then WorksheetPickerExpectedForTest = CStr(lst.List(0, 1))
End Function
'@)
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('mProduction').CodeModule.AddFromString(@'
Public Function WorksheetPickerExpectedForTest(ByVal field As String) As String
    If mProcessItemPicker Is Nothing Then Exit Function
    WorksheetPickerExpectedForTest = CStr(mProcessItemPicker.WorksheetPickerExpectedForTest(field))
End Function
'@)
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Sub WorksheetPickerAddPairForTest()
    mBtnProcessWorksheetAddAlternative_Click
End Sub
'@)
    $module=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $module.InsertLines(1,@'
Private mPickerItem As Range
Private mPickerSku As Range
Private mPickerExpectedName As String
Private mPickerExpectedSku As String
Private mPickerOtherCells As String
Private mPickerHeaders As String
Private mPickerOutput As Boolean
'@)
    $module.AddFromString(@'
Public Function WorksheetPickerPrepare(ByVal workbookName As String, ByVal mode As String, ByVal pairNumber As Long, ByVal outputRow As Boolean) As Boolean
    Dim column As ListColumn, custom As ListColumn, pair As Long, expectedHeader As String
    On Error GoTo Failed
    If Not WorksheetHeadersPrepare(workbookName, "Canonical") Then Exit Function
    mForm.WorksheetPickerAddPairForTest
    If HeaderTestColumn("Accepted SKU 5") Is Nothing Then Exit Function
    Set custom = mHeaderTable.ListColumns.Add(2)
    custom.Name = "Operator Custom": custom.DataBodyRange.Value2 = "PICKER_CUSTOM_VALUE"
    Set custom = mHeaderTable.ListColumns.Add(3)
    custom.Name = "Operator Formula": custom.DataBodyRange.Formula = "=""PICKER_CUSTOM_FORMULA"""
    For pair = 1 To 5
        HeaderTestColumn("Acceptable Managed Item " & CStr(pair)).DataBodyRange.Cells(1, 1).Value2 = "BEFORE_ITEM_" & CStr(pair)
        HeaderTestColumn("Accepted SKU " & CStr(pair)).DataBodyRange.Cells(1, 1).Value2 = "BEFORE_SKU_" & CStr(pair)
    Next pair
    mPickerOutput = outputRow
    If outputRow Then HeaderTestColumn("Record Type").DataBodyRange.Cells(1, 1).Value2 = "OUTPUT"
    For Each column In mHeaderTable.ListColumns
        If Left$(column.Name, 8) <> "Operator" Then
            Select Case mode
                Case "Normalized": expectedHeader = " " & LCase$(column.Name) & " "
                Case "Lower": expectedHeader = LCase$(column.Name)
                Case "Trimmed": expectedHeader = " " & column.Name & " "
                Case "Canonical": expectedHeader = column.Name
                Case Else: Exit Function
            End Select
            column.Name = "PICKER_TEMP_HEADER"
            column.Name = expectedHeader
            If StrComp(column.Name, expectedHeader, vbBinaryCompare) <> 0 Then Exit Function
        End If
    Next column
    Set mPickerItem = HeaderTestColumn("Acceptable Managed Item " & CStr(pairNumber)).DataBodyRange.Cells(1, 1)
    Set mPickerSku = HeaderTestColumn("Accepted SKU " & CStr(pairNumber)).DataBodyRange.Cells(1, 1)
    mPickerOtherCells = WorksheetPickerOtherSnapshot()
    mPickerHeaders = HeaderNamesSnapshot()
    mHeaderTable.Parent.Parent.Save
    WorksheetPickerPrepare = True
    Exit Function
Failed:
    WorksheetPickerPrepare = False
End Function
Public Function WorksheetPickerOpen() As Boolean
    On Error GoTo Failed
    mProduction.ShowProductionProcessItemSearch mPickerItem
    If Not mProduction.ProductionProcessItemSearchVisibleForTest() Then Exit Function
    If mProduction.ProductionProcessItemSearchResultCountForTest() = 0 Then Exit Function
    mPickerExpectedName = mProduction.WorksheetPickerExpectedForTest("Name")
    mPickerExpectedSku = mProduction.WorksheetPickerExpectedForTest("Sku")
    WorksheetPickerOpen = (mPickerExpectedName <> "" And mPickerExpectedSku <> "")
Failed:
End Function
Public Function WorksheetPickerCommit() As Boolean
    WorksheetPickerCommit = mProduction.CommitFirstProductionProcessItemSearchResultForTest()
End Function
Private Function WorksheetPickerOtherSnapshot() As String
    Dim cell As Range
    For Each cell In mHeaderTable.DataBodyRange.Cells
        If cell.Address <> mPickerItem.Address And cell.Address <> mPickerSku.Address Then
            If Not (mPickerOutput And cell.Row = mPickerItem.Row And (cell.Column = HeaderTestColumn("Name").Range.Column Or cell.Column = HeaderTestColumn("Output SKU").Range.Column)) Then
                WorksheetPickerOtherSnapshot = WorksheetPickerOtherSnapshot & Len(CStr(cell.Formula)) & ":" & CStr(cell.Formula) & "|"
            End If
        End If
    Next cell
End Function
Public Function WorksheetPickerState(ByVal field As String) As Boolean
    On Error GoTo Failed
    Select Case field
        Case "SelectedItem": WorksheetPickerState = (CStr(mPickerItem.Value2) = mPickerExpectedName)
        Case "SelectedSku": WorksheetPickerState = (CStr(mPickerSku.Value2) = mPickerExpectedSku)
        Case "OtherCells": WorksheetPickerState = (WorksheetPickerOtherSnapshot() = mPickerOtherCells)
        Case "Headers": WorksheetPickerState = (HeaderNamesSnapshot() = mPickerHeaders)
        Case "OutputSku"
            WorksheetPickerState = True
            If mPickerOutput Then WorksheetPickerState = (CStr(HeaderTestColumn("Output SKU").DataBodyRange.Cells(1, 1).Value2) = mPickerExpectedSku)
    End Select
Failed:
End Function
Public Sub WorksheetPickerRelease()
    mProduction.CloseProductionProcessItemSearchForTest
    Set mPickerItem = Nothing: Set mPickerSku = Nothing
End Sub
'@)
}

function Test-ProcessWorksheetPicker($Fixture) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin-seeded picker fixture unavailable; not product RED.'}
    $inventory=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    foreach($open in @($excel.Workbooks|Where-Object{$_.FullName -ceq $inventory})){$open.Close($false)}
    SelectTarget $Fixture 'config-producer'
    $pins=@{};foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Filter '*.xlsb' -File){$pins[$file.FullName]=Hash $file.FullName}
    $cases=@(@('Canonical',1,$false),@('Canonical',2,$false),@('Canonical',5,$false),@('Lower',2,$false),@('Trimmed',2,$false),@('Normalized',1,$false),@('Normalized',2,$false),@('Normalized',5,$false),@('Normalized',1,$true))
    foreach($case in $cases){
        $mode=[string]$case[0];$pair=[int]$case[1];$output=[bool]$case[2]
        $label=$mode+'.Pair'+$pair+$(if($output){'.Output'}else{'.Input'})
        $book=$null;$decoy=$null
        try {
            $book=$excel.Workbooks.Add();$path=Join-Path $runRoot ('picker-'+$label+'.xlsb');$book.SaveAs($path,50)
            [void](Probe 'OpenDesigner' @($book.Name))
            if(-not [bool](Probe 'WorksheetPickerPrepare' @($book.Name,$mode,$pair,$output))){throw 'Typed picker workbook fixture unavailable; not product RED.'}
            $pin=Hash $path
            $decoy=$excel.Workbooks.Add();$decoy.Worksheets.Item(1).Range('A1').Value2='DECOY_UNCHANGED';$decoy.Activate()
            if(-not [bool](Probe 'WorksheetPickerOpen')){throw 'Actual picker with nonempty seeded selection unavailable; not product RED.'}
            if($CaptureEvidence -and $mode -eq 'Normalized' -and $pair -eq 2){
                $caption=[string](Run 'invSys.Core.xlam' 'modItemSearch.ResolveSearchCaption' @('production','item'))
                CaptureOwnedFormByCaptionEvidence $caption 'process-picker-normalized-before-commit.png'
            }
            $decoy.Activate()
            Check ('ProcessPicker.'+$label+'.ActualCommitHandler') ([bool](Probe 'WorksheetPickerCommit'))
            foreach($field in @('SelectedItem','SelectedSku','OtherCells','Headers','OutputSku')){Check ('ProcessPicker.'+$label+'.'+$field) ([bool](Probe 'WorksheetPickerState' @($field)))}
            Check ('ProcessPicker.'+$label+'.DecoyPreserved') ($decoy.Worksheets.Count -eq 1 -and [string]$decoy.Worksheets.Item(1).Range('A1').Value2 -ceq 'DECOY_UNCHANGED' -and $decoy.Worksheets.Item(1).ListObjects.Count -eq 0)
            Check ('ProcessPicker.'+$label+'.NoImplicitSave') ((Hash $path) -ceq $pin)
        } finally {
            try{[void](Probe 'WorksheetPickerRelease');[void](Probe 'WorksheetHeadersRelease');[void](Probe 'CloseDesigner')}catch{}
            if($null -ne $decoy){$decoy.Close($false)}
            if($null -ne $book){$book.Close($false)}
        }
    }
    Check 'ProcessPicker.CanonicalAuthorityBytesPreserved' (@($pins.Keys|Where-Object{(Hash $_) -cne $pins[$_]}).Count -eq 0)
}
