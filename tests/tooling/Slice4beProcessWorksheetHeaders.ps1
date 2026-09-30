# D14/D15: actual packaged Send and Retrieve handlers; disposable worksheet data
# stays in memory. Adapters are added only to the unsaved test package instance.
function Install-ProcessWorksheetHeadersProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Sub WorksheetHeadersCreateForTest()
    ClearProcessDraft True
    mBtnProcessWorksheetCreate_Click
End Sub
Public Function WorksheetHeadersRetrieveForTest() As Boolean
    mBtnProcessWorksheetRetrieve_Click
    WorksheetHeadersRetrieveForTest = (InStr(1, mTxtStatus.Text, "Process Name is required on the worksheet.", vbBinaryCompare) > 0)
End Function
Public Function WorksheetHeadersImportForTest() As Boolean
    mBtnProcessWorksheetRetrieve_Click
    WorksheetHeadersImportForTest = (InStr(1, mTxtStatus.Text, "Retrieved 1 selected Process table", vbBinaryCompare) > 0 And mLstProcessRequirements.ListCount = 3 And mLstProcessOutputs.ListCount = 1)
    If Not WorksheetHeadersImportForTest Then Exit Function
    WorksheetHeadersImportForTest = (UCase$(CStr(mLstProcessRequirements.List(0, 5))) = "LB" And UCase$(CStr(mLstProcessRequirements.List(1, 5))) = "EA" And UCase$(CStr(mLstProcessRequirements.List(2, 5))) = "EA" And Abs(CDbl(mLstProcessRequirements.List(0, 4)) - 4.5) < 0.001 And Abs(CDbl(mLstProcessRequirements.List(1, 4)) - 2#) < 0.001 And UCase$(CStr(mLstProcessOutputs.List(0, 8))) = "EA")
End Function
Public Sub WorksheetHeadersShowForTest()
    mPages.Value = 0
    Me.Show vbModeless
End Sub
'@)
    $module=$project.VBComponents.Item('TestProductionDesigner').CodeModule
    $module.InsertLines(1,@'
Private mHeaderTable As ListObject
Private mHeaderNames As String
Private mHeaderIds As String
Private mCustomColumn As ListColumn
Private mCustomFormulas As String
Private mCustomPosition As Long
Private mHeaderCase As String
'@)
    $module.AddFromString(@'
Public Function WorksheetHeadersPrepare(ByVal workbookName As String, ByVal caseName As String) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject, count As Long
    Dim column As ListColumn, report As String, values As Variant, names As Variant
    On Error GoTo Failed
    Set wb = Application.Workbooks(workbookName)
    mForm.WorksheetHeadersCreateForTest
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If Left$(lo.Name, 15) = "invSys_Process_" Then
                Set mHeaderTable = lo
                count = count + 1
            End If
        Next lo
    Next ws
    If count <> 1 Then Exit Function
    If mHeaderTable.ListRows.Count < 3 Then Exit Function
    mHeaderCase = caseName
    If Left$(caseName, 6) = "Import" Then
        If Not modProductionProcessWorksheet.PopulateFormulationExampleForTest(wb, mHeaderTable.Name, True, report) Then Exit Function
        mHeaderTable.Parent.Cells(mHeaderTable.HeaderRowRange.Row - 4, 2).Value2 = "Header import fixture"
    End If
    mHeaderIds = HeaderColumnSnapshot(mHeaderTable.ListColumns("ID"))
    Set mCustomColumn = Nothing
    Select Case caseName
        Case "Canonical", "ImportCanonical"
        Case "CustomValue", "CustomFormula", "BeforeId", "ImportShifted"
            If caseName = "BeforeId" Then count = 2 Else count = 11
            Set mCustomColumn = mHeaderTable.ListColumns.Add(count)
            mCustomColumn.Name = "Operator Custom"
            If caseName = "CustomFormula" Then
                mCustomColumn.DataBodyRange.Formula = "=""LOCAL_CUSTOM_FORMULA"""
            Else
                mCustomColumn.DataBodyRange.Value2 = "LOCAL_CUSTOM_VALUE"
            End If
            mCustomPosition = mCustomColumn.Index
            mCustomFormulas = HeaderColumnSnapshot(mCustomColumn)
        Case "NormalizedRequirement"
            With mHeaderTable.ListColumns("Requirement ID")
                .Name = " requirement id "
                .DataBodyRange.ClearContents
            End With
        Case "NormalizedAll", "ImportNormalized"
            For Each column In mHeaderTable.ListColumns
                column.Name = " " & LCase$(column.Name) & " "
            Next column
            HeaderTestColumn("Requirement ID").DataBodyRange.ClearContents
        Case "Reordered"
            ' Swap complete managed fields; retained custom columns are separate cases.
            values = mHeaderTable.ListColumns("Record Type").DataBodyRange.Value2
            names = mHeaderTable.ListColumns("Name").DataBodyRange.Value2
            mHeaderTable.ListColumns("Record Type").Name = "Temporary field swap"
            mHeaderTable.ListColumns("Name").Name = "Record Type"
            mHeaderTable.ListColumns("Temporary field swap").Name = "Name"
            mHeaderTable.ListColumns("Record Type").DataBodyRange.Value2 = values
            mHeaderTable.ListColumns("Name").DataBodyRange.Value2 = names
            mHeaderTable.ListColumns("Requirement ID").DataBodyRange.ClearContents
        Case Else
            Exit Function
    End Select
    mHeaderNames = HeaderNamesSnapshot()
    wb.Save
    wb.Activate
    mHeaderTable.Parent.Activate
    mHeaderTable.DataBodyRange.Cells(1, 1).Select
    WorksheetHeadersPrepare = True
    Exit Function
Failed:
    WorksheetHeadersPrepare = False
End Function
Private Function HeaderTestColumn(ByVal header As String) As ListColumn
    Dim column As ListColumn
    For Each column In mHeaderTable.ListColumns
        If LCase$(Trim$(column.Name)) = LCase$(header) Then
            Set HeaderTestColumn = column
            Exit Function
        End If
    Next column
End Function
Private Function HeaderColumnSnapshot(ByVal column As ListColumn) As String
    Dim cell As Range
    For Each cell In column.DataBodyRange.Cells
        HeaderColumnSnapshot = HeaderColumnSnapshot & Len(CStr(cell.Formula)) & ":" & CStr(cell.Formula) & "|"
    Next cell
End Function
Private Function HeaderNamesSnapshot() As String
    Dim column As ListColumn
    For Each column In mHeaderTable.ListColumns
        HeaderNamesSnapshot = HeaderNamesSnapshot & Len(column.Name) & ":" & column.Name & "|"
    Next column
End Function
Public Function WorksheetHeadersRetrieve() As Boolean
    mHeaderTable.Parent.Parent.Activate
    mHeaderTable.Parent.Activate
    mHeaderTable.DataBodyRange.Cells(1, 1).Select
    WorksheetHeadersRetrieve = mForm.WorksheetHeadersRetrieveForTest()
End Function
Public Function WorksheetHeadersImport() As Boolean
    WorksheetHeadersImport = mForm.WorksheetHeadersImportForTest()
End Function
Public Sub WorksheetHeadersShow()
    mForm.WorksheetHeadersShowForTest
End Sub
Public Function WorksheetHeadersState(ByVal checkName As String) As Boolean
    Dim column As ListColumn, requirement As ListColumn, rowIndex As Long
    On Error GoTo Failed
    Select Case checkName
        Case "TableRetained"
            WorksheetHeadersState = (mHeaderTable.ListRows.Count >= 3)
        Case "HeadersAndPositions"
            WorksheetHeadersState = (HeaderNamesSnapshot() = mHeaderNames)
        Case "IdentityPreserved"
            WorksheetHeadersState = (HeaderColumnSnapshot(HeaderTestColumn("ID")) = mHeaderIds)
        Case "CustomPreserved"
            If mCustomColumn Is Nothing Then
                WorksheetHeadersState = True
            Else
                WorksheetHeadersState = (mCustomColumn.Index = mCustomPosition And HeaderColumnSnapshot(mCustomColumn) = mCustomFormulas)
            End If
        Case "ManagedRequirementRestored"
            For Each column In mHeaderTable.ListColumns
                If LCase$(Trim$(column.Name)) = "requirement id" Then Set requirement = column
            Next column
            If requirement Is Nothing Then Exit Function
            For rowIndex = 1 To mHeaderTable.ListRows.Count
                If CStr(HeaderTestColumn("Record Type").DataBodyRange.Cells(rowIndex, 1).Value2) = "INPUT" Then
                    If Not requirement.DataBodyRange.Cells(rowIndex, 1).HasFormula Then Exit Function
                    If CStr(requirement.DataBodyRange.Cells(rowIndex, 1).Value2) <> CStr(HeaderTestColumn("ID").DataBodyRange.Cells(rowIndex, 1).Value2) Then Exit Function
                End If
            Next rowIndex
            WorksheetHeadersState = True
    End Select
Failed:
End Function
Public Sub WorksheetHeadersRelease()
    Set mCustomColumn = Nothing
    Set mHeaderTable = Nothing
End Sub
'@)
}

function Test-ProcessWorksheetHeaders($Fixture) {
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function InventoryBusinessState {
        $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
        $source=$null;$owned=$false
        foreach($open in $excel.Workbooks){if([string]$open.FullName -ieq $path){$source=$open;break}}
        if($null -eq $source){$source=$excel.Workbooks.Open($path,0,$true);$owned=$true}
        try {
            $rows=@()
            foreach($sheet in $source.Worksheets){foreach($table in $sheet.ListObjects){
                if([string]$table.Name -cin @('tblInventoryLog','tblAppliedEvents','tblInventoryEntities','tblSkuBalance','tblLocationBalance','tblSkuCatalog')){
                    $rows+=@{Name=[string]$table.Name;Values=$table.Range.Value2}
                }
            }}
            if($rows.Count -ne 6){throw 'Six Inventory business tables required for import preservation fixture.'}
            return ($rows|Sort-Object Name|ConvertTo-Json -Depth 8 -Compress)
        } finally {if($owned){$source.Close($false)}}
    }
    SelectTarget $Fixture 'config-producer'
    $pins=@{}
    $authorityTrace=[Collections.Generic.List[object]]::new()
    foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -File -Filter '*.xlsb'){$pins[$file.FullName]=Hash $file.FullName}
    foreach($case in @('Canonical','CustomValue','CustomFormula','BeforeId','NormalizedRequirement','NormalizedAll','Reordered')){
        $book=$null;$decoy=$null
        try {
            $book=$excel.Workbooks.Add();$path=Join-Path $runRoot ('process-headers-'+$case+'.xlsb');$book.SaveAs($path,50)
            $decoy=$excel.Workbooks.Add();$decoySheets=$decoy.Worksheets.Count
            [void](Probe 'OpenDesigner' @($book.Name));$decoy.Activate()
            if(-not [bool](Probe 'WorksheetHeadersPrepare' @($book.Name,$case))){throw 'Process worksheet fixture setup unavailable; not behavioral RED.'}
            Check ('ProcessHeaders.'+$case+'.SendUsesCapturedWorkbook') ($decoy.Worksheets.Count -eq $decoySheets -and $decoy.Worksheets.Item(1).ListObjects.Count -eq 0)
            $diskPin=Hash $path
            Check ('ProcessHeaders.'+$case+'.ActualRetrieveRejectsMissingName') ([bool](Probe 'WorksheetHeadersRetrieve'))
            foreach($check in @('TableRetained','HeadersAndPositions','IdentityPreserved','CustomPreserved','ManagedRequirementRestored')){
                Check ('ProcessHeaders.'+$case+'.'+$check) ([bool](Probe 'WorksheetHeadersState' @($check)))
            }
            Check ('ProcessHeaders.'+$case+'.NoImplicitSave') ((Hash $path) -ceq $diskPin)
            if($CaptureEvidence -and $case -in @('CustomValue','NormalizedRequirement')){
                $excel.ScreenUpdating=$true;$excel.Visible=$true;$book.Windows.Item(1).Visible=$true;$book.Activate();$book.Worksheets.Item('invSys Process Editor').Activate()
                Start-Sleep -Milliseconds 400
                Initialize-SettingsCapture
                [InvSysSettingsCapture]::SaveVisibleWindow([IntPtr]$excel.Hwnd,(Join-Path $reportRoot ('process-headers-'+$case+'.png')))
                [void](Probe 'WorksheetHeadersShow')
                CaptureOwnedFormByCaptionEvidence 'Production' ('process-headers-'+$case+'-status.png')
            }
        } finally {
            try{[void](Probe 'WorksheetHeadersRelease');[void](Probe 'CloseDesigner')}catch{}
            if($null -ne $decoy){$decoy.Close($false)}
            if($null -ne $book){$book.Close($false)}
        }
    }
    Check 'ProcessHeaders.CanonicalAuthorityBytesPreserved' (@($pins.Keys|Where-Object{(Hash $_) -cne $pins[$_]}).Count -eq 0)
    $inventoryBefore=InventoryBusinessState
    # Successful explicit Retrieve removes the selected staging table under D15.
    # Verify actual mixed-UOM import separately from retained-table preservation.
    foreach($case in @('ImportCanonical','ImportShifted','ImportNormalized')){
        $book=$null
        try {
            $book=$excel.Workbooks.Add();$book.SaveAs((Join-Path $runRoot ($case+'.xlsb')),50)
            [void](Probe 'OpenDesigner' @($book.Name))
            if(-not [bool](Probe 'WorksheetHeadersPrepare' @($book.Name,$case))){throw 'Process import fixture setup unavailable; not behavioral RED.'}
            Check ('ProcessHeaders.'+$case+'.ActualRetrieveImportsMixedUom') ([bool](Probe 'WorksheetHeadersImport'))
            Check ('ProcessHeaders.'+$case+'.OnlyConfirmedTableRemoved') ($book.Worksheets.Item('invSys Process Editor').ListObjects.Count -eq 0)
            foreach($file in $pins.Keys){
                $kind=switch -Regex ([IO.Path]::GetFileName($file)){'\.Config\.'{'Config';break} '\.Auth\.'{'Auth';break} '\.Inventory\.'{'Inventory';break} '\.Designs\.'{'Designs';break} default{'Other generated workbook'}}
                $authorityTrace.Add([pscustomobject]@{Case=$case;Kind=$kind;BytesPreserved=((Hash $file) -ceq $pins[$file])})
            }
            $authorityTrace|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'import-authority-checkpoints.json')
        } finally {
            try{[void](Probe 'WorksheetHeadersRelease');[void](Probe 'CloseDesigner')}catch{}
            if($null -ne $book){$book.Close($false)}
        }
    }
    # RunBatch owns locks/inboxes/publication. Byte identity is applicable to
    # Auth/Config here; inventory business tables must remain logically exact.
    Check 'ProcessHeaders.ImportPreservesAuthConfigBytes' (@($pins.Keys|Where-Object{($_ -like '*.Auth.xlsb' -or $_ -like '*.Config.xlsb') -and (Hash $_) -cne $pins[$_]}).Count -eq 0)
    Check 'ProcessHeaders.ImportPreservesInventoryBusinessState' ((InventoryBusinessState) -ceq $inventoryBefore)
}
