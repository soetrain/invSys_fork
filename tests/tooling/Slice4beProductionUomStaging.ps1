# Existing captured-workbook/unknown-column contract through the actual Send handler.
function Install-ProductionUomStagingProbe {
    $project=$packages['invSys.Operations.xlam'].VBProject
    $project.VBComponents.Item('frmProduction').CodeModule.AddFromString(@'
Public Function UomSendForTest() As String
    On Error GoTo Failed
    mBtnUomCatalogSend_Click
    UomSendForTest = mTxtStatus.Text
    Exit Function
Failed:
    UomSendForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Public Function UomRetrieveForTest() As String
    On Error GoTo Failed
    mBtnUomCatalogRetrieve_Click
    UomRetrieveForTest = mTxtStatus.Text
    Exit Function
Failed:
    UomRetrieveForTest = "HANDLER_ERROR|" & CStr(Err.Number)
End Function
Public Sub ShowUomForTest()
    mPages.Value = 5
    Me.Show vbModeless
End Sub
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function SendUom() As String
    SendUom = mForm.UomSendForTest()
End Function
Public Function RetrieveUom() As String
    RetrieveUom = mForm.UomRetrieveForTest()
End Function
Public Sub ShowUom()
    mForm.ShowUomForTest
End Sub
'@)
}

function Test-ProductionUomStaging($Fixture) {
    function Column($Table,[string]$Name){
        foreach($column in $Table.ListColumns){if([string]$column.Name -ceq $Name){return $column}}
        return $null
    }
    $book=$null;$decoy=$null
    $configPin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    SelectTarget $Fixture 'config-producer'
    $canary='LOCALUOM'+[guid]::NewGuid().ToString('N')
    $bookPath=Join-Path $runRoot 'uom-staging-operator.xlsb'
    try {
        $book=$excel.Workbooks.Add()
        $book.SaveAs($bookPath,50)
        $book.Close($false)
        $workbookPin=(Get-FileHash -LiteralPath $bookPath).Hash
        $book=$excel.Workbooks.Open($bookPath,0,$false)
        $decoy=$excel.Workbooks.Add()
        [void](Run 'invSys.Operations.xlam' 'TestProductionDesigner.OpenDesigner' @($book.Name))
        $decoy.Activate()
        $first=[string](Run 'invSys.Operations.xlam' 'TestProductionDesigner.SendUom')
        Check 'UomStaging.Initial.ActualHandlerSucceeds' ($first -notlike 'HANDLER_ERROR*' -and $first.Contains('sent to the captured workbook'))
        $sheet=$book.Worksheets.Item('invSys UOM Catalog')
        $table=$sheet.ListObjects.Item('tblInvSysUomCatalog')
        $rowCount=$table.ListRows.Count
        Check 'UomStaging.Initial.RequiredHeaders' ($table.ListColumns.Count -eq 7 -and $rowCount -gt 0 -and $null -ne (Column $table 'UOM') -and $null -ne (Column $table 'Units Per Base UOM'))
        $custom=$table.ListColumns.Add(2);$custom.Name='Operator Custom'
        $custom.DataBodyRange.Value2=$canary
        $formula=$table.ListColumns.Add(3);$formula.Name='Operator Formula'
        $formula.DataBodyRange.FormulaR1C1='=ROW()'
        $sheet.Range('K1').Value2=$canary
        (Column $table 'Notes').DataBodyRange.Cells.Item(1,1).Value2=$canary
        $customIndex=$custom.Index;$formulaIndex=$formula.Index
        $formulaText=[string]$formula.DataBodyRange.Cells.Item(1,1).FormulaR1C1
        $decoyCount=$decoy.Worksheets.Count
        $decoy.Activate()
        $repeat=[string](Run 'invSys.Operations.xlam' 'TestProductionDesigner.SendUom')
        Check 'UomStaging.Repeat.ActualHandlerSucceeds' ($repeat -notlike 'HANDLER_ERROR*' -and $repeat -notlike '*failed*' -and $repeat.Contains('captured workbook'))
        $table=$sheet.ListObjects.Item('tblInvSysUomCatalog')
        $custom=Column $table 'Operator Custom';$formula=Column $table 'Operator Formula'
        Check 'UomStaging.Repeat.UnknownColumnsRetained' ($null -ne $custom -and $null -ne $formula)
        $valuesPreserved=$null -ne $custom
        if($valuesPreserved){
            $valuesPreserved=$custom.DataBodyRange.Rows.Count -eq $rowCount
            for($row=1;$row -le $rowCount;$row++){$valuesPreserved=$valuesPreserved -and [string]$custom.DataBodyRange.Cells.Item($row,1).Value2 -ceq $canary}
        }
        Check 'UomStaging.Repeat.UnknownValuesRetained' $valuesPreserved
        Check 'UomStaging.Repeat.UnknownFormulaRetained' ($null -ne $formula -and [string]$formula.DataBodyRange.Cells.Item(1,1).FormulaR1C1 -ceq $formulaText)
        Check 'UomStaging.Repeat.UnknownColumnOrderRetained' ($null -ne $custom -and $null -ne $formula -and $custom.Index -eq $customIndex -and $formula.Index -eq $formulaIndex)
        Check 'UomStaging.Repeat.UnrelatedWorksheetCellRetained' ([string]$sheet.Range('K1').Value2 -ceq $canary)
        Check 'UomStaging.Repeat.ManagedCatalogRowsRetained' ($table.ListRows.Count -eq $rowCount -and $null -ne (Column $table 'UOM') -and $null -ne (Column $table 'Notes'))
        Check 'UomStaging.Repeat.ManagedDraftEditRetained' ([string](Column $table 'Notes').DataBodyRange.Cells.Item(1,1).Value2 -ceq $canary)
        $decoyUnchanged=$decoy.Worksheets.Count -eq $decoyCount
        foreach($otherSheet in $decoy.Worksheets){$decoyUnchanged=$decoyUnchanged -and [string]$otherSheet.Name -cne 'invSys UOM Catalog'}
        Check 'UomStaging.Repeat.UsesCapturedWorkbook' $decoyUnchanged
        Check 'UomStaging.ConfigBytesPreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
        if($CaptureEvidence){
            $excel.Visible=$true;$book.Activate();$sheet.Activate()
            Initialize-SettingsCapture
            [InvSysSettingsCapture]::SaveVisibleWindow([IntPtr]$excel.Hwnd,(Join-Path $reportRoot 'uom-draft-preserved.png'))
            [void](Run 'invSys.Operations.xlam' 'TestProductionDesigner.ShowUom')
            CaptureOwnedFormByCaptionEvidence 'Production' 'uom-draft-reused-status.png'
        }
    } finally {
        try{[void](Run 'invSys.Operations.xlam' 'TestProductionDesigner.CloseDesigner')}catch{}
        if($null -ne $decoy){$decoy.Close($false)}
        if($null -ne $book){$book.Close($false)}
    }
    Check 'UomStaging.NoImplicitWorkbookSave' ((Get-FileHash -LiteralPath $bookPath).Hash -ceq $workbookPin)
    Test-ProductionUomHeaderAndReopen $Fixture
}

function Test-ProductionUomHeaderAndReopen($Fixture) {
    SelectTarget $Fixture 'config-producer'
    foreach($case in @('Retrieve','RetrieveGap','RetrieveGapSaved','Reopen','MissingHeader','DuplicateHeader','UnownedSheet','MarkerCollision','MarkerWrongSheet','MarkerBroken')){
        $savedPin=$null
        $book=$excel.Workbooks.Add()
        $canary='LOCALUOM'+[guid]::NewGuid().ToString('N')
        try {
            [void](Run 'invSys.Operations.xlam' 'TestProductionDesigner.OpenDesigner' @($book.Name))
            [void](Run 'invSys.Operations.xlam' 'TestProductionDesigner.SendUom')
            $sheet=$book.Worksheets.Item('invSys UOM Catalog')
            $table=$sheet.ListObjects.Item('tblInvSysUomCatalog')
            $extra=$table.ListColumns.Add(2);$extra.Name='Operator Custom';$extra.DataBodyRange.Value2=$canary
            $formula=$table.ListColumns.Add(3);$formula.Name='Operator Formula';$formula.DataBodyRange.FormulaR1C1='=ROW()'
            $notes=$table.ListColumns.Item('Notes');$notes.DataBodyRange.Cells.Item(1,1).Value2=$canary
            # Notes are local; toggle a published field so the version assertion
            # proves a real catalog change rather than a no-op publication.
            $enabled=$table.ListColumns.Item('Enabled').DataBodyRange.Cells.Item(3,1)
            $enabled.Formula=if([bool]$enabled.Value2){'=FALSE()'}else{'=TRUE()'}
            $dimension=$table.ListColumns.Item('Dimension')
            $dimensionValues=$dimension.DataBodyRange.Value2
            $noteValues=$notes.DataBodyRange.Value2
            $dimension.Name='Temporary swap';$notes.Name='Dimension';$dimension.Name='Notes'
            for($row=1;$row -le $table.ListRows.Count;$row++){
                $dimension.DataBodyRange.Cells.Item($row,1).Value2=[string]$noteValues.GetValue($row,1)
                $notes.DataBodyRange.Cells.Item($row,1).Value2=[string]$dimensionValues.GetValue($row,1)
            }
            $sheet.Range('K1').Value2=$canary
            $table.ListColumns.Item('UOM').Name=' uom '
            $table.ListColumns.Item('Dimension').Name=' DIMENSION '
            if($case -like 'RetrieveGap*'){
                # Blank rows are valid staging. Preserve the complete table extent
                # through actual Retrieve/unlist and Edit, not only cell values.
                $blank=$table.ListRows.Add(2)
                $blank.Range.ClearContents()
            }
            $range=$table.Range
            $address=$range.Address()
            $beforeCells=$range.Formula|ConvertTo-Json -Compress -Depth 5
            $version=[long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('UomConversionCatalogVersion',1))
            $configPin=(Get-FileHash -LiteralPath $Fixture.Config).Hash
            if($case -eq 'Retrieve' -or $case -like 'RetrieveGap*'){
                $sheet.Activate();$table.DataBodyRange.Cells.Item(1,1).Select()
                $result=[string](Run 'invSys.Operations.xlam' 'TestProductionDesigner.RetrieveUom')
                Check "UomStaging.$case.NormalizedHeadersAndExtraColumns" ($result -notlike '*failed*' -and $result -notlike 'HANDLER_ERROR*' -and $sheet.ListObjects.Count -eq 0)
                Check "UomStaging.$case.PublishesOneVersion" ([long](Run 'invSys.Core.xlam' 'modConfig.GetLong' @('UomConversionCatalogVersion',1)) -eq $version+1)
                Check "UomStaging.$case.AllStagingCellsPreserved" (($sheet.Range($address).Formula|ConvertTo-Json -Compress -Depth 5) -ceq $beforeCells)
                Check "UomStaging.$case.UnrelatedCellPreserved" ([string]$sheet.Range('K1').Value2 -ceq $canary)
                if($sheet.ListObjects.Count -eq 0){
                    if($case -eq 'RetrieveGapSaved'){
                        [void](Run 'invSys.Operations.xlam' 'TestProductionDesigner.CloseDesigner')
                        $savedPath=Join-Path $runRoot 'uom-gap-draft.xlsb'
                        $book.SaveAs($savedPath,50);$book.Close($false)
                        $savedPin=(Get-FileHash -LiteralPath $savedPath).Hash
                        $book=$excel.Workbooks.Open($savedPath,0,$false)
                        $sheet=$book.Worksheets.Item('invSys UOM Catalog')
                        [void](Run 'invSys.Operations.xlam' 'TestProductionDesigner.OpenDesigner' @($book.Name))
                    }
                    $result=[string](Run 'invSys.Operations.xlam' 'TestProductionDesigner.SendUom')
                    $reopened=$sheet.ListObjects.Count -eq 1 -and ($sheet.Range($address).Formula|ConvertTo-Json -Compress -Depth 5) -ceq $beforeCells
                } else {$reopened=$false}
                Check "UomStaging.$case.ThenEditRetainsDraft" $reopened
                Check "UomStaging.$case.ThenEditRestoresCompleteExtent" ($sheet.ListObjects.Count -eq 1 -and $sheet.ListObjects.Item(1).Range.Address() -ceq $address)
            } elseif($case -eq 'Reopen'){
                # Unlisting is the successful Retrieve postcondition; isolate reopen
                # even while the independent retrieval assertion remains RED.
                $table.Unlist()
                $result=[string](Run 'invSys.Operations.xlam' 'TestProductionDesigner.SendUom')
                Check 'UomStaging.Reopen.ActualHandlerSucceeds' ($result -notlike '*failed*' -and $result -notlike 'HANDLER_ERROR*' -and $sheet.ListObjects.Count -eq 1)
                Check 'UomStaging.Reopen.AllStagingCellsPreserved' (($sheet.Range($address).Formula|ConvertTo-Json -Compress -Depth 5) -ceq $beforeCells)
                Check 'UomStaging.Reopen.OriginalExtentRetained' ($sheet.ListObjects.Count -eq 1 -and $sheet.ListObjects.Item(1).Range.Address() -ceq $address)
                Check 'UomStaging.Reopen.UnrelatedCellPreserved' ([string]$sheet.Range('K1').Value2 -ceq $canary)
                Check 'UomStaging.Reopen.ConfigBytesPreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
            } else {
                if($case -eq 'MissingHeader'){$table.ListColumns.Item('Notes').Name='Local Notes'}
                if($case -eq 'DuplicateHeader'){$extra.Name='UOM'}
                if($case -eq 'UnownedSheet'){$table.Unlist();$sheet.Range('A4').Value2='Local material'}
                if($case -like 'Marker*'){
                    $marker=$sheet.Names.Add("'invSys UOM Catalog'!_invSysUomDraftExtent","='invSys UOM Catalog'!"+$address,$false)
                    $marker.Comment=if($case -eq 'MarkerCollision'){'Operator-owned name'}else{'invSys.UomDraftExtent.v1'}
                    if($case -eq 'MarkerWrongSheet'){$marker.RefersTo="='"+$book.Worksheets.Item(1).Name+"'!"+$address}
                    if($case -eq 'MarkerBroken'){$marker.RefersTo='=#REF!'}
                    $markerPin=[string]$marker.RefersTo+'|'+[string]$marker.Comment+'|'+[string]$marker.Visible
                }
                $beforeCells=$sheet.UsedRange.Formula|ConvertTo-Json -Compress -Depth 5
                $tables=$sheet.ListObjects.Count
                if($case -ne 'UnownedSheet'){
                    $sheet.Activate();$table.DataBodyRange.Cells.Item(1,1).Select()
                    $result=[string](Run 'invSys.Operations.xlam' 'TestProductionDesigner.RetrieveUom')
                    Check "UomStaging.$case.RetrieveRejected" ($result -like '*failed*' -and $result -notlike 'HANDLER_ERROR*')
                    Check "UomStaging.$case.RetrievePreservesStagingAndConfig" (($sheet.UsedRange.Formula|ConvertTo-Json -Compress -Depth 5) -ceq $beforeCells -and $sheet.ListObjects.Count -eq $tables -and (Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
                }
                $result=[string](Run 'invSys.Operations.xlam' 'TestProductionDesigner.SendUom')
                Check "UomStaging.$case.SendRejected" ($result -like '*failed*' -and $result -notlike 'HANDLER_ERROR*')
                Check "UomStaging.$case.SendPreservesCells" (($sheet.UsedRange.Formula|ConvertTo-Json -Compress -Depth 5) -ceq $beforeCells -and $sheet.ListObjects.Count -eq $tables)
                Check "UomStaging.$case.ConfigBytesPreserved" ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
                if($case -like 'Marker*'){
                    Check "UomStaging.$case.MarkerPreserved" (([string]$marker.RefersTo+'|'+[string]$marker.Comment+'|'+[string]$marker.Visible) -ceq $markerPin)
                }
            }
        } finally {
            try{[void](Run 'invSys.Operations.xlam' 'TestProductionDesigner.CloseDesigner')}catch{}
            $book.Close($false)
        }
        if($case -eq 'RetrieveGapSaved' -and $null -ne $savedPin){
            Check 'UomStaging.RetrieveGapSaved.NoImplicitSave' ((Get-FileHash -LiteralPath $savedPath).Hash -ceq $savedPin)
        }
    }
}
