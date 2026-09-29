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
'@)
    $project.VBComponents.Item('TestProductionDesigner').CodeModule.AddFromString(@'
Public Function SendUom() As String
    SendUom = mForm.UomSendForTest()
End Function
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
        $customIndex=$custom.Index;$formulaIndex=$formula.Index
        $formulaText=[string]$formula.DataBodyRange.Cells.Item(1,1).FormulaR1C1
        $decoyCount=$decoy.Worksheets.Count
        $decoy.Activate()
        $repeat=[string](Run 'invSys.Operations.xlam' 'TestProductionDesigner.SendUom')
        Check 'UomStaging.Repeat.ActualHandlerSucceeds' ($repeat -notlike 'HANDLER_ERROR*' -and $repeat.Contains('sent to the captured workbook'))
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
        $decoyUnchanged=$decoy.Worksheets.Count -eq $decoyCount
        foreach($otherSheet in $decoy.Worksheets){$decoyUnchanged=$decoyUnchanged -and [string]$otherSheet.Name -cne 'invSys UOM Catalog'}
        Check 'UomStaging.Repeat.UsesCapturedWorkbook' $decoyUnchanged
        Check 'UomStaging.ConfigBytesPreserved' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configPin)
    } finally {
        try{[void](Run 'invSys.Operations.xlam' 'TestProductionDesigner.CloseDesigner')}catch{}
        if($null -ne $decoy){$decoy.Close($false)}
        if($null -ne $book){$book.Close($false)}
    }
    Check 'UomStaging.NoImplicitWorkbookSave' ((Get-FileHash -LiteralPath $bookPath).Hash -ceq $workbookPin)
}
