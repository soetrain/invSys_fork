# D8-INV-CONFIG-01: actual packaged actions and independent persisted Config reads.
. (Join-Path $PSScriptRoot 'Slice4beInventoryCatalogCompletion.ps1')
function Install-InventoryConfigPreservationProbe {
    Install-InventoryCatalogCompletionProbe
}

function Set-InventoryConfigPresentation($Fixture) {
    $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
    try {
        $sheet=$book.Worksheets.Item(1)
        $sheet.Columns.Item(1).ColumnWidth=43
        $sheet.Rows.Item(1).Font.Bold=$false
        $book.Save()
    } finally {$book.Close($false)}
}

function Test-InventoryConfigPresentation($Fixture,[string]$Id) {
    $hash=Get-AdminInventoryEntryHash $Fixture.Config
    $book=$excel.Workbooks.Open($Fixture.Config,0,$true)
    try {
        $sheet=$book.Worksheets.Item(1)
        Check ($Id+'.CustomWidthPreserved') ([double]$sheet.Columns.Item(1).ColumnWidth -eq 43)
        Check ($Id+'.CustomHeaderFontPreserved') ($sheet.Rows.Item(1).Font.Bold -eq $false)
    } finally {$book.Close($false)}
    Check ($Id+'.PresentationInspectionReadOnly') ($hash -ceq (Get-AdminInventoryEntryHash $Fixture.Config))
}

function Test-InventoryConfigPreservation($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    SelectTarget $Fixture
    # Owning Seed initializes inventory and all three inbox roles before observation.
    $csv=Join-Path $runRoot 'config-preservation-fixture.csv'
    @('ITEM_CODE,ITEM,UOM,LOCATION,QTY,CONDITION','DEMO-CONFIG-BASE,Config baseline,LB,CLEARVIEW,3,GOOD')|Set-Content $csv -Encoding ASCII
    if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareExactQuantityFixture' @($Fixture.Warehouse,$csv))){throw 'Owning seed fixture unavailable; not RED.'}
    foreach($case in @('ModalAdd','WorksheetAdd','ModalCatalogEdit','WorksheetCatalogEdit')){
        SelectTarget $Fixture
        $id='InventoryConfigPreservation.'+$case
        $edit=$case -like '*Edit'
        $worksheet=$case -like 'Worksheet*'
        $sku='DEMO-CATALOG-COMPLETION-'+$case.ToUpperInvariant()
        if($edit){
            $book=$excel.Workbooks.Open((Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')),0,$false)
            try {
                $prepared=[bool](Run 'invSys.Inventory.Domain.xlam' 'modInventoryApply.TestCatalogOnlyFixture' @($book.Name,$sku))
                if(-not $prepared){throw 'Catalog-only fixture unavailable; not RED.'}
                $book.Save()
            } finally {$book.Close($false)}
            Check ($id+'.CatalogOnlyHasNoEntity') ((Get-CatalogCompletionEntityStates $Fixture $sku).Count -eq 0)
        }
        $hostBook=$excel.Workbooks.Add()
        $hostBook.SaveAs((Join-Path $runRoot ($case+'.xlsx')),51)
        try {
            if($worksheet){
                if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))){throw 'Worksheet fixture unavailable; not RED.'}
                $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
                if($edit){
                    foreach($pair in @{Action='EDIT';'Item Code'=$sku;'Quantity'=7;'Edit Reason'='Config preservation completion'}.GetEnumerator()){
                        $table.ListColumns.Item($pair.Key).DataBodyRange.Cells.Item(1,1).Value2=$pair.Value
                    }
                }
                $extra=$table.ListColumns.Add();$extra.Name='USER NOTE';$extra.DataBodyRange.Cells.Item(1,1).Value2='Preserved extension'
                $hostBook.Save()
            }
            Set-InventoryConfigPresentation $Fixture
            $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $auth=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
            [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.Arm' @($false))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$(if($edit){$sku}else{''})))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetQuantityTarget' @(7.0))
            $action=if($worksheet){'Upload'}elseif($edit){'Edit'}else{'Add'}
            $mode=if($edit -and -not $worksheet){'SetQuantity'}else{'Authorized'}
            $phase=if($edit -and -not $worksheet){'EditReason'}else{'ModalReturn'}
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,$action,$mode,$phase,$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq 'True'
            Check ($id+'.ActualHandlerAndCapturedOwnerReached') $ready
            if(-not $ready){throw 'Packaged action unavailable; not RED.'}
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $newKeys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
            $target=if($edit){7}else{0}
            Check ($id+'.ExactEntityAndQuantitySaved') ($newKeys.Count -eq 1 -and $after.EntityQuantities[$newKeys[0]] -eq $target -and $after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1)
            if($edit){
                $states=Get-CatalogCompletionEntityStates $Fixture $sku
                Check ($id+'.ExactCatalogItemCompleted') ($states.Count -eq 1 -and @($states.Values)[0].Quantity -eq 7 -and @($states.Values)[0].State -ceq 'ACTIVE')
            }
            if($worksheet){
                Check ($id+'.ImportedStatus') ([string]$table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'IMPORTED')
                Check ($id+'.UnknownColumnPreserved') ([string]$table.ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Preserved extension')
            }
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            if($saveFacts.Count -ne 3){throw 'Config save observer unavailable.'}
            $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
            Check ($id+'.NoUnnecessaryConfigSave') ([int]$saveFacts[0] -eq 0)
            Check ($id+'.ConfigBytesPreserved') ($configBefore.FileHash -ceq $configAfter.FileHash)
            Test-InventoryConfigPresentation $Fixture $id
            Check ($id+'.AuthBytesPreserved') ($auth -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))))
            $savePhase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
            Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $case $savePhase
        } finally {$hostBook.Close($false)}
    }
    Test-InventoryConfigDiagnosticPreservation
}
