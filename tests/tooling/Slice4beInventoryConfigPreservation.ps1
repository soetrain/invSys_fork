# D8-INV-CONFIG-01: actual packaged actions and independent persisted Config reads.
. (Join-Path $PSScriptRoot 'Slice4beInventoryCatalogCompletion.ps1')
. (Join-Path $PSScriptRoot 'Slice4beInventoryConfigRetained.ps1')
function Install-InventoryConfigPreservationProbe {
    Install-InventoryCatalogCompletionProbe
    $packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestInventoryCatalogCompletion').CodeModule.AddFromString(@'
Public Function PrepareCatalogEditTable(ByVal host As String, ByVal tableName As String, ByVal sku As String) As Boolean
    Dim wb As Workbook, ws As Worksheet, lo As ListObject
    Set wb = Application.Workbooks(host)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If lo.Name = tableName And lo.Name Like "invSys_Inventory_*" Then
                lo.ListColumns("Action").DataBodyRange.Cells(1, 1).Value2 = "EDIT"
                lo.ListColumns("Item Code").DataBodyRange.Cells(1, 1).Value2 = sku
                lo.ListColumns("Quantity").DataBodyRange.Cells(1, 1).Value2 = 7#
                lo.ListColumns("Edit Reason").DataBodyRange.Cells(1, 1).Value2 = "Config preservation completion"
                PrepareCatalogEditTable = (lo.ListColumns("Quantity").DataBodyRange.Cells(1, 1).Value2 = 7#)
                Exit Function
            End If
        Next lo
    Next ws
End Function
Public Function PrepareMissingDefault(ByVal bookName As String) As Boolean
    Dim book As Workbook, sheet As Worksheet, table As ListObject, column As ListColumn
    Dim custom As ListColumn
    Set book = Application.Workbooks(bookName)
    For Each sheet In book.Worksheets
        For Each table In sheet.ListObjects
            If table.Name = "tblWarehouseConfig" Then
                If table.ListRows.Count <> 1 Then Exit Function
                For Each column In table.ListColumns
                    If column.Name = "USER CONFIG NOTE" Then Set custom = column
                Next column
                If custom Is Nothing Then
                    Set custom = table.ListColumns.Add
                    custom.Name = "USER CONFIG NOTE"
                End If
                custom.DataBodyRange.Cells(1, 1).Value2 = "Retained Config extension"
                table.ListColumns("Timezone").DataBodyRange.Cells(1, 1).ClearContents
                PrepareMissingDefault = True
                Exit Function
            End If
        Next table
    Next sheet
End Function
'@)
    Install-InventoryConfigRetainedProbe
}

function Set-InventoryConfigPresentation($Fixture,[double]$Width=43,[string]$SheetName='') {
    $book=$excel.Workbooks.Open($Fixture.Config,0,$false)
    try {
        if($SheetName){$sheet=$book.Worksheets.Item($SheetName)}else{$sheet=$book.Worksheets.Item(1)}
        $sheet.Columns.Item(1).ColumnWidth=$Width
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
    if($AdminInventoryConfigPreservationCase -ceq 'RetainedShared'){
        Test-InventoryConfigRetainedShared $Fixture
        return
    }
    $missingDefault=$AdminInventoryConfigPreservationCase -ceq 'MissingDefault'
    foreach($case in @('ModalAdd','WorksheetAdd','ModalCatalogEdit','WorksheetCatalogEdit')){
        if(-not $missingDefault -and $AdminInventoryConfigPreservationCase -ne 'All' -and $case -cne $AdminInventoryConfigPreservationCase){continue}
        SelectTarget $Fixture
        $id='InventoryConfigPreservation.'+$case
        if($missingDefault){$id+='MissingDefault'}
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
                    # Keep mixed string/numeric fixture writes inside VBA to avoid the
                    # observed PowerShell COM numeric-to-String assignment failure.
                    if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.PrepareCatalogEditTable' @($hostBook.Name,$table.Name,$sku))){throw 'Catalog EDIT fixture unavailable; not RED.'}
                }
                $extra=$table.ListColumns.Add();$extra.Name='USER NOTE';$extra.DataBodyRange.Cells.Item(1,1).Value2='Preserved extension'
                $hostBook.Save()
            }
            Set-InventoryConfigPresentation $Fixture
            if($missingDefault){
                # Only this run's disposable Config is changed, before observation.
                $configBook=$excel.Workbooks.Open($Fixture.Config,0,$false)
                try{
                    if(-not [bool](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.PrepareMissingDefault' @($configBook.Name))){throw 'Missing-default fixture unavailable; not RED.'}
                    $configBook.Save()
                }finally{$configBook.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($configBook);$configBook=$null}
            }
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
            if($missingDefault){
                $headers=@($configBefore.Cells.Keys|Where-Object {$_ -clike 'tblWarehouseConfig|1|*' -and $configBefore.Cells[$_] -ceq 'System.String|Timezone'})
                if($headers.Count -ne 1){throw 'Timezone observation unavailable; not RED.'}
                $expectedCell=$headers[0].Replace('|1|','|2|')
                $changes=@(@($configBefore.Cells.Keys)+@($configAfter.Cells.Keys)|Select-Object -Unique|Where-Object {
                    -not $configBefore.Cells.ContainsKey($_) -or -not $configAfter.Cells.ContainsKey($_) -or $configBefore.Cells[$_] -cne $configAfter.Cells[$_]
                })
                Check ($id+'.OnlyMissingDefaultPersisted') ($changes.Count -eq 1 -and $changes[0] -ceq $expectedCell -and $configBefore.Cells[$expectedCell] -ceq '<EMPTY>' -and $configAfter.Cells[$expectedCell] -ceq 'System.String|UTC')
                Check ($id+'.ConfigShapeAndUnknownColumnPreserved') ($configBefore.Shapes -ceq $configAfter.Shapes -and @($configAfter.Cells.Values|Where-Object {$_ -ceq 'System.String|Retained Config extension'}).Count -eq 1)
                Check ($id+'.OneNecessaryConfigSave') ([int]$saveFacts[0] -eq 1 -and [int]$saveFacts[1] -eq 1 -and [int]$saveFacts[2] -eq 0)
            }else{
                Check ($id+'.NoUnnecessaryConfigSave') ([int]$saveFacts[0] -eq 0)
                Check ($id+'.ConfigBytesPreserved') ($configBefore.FileHash -ceq $configAfter.FileHash)
            }
            Test-InventoryConfigPresentation $Fixture $id
            Check ($id+'.AuthBytesPreserved') ($auth -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))))
            $savePhase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
            $diagnosticCase=if($missingDefault){$case+'MissingDefault'}else{$case}
            Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $diagnosticCase $savePhase
        } finally {$hostBook.Close($false)}
    }
    Test-InventoryConfigDiagnosticPreservation
}
