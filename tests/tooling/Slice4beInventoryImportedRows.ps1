# Existing successful row status is an untrusted projection. Exercise the same
# Upload handler with another pending row so an invalid skip cannot hide behind
# an otherwise empty table. Canonical inventory is always owning-service created.
function Install-InventoryImportedRowsProbe {
    $packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function ImportedRowRefusalForTest() As Boolean
    ImportedRowRefusalForTest = (InStr(1, mLastNotice, "No inventory or catalog changes were made.", vbTextCompare) > 0 _
        And InStr(1, mLastNotice, "receipt", vbTextCompare) > 0 _
        And InStr(1, mLastNotice, "Uploaded ", vbTextCompare) = 0)
End Function
'@)
}
function Test-InventoryImportedRows($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    foreach($case in @('ConfirmedPreviousAdd','ForgedImported','ChangedImportedQuantity','MissingImportedReceipt','ChangedImportedKey','ChangedImportedSku')){
        SelectTarget $Fixture
        $id='InventoryImportedRow.'+$case
        $hostBook=$excel.Workbooks.Add();$hostPath=Join-Path $runRoot ('imported-row-'+$case+'.xlsx');$hostBook.SaveAs($hostPath,51)
        try{
            $prepared=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))
            Check ($id+'.OperatorTablePrepared') $prepared
            if(-not $prepared){throw 'Owned table unavailable; not RED.'}
            $table=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})[0]
            $extra=$table.ListColumns.Add();$extra.Name='USER NOTE';$extra.DataBodyRange.Cells.Item(1,1).Value2='Retained projection extension'
            $hostBook.Save()
            $firstEvent='';$firstKey='';$firstResult=''
            if($case -cne 'ForgedImported'){
                $initial=Get-AdminInventoryEntryFacts $Fixture -Persisted
                [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('ImportedSetup',$hostBook.Name))
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
                [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
                $created=Get-AdminInventoryEntryFacts $Fixture -Persisted
                $newKeys=@($created.EntityQuantities.Keys|Where-Object {-not $initial.EntityQuantities.ContainsKey($_)})
                $firstResult=[string]$table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2
                $eventMatch=[regex]::Match($firstResult,'(?im)^Event:\s*([^\r\n]+)')
                $valid=$newKeys.Count -eq 1 -and $created.EntityQuantities[$newKeys[0]] -eq 0 -and $created.LogRows -eq $initial.LogRows+1 -and $created.AppliedRows -eq $initial.AppliedRows+1 -and $eventMatch.Success -and [string]$table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'IMPORTED'
                Check ($id+'.OwningPriorAddIndependentlySaved') $valid
                if(-not $valid){throw 'Prior owning ADD fixture unavailable; not RED.'}
                $firstKey=[string]$newKeys[0];$firstEvent=$eventMatch.Groups[1].Value.Trim()
            }
            # Copy only user business input into a second pending ADD. Generated
            # Item Code and all status/result fields remain empty for that row.
            foreach($header in @('Action','Item Name','UOM','Qty Mode','Quantity','Default Location','Category','Description','Vendor(s)','Vendor Code','External Code','Picture Path/URL')){
                # Use one stable COM setter type across this mixed-column loop.
                $table.ListColumns.Item($header).DataBodyRange.Cells.Item(2,1).Value2=[string]$table.ListColumns.Item($header).DataBodyRange.Cells.Item(1,1).Value2
            }
            $table.ListColumns.Item('Item Name').DataBodyRange.Cells.Item(2,1).Value2='Pending second fixture item'
            switch -CaseSensitive ($case){
                'ForgedImported' {
                    $table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2='IMPORTED'
                    $table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2='Unproved success label'
                }
                'ChangedImportedQuantity' {$table.ListColumns.Item('Quantity').DataBodyRange.Cells.Item(1,1).Value2=5.0}
                'MissingImportedReceipt' {$table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2=''}
                'ChangedImportedKey' {$table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2=$firstResult.Replace($firstKey,'UNCONFIRMED-RESULT-KEY')}
                'ChangedImportedSku' {$table.ListColumns.Item('Item Code').DataBodyRange.Cells.Item(1,1).Value2='DEMO-ALTERED-IMPORTED-CHOICE'}
            }
            $hostBook.Save();$table.DataBodyRange.Cells.Item(1,1).Select()
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $authBefore=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
            [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
            [void](Run 'invSys.Admin.xlam' 'TestInventoryWorksheetResult.Arm' @('ImportedRetry',$hostBook.Name))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Upload','Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualUploadHandlerReached') $ready
            if(-not $ready){throw 'Imported-row Upload action unavailable; not RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $newKeys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
            if($case -ceq 'ConfirmedPreviousAdd'){
                Check ($id+'.OnlyPendingRowAppliedExactlyOnce') ($newKeys.Count -eq 1 -and $after.EntityQuantities[$newKeys[0]] -eq 0 -and $after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1 -and $after.EntityQuantities[$firstKey] -eq 0)
                Check ($id+'.ConfirmedReceiptPreserved') ([string]$table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2 -ceq $firstResult -and [string]$table.ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'IMPORTED')
            }else{
                Check ($id+'.WholeSelectionRefusedBeforeCanonicalMutation') ($after.FileHash -ceq $before.FileHash -and $after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows -and $newKeys.Count -eq 0)
                Check ($id+'.UnconfirmedReceiptReportedHonestly') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.ImportedRowRefusalForTest'))
                if($firstEvent -ne '' -and $case -cne 'MissingImportedReceipt'){
                    $result=[string]$table.ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2
                    Check ($id+'.ExistingExactEventRetainedForReconciliation') $result.Contains($firstEvent)
                }
            }
            Check ($id+'.UnknownProjectionColumnPreserved') ([string]$table.ListColumns.Item('USER NOTE').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Retained projection extension')
            $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            if($saveFacts.Count -ne 3){throw 'Config save observer unavailable.'}
            $phase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
            Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $case $phase
            [pscustomobject]@{Case=$case;AuthBytesPreserved=($authBefore -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))))}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'imported-row-authority-observations.jsonl')
        }finally{$hostBook.Close($false)}
    }
    Test-InventoryConfigDiagnosticPreservation
}
