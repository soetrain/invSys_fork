# Near-miss references exist only in disposable projections of Admin-seeded data.
function Install-ProductionPrintIdentityProbe {
    $module=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('mProduction').CodeModule
    $module.AddFromString(@'
Public Function PrintRecallLogForTest(ByVal workbookName As String) As String
    Dim source As Worksheet, outputs As ListObject, inventory As ListObject, note As String
    Set source = Application.Workbooks(workbookName).Worksheets("Production")
    Set outputs = source.ListObjects("ProductionOutput")
    Set inventory = GetInvSysTableFromWorkbook(source.Parent)
    ApplyRecallCodesForOutput source, outputs, inventory, note
    Dim log As ListObject
    Set log = source.Parent.Worksheets("BatchCodesLog").ListObjects("Table48")
    If log.ListRows.Count <> 1 Then Err.Raise vbObjectError + 263, , "Recall log fixture did not append once."
    PrintRecallLogForTest = CStr(log.ListColumns("LOCATION").DataBodyRange.Cells(1, 1).Value)
End Function
'@)
}

function Test-ProductionPrintIdentity($Fixture,$Book,$Decoy){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function Fingerprint($Sheet){ConvertTo-Json -Compress -Depth 10 -InputObject @($Sheet.UsedRange.Formula)}
    function Pins {$result=@{};foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object Name -NotLike '~$*'){$result[$file.FullName]=Hash $file.FullName};return $result}
    function Preserved($Before){$after=Pins;if($after.Count -ne $Before.Count){return $false};foreach($path in $Before.Keys){if(-not $after.ContainsKey($path) -or $after[$path] -cne $Before[$path]){return $false}};return $true}
    $key=[string]$Book.Worksheets.Item('Production').ListObjects.Item('ProductionOutput').ListColumns.Item('System_Key').DataBodyRange.Cells.Item(1,1).Value2
    $location=[string]$Book.Worksheets.Item('InventoryManagement').Range('C2').Value2
    $nearKey=$key.ToLowerInvariant();if($nearKey -ceq $key){$nearKey=$key.ToUpperInvariant()}
    if(-not $key -or -not $location -or $nearKey -ceq $key){throw 'Admin-seeded identity fixture unavailable; not product RED.'}
    $foreign=Fingerprint $Decoy.Worksheets.Item('InventoryManagement')
    foreach($case in @('Exact','CaseMismatch','OutputPadding','InventoryPadding','ShiftedTable','ReorderedTable')){
        $work=$null;$priorVisible=$excel.Visible
        try{
            SelectTarget $Fixture 'config-producer'
            $label='PrintIdentity.'+$case;$path=Join-Path $runRoot ($label+'.xlsb')
            if(Test-Path -LiteralPath $path){throw 'Preserve existing identity fixture.'}
            $Book.SaveCopyAs($path);$saved=Hash $path;$work=$excel.Workbooks.Open($path,0,$false)
            $source=$work.Worksheets.Item('Production');$inventory=$work.Worksheets.Item('InventoryManagement')
            if($inventory.ListObjects.Count -ne 0){throw 'Unlisted local projection prerequisite missing.'}
            $reference=$key;if($case -ceq 'CaseMismatch'){$reference=$nearKey};if($case -ceq 'OutputPadding'){$reference=' '+$key+' '}
            $source.ListObjects.Item('ProductionOutput').ListColumns.Item('System_Key').DataBodyRange.Cells.Item(1,1).Value2=$reference
            $address='A1:D2';if($case -in @('ShiftedTable','ReorderedTable')){$address='F4:I5';$inventory.Range('A1:D2').Copy($inventory.Range('F4'))}
            $range=$inventory.Range($address)
            if($case -ceq 'InventoryPadding'){$range.Cells.Item(2,1).Value2=' '+$key+' '}
            if($case -ceq 'ReorderedTable'){
                $range.Cells.Item(1,1).Value2=' location ';$range.Cells.Item(1,2).Value2='LOCAL_NOTE'
                $range.Cells.Item(1,3).Value2='System_Key';$range.Cells.Item(1,4).Value2='ITEM_CODE'
                $range.Cells.Item(2,1).Value2=$location;$range.Cells.Item(2,2).Formula='=12+3'
                $range.Cells.Item(2,3).Value2=$key;$range.Cells.Item(2,4).Value2='SEED-PROJECTION'
            }
            $table=$inventory.ListObjects.Add(1,$range,$null,1);$table.Name='invSys'
            $sourceBefore=Fingerprint $source;$inventoryBefore=Fingerprint $inventory;$authority=Pins
            [void](Probe 'OpenDesigner' @($work.Name));[void](Probe 'RunLocalShowAndCapture' @($work.Name,'PRINT'))
            [void](Probe 'ResetPrintPreviewForTest');[void](Probe 'ResetPrintInventoryForTest');$Decoy.Activate()
            Check ($label+'.ActualHandlerReturned') ([bool](Probe 'PrintAct'))
            Check ($label+'.OneOwnerEntry') ([int](Probe 'PrintOwnerEntries') -eq 1)
            Check ($label+'.OneReportRead') ([int](Probe 'PrintReportReads') -eq 1)
            Check ($label+'.OnePreviewBoundary') ([int](Probe 'PrintPreviewCountForTest') -eq 1)
            $expected=if($case -in @('Exact','ShiftedTable','ReorderedTable')){$location}else{''}
            Check ($label+'.ExactLocation') ([string](Probe 'PrintLocationForTest') -ceq $expected)
            Check ($label+'.TruthfulStatus') ([string](Probe 'PrintStatus') -ceq 'Print preview closed.')
            Check ($label+'.SourceKeyAndCustomFieldsPreserved') ((Fingerprint $source) -ceq $sourceBefore)
            Check ($label+'.InventoryProjectionPreserved') ((Fingerprint $inventory) -ceq $inventoryBefore)
            Check ($label+'.DecoyPreserved') ((Fingerprint $Decoy.Worksheets.Item('InventoryManagement')) -ceq $foreign)
            Check ($label+'.AuthorityPreserved') (Preserved $authority)
            # Supplement the actual Print handler with the shared recall-log caller.
            $logSheet=$work.Worksheets.Add();$logSheet.Name='BatchCodesLog'
            $headers=@('RECIPE','RECIPE_ID','PROCESS','OUTPUT','LOCATION')
            for($column=0;$column -lt $headers.Count;$column++){$logSheet.Cells.Item(1,$column+1).Value2=$headers[$column]}
            $logTable=$logSheet.ListObjects.Add(1,$logSheet.Range('A1:E2'),$null,1);$logTable.Name='Table48';$logTable.ListRows.Item(1).Delete()
            $boxes=@($source.Shapes|Where-Object Name -CEQ 'CHK_RECALL_1')
            if($boxes.Count -gt 1){throw 'Ambiguous recall checkbox fixture.'}
            if($boxes.Count){$checkbox=$boxes[0]}else{$checkbox=$source.Shapes.AddFormControl(1,5,5,12,12);$checkbox.Name='CHK_RECALL_1'}
            $checkbox.ControlFormat.Value=1
            $recall=$source.ListObjects.Item('ProductionOutput').ListColumns.Item('RECALL CODE').DataBodyRange.Cells.Item(1,1)
            $priorRecall=[string]$recall.Value2;$recall.Value2=''
            $logLocation=[string](Run 'invSys.Operations.xlam' 'mProduction.PrintRecallLogForTest' @($work.Name))
            Check ($label+'.RecallLogExactLocation') ($logLocation -ceq $expected)
            Check ($label+'.RecallLogGeneratedCode') ([string]$recall.Value2 -clike 'RC-*')
            $recall.Value2=$priorRecall
            Check ($label+'.RecallLogPreservesOtherSourceFields') ((Fingerprint $source) -ceq $sourceBefore)
            Check ($label+'.RecallLogPreservesAuthority') (Preserved $authority)
            [void](Probe 'CloseDesigner')
            if($case -in @('ShiftedTable','ReorderedTable')){
                $excel.Visible=$true;$work.Activate();$work.Worksheets.Item('RecallCodesPrint').Activate()
                Initialize-SettingsCapture
                [InvSysSettingsCapture]::SaveVisibleWindow([IntPtr]$work.Windows.Item(1).Hwnd,(Join-Path $reportRoot ($label.ToLowerInvariant()+'.png')))
            }
            $work.Close($false);$work=$null
            Check ($label+'.SavedFixtureBytesPreserved') ((Hash $path) -ceq $saved)
        }finally{
            [void](Probe 'CloseDesigner')
            if($null -ne $work){$work.Close($false)}
            $excel.Visible=$priorVisible
        }
    }
}
