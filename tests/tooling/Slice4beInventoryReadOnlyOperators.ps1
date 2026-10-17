function Install-InventoryReadOnlyOperatorProbe {
    $project=$packages['invSys.Admin.xlam'].VBProject
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryReadOnlySave'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Public Sub SaveReturn(ByVal book As Workbook)
    If book.ReadOnly Then Err.Raise vbObjectError + 2150, , "Owned readonly save refusal."
    book.Save
End Sub
'@)
    $owner=$project.VBComponents.Item('modAdminInventoryWorksheet').CodeModule
    foreach($name in @('CreateInventoryWorksheetTable','SaveInventoryWorksheetWorkbook')){
        $start=$owner.ProcStartLine($name,0);$count=$owner.ProcCountLines($name,0)
        $body=$owner.Lines($start,$count)
        if([regex]::Matches($body,'(?i)\bwb\.Save\b').Count -ne 1){throw 'Readonly fixture save boundary unavailable.'}
        $body=[regex]::Replace($body,'(?i)\bwb\.Save\b','TestInventoryReadOnlySave.SaveReturn wb')
        $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
    }
    [pscustomobject]@{ReadOnlySaveSeam=$true;WritableControlUsesNormalSave=$true;NoReadOnlyOSSaveAttempted=$true}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'readonly-save-seam.json')
}

# A readonly operator must refuse before in-memory staging/structure mutation.
function Test-InventoryReadOnlyOperators($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    SelectTarget $Fixture
    $control=$excel.Workbooks.Add();$control.SaveAs((Join-Path $runRoot 'readonly-control.xlsx'),51)
    try{
        $control.Activate();$key=[string](Run 'invSys.Core.xlam' 'modInventoryCommands.Capture' @($control.Name))
        Check 'InventoryReadOnlyOperator.WritableControlCaptures' ($key -ne '')
        if($key){[void](Run 'invSys.Core.xlam' 'modInventoryCommands.Release' @($key))}
    }finally{$control.Close($false)}
    foreach($case in @('WritableCreate','ReadOnlyCreate','ReadOnlyUpload')){
        SelectTarget $Fixture
        $readonly=$case -cne 'WritableCreate';$action=if($case -ceq 'ReadOnlyUpload'){'Upload'}else{'Create'}
        $hostPath=Join-Path $runRoot ('readonly-operator-'+$case+'.xlsx')
        $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs($hostPath,51)
        if($action -ceq 'Upload'){
            $prepared=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareBoundaryTable' @($hostBook.Name))
            if(-not $prepared){throw 'Owned readonly staging fixture unavailable.'}
        }
        if($readonly){$hostBook.Close($false);$hostBook=$excel.Workbooks.Open($hostPath,0,$true)}
        try{
            $id='InventoryReadOnlyOperator.'+$case
            $hostBook.Activate()
            $tables=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})
            if($action -ceq 'Upload'){$tables[0].DataBodyRange.Cells.Item(1,1).Select()}
            $sheetsBefore=$hostBook.Worksheets.Count;$tablesBefore=$tables.Count;$rowBefore=''
            if($action -ceq 'Upload'){$rowBefore=[string]$tables[0].ListColumns.Item('Item Code').DataBodyRange.Cells.Item(1,1).Value2+'|'+[string]$tables[0].ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2+'|'+[string]$tables[0].ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2}
            if($readonly){
                $key=[string](Run 'invSys.Core.xlam' 'modInventoryCommands.Capture' @($hostBook.Name))
                Check ($id+'.CoreCaptureRefusesReadOnlyOwner') ($key -eq '')
                if($key){[void](Run 'invSys.Core.xlam' 'modInventoryCommands.Release' @($key))}
            }
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted;$hash=Get-AdminInventoryEntryHash $hostPath
            $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $authBefore=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
            [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
            Check ($id+'.FixtureAccessMatchesCase') ([bool]$hostBook.ReadOnly -eq $readonly)
            Check ($id+'.SignedInControl') ([bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,''))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,$action,'Authorized','ModalReturn',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            Check ($id+'.ActualRibbonAttemptReturned') $true
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $handler=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[2] -ceq '1' -and $facts[3] -ceq 'True'
            if(-not $readonly){Check ($id+'.WritableActualHandlerAndOwnerReached') $handler}
            [pscustomobject]@{Case=$case;ReadOnly=[bool]$hostBook.ReadOnly;ActualHandlerReached=$handler}|
                ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'readonly-operator-entry.jsonl')
            $tables=@($hostBook.Worksheets|ForEach-Object {$_.ListObjects}|Where-Object {$_.Name -like 'invSys_Inventory_*'})
            if($case -ceq 'WritableCreate'){$memory=$hostBook.Worksheets.Count -ge $sheetsBefore -and $tables.Count -eq $tablesBefore+1}
            elseif($case -ceq 'ReadOnlyCreate'){$memory=$hostBook.Worksheets.Count -eq $sheetsBefore -and $tables.Count -eq $tablesBefore}
            else{
                $rowAfter=[string]$tables[0].ListColumns.Item('Item Code').DataBodyRange.Cells.Item(1,1).Value2+'|'+[string]$tables[0].ListColumns.Item('Upload Status').DataBodyRange.Cells.Item(1,1).Value2+'|'+[string]$tables[0].ListColumns.Item('Upload Result').DataBodyRange.Cells.Item(1,1).Value2
                $memory=$rowBefore -ceq $rowAfter
            }
            Check ($id+'.ExpectedMemoryOutcome') $memory
            Check ($id+'.ExpectedPersistedOutcome') (((Get-AdminInventoryEntryHash $hostPath) -cne $hash) -eq (-not $readonly))
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            Check ($id+'.CanonicalSourceUnchanged') ($after.FileHash -ceq $before.FileHash -and $after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows)
            $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            if($saveFacts.Count -ne 3){throw 'Config save observations unavailable.'}
            $phase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
            Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $case $phase
            [pscustomobject]@{Case=$case;AuthBytesPreserved=($authBefore -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))))}|
                ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'readonly-authority-observations.jsonl')
        }finally{$hostBook.Close($false)}
    }
    Test-InventoryConfigDiagnosticPreservation
}
