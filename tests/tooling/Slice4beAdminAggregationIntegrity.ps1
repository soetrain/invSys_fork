# Additional D8/D17 integrity cases on Admin-generated disposable fixtures.
# Values stay in memory. Reports contain fixed case names and booleans only.
function Test-AggregationIntegrity($Fixture,$Other,$HostBook) {
    function Call([string]$Name,[object[]]$Values=@()){Run 'invSys.Admin.xlam' ('TestAggregation.'+$Name) $Values}
    function Core([string]$Name,[object[]]$Values=@()){Run 'invSys.Core.xlam' ('TestAggregationBoundary.'+$Name) $Values}
    function Hash([string]$Path){
        if(-not(Test-Path -LiteralPath $Path)){return ''}
        $stream=[IO.File]::Open($Path,'Open','Read',([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
        try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
    }
    function ResolveKnownAttempt {
        $pending=@(Get-ChildItem -LiteralPath (Split-Path -Parent $Fixture.Output) -File -Filter 'invSys.Aggregation.*.retry.json' -ErrorAction SilentlyContinue)
        if(-not $pending.Count){return}
        SelectTarget $Fixture
        $context=[string](Core 'CaptureDirect' @($HostBook.Name,$Fixture.Warehouse,$Fixture.Config,$Fixture.Snapshot))
        if(-not [bool](Core 'ResolveKnown' @($context,$Fixture.Snapshot))){throw 'Known-state integrity attempt could not be resolved; case preparation stopped.'}
    }
    $owned=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    foreach($path in @($Fixture.Output,$Fixture.Snapshot,$Other.Snapshot)){
        if(-not [IO.Path]::GetFullPath($path).StartsWith($owned,[StringComparison]::OrdinalIgnoreCase)){throw 'Integrity fixture escapes its owned root.'}
    }
    $sourceBytes=[IO.File]::ReadAllBytes($Fixture.Snapshot)
    $sourcePins=@{}
    foreach($f in @($Fixture,$Other)){
        foreach($path in @($f.Config,(Join-Path $f.Root ($f.Warehouse+'.invSys.Auth.xlsb')),(Join-Path $f.Root ($f.Warehouse+'.invSys.Data.Inventory.xlsb')),$f.Snapshot)){$sourcePins[$path]=Hash $path}
    }
    # Assert capture before closing the original so an opening redirection cannot
    # masquerade as a valid closed-original-workbook precondition.
    SelectTarget $Fixture
    $closed=$excel.Workbooks.Add();$closed.Windows.Item(1).Visible=$true
    if(-not [bool](Call 'OpenForm' @($closed.Name)) -or [int](Call 'SelectSource' @($Fixture.Warehouse)) -ne 1){throw 'Actual closed-original entry unavailable.'}
    Check 'Aggregation.Integrity.RibbonEntry.CapturesExactOriginalWorkbook' ([bool](Call 'SourceIs' @($closed.Name)))
    $closed.Close($false);$HostBook.Activate();$sheets=$HostBook.Worksheets.Count
    [void](Call 'Aggregate')
    Check 'Aggregation.Integrity.ClosedOriginal.RefusedBeforeWriter' ([int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls') -eq 0)
    Check 'Aggregation.Integrity.ClosedOriginal.OtherWorkbookPreserved' ($HostBook.Worksheets.Count -eq $sheets)
    [void](Call 'CloseForm')
    foreach($case in @('MultiSource','MixedChangedSelection','OrphanQuantity','NumericTextSku','ExternallyHeldOutput','UnverifiableOverlap','LocalFormulaReorder','MissingCandidateStatus','WrongCandidateProvenance','InspectionRefusal')){
        $held=$null
        try{
            SelectTarget $Fixture
            [void](Core 'Arm' @('','','','',''))
            if(-not [bool](Run 'invSys.Core.xlam' 'TestAggregationRoots.NormalizedControl' @($Fixture.Snapshot,$Fixture.Output))){throw 'Integrity output prerequisite failed; not behavioral RED.'}
            if($case -ceq 'OrphanQuantity'){
                $book=$excel.Workbooks.Open($Fixture.Snapshot,0,$false)
                try{
                    $table=Table $book 'tblInventorySnapshot';$row=$table.ListRows.Add()
                    foreach($name in @('WarehouseId','System_Key','SKU','QtyOnHand','LastAppliedAtUTC')){$row.Range.Cells.Item(1,$table.ListColumns.Item($name).Index).ClearContents()|Out-Null}
                    $row.Range.Cells.Item(1,$table.ListColumns.Item('QtyOnHand').Index).Value2=[double]1
                    $book.Save()
                }finally{$book.Close($false)}
            }
            if($case -ceq 'NumericTextSku'){
                $book=$excel.Workbooks.Open($Fixture.Snapshot,0,$false)
                try{
                    $table=Table $book 'tblInventorySnapshot';$row=$table.ListRows.Item(1)
                    $numericEntityKey=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2
                    $sku=$row.Range.Cells.Item(1,$table.ListColumns.Item('SKU').Index)
                    $sku.NumberFormat='@';$sku.Value2='001';$book.Save()
                }finally{$book.Close($false)}
            }
            if($case -ceq 'LocalFormulaReorder'){
                $book=$excel.Workbooks.Open($Fixture.Output,0,$false)
                try{
                    $table=Table $book 'tblGlobalInventorySnapshot'
                    if($table.ListRows.Count -lt 2){throw 'Formula reorder needs multiple generated entities; not behavioral RED.'}
                    $column=$table.ListColumns.Add();$column.Name='Local Entity Formula'
                    $offset=$table.ListColumns.Item('System_Key').Index-$column.Index
                    $column.DataBodyRange.FormulaR1C1='=RC['+$offset+']'
                    $table.Sort.SortFields.Clear()
                    [void]$table.Sort.SortFields.Add($table.ListColumns.Item('System_Key').Range,0,2)
                    $table.Sort.Header=1;$table.Sort.Apply();$book.Save()
                }finally{$book.Close($false)}
            }
            $before=Hash $Fixture.Output
            if($case -ceq 'ExternallyHeldOutput'){
                $held=[IO.File]::Open($Fixture.Output,'Open','Read',[IO.FileShare]::ReadWrite)
            }
            if(-not [bool](Call 'OpenForm' @($HostBook.Name))){throw 'Actual integrity form entry unavailable.'}
            $selected=[int](Call 'SelectSource' @($Fixture.Warehouse))
            if($case -ne 'OrphanQuantity' -and $selected -ne 1){throw 'Integrity source-selection prerequisite unavailable.'}
            if($case -cin @('MultiSource','MixedChangedSelection')){
                if([int](Call 'AddSource' @($Other.Warehouse)) -ne 2){throw 'Two actual source selections were not staged.'}
            }
            if($case -ceq 'MixedChangedSelection'){
                [IO.File]::WriteAllBytes($Fixture.Snapshot,[IO.File]::ReadAllBytes($Other.Snapshot))
            }
            $fault=switch($case){'MissingCandidateStatus'{'CANDIDATE_STATUS_REMOVED'}'WrongCandidateProvenance'{'CANDIDATE_SOURCE_WRONG'}'InspectionRefusal'{'SIGN_OUT'}'UnverifiableOverlap'{'IDENTITY_UNAVAILABLE'}default{''}}
            if($fault){
                $point=if($case -eq 'InspectionRefusal'){'AFTER_INSPECTION'}else{'AFTER_SAVE'}
                [void](Core 'Arm' @($fault,$point,'','',$Fixture.Output))
            }
            [void](Call 'Aggregate')
            $success=[bool](Call 'OwnerSuccess');$calls=[int](Run 'invSys.Core.xlam' 'TestAggregationRoots.Calls')
            if($fault){
                if(-not [bool](Core 'Injected')){throw 'Candidate integrity fault was not reached; not behavioral RED.'}
                Check ('Aggregation.Integrity.'+$case+'.FaultInjected') $true
            }
            if($case -cin @('MultiSource','LocalFormulaReorder','NumericTextSku')){
                Check ('Aggregation.Integrity.'+$case+'.AuthorizedCompletePersistence') $success
                if($success){
                    $book=$excel.Workbooks.Open($Fixture.Output,0,$true)
                    try{
                        $table=Table $book 'tblGlobalInventorySnapshot'
                        if($case -ceq 'MultiSource'){
                            $actual=[Collections.Generic.Dictionary[string,double]]::new([StringComparer]::Ordinal)
                            foreach($row in $table.ListRows){$id=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('WarehouseId').Index).Value2+'|'+[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2;$actual.Add($id,[double]$row.Range.Cells.Item(1,$table.ListColumns.Item('QtyOnHand').Index).Value2)}
                        }elseif($case -ceq 'LocalFormulaReorder'){
                            $preserved=$true
                            foreach($row in $table.ListRows){$preserved=$preserved -and ([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('Local Entity Formula').Index).Value2 -ceq [string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2)}
                            Check 'Aggregation.Integrity.LocalFormulaReorder.ValuesFollowEntity' $preserved
                        }else{
                            $exact=$false
                            foreach($row in $table.ListRows){
                                if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2 -ceq $numericEntityKey){$exact=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('SKU').Index).Value2 -ceq '001'}
                            }
                            Check 'Aggregation.Integrity.NumericTextSku.ExactTextPreserved' $exact
                        }
                    }finally{$book.Close($false)}
                    if($case -ceq 'MultiSource'){
                        $expected=[Collections.Generic.Dictionary[string,double]]::new([StringComparer]::Ordinal)
                        foreach($f in @($Fixture,$Other)){
                            $book=$excel.Workbooks.Open($f.Snapshot,0,$true)
                            try{$table=Table $book 'tblInventorySnapshot';foreach($row in $table.ListRows){$id=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('WarehouseId').Index).Value2+'|'+[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2;$expected.Add($id,[double]$row.Range.Cells.Item(1,$table.ListColumns.Item('QtyOnHand').Index).Value2)}}finally{$book.Close($false)}
                        }
                        $exact=$actual.Count -eq $expected.Count
                        foreach($id in $expected.Keys){$exact=$exact -and $actual.ContainsKey($id) -and $actual[$id] -eq $expected[$id]}
                        Check 'Aggregation.Integrity.MultiSource.ExactWholeSelectedSet' $exact
                    }
                }
            }else{
                Check ('Aggregation.Integrity.'+$case+'.NoFalseCompletion') (-not $success)
                Check ('Aggregation.Integrity.'+$case+'.ExistingFeedPreserved') ((Hash $Fixture.Output) -ceq $before)
                if(-not $fault -or $case -eq 'UnverifiableOverlap'){Check ('Aggregation.Integrity.'+$case+'.RefusedBeforeWriter') ($calls -eq 0)}
                if($case -eq 'InspectionRefusal'){
                    Check 'Aggregation.Integrity.InspectionRefusal.NoInventedPriorAttempt' ([bool](Call 'NoInventedRetry'))
                }
            }
        }finally{
            if($null -ne $held){$held.Dispose()}
            [void](Call 'CloseForm');[void](Core 'Arm' @('','','','',''))
            [IO.File]::WriteAllBytes($Fixture.Snapshot,$sourceBytes)
        }
        ResolveKnownAttempt
        $preserved=$true
        foreach($path in $sourcePins.Keys){$preserved=$preserved -and (Hash $path) -ceq $sourcePins[$path]}
        Check ('Aggregation.Integrity.'+$case+'.SourceAuthoritiesAndSnapshotsPreserved') $preserved
    }
}
