# Execute source Core classes/facade and source Admin catalog helper in unsaved
# probes. The save-return seam tests persisted proof; ordinary Save is the control.
function Install-InventoryCatalogRealProbe {
    $project=$packages['invSys.Core.xlam'].VBProject
    foreach($name in @('cInventoryTableInput','cInventoryCommandInputs','cInventoryCreatePlan','cInventoryRetirementPlan')){
        $text=Get-Content (Join-Path $repo ('src/Core/ClassModules/'+$name+'.cls')) -Raw
        $component=$project.VBComponents.Add(2);$component.Name=$name
        $component.CodeModule.AddFromString($text.Substring($text.IndexOf('Option Explicit')))
    }
    foreach($name in @('modInventoryCommands','modAggregationFiles')){
        $text=Get-Content (Join-Path $repo ('src/Core/Modules/'+$name+'.bas')) -Raw
        $module=$project.VBComponents.Item($name).CodeModule
        $module.DeleteLines(1,$module.CountOfLines)
        $module.AddFromString($text.Substring($text.IndexOf('Option Explicit')))
    }
    $admin=$packages['invSys.Admin.xlam'].VBProject
    $source=$admin.VBComponents.Item('TestCatalogOwner').CodeModule
    $text=$source.Lines(1,$source.CountOfLines)
    $text=$text.Substring(0,$text.IndexOf('Public Function PrepareCatalogFixture'))
    if([regex]::Matches($text,'(?m)^    wb\.Save\s*$').Count -ne 1){throw 'Catalog save seam changed.'}
    $text=$text.Replace('TestCatalogPhase.','modInventoryCommands.').Replace('    wb.Save','    TestCatalogReal.SaveReturn wb')
    $owner=$admin.VBComponents.Add(1);$owner.Name='TestCatalogRealOwner';$owner.CodeModule.AddFromString($text)
    $probe=$admin.VBComponents.Add(1);$probe.Name='TestCatalogReal'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mMode As String
Public Sub SaveReturn(ByVal book As Workbook)
    If mMode = "MemoryOnlyMarkedSaved" Then book.Saved = True Else book.Save
    If mMode = "SavedAuthorityLost" Then modAuth.SignOut
End Sub
Public Function Execute(ByVal host As String, ByVal warehouse As String, ByVal sku As String, ByVal mode As String) As String
    Dim key As String, fields As Object, report As String, afterReport As String
    Dim succeeded As Boolean, partial As Boolean, valid As Boolean
    mMode = mode
    Application.Workbooks(host).Activate
    key = modInventoryCommands.Capture(host)
    If key <> "" Then
        Set fields = CreateObject("Scripting.Dictionary")
        fields.Add "PHASE CUSTOM", "planned custom"
        fields.Add "EDIT_HISTORY_APPEND", "Catalog phase test"
        succeeded = TestCatalogRealOwner.UpdateInventoryItemCatalogForWarehouse(warehouse, sku, _
            "Catalog phase updated", "LB", "CLEARVIEW", "planned description", "planned vendor", _
            "planned code", "planned category", "planned external", "one.png;two.png", fields, report, key)
        partial = (InStr(1, report, "partial", vbTextCompare) > 0)
        valid = modInventoryCommands.CheckForm(key, afterReport)
        modInventoryCommands.Release key
    End If
    Execute = CStr(key <> "") & "|" & CStr(succeeded) & "|" & CStr(partial) & "|" & CStr(valid)
End Function
'@)
    [pscustomobject]@{SourceComponentOnly=$true;ActualCoreFacade=$true;ActualInputClasses=$true;
        SaveReturnSeam=$true;PackagedAction=$false;XlamBuilt=$false}|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'catalog-real-source-scope.json')
}

function Test-InventoryCatalogReal($Fixture) {
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    Check 'CatalogReal.GeneratedFixtureSeedApplied' $seed.StartsWith('OK|')
    if(-not $seed.StartsWith('OK|')){throw 'Generated real-Core fixture unavailable.'}
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    foreach($book in @($excel.Workbooks)){if($book.FullName -ieq $path){$book.Close($false)}}
    $baseline=Join-Path $runRoot 'catalog-real-generated-baseline.xlsb';Copy-Item $path $baseline
    $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot 'catalog-real-host.xlsx'),51)
    try{
        foreach($mode in @('Authorized','MemoryOnlyMarkedSaved','SavedAuthorityLost')){
            Copy-Item $baseline $path -Force
            SelectTarget $Fixture
            $book=$excel.Workbooks.Open($path,0,$false)
            try{
                $sku=[string](Run 'invSys.Admin.xlam' 'TestCatalogOwner.PrepareCatalogFixture' @($book.Name,$false));$book.Save()
                if(-not $sku){throw 'Generated catalog row unavailable.'}
                $hash=Get-AdminInventoryEntryHash $path
                $facts=([string](Run 'invSys.Admin.xlam' 'TestCatalogReal.Execute' @($hostBook.Name,$Fixture.Warehouse,$sku,$mode))).Split('|')
                if($facts.Count -ne 4){throw 'Real catalog facts unavailable.'}
                $id='CatalogReal.'+$mode
                Check ($id+'.CapturedSourceContext') ($facts[0] -eq 'True')
                if($mode -eq 'Authorized'){
                    Check ($id+'.Succeeded') ($facts[1] -eq 'True')
                    Check ($id+'.ConfirmedContextRemainsValid') ($facts[3] -eq 'True')
                }else{
                    Check ($id+'.NotSuccessful') ($facts[1] -eq 'False')
                    Check ($id+'.HonestPartial') ($facts[2] -eq 'True')
                    Check ($id+'.OldContextCannotContinue') ($facts[3] -eq 'False')
                }
            }finally{$book.Close($false)}
            $changed=(Get-AdminInventoryEntryHash $path) -cne $hash
            Check ($id+'.ExpectedPersistedOutcome') ($changed -eq ($mode -ne 'MemoryOnlyMarkedSaved'))
            $coldHash=Get-AdminInventoryEntryHash $path
            $cold=$excel.Workbooks.Open($path,0,$true)
            try{
                $table=Table $cold 'tblSkuCatalog'
                $updated=$table.ListColumns.Item('ITEM').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Catalog phase updated'
                Check ($id+'.IndependentPersistedFields') ($updated -eq ($mode -ne 'MemoryOnlyMarkedSaved'))
            }finally{$cold.Close($false)}
            Check ($id+'.ColdReadPreservedBytes') ((Get-AdminInventoryEntryHash $path) -ceq $coldHash)
        }
    }finally{$hostBook.Close($false)}
}
