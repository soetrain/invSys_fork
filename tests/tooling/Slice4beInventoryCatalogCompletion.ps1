# Actual modal EDIT protects the explicitly retained catalog-only completion case.
# Synthetic catalog fixture uses the Domain catalog routine in an Admin-generated
# disposable warehouse; no business import, ledger row or entity key is fabricated.
function Install-InventoryCatalogCompletionProbe {
    $packages['invSys.Inventory.Domain.xlam'].VBProject.VBComponents.Item('modInventoryApply').CodeModule.AddFromString(@'
Public Function TestCatalogOnlyFixture(ByVal bookName As String, ByVal sku As String) As Boolean
    Dim item As Object, book As Workbook, table As ListObject
    Set book = Application.Workbooks(bookName)
    Set table = FindListObjectByNameApply(book, "tblSkuCatalog")
    If table Is Nothing Or Left$(sku, 24) <> "DEMO-CATALOG-COMPLETION-" Then Exit Function
    Set item = CreateObject("Scripting.Dictionary")
    item("SKU") = sku: item("ITEM_CODE") = sku: item("ITEM") = "Synthetic catalog-only fixture"
    item("UOM") = "LB": item("LOCATION") = "CLEARVIEW": item("CATALOG_STATE") = "ACTIVE"
    EnsureSkuCatalogFromPayloadLineApply book, item
    TestCatalogOnlyFixture = (FindRowByColumnValueApply(table, "SKU", sku) > 0)
End Function
'@)
    $project=$packages['invSys.Admin.xlam'].VBProject
    $project.VBComponents.Item('TestInventoryCommands').CodeModule.AddFromString(@'
Public Function CatalogCompletionPartialForTest() As Boolean
    CatalogCompletionPartialForTest = (InStr(1, mLastNotice, "partial", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "Event:", vbTextCompare) > 0 And _
        InStr(1, mLastNotice, "saved", vbTextCompare) > 0)
End Function
'@)
    $probe=$project.VBComponents.Add(1);$probe.Name='TestInventoryCatalogCompletion'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mLose As Boolean, mReturned As Boolean, mKey As String
Public Sub Arm(ByVal lose As Boolean)
    mLose = lose: mReturned = False: mKey = ""
End Sub
Public Function QueueUnrelated(ByVal warehouse As String, ByVal key As String, ByVal sku As String) As Boolean
    Dim items As New Collection, item As Object, payload As String, eventId As String, report As String
    Set item = modRoleEventWriter.CreateInventoryEntityPayloadItem(key, sku, 1, "CLEARVIEW", "GOOD", "", "Unrelated pending fixture")
    item("IoType") = "ADJUST": item("PreviousQty") = 3#: item("TargetQty") = 4#: item("Reason") = "Unrelated pending fixture"
    items.Add item
    payload = modRoleEventWriter.BuildPayloadJsonFromCollection(items)
    QueueUnrelated = modRoleEventWriter.QueueAdminInventoryAdjustEvent(warehouse, "S1", "config-admin", _
        payload, "Unrelated pending fixture", 0, eventId, report, "")
    QueueUnrelated = QueueUnrelated And eventId <> ""
End Function
Public Sub RecordKey(ByVal key As String)
    mKey = key
End Sub
Public Sub AfterProcessor()
    mReturned = True
    If mLose Then modAuth.SignOut
End Sub
Public Function Facts(ByVal expectedKey As String) As String
    Facts = CStr(mReturned) & "|" & CStr(mKey <> "" And mKey = expectedKey)
End Function
'@)
    $componentName='modAdmin';$procedureName='CreateFirstInventoryEntityForCatalogItemAdmin'
    foreach($component in $project.VBComponents){
        if($component.Name -ceq 'modAdminInventoryCreate'){$componentName=$component.Name;$procedureName='CompleteCatalogItem';break}
    }
    $owner=$project.VBComponents.Item($componentName).CodeModule
    $start=$owner.ProcStartLine($procedureName,0)
    $count=$owner.ProcCountLines($procedureName,0)
    $body=$owner.Lines($start,$count)
    foreach($point in @(
        @('(?im)^[\t ]*systemKey = modRoleEventWriter\.CreateSystemKey\(\)[\t ]*\r?$','    TestInventoryCatalogCompletion.RecordKey systemKey'),
        @('(?im)^[\t ]*processedCount = modProcessor\.RunBatch\(warehouseId, 0, batchReport(?:, eventIdOut, applied, True)?\)[\t ]*\r?$','    TestInventoryCatalogCompletion.AfterProcessor')
    )){
        if([regex]::Matches($body,$point[0]).Count -ne 1){throw 'Catalog completion owning boundary unavailable.'}
        $mark=$point[1]
        $body=[regex]::Replace($body,$point[0],[Text.RegularExpressions.MatchEvaluator]{param($m) $m.Value.TrimEnd("`r")+"`r`n"+$mark})
    }
    $owner.DeleteLines($start,$count);$owner.InsertLines($start,$body)
}
function Get-CatalogCompletionEntityStates($Fixture,$Sku) {
    $book=$excel.Workbooks.Open((Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')),0,$true)
    try{
        $table=Table $book 'tblInventoryEntities';$states=@{}
        foreach($row in $table.ListRows){
            if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('SKU').Index).Value2 -ceq $Sku){
                $key=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('System_Key').Index).Value2
                $states[$key]=[pscustomobject]@{State=[string]$row.Range.Cells.Item(1,$table.ListColumns.Item('InventoryState').Index).Value2;Quantity=[double]$row.Range.Cells.Item(1,$table.ListColumns.Item('QtyOnHand').Index).Value2}
            }
        }
        return $states
    }finally{$book.Close($false)}
}
function Test-InventoryCatalogCompletionResults($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    foreach($case in @('Authorized','AuthorityLost')){
        SelectTarget $Fixture
        $sku='DEMO-CATALOG-COMPLETION-'+$case.ToUpperInvariant()
        $id='InventoryCatalogCompletion.'+$case
        $unrelatedSku=$sku+'-OTHER'
        $csv=Join-Path $runRoot ('catalog-completion-'+$case+'.csv')
        @('ITEM_CODE,ITEM,UOM,LOCATION,QTY,CONDITION',($unrelatedSku+',Unrelated pending fixture,LB,CLEARVIEW,3,GOOD'))|Set-Content $csv -Encoding ASCII
        $created=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.PrepareExactQuantityFixture' @($Fixture.Warehouse,$csv))
        Check ($id+'.OwningCreateFixtureApplied') $created
        if(-not $created){throw 'Owning creation fixture unavailable; not RED.'}
        $fixtureBefore=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $book=$excel.Workbooks.Open((Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')),0,$false)
        try{
            $prepared=[bool](Run 'invSys.Inventory.Domain.xlam' 'modInventoryApply.TestCatalogOnlyFixture' @($book.Name,$sku))
            if(-not $prepared){throw 'Synthetic catalog fixture routine unavailable; not RED.'}
            $book.Save()
        }finally{$book.Close($false)}
        Check ($id+'.DomainCatalogFixturePrepared') $prepared
        $fixtureAfter=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $states=Get-CatalogCompletionEntityStates $Fixture $sku
        Check ($id+'.NoEntityOrEventFabricatedForCatalogFixture') ($states.Count -eq 0 -and $fixtureAfter.DistinctKeys -eq $fixtureBefore.DistinctKeys -and $fixtureAfter.LogRows -eq $fixtureBefore.LogRows -and $fixtureAfter.AppliedRows -eq $fixtureBefore.AppliedRows)
        if($states.Count -ne 0){throw 'Catalog-only fixture unexpectedly has an entity; not RED.'}
        $unrelatedStates=Get-CatalogCompletionEntityStates $Fixture $unrelatedSku
        if($unrelatedStates.Count -ne 1){throw 'Expected owning unrelated fixture entity.'}
        $unrelatedKey=@($unrelatedStates.Keys)[0]
        $queued=[bool](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.QueueUnrelated' @($Fixture.Warehouse,$unrelatedKey,$unrelatedSku))
        Check ($id+'.UnrelatedPendingEventQueued') $queued
        if(-not $queued){throw 'Unrelated pending queue fixture unavailable; not RED.'}
        $hostBook=$excel.Workbooks.Add();$hostBook.SaveAs((Join-Path $runRoot ('catalog-completion-'+$case+'.xlsx')),51)
        try{
            $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $authBefore=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
            [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.Arm' @(($case -ceq 'AuthorityLost')))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetSelection' @($Other.Root,$sku))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.SetQuantityTarget' @(7.0))
            [void](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Execute' @($hostBook.Name,'Edit','SetQuantity','EditReason',$Fixture.Root,$Fixture.Warehouse,$Fixture.Secret))
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryCommands.Facts')).Split('|')
            $ready=$facts.Count -eq 5 -and $facts[0] -ceq 'True' -and $facts[1] -ceq 'True' -and $facts[2] -ceq '1'
            Check ($id+'.ActualEditHandlerAndReasonReached') $ready
            if(-not $ready){throw 'Catalog completion operator action unavailable; not RED.'}
            Check ($id+'.OriginalOperatorCaptured') ($facts[3] -ceq 'True')
            $shows=[long](Run 'invSys.Admin.xlam' 'TestInventoryCommands.OwnResultShowsForTest')
            Check ($id+'.ExpectedContinuation') ($shows -eq $(if($case -ceq 'Authorized'){2}else{1}))
            $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
            $newKeys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
            $newKey=if($newKeys.Count -eq 1){[string]$newKeys[0]}else{''}
            $boundary=([string](Run 'invSys.Admin.xlam' 'TestInventoryCatalogCompletion.Facts' @($newKey))).Split('|')
            if($boundary.Count -ne 2){throw 'Catalog completion observer unavailable.'}
            Check ($id+'.ProcessorReturned') ($boundary[0] -ceq 'True')
            Check ($id+'.OwningGeneratedKeyPreserved') ($boundary[1] -ceq 'True')
            Check ($id+'.ExactNewEntityAndQuantitySaved') ($newKeys.Count -eq 1 -and $after.EntityQuantities[$newKey] -eq 7 -and $after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1)
            $same=$true
            foreach($key in $before.EntityQuantities.Keys){if(-not $after.EntityQuantities.ContainsKey($key) -or $after.EntityQuantities[$key] -ne $before.EntityQuantities[$key]){$same=$false}}
            Check ($id+'.ExistingEntityIdentitiesAndQuantitiesPreserved') $same
            $saved=Get-CatalogCompletionEntityStates $Fixture $sku
            Check ($id+'.ExactActiveEntitySaved') ($saved.Count -eq 1 -and $saved.ContainsKey($newKey) -and $saved[$newKey].State -ceq 'ACTIVE' -and $saved[$newKey].Quantity -eq 7)
            if($case -ceq 'AuthorityLost'){Check ($id+'.HonestSavedPartialReported') ([bool](Run 'invSys.Admin.xlam' 'TestInventoryCommands.CatalogCompletionPartialForTest'))}
            $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
            $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
            if($saveFacts.Count -ne 3){throw 'Config save observer unavailable.'}
            $phase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
            Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $case $phase
            [pscustomobject]@{Case=$case;AuthBytesPreserved=($authBefore -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))))}|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'catalog-completion-authority-observations.jsonl')
        }finally{$hostBook.Close($false)}
    }
    Test-InventoryConfigDiagnosticPreservation
}
