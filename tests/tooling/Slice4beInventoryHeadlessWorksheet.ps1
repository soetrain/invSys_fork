# Explicit provisioned service calls remain distinct from interactive sessions.
function Install-InventoryHeadlessWorksheetProbe {
    $probe=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Add(1);$probe.Name='TestInventoryHeadlessWorksheet'
    $probe.CodeModule.AddFromString(@'
Option Explicit
Private mEntered As Boolean
Public Function Grant(ByVal warehouse As String, ByVal user As String) As Boolean
    If Not modAuth.LoadAuth(warehouse) Then Exit Function
    Grant = modAuth.HasProvisionedCapabilityForSystem("ADMIN_MAINT", user, warehouse, "S1")
End Function
Public Function Execute(ByVal warehouse As String, ByVal user As String, ByVal sku As String, _
        Optional ByVal action As String = "ADD") As String
    Dim row As Object, fields As Object, key As Variant, report As String, succeeded As Boolean
    Set row = CreateObject("Scripting.Dictionary")
    row("TableRow") = 1: row("Action") = action: row("ItemCode") = sku
    row("ItemName") = "Provisioned service fixture": row("UOM") = "LB"
    row("QtyMode") = "COUNTED": row("Quantity") = 0#: row("HasQuantity") = True
    row("Location") = "CLEARVIEW"
    For Each key In Array("Category", "Description", "Vendors", "VendorCode", "ExternalCode", "Picture", "EditReason")
        row(CStr(key)) = ""
    Next key
    If action = "EDIT" Then row("ItemName") = "Provisioned catalog edit": row("EditReason") = "Explicit service edit"
    Set fields = CreateObject("Scripting.Dictionary")
    fields("TRACK_QTY") = "TRUE": fields("ITEM_KIND") = "INVENTORY"
    Set row("CustomFields") = fields
    mEntered = True
    succeeded = modAdmin.ApplyInventoryWorksheetRecordForWarehouse(warehouse, "S1", user, row, report)
    Execute = CStr(mEntered) & "|" & CStr(succeeded) & "|" & CStr(InStr(1, report, "Event:", vbTextCompare) > 0)
End Function
Public Function CatalogWithoutActor(ByVal warehouse As String, ByVal sku As String) As Boolean
    Dim report As String
    CatalogWithoutActor = modAdmin.UpdateInventoryItemCatalogForWarehouse(warehouse, sku, _
        "Missing actor must refuse", "LB", "CLEARVIEW", , , , , , , , report)
End Function
'@)
}
function Get-HeadlessCatalogItem($Fixture,$Sku) {
    $path=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $book=$excel.Workbooks.Open($path,0,$true)
    try{
        $table=Table $book 'tblSkuCatalog'
        foreach($row in $table.ListRows){
            if([string]$row.Range.Cells.Item(1,$table.ListColumns.Item('ITEM_CODE').Index).Value2 -ceq $Sku){
                return [string]$row.Range.Cells.Item(1,$table.ListColumns.Item('ITEM').Index).Value2
            }
        }
        throw 'Generated catalog fixture missing.'
    }finally{$book.Close($false)}
}
function Test-InventoryHeadlessWorksheet($Fixture,$Other) {
    $script:InventoryConfigReadChecks=[Collections.Generic.List[object]]::new()
    foreach($case in @('ProvisionedAdminSignedOut','DeniedReaderSignedOut')){
        SelectTarget $Fixture
        $authorized=$case -ceq 'ProvisionedAdminSignedOut';$user=if($authorized){'config-admin'}else{'config-reader'}
        $id='InventoryHeadlessWorksheet.'+$case
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        Check ($id+'.NoInteractiveSession') (-not [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'))
        $grant=[bool](Run 'invSys.Admin.xlam' 'TestInventoryHeadlessWorksheet.Grant' @($Fixture.Warehouse,$user))
        Check ($id+'.ProvisionedGrantMatchesCase') ($grant -eq $authorized)
        if($grant -ne $authorized){throw 'Provisioned fixture grant unavailable; not product RED.'}
        $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
        $authBefore=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
        [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
        $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryHeadlessWorksheet.Execute' @($Fixture.Warehouse,$user,('DEMO-HEADLESS-'+$case.ToUpperInvariant())))).Split('|')
        Check ($id+'.ActualProvisionedServiceInvoked') ($facts.Count -eq 3 -and $facts[0] -ceq 'True')
        if($facts.Count -ne 3){throw 'Provisioned service facts unavailable.'}
        Check ($id+'.ExpectedServiceOutcome') (($facts[1] -ceq 'True') -eq $authorized)
        $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $newKeys=@($after.EntityQuantities.Keys|Where-Object {-not $before.EntityQuantities.ContainsKey($_)})
        if($authorized){
            Check ($id+'.ExactZeroCreateSaved') ($newKeys.Count -eq 1 -and $after.EntityQuantities[$newKeys[0]] -eq 0 -and $after.LogRows -eq $before.LogRows+1 -and $after.AppliedRows -eq $before.AppliedRows+1)
            Check ($id+'.ExactEventOutcomeReported') ($facts[2] -ceq 'True')
        }else{
            Check ($id+'.NoCanonicalOutput') ($newKeys.Count -eq 0 -and $after.FileHash -ceq $before.FileHash)
        }
        $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
        $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
        if($saveFacts.Count -ne 3){throw 'Config save facts unavailable.'}
        $phase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
        if(-not $authorized){Check ($id+'.NoConfigSaveBeforeDeniedAuthority') ($phase.Calls -eq 0)}
        Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $case $phase
        [pscustomobject]@{Case=$case;AuthBytesPreserved=($authBefore -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))))}|
            ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'headless-authority-observations.jsonl')
    }
    $sku='DEMO-HEADLESS-PROVISIONEDADMINSIGNEDOUT'
    foreach($case in @('ReaderCatalogEdit','MissingActorCatalog','AdminCatalogEdit','ReaderQuantityNoop','MissingActorQuantityNoop','AdminQuantityNoop')){
        SelectTarget $Fixture
        [void](Run 'invSys.Core.xlam' 'modAuth.SignOut')
        $id='InventoryHeadlessWorksheet.'+$case
        Check ($id+'.NoInteractiveSession') (-not [bool](Run 'invSys.Core.xlam' 'modAuth.IsSignedIn'))
        $before=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $itemBefore=Get-HeadlessCatalogItem $Fixture $sku
        $configBefore=Get-InventoryConfigDiagnosticSnapshot $Fixture
        $authBefore=Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))
        [void](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Reset')
        $authorized=$case.StartsWith('Admin');$user=if($authorized){'config-admin'}elseif($case.StartsWith('Reader')){'config-reader'}else{''}
        if($case -ceq 'MissingActorCatalog'){
            $succeeded=[bool](Run 'invSys.Admin.xlam' 'TestInventoryHeadlessWorksheet.CatalogWithoutActor' @($Fixture.Warehouse,$sku))
        }elseif($case.EndsWith('QuantityNoop')){
            # Public Core service call, target equals the owning zero-create quantity.
            $invocationArguments=@($Fixture.Warehouse,'S1',$user,$sku,0.0,'Explicit service no-op','')
            $succeeded=[bool](Run 'invSys.Core.xlam' 'modInventoryQuantityCommands.SetSingleEntityQuantity' $invocationArguments)
        }else{
            $facts=([string](Run 'invSys.Admin.xlam' 'TestInventoryHeadlessWorksheet.Execute' @($Fixture.Warehouse,$user,$sku,'EDIT'))).Split('|')
            if($facts.Count -ne 3 -or $facts[0] -cne 'True'){throw 'Provisioned EDIT service facts unavailable.'}
            $succeeded=$facts[1] -ceq 'True'
        }
        Check ($id+'.ExpectedServiceOutcome') ($succeeded -eq $authorized)
        $after=Get-AdminInventoryEntryFacts $Fixture -Persisted
        $itemAfter=Get-HeadlessCatalogItem $Fixture $sku
        Check ($id+'.NoQuantityOrEventMutation') ($after.LogRows -eq $before.LogRows -and $after.AppliedRows -eq $before.AppliedRows -and $after.DistinctKeys -eq $before.DistinctKeys)
        if($case -ceq 'AdminCatalogEdit'){
            Check ($id+'.ExactCatalogEditSaved') ($itemAfter -ceq 'Provisioned catalog edit')
        }else{
            Check ($id+'.CatalogAndInventoryPreserved') ($itemAfter -ceq $itemBefore -and $after.FileHash -ceq $before.FileHash)
        }
        $configAfter=Get-InventoryConfigDiagnosticSnapshot $Fixture
        $saveFacts=([string](Run 'invSys.Core.xlam' 'TestInventoryConfigSave.Facts')).Split('|')
        if($saveFacts.Count -ne 3){throw 'Config save facts unavailable.'}
        $phase=[pscustomobject]@{Calls=[int]$saveFacts[0];ByteChangesDuringSave=[int]$saveFacts[1];UnavailableHashComparisons=[int]$saveFacts[2];Procedure='modConfig.SaveConfigWorkbookIfWritable'}
        Check ($id+'.NoConfigSetupSave') ($phase.Calls -eq 0)
        Write-InventoryConfigDiagnosticComparison $configBefore $configAfter $case $phase
        [pscustomobject]@{Case=$case;AuthBytesPreserved=($authBefore -ceq (Get-AdminInventoryEntryHash (Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb'))))}|
            ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'headless-authority-observations.jsonl')
    }
    Test-InventoryConfigDiagnosticPreservation
}
