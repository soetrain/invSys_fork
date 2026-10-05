# Captured inventory lookup through the actual Print Recall handler and report owner.
function Install-ProductionPrintInventoryProbe($Module){
    $start=$Module.ProcBodyLine('RenderRecallCodesReport',0)
    $Module.InsertLines($start+1,'    TestProductionDesigner.PrintInventoryHit wsProd, invLo')
    $adapter=$packages['invSys.Operations.xlam'].VBProject.VBComponents.Item('TestProductionDesigner').CodeModule
    $adapter.InsertLines(1,'Private mPrintInventoryReads As String, mPrintLocation As String')
    $start=$adapter.ProcBodyLine('PrintPreviewForTest',0)
    $adapter.InsertLines($start+1,'    mPrintLocation = CStr(report.ListObjects("RecallCodesReport").ListColumns("LOCATION").DataBodyRange.Cells(1, 1).Value)')
    $adapter.AddFromString(@'
Public Sub PrintInventoryHit(ByVal source As Worksheet, ByVal inventory As ListObject)
    Dim kind As String
    kind = "None"
    If Not inventory Is Nothing Then
        kind = "Foreign"
        If inventory.Parent.Parent Is source.Parent Then kind = "Captured"
    End If
    mPrintInventoryReads = mPrintInventoryReads & kind & "|"
End Sub
Public Sub ResetPrintInventoryForTest()
    mPrintInventoryReads = "": mPrintLocation = ""
End Sub
Public Function PrintInventoryReadsForTest() As String
    PrintInventoryReadsForTest = mPrintInventoryReads
End Function
Public Function PrintLocationForTest() As String
    PrintLocationForTest = mPrintLocation
End Function
'@)
}

function Test-ProductionPrintInventory($Fixture,$Book,$Sheet,$Decoy){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Hash([string]$Path){$stream=[IO.File]::Open($Path,'Open','Read','ReadWrite');try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}}
    function InventoryPins {
        $result=@{}
        foreach($file in Get-ChildItem -LiteralPath $Fixture.Root -Recurse -File|Where-Object Name -NotLike '~$*'){$result[$file.FullName]=Hash $file.FullName}
        return $result
    }
    function InventoryPreserved($Before){
        $after=InventoryPins
        if($after.Count -ne $Before.Count){return $false}
        foreach($path in $Before.Keys){if(-not $after.ContainsKey($path) -or $after[$path] -cne $Before[$path]){return $false}}
        return $true
    }
    function Fingerprint($Worksheet){ConvertTo-Json -Compress -Depth 10 -InputObject @($Worksheet.UsedRange.Formula)}
    SelectTarget $Fixture
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    if(-not $seed.StartsWith('OK|')){throw 'Admin Seed prerequisite unavailable; not product RED.'}
    $authorityPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Data.Inventory.xlsb')
    $pins=InventoryPins
    $existing=@($excel.Workbooks|Where-Object {$_.FullName -ieq $authorityPath})
    if($existing.Count -gt 1){throw 'Duplicate owned fixture authority; not product RED.'}
    $opened=$existing.Count -eq 0
    if($opened){$authority=$excel.Workbooks.Open($authorityPath,0,$true)}else{$authority=$existing[0]}
    try{
        $entities=Table $authority 'tblInventoryEntities'
        if($null -eq $entities.DataBodyRange){throw 'Seed entity unavailable; not product RED.'}
        $key=[string]$entities.ListColumns.Item('System_Key').DataBodyRange.Cells.Item(1,1).Value2
        $location=[string]$entities.ListColumns.Item('Location').DataBodyRange.Cells.Item(1,1).Value2
        if(-not $key -or -not $location){throw 'Seed identity/location unavailable; not product RED.'}
    }finally{if($opened){$authority.Close($false)}}
    Check 'PrintInventory.ReadOnlySeedInspection' (InventoryPreserved $pins)
    SelectTarget $Fixture 'config-producer'
    $output=$Sheet.ListObjects.Item('ProductionOutput')
    $keyColumn=$output.ListColumns.Add();$keyColumn.Name='System_Key'
    $keyColumn.DataBodyRange.Cells.Item(1,1).Value2=$key
    $local=$Book.Worksheets.Add();$local.Name='InventoryManagement'
    $foreign=$Decoy.Worksheets.Add();$foreign.Name='InventoryManagement'
    foreach($ws in @($local,$foreign)){
        $ws.Range('A1').Value2='System_Key';$ws.Range('B1').Value2='ITEM_CODE';$ws.Range('C1').Value2='LOCATION';$ws.Range('D1').Value2='LOCAL_NOTE'
        $ws.Range('A2').Value2=$key;$ws.Range('B2').Value2='SEED-PROJECTION';$ws.Range('C2').Value2=$location;$ws.Range('D2').Formula='=12+3'
        $table=$ws.ListObjects.Add(1,$ws.Range('A1:D2'),$null,1);$table.Name='invSys'
    }
    $foreign.Range('C2').Value2='DECOY-LOCATION'
    $source=Fingerprint $Sheet;$localPin=Fingerprint $local;$foreignPin=Fingerprint $foreign
    foreach($case in @('Captured','MissingSheet','AliasCaptured','MissingTable')){
        $local.Name='InventoryManagement'
        if($case -ceq 'MissingSheet'){$local.Name='UnavailableInventory'}
        if($case -ceq 'AliasCaptured'){$local.Name='Inventory Management'}
        if($case -ceq 'MissingTable'){$local.ListObjects.Item('invSys').Unlist()}
        [void](Probe 'OpenDesigner' @($Book.Name));[void](Probe 'RunLocalShowAndCapture' @($Book.Name,'PRINT'))
        $Decoy.Activate();[void](Probe 'ResetPrintInventoryForTest');[void](Probe 'ResetPrintPreviewForTest')
        $label='PrintInventory.'+$case
        Check ($label+'.ActualHandlerReturned') ([bool](Probe 'PrintAct'))
        $available=$case -in @('Captured','AliasCaptured')
        $expected=if($available){'Captured|Captured|'}else{'None|None|'}
        $expectedLocation=if($available){$location}else{''}
        # The operator's owner runs once; explicitly retain the separate diagnostic API.
        $diagnostic=[string](Run 'invSys.Operations.xlam' 'mProduction.GetRecallPrintDiagnostic')
        Check ($label+'.ExplicitDiagnosticAvailable') ($diagnostic.StartsWith('OK; Sheet=RecallCodesPrint; Rows=1'))
        Check ($label+'.BothReadsStayCaptured') ([string](Probe 'PrintInventoryReadsForTest') -ceq $expected)
        Check ($label+'.PreviewBoundaryOnce') ([int](Probe 'PrintPreviewCountForTest') -eq 1)
        Check ($label+'.PreviewLocation') ([string](Probe 'PrintLocationForTest') -ceq $expectedLocation)
        $report=$Book.Worksheets.Item('RecallCodesPrint').ListObjects.Item('RecallCodesReport')
        Check ($label+'.DiagnosticLocation') ([string]$report.ListColumns.Item('LOCATION').DataBodyRange.Cells.Item(1,1).Value2 -ceq $expectedLocation)
        Check ($label+'.OutputAndExactKeyPreserved') ((Fingerprint $Sheet) -ceq $source)
        Check ($label+'.LocalProjectionPreserved') ((Fingerprint $local) -ceq $localPin)
        Check ($label+'.DecoyPreserved') ((Fingerprint $foreign) -ceq $foreignPin -and $Decoy.Worksheets.Count -eq 2)
        CaptureOwnedFormByCaptionEvidence 'Production' ('print-inventory-'+$case.ToLowerInvariant()+'.png')
        [void](Probe 'CloseDesigner')
    }
    Check 'PrintInventory.AuthorityPreserved' (InventoryPreserved $pins)
}
