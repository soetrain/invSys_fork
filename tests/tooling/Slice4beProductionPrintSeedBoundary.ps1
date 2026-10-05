# Native-failure isolation, not a substitute for the full Print acceptance gate.
function Test-ProductionPrintSeedBoundary($Fixture,$Other,$Book,$Decoy,[string]$Mode){
    function Probe([string]$Method,[object[]]$Values=@()){Run 'invSys.Operations.xlam' ('TestProductionDesigner.'+$Method) $Values}
    function Mark([string]$Stage){
        [pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');Mode=$Mode;Stage=$Stage}|
            ConvertTo-Json -Compress|Add-Content -LiteralPath (Join-Path $reportRoot 'print-seed-boundary.jsonl')
    }
    function Hash([string]$Path){
        $stream=[IO.File]::Open($Path,'Open','Read','ReadWrite')
        try{(Get-FileHash -InputStream $stream).Hash}finally{$stream.Dispose()}
    }
    $label='Diagnostic.PrintSeedBoundary.'+$Mode
    Mark 'Start'
    $otherPins=RestartPins $Other.Root
    $saved=Hash $Book.FullName
    $source=ConvertTo-Json -Compress -Depth 10 -InputObject @($Book.Worksheets.Item('Production').UsedRange.Formula)
    $foreign=ConvertTo-Json -Compress -Depth 10 -InputObject @($Decoy.Worksheets.Item('Production').UsedRange.Formula)
    if($Mode -ceq 'BareSeed'){
        Check ($label+'.ProductionProbeAbsent') (@($packages['invSys.Operations.xlam'].VBProject.VBComponents|Where-Object Name -CEQ 'TestProductionDesigner').Count -eq 0)
    }else{
        Check ($label+'.NoProductionFormLoaded') ([int](Probe 'PrintLoadedFormCountForTest') -eq 0)
    }
    Mark 'BeforeSelectTarget'
    SelectTarget $Fixture
    Mark 'BeforeSeed'
    $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    Mark 'AfterSeed'
    Check ($label+'.PackagedSeedAcknowledged') ($seed.StartsWith('OK|'))
    Check ($label+'.SavedOperatorPreserved') ((Hash $Book.FullName) -ceq $saved)
    Check ($label+'.CapturedSourcePreserved') ((ConvertTo-Json -Compress -Depth 10 -InputObject @($Book.Worksheets.Item('Production').UsedRange.Formula)) -ceq $source)
    Check ($label+'.DecoyPreserved') ((ConvertTo-Json -Compress -Depth 10 -InputObject @($Decoy.Worksheets.Item('Production').UsedRange.Formula)) -ceq $foreign)
    Check ($label+'.OtherWarehousePreserved') (RestartPinsEqual $otherPins $Other.Root)
    Mark 'Done'
}
