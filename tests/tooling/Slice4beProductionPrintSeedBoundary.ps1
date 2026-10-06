# Native-failure isolation, not a substitute for the full Print acceptance gate.
# SeedOnly also protects Config policy/history through the actual Admin callback.
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
    function ConfigExtras {
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$true)
        try{
            $state=[ordered]@{}
            foreach($sheet in $cfg.Worksheets){
                if($sheet.Name -cnotin @('WarehouseConfig','StationConfig')){
                    $state[$sheet.Name]=@($sheet.UsedRange.Formula)
                }
            }
            ConvertTo-Json -Compress -Depth 12 -InputObject $state
        }finally{$cfg.Close($false)}
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
    if($Mode -ceq 'SeedOnly'){
        foreach($enabled in @($true,$false)){
            if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.ReadPolicyForTest' @($enabled))){throw 'Seed policy fixture unavailable; not product RED.'}
        }
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$false)
        try{
            $table=Table $cfg 'tblWarehouseConfig'
            if(@($table.ListColumns|Where-Object Name -CEQ 'Timezone').Count){throw 'Missing optional-header prerequisite absent; not product RED.'}
            $extra=$cfg.Worksheets.Add();$extra.Name='OperatorNotes'
            $extra.Cells.Item(1,1).Value2='SEED-PRESERVE';$extra.Cells.Item(2,1).Formula='=2+3'
            $cfg.Save()
        }finally{$cfg.Close($false)}
        $extrasBefore=ConfigExtras
        $policyBefore=[string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_RUN_PRINT'))
        if($policyBefore -cne 'True|False|False|2'){throw 'Saved disabled policy prerequisite absent; not product RED.'}
    }
    Mark 'BeforeSeed'
    if($Mode -ceq 'SeedOnly'){
        $seed=[string](Run 'invSys.Admin.xlam' 'modAdmin.RunDemoInventoryActionCallbackForAutomation' @($Fixture.Warehouse,'S1','config-admin','SEED'))
    }else{
        $seed=[string](Run 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($Fixture.Warehouse,'S1','config-admin'))
    }
    Mark 'AfterSeed'
    Check ($label+'.PackagedSeedAcknowledged') ($seed.StartsWith('OK|'))
    if($Mode -ceq 'SeedOnly'){
        Check ($label+'.PolicyHistoryAndCustomSheetsPreserved') ((ConfigExtras) -ceq $extrasBefore)
        Check ($label+'.DisabledPolicyStillEffective') ([string](Run 'invSys.Core.xlam' 'TestShippingCatalog.Policy' @('PRODUCTION_RUN_PRINT')) -ceq $policyBefore)
        $cfg=$excel.Workbooks.Open($Fixture.Config,0,$true)
        try{
            $table=Table $cfg 'tblWarehouseConfig'
            Check ($label+'.UnknownManagedTableColumnPreserved') ($table.ListColumns.Item('Operator Extra').DataBodyRange.Cells.Item(1,1).Value2 -ceq 'preserve')
        }finally{$cfg.Close($false)}
    }
    Check ($label+'.SavedOperatorPreserved') ((Hash $Book.FullName) -ceq $saved)
    Check ($label+'.CapturedSourcePreserved') ((ConvertTo-Json -Compress -Depth 10 -InputObject @($Book.Worksheets.Item('Production').UsedRange.Formula)) -ceq $source)
    Check ($label+'.DecoyPreserved') ((ConvertTo-Json -Compress -Depth 10 -InputObject @($Decoy.Worksheets.Item('Production').UsedRange.Formula)) -ceq $foreign)
    Check ($label+'.OtherWarehousePreserved') (RestartPinsEqual $otherPins $Other.Root)
    Mark 'Done'
}
