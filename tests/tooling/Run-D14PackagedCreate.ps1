[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$DeployRoot,
    [Parameter(Mandatory=$true)][string]$OutputRoot
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before D14 packaged test.'}
if(Test-Path -LiteralPath $OutputRoot){throw 'D14 output root already exists.'}
$deploy = (Resolve-Path -LiteralPath $DeployRoot).Path
$packageNames = @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')
foreach($name in $packageNames){if(-not(Test-Path -LiteralPath (Join-Path $deploy $name))){throw "Missing packaged input: $name"}}

New-Item -ItemType Directory -Path $OutputRoot | Out-Null
$rows = [System.Collections.Generic.List[object]]::new()
$packages = @{}
$excel = $null
$inventory = $null
$snapshot = $null
$operator = $null
$excelPid = $null
$quitIssued = $false
$start = [datetime]::UtcNow

function Release-ComObject([object]$Object){
    if($null -ne $Object){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)}catch{}}
}
function Run-Package([string]$Package,[string]$Macro,[object[]]$Arguments=@()){
    $target = "'$($packages[$Package].Name)'!$Macro"
    switch($Arguments.Count){
        0 {return $excel.Run($target)}
        1 {return $excel.Run($target,$Arguments[0])}
        2 {return $excel.Run($target,$Arguments[0],$Arguments[1])}
        3 {return $excel.Run($target,$Arguments[0],$Arguments[1],$Arguments[2])}
        4 {return $excel.Run($target,$Arguments[0],$Arguments[1],$Arguments[2],$Arguments[3])}
        5 {return $excel.Run($target,$Arguments[0],$Arguments[1],$Arguments[2],$Arguments[3],$Arguments[4])}
        6 {return $excel.Run($target,$Arguments[0],$Arguments[1],$Arguments[2],$Arguments[3],$Arguments[4],$Arguments[5])}
        default {throw 'Unsupported packaged argument count.'}
    }
}
function Check([string]$Name,[bool]$Passed){
    $rows.Add([pscustomobject]@{Name=$Name;Passed=$Passed})
    Write-Output ("{0}: {1}" -f $Name,$(if($Passed){'PASS'}else{'FAIL'}))
}
function Table([object]$Book,[string]$Name){
    $sheets = $Book.Worksheets
    for($i=1;$i -le $sheets.Count;$i++){
        $sheet = $sheets.Item($i)
        try{$found=$sheet.ListObjects.Item($Name);if($null -ne $found){return $found}}catch{}
    }
    throw "Required packaged table missing: $Name"
}
function Column([object]$ListObject,[string]$Name){
    $cols=$ListObject.ListColumns
    for($i=1;$i -le $cols.Count;$i++){
        if([string]::Equals(([string]$cols.Item($i).Name).Trim(),$Name,[StringComparison]::OrdinalIgnoreCase)){return $i}
    }
    return 0
}
function CountRows([object]$ListObject){if($null -eq $ListObject.DataBodyRange){return 0};return [int]$ListObject.ListRows.Count}
function Cell([object]$ListObject,[int]$Row,[string]$Header){
    $col=Column $ListObject $Header
    if($col -le 0 -or $Row -lt 1 -or $Row -gt (CountRows $ListObject)){return $null}
    return $ListObject.DataBodyRange.Cells.Item($Row,$col).Value2
}
function HasHeaders([object]$ListObject,[string[]]$Headers){
    foreach($header in $Headers){if((Column $ListObject $header) -le 0){return $false}}
    return $true
}
function NoRowHeader([object[]]$Books){
    foreach($book in $Books){
        $sheets=$book.Worksheets
        for($i=1;$i -le $sheets.Count;$i++){
            $tables=$sheets.Item($i).ListObjects
            for($j=1;$j -le $tables.Count;$j++){if((Column $tables.Item($j) 'ROW') -gt 0){return $false}}
        }
    }
    return $true
}
function KeyState([object]$ListObject,[bool]$RequireGood){
    $keys=[System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
    $valid=(Column $ListObject 'System_Key') -gt 0 -and (CountRows $ListObject) -gt 0
    if($RequireGood){$valid=$valid -and (Column $ListObject 'Condition') -gt 0}
    if($valid){
        for($i=1;$i -le (CountRows $ListObject);$i++){
            $key=([string](Cell $ListObject $i 'System_Key')).Trim()
            if([string]::IsNullOrWhiteSpace($key) -or -not $keys.Add($key)){$valid=$false}
            if($RequireGood -and ([string](Cell $ListObject $i 'Condition')).Trim().ToUpperInvariant() -cne 'GOOD'){$valid=$false}
        }
    }
    return [pscustomobject]@{Valid=$valid;Keys=$keys}
}
function FindRowByKey([object]$ListObject,[string]$Key){
    for($i=1;$i -le (CountRows $ListObject);$i++){
        if([string]::Equals([string](Cell $ListObject $i 'System_Key'),$Key,[StringComparison]::Ordinal)){return $i}
    }
    return 0
}

try{
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$true
    $excel.DisplayAlerts=$false
    $live=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
    if($live.Count -ne 1){throw 'The focused test did not own exactly one Excel process.'}
    $excelPid=$live[0].Id
    foreach($name in $packageNames){$packages[$name]=$excel.Workbooks.Open((Join-Path $deploy $name))}
    $warehouseId='D14'+[guid]::NewGuid().ToString('N').Substring(0,8).ToUpperInvariant()
    $warehouseRoot=Join-Path $OutputRoot 'warehouse'
    $snapshotPath=Join-Path $warehouseRoot ($warehouseId+'.invSys.Snapshot.Inventory.xlsb')
    [void](Run-Package 'invSys.Core.xlam' 'modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride' @((Join-Path $deploy 'templates')))
    $boot=[bool](Run-Package 'invSys.Admin.xlam' 'modAdminConsole.BootstrapWarehouseLocalAdmin' @($warehouseId,'D14 packaged test','S1','d14-test-admin',$warehouseRoot,''))
    if(-not $boot){throw 'Packaged Admin bootstrap failed; not behavioral RED.'}
    $operatorPath=[string](Run-Package 'invSys.Core.xlam' 'modWarehouseBootstrap.GetLastWarehouseOperatorWorkbookPath')
    if(-not(Test-Path -LiteralPath $operatorPath)){throw 'Packaged bootstrap did not provide an operator workbook.'}
    [void](Run-Package 'invSys.Core.xlam' 'modRuntimeWorkbooks.SetCoreDataRootOverride' @($warehouseRoot))
    $seed=[string](Run-Package 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($warehouseId,'S1','d14-test-admin'))
    if(-not $seed.StartsWith('OK|',[StringComparison]::OrdinalIgnoreCase)){throw 'Packaged Admin seed failed; not behavioral RED.'}
    $inventory=$excel.Workbooks.Open((Join-Path $warehouseRoot ($warehouseId+'.invSys.Data.Inventory.xlsb')))
    $operator=$excel.Workbooks.Open($operatorPath)
    $generated=[bool](Run-Package 'invSys.Core.xlam' 'modWarehouseSync.GenerateWarehouseSnapshot' @($warehouseId,$inventory,$snapshotPath))
    $refreshed=[bool](Run-Package 'invSys.Core.xlam' 'modOperatorReadModel.RefreshInventoryReadModelForWorkbook' @($operator,$warehouseId,'LOCAL'))
    if(-not $generated -or -not $refreshed){throw 'Packaged snapshot/refresh prerequisite failed; not behavioral RED.'}
    $snapshot=$excel.Workbooks.Open($snapshotPath)
    $entities=Table $inventory 'tblInventoryEntities'
    $snap=Table $snapshot 'tblInventorySnapshot'
    $display=Table $operator 'invSys'
    $required=@('System_Key','SKU','QtyOnHand','Location','Condition')
    $schema=(HasHeaders $entities $required) -and (HasHeaders $snap $required) -and (HasHeaders $display ($required+@('LastRefreshUTC','SnapshotId','SourceType','IsStale'))) -and (NoRowHeader @($inventory,$snapshot,$operator))
    Check 'D14.GeneratedSchemas.ManagedHeadersNoROW' $schema
    $before=KeyState $entities $true
    Check 'D14.Seed.UniqueKeysConditionGood' ($before.Valid -and $before.Keys.Count -gt 0)
    $snapshotKeys=KeyState $snap $true
    $operatorKeys=KeyState $display $true
    Check 'D14.RoundTrip.PreservesSystemKey' ($before.Valid -and $snapshotKeys.Valid -and $operatorKeys.Valid -and $before.Keys.SetEquals($snapshotKeys.Keys) -and $before.Keys.SetEquals($operatorKeys.Keys))
    $firstKeys=[string[]]@($before.Keys)
    $seedAgain=[string](Run-Package 'invSys.Admin.xlam' 'modAdminConsole.SeedDemoInventoryForAutomation' @($warehouseId,'S1','d14-test-admin'))
    $after=KeyState (Table $inventory 'tblInventoryEntities') $true
    $log=Table $inventory 'tblInventoryLog'
    $migration=$false
    for($i=1;$i -le (CountRows $log);$i++){if([string](Cell $log $i 'EventType') -ceq 'MIGRATION_SEED'){$migration=$true}}
    Check 'D14.AdminSeed.RepeatedCreatesUniqueKeysNoMigration' ($seedAgain.StartsWith('OK|',[StringComparison]::OrdinalIgnoreCase) -and $before.Keys.Count -eq 24 -and $after.Valid -and $after.Keys.Count -eq 24 -and $after.Keys.SetEquals($before.Keys) -and -not $migration)
    $key=$firstKeys[0]
    if((Column $display 'Custom_Local_Note') -eq 0){$newColumn=$display.ListColumns.Add();$newColumn.Name='Custom_Local_Note'}
    $row=FindRowByKey $display $key
    if($row -le 0){throw 'The target System_Key is missing before custom-column test.'}
    $display.DataBodyRange.Cells.Item($row,(Column $display 'Custom_Local_Note')).Value2='PRESERVE-ME'
    $refreshAgain=[bool](Run-Package 'invSys.Core.xlam' 'modOperatorReadModel.RefreshInventoryReadModelForWorkbook' @($operator,$warehouseId,'LOCAL'))
    $display=Table $operator 'invSys'
    $row=FindRowByKey $display $key
    $retained=$refreshAgain -and $row -gt 0 -and [string](Cell $display $row 'Custom_Local_Note') -ceq 'PRESERVE-ME'
    Check 'D14.OperatorRefresh.PreservesCustomColumn' $retained
    $operator.Save()
    $operator.Close($false)
    Release-ComObject $operator
    $operator=$excel.Workbooks.Open($operatorPath)
    $display=Table $operator 'invSys'
    $row=FindRowByKey $display $key
    Check 'D14.OperatorReopen.PreservesCustomColumn' ($row -gt 0 -and [string](Cell $display $row 'Custom_Local_Note') -ceq 'PRESERVE-ME')
}finally{
    foreach($book in @($operator,$snapshot,$inventory)){if($null -ne $book){try{$book.Close($false)}catch{};Release-ComObject $book}}
    foreach($name in @($packageNames)[($packageNames.Count-1)..0]){if($packages.ContainsKey($name)){try{$packages[$name].Close($false)}catch{};Release-ComObject $packages[$name]}}
    if($null -ne $excel){try{$excel.Quit();$quitIssued=$true}catch{};Release-ComObject $excel}
    ([pscustomobject]@{Utc=[datetime]::UtcNow.ToString('o');StartedUtc=$start.ToString('o');ExcelPid=$excelPid;QuitIssued=$quitIssued;Checks=$rows;Passed=@($rows|Where-Object Passed).Count;Failed=@($rows|Where-Object {-not $_.Passed}).Count}|ConvertTo-Json -Depth 5)|Set-Content -LiteralPath (Join-Path $OutputRoot 'result.json')
}
if($rows.Count -ne 6 -or @($rows|Where-Object {-not $_.Passed}).Count -ne 0 -or -not $quitIssued){throw 'D14 packaged worker did not complete its six assertions and Quit.'}
Write-Output 'D14 packaged worker: six assertions passed; Quit issued. Controller must verify native exit.'
