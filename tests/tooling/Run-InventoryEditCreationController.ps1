param(
    [ValidateSet('RED','GREEN')][string]$Phase='GREEN',
    [ValidateSet('All','Safety','QuantityPhase','InterveningInputs','QueuedInputs','QueuedLocation','QueuedQuantityMode','RetainedReceipts','Success','QueueFailure','ApplyFailure','ResultSaveFailure','AuthorityLost','ZeroQuantity','DuplicateRows','RetiredCatalog','BeforeCatalogSaveThrows','AfterCatalogSaveThrows','ReplacementSession','OtherWarehouseAdmin','RetiredHistory','InterveningEntity','InterveningCatalog','InterveningLocation','InterveningQuantityMode')][string]$Case='ApplyFailure',
    [string]$Deploy='deploy/validation-inventory-config-03',
    [string]$CodeRoot=''
)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$sourceRoot=Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
if(-not $CodeRoot){$CodeRoot=if((Split-Path $sourceRoot -Leaf) -ceq 'invSys_fork'){$sourceRoot}else{Join-Path (Split-Path $sourceRoot -Parent) 'invSys_fork'}}
$CodeRoot=(Resolve-Path -LiteralPath $CodeRoot).Path
if([IO.Path]::IsPathRooted($Deploy)){throw 'Deploy must be relative to the code root.'}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be closed.'}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$root=Join-Path $CodeRoot ('reports/runtime/inventory-edit-creation/'+$Phase.ToLowerInvariant()+'-'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$settings=Get-InvSysTestSettingsSnapshot
$deployPath=Join-Path $CodeRoot $Deploy
$pins=@(Get-ChildItem -LiteralPath $deployPath -File -Filter '*.xlam'|ForEach-Object{[pscustomobject]@{File=$_.FullName;Hash=(Get-FileHash -LiteralPath $_.FullName).Hash}})
if($pins.Count -ne 5){throw 'Five packages required.'}
$pins|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root 'package-pins.json')
$code=1;$start=[DateTimeOffset]::UtcNow
Write-Output ('Inventory recording controller: '+$root)
try {
    & powershell.exe -NoProfile -NonInteractive -File (Join-Path $PSScriptRoot 'Test-Slice4beConfigCommands.ps1') -RepoRoot $CodeRoot -DeployRoot $Deploy -Phase $Phase -AdminInventoryEntryOnly -AdminInventoryEditCreationOnly -AdminInventoryEditCreationCase $Case -CompileEvaluationProbesForTest -TraceInventoryCallsForTest -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40 -DesktopMonitorReport (Join-Path $CodeRoot 'reports/runtime/inventory-config-preservation-desktop.json') *> (Join-Path $root 'worker.log')
    $code=$LASTEXITCODE
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $same=@($pins|Where-Object{(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$code;ExcelClosed=$true;SettingsRestored=$restored;PackagesPreserved=$same;FullAcceptance=$false}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $root 'closure.json')
    if(-not($restored -and $same)){throw 'Preservation failed.'}
}
Get-Content -LiteralPath (Join-Path $root 'closure.json')
exit $code
