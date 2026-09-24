[CmdletBinding()]
param([string]$RepoRoot='.',[string]$DeployRoot='deploy/validation-production-quiet-boundary',
      [ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the isolated palette gate.'}
. (Join-Path $repo 'tests/tooling/Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$root=Join-Path $repo ('reports/runtime/production-palette/'+[guid]::NewGuid().ToString('N'))
$packages=Join-Path $root 'packages'
New-Item -ItemType Directory -Path $packages|Out-Null
$pins=@()
foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')){
    $source=Join-Path (Join-Path $repo $DeployRoot) $name
    $pins+=[pscustomobject]@{File=$source;Hash=(Get-FileHash -LiteralPath $source).Hash}
    Copy-Item -LiteralPath $source -Destination (Join-Path $packages $name)
}
$started=[DateTimeOffset]::UtcNow.ToString('o');$result=1
Write-Output ('Palette '+$Phase+': '+$root)
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1') -RepoRoot $repo -DeployRoot $packages.Substring($repo.Length+1) -OutputDirectory $root.Substring($repo.Length+1) -CallbackFilter Production -WorkbookState NoEligible -ProductionPaletteProbe *> (Join-Path $root 'worker.log')
    $result=$LASTEXITCODE
} finally {
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    $preserved=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{Phase=$Phase;ExitCode=$result;StartUTC=$started;EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExcelClosed=$true;SettingsRestored=$restored;PackagesPreserved=$preserved;NormalUnassistedClosureProven=$false;CleanupScope='Existing launcher harness may terminate its owned process after Quit.'}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not $restored -or -not $preserved){throw 'Palette gate preservation failed.'}
}
Write-Output ('Palette terminal: '+$result)
exit $result
