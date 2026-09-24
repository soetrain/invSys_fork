[CmdletBinding()]
param([string]$RepoRoot='.',[string]$DeployRoot='deploy/validation-guide-layout-normalized',
      [ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be closed before the isolated probe.'}
. (Join-Path $repo 'tests/tooling/Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$root=Join-Path $repo ('reports/runtime/production-quiet-boundary/'+[guid]::NewGuid().ToString('N'))
$packages=Join-Path $root 'packages'
New-Item -ItemType Directory -Path $packages|Out-Null
$pins=@()
foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')){
    $source=Join-Path (Join-Path $repo $DeployRoot) $name
    $pins+= [pscustomobject]@{File=$source;Hash=(Get-FileHash -LiteralPath $source).Hash}
    Copy-Item -LiteralPath $source -Destination (Join-Path $packages $name)
}
$report=Join-Path $repo 'tests/unit/phase6_live_role_workflow_results.md'
$reportBefore=[IO.File]::ReadAllBytes($report)
$started=[DateTimeOffset]::UtcNow.ToString('o');$exitCode=1
Write-Output ('Production boundary '+$Phase+': '+$root)
try {
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $repo 'tools/validate_phase6_live_role_workflows.ps1') -RepoRoot $repo -DeployRoot ($packages.Substring($repo.Length+1)) -CheckProductionQuietBoundary *> (Join-Path $root 'worker.log')
    $exitCode=$LASTEXITCODE
    Copy-Item -LiteralPath $report -Destination (Join-Path $root 'results.md')
} finally {
    # Retain private settings in memory until the actual worker/Excel exits.
    Wait-RecordingCleanup -Creator $null -Worker $null
    $restored=Restore-InvSysTestSettingsSnapshot $settings
    [IO.File]::WriteAllBytes($report,$reportBefore)
    $preserved=@($pins|Where-Object {(Get-FileHash -LiteralPath $_.File).Hash -cne $_.Hash}).Count -eq 0
    [pscustomobject]@{Phase=$Phase;ExitCode=$exitCode;StartUTC=$started;EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExcelClosed=$true;SettingsRestored=$restored;PackagesPreserved=$preserved;TrackedReportRestored=([Convert]::ToBase64String([IO.File]::ReadAllBytes($report)) -ceq [Convert]::ToBase64String($reportBefore))}|ConvertTo-Json|Set-Content (Join-Path $root 'closure.json')
    if(-not $restored -or -not $preserved){throw 'Probe preservation failed.'}
}
Write-Output ('Production boundary terminal: '+$exitCode)
exit $exitCode
