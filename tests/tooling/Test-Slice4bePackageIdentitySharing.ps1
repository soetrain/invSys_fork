[CmdletBinding()]
param([Parameter(Mandatory=$true)][string]$DeployRoot,[ValidateSet('RED','GREEN')][string]$Phase='RED')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'Slice4beActivityAssertions.ps1')
$root=Join-Path 'reports/runtime/package-identity-sharing' ([guid]::NewGuid().ToString('N'))
$copies=Join-Path $root 'packages';New-Item -ItemType Directory -Path $copies -Force|Out-Null
$results=[Collections.Generic.List[object]]::new();$locks=[Collections.Generic.List[IO.FileStream]]::new()
function Check([string]$Name,[bool]$Passed){$results.Add([pscustomobject]@{Check=$Name;Passed=$Passed});Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))}
$pins=@(foreach($name in @('Core','Inventory.Domain','Designs.Domain','Operations','Admin')){
    $file=Join-Path $DeployRoot ('invSys.'+$name+'.xlam')
    $copy=Join-Path $copies ([IO.Path]::GetFileName($file));Copy-Item -LiteralPath $file -Destination $copy
    [pscustomobject]@{Source=$file;Copy=$copy;Hash=(Get-FileHash -LiteralPath $file).Hash}
})
Write-Output ('Package sharing evidence: '+$root)
try {
    Check 'PackageIdentity.ClosedCopiesValid' (Test-Slice4bePackageIdentity $copies)
    foreach($pin in $pins){$locks.Add([IO.FileStream]::new($pin.Copy,[IO.FileMode]::Open,[IO.FileAccess]::ReadWrite,[IO.FileShare]::ReadWrite))}
    Check 'PackageIdentity.WritableOpenCopiesStillReadable' (Test-Slice4bePackageIdentity $copies)
} finally {
    foreach($handle in $locks){$handle.Dispose()}
    Check 'PackageIdentity.CopyBytesPreserved' (@($pins|Where-Object {(Get-FileHash -LiteralPath $_.Copy).Hash -cne $_.Hash}).Count -eq 0)
    Check 'PackageIdentity.SourceBytesPreserved' (@($pins|Where-Object {(Get-FileHash -LiteralPath $_.Source).Hash -cne $_.Hash}).Count -eq 0)
    $results|ConvertTo-Json|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
}
if(@($results|Where-Object {-not $_.Passed}).Count){exit 1}
