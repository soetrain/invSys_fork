[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$ReleaseId,
    [string]$CacheRoot = (Join-Path $env:LOCALAPPDATA "invSys\\Addins"),
    [string]$RegisterScriptPath = (Join-Path $PSScriptRoot "register_current_addins.ps1"),
    [string]$ExcelProcessName = "EXCEL",
    [string]$ExcelOptionsKey = "HKCU:\\Software\\Microsoft\\Office\\16.0\\Excel\\Options",
    [string]$AddinManagerKey = "HKCU:\\Software\\Microsoft\\Office\\16.0\\Excel\\Add-in Manager",
    [Parameter(Mandatory = $true)][ValidateSet("FailedUpdate", "OperatorIssue", "ApprovedCorrectiveAction", "CompatibilityRollback")][string]$ReasonCode,
    [switch]$ConfirmRollback
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "invsys_release_common.ps1")
if (-not $ConfirmRollback) { throw "Rollback is deliberate. Re-run with -ConfirmRollback after choosing the target release." }
$principal = [Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()
if (-not $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { throw "Rollback requires a local Windows administrator." }
if (Get-Process -Name $ExcelProcessName -ErrorAction SilentlyContinue) { throw "Close all Excel windows before rolling back invSys add-ins." }
Assert-InvSysReleaseId -ReleaseId $ReleaseId
$release = Join-Path (Join-Path $CacheRoot "Releases") $ReleaseId
[void](Assert-InvSysReleaseManifest -ReleaseDirectory $release -ExpectedReleaseId $ReleaseId)
$priorReleaseId = ""
$pointerPath = Join-Path $CacheRoot "current-release.json"
if (Test-Path -LiteralPath $pointerPath) { $priorReleaseId = [string](Get-InvSysReleasePointer -Root $CacheRoot).releaseId }
& $RegisterScriptPath -AddinsRoot $release -ExcelOptionsKey $ExcelOptionsKey -AddinManagerKey $AddinManagerKey -ExcelProcessName $ExcelProcessName
Write-InvSysJsonAtomic -Path $pointerPath -Value ([ordered]@{ releaseId = $ReleaseId; manifest = "Releases/$ReleaseId/release-manifest.json"; priorReleaseId = $priorReleaseId; rollbackReasonCode = $ReasonCode; rolledBackAtUtc = [DateTime]::UtcNow.ToString("o") })
Write-InvSysUpdateStatus -CacheRoot $CacheRoot -Status "ROLLED_BACK" -ReleaseId $ReleaseId -Detail "Local administrator selected a cached, hash-verified release. ReasonCode: $ReasonCode"
Write-Host "INVSYS_ROLLBACK_APPLIED"
Write-Host "ReleaseId=$ReleaseId"
