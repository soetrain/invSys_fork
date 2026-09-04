[CmdletBinding()]
param()

$ErrorActionPreference = "Stop"
Set-StrictMode -Version 2.0

$repo = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path

function Read-Source([string]$relativePath) {
    Get-Content -Raw -LiteralPath (Join-Path $repo $relativePath)
}

$role = Read-Source "src\Core\Modules\modRoleEventWriter.bas"
$nas = Read-Source "src\Core\Modules\modNasConnection.bas"
$form = Read-Source "src\Core\Forms\frmWarehouseConnection.frm"
$failures = [System.Collections.Generic.List[string]]::new()
$passes = [System.Collections.Generic.List[string]]::new()

function Check([string]$name, [bool]$passed, [string]$contract) {
    if ($passed) {
        $passes.Add($name)
        Write-Host "PASS $name"
    } else {
        $failures.Add("${name}: ${contract}")
        Write-Host "FAIL $name - $contract"
    }
}

$connectHandler = [regex]::Match($role, '(?ms)^Public Sub ConnectWarehouseStorageForCapability.*?^End Sub').Value
$firstUseHelper = [regex]::Match($role, '(?ms)^Private Function ShouldPromptForFirstUseServerConnectionRole.*?^End Function').Value

Check "PublicHandler.FirstUseFallback" (
    $connectHandler.Contains('ShouldPromptForFirstUseServerConnectionRole(statusText)') -and
    $connectHandler.Contains('ShowWarehouseConnectionPromptForTarget(FirstUseServerConnectionPromptRole(), requireStationInbox)')
) "The public Server Sign In handler must offer the Core connection form when the current profile has no remembered NAS root."

Check "FirstUseFallback.OnlyNoRememberedRoot" (
    $firstUseHelper.Contains('no remembered nas root') -and
    -not $firstUseHelper.Contains('credential rejected')
) "First-use prompting must be limited to the no-remembered-root state; the existing credential-recovery path remains distinct."

Check "FirstUseFallback.NonAdminGuidance" (
    $connectHandler.Contains('Use Server Sign In to enter an authorized NAS warehouse root') -and
    -not $connectHandler.Contains('Use Admin/setup to add or repair the warehouse server root')
) "A non-Admin first-use failure must direct the user to Server Sign In, not require Admin setup."

Check "ConnectionForm.ScansExistingWindowsSession" (
    $form.Contains('statusCode = modNasConnection.TryRevalidateRememberedRoot(rootPath)') -and
    $form.Contains('statusCode = modNasConnection.SelectWarehouseTarget(') -and
    $form.Contains('mTxtPassword.Value = vbNullString')
) "The shared form must scan with current Windows access, select one target through Core, and clear the entered password."

Check "ConnectionRoot.NoAdminOnlyMessage" (
    -not $nas.Contains('Use Admin > Add Warehouse Root or setup to save the server path.') -and
    -not $nas.Contains('Use Admin/setup to save the server path.')
) "No-root Core status text must not claim that only Admin can establish the first connection."

Write-Host "RESULT passed=$($passes.Count) failed=$($failures.Count)"
if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Host "  $_" }
    exit 1
}
