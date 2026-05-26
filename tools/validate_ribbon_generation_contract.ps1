[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path $RepoRoot).Path
$buildPath = Join-Path $repo "tools/build-xlam.ps1"
$validatorPath = Join-Path $repo "tools/validate_phase6_packaged_ribbon.ps1"
$resultPath = Join-Path $repo "tests/unit/phase6_ribbon_generation_contract_results.md"

$checks = New-Object 'System.Collections.Generic.List[object]'

function Add-Check {
    param(
        [string]$Name,
        [bool]$Passed,
        [string]$Detail = ""
    )
    $checks.Add([pscustomobject]@{
        Name = $Name
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

$buildText = Get-Content -LiteralPath $buildPath -Raw
$validatorText = Get-Content -LiteralPath $validatorPath -Raw
$runtimeStatusText = Get-Content -LiteralPath (Join-Path $repo "src/Core/Modules/modRibbonRuntimeStatus.bas") -Raw
$adminConsoleText = Get-Content -LiteralPath (Join-Path $repo "src/Admin/Modules/modAdminConsole.bas") -Raw
$authText = Get-Content -LiteralPath (Join-Path $repo "src/Core/Modules/modAuth.bas") -Raw

Add-Check "Build.GetEnabledXml" ($buildText.Contains('getEnabled="RibbonRequiredCapabilityGetEnabled"')) "RequiredCapability buttons emit getEnabled."
Add-Check "Build.GetEnabledCallback" ($buildText.Contains('Public Sub RibbonRequiredCapabilityGetEnabled')) "Generated callback exists."
Add-Check "Build.GetEnabledCached" ($buildText.Contains('CanCurrentUserPerformCapabilityCached')) "Ribbon getEnabled uses cached auth/target state."
Add-Check "Build.ReceivingCapability" ($buildText.Contains('RequiredCapability = "RECEIVE_POST"')) "Receiving buttons declare capability."
Add-Check "Build.ShippingCapability" ($buildText.Contains('RequiredCapability = "SHIP_POST"')) "Shipping buttons declare capability."
Add-Check "Build.ProductionCapability" ($buildText.Contains('RequiredCapability = "PROD_POST"')) "Production buttons declare capability."
Add-Check "Build.RoleConnectServerButtons" ($buildText.Contains('btnReceivingConnectServer') -and $buildText.Contains('btnShippingConnectServer') -and $buildText.Contains('btnProductionConnectServer')) "Role ribbons expose Connect Server buttons."
Add-Check "Build.RoleSignOutButtons" ($buildText.Contains('btnReceivingSignOut') -and $buildText.Contains('btnShippingSignOut') -and $buildText.Contains('btnProductionSignOut')) "Role ribbons expose Sign Out buttons."
Add-Check "Build.SignInLabelCallback" ($buildText.Contains('returnedVal = "Sign In"') -and $buildText.Contains('RibbonCurrentUserGetLabel')) "Current user button acts as Sign In while signed out."
Add-Check "Build.UserLabelUsesDisplayName" ($buildText.Contains('modAuth.GetCurrentUserDisplayName()') -and -not $buildText.Contains('returnedVal = "User: " & userId')) "Ribbon user label uses display name, not account id."
Add-Check "Build.RuntimeContextNoSignIn" (-not $buildText.Contains('btnRuntimeSetUser') -and -not $buildText.Contains('RibbonRuntimeUserPrompt')) "Runtime Context is informational and does not expose separate Sign In."
Add-Check "Build.ServerStatusLabelControl" ($buildText.Contains('<labelControl id=""{0}"" getLabel=""{1}""/>') -and $buildText.Contains('RibbonServerStatusGetLabel')) "Role ribbons emit server status label controls."
$roleEventText = Get-Content -LiteralPath (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas") -Raw
Add-Check "Build.RuntimeReferencesNormalXlams" ($buildText.Contains('Get-DeployedOutputPath -Project $referenceProject -OutputDir $outputDir') -and -not $buildText.Contains('Published reference copy')) "Built operator XLAMs reference normal deployed XLAM outputs."
Add-Check "Core.RoleConnectNonModal" ($roleEventText.Contains('ResolveRoleWarehouseTarget(requireNasTarget, statusCode)') -and $roleEventText.Contains('TryApplyRememberedWarehouseTarget') -and -not $roleEventText.Contains('ConnectWarehouseStorageForCapability(Optional ByVal requiredCapability As String = "")' + [Environment]::NewLine + '    If modNasConnection.EnsureWarehouseTargetInteractive')) "Role Connect Server resolves without opening the warehouse connection form."
Add-Check "Core.ConnectServerRootOnly" ($roleEventText.Contains('ConnectKnownWarehouseServer(connectedRoot, statusText)') -and -not $roleEventText.Contains('No selected NAS warehouse target was found')) "Connect Server validates the saved server root without requiring Send To."
Add-Check "Core.SignOutClearsPersistedUser" ($roleEventText.Contains('SetCurrentUserId vbNullString') -and $authText.Contains('mCurrentUserDisplayName = vbNullString')) "Sign Out clears live auth and persisted current-user state."
Add-Check "Core.AuthStoresDisplayName" ($authText.Contains('userInfo("DisplayName")') -and $authText.Contains('Public Function GetCurrentUserDisplayName()')) "Auth cache stores and exposes signed-in display name."
Add-Check "Core.RuntimeContextShowsUserId" ($runtimeStatusText.Contains('"User ID: " & ValueOrPlaceholderStatus(ResolveRuntimeUserStatus())')) "Runtime Context shows signed-in account id."
Add-Check "Core.RememberedTargetUsesConfigAuth" ($runtimeStatusText.Contains('TryRevalidateRememberedRoot(targetRoot)') -and -not $runtimeStatusText.Contains('warehouseId & ".invSys.Data.Inventory.xlsb"')) "Remembered server reconnect requires config/auth, not a local inventory workbook."
Add-Check "Admin.DirectoryReadsNasRoots" ($adminConsoleText.Contains('modNasConnection.GetKnownWarehouseTargetRoots()')) "Admin View Warehouses includes NAS roots remembered by Connect Server."
Add-Check "Core.SendToScansConnectedRoots" ($runtimeStatusText.Contains('If modNasConnection.IsConnected() Then AddKnownServerConfigTargetsStatus targets, seen') -and $runtimeStatusText.Contains('Public Sub InvalidateWarehouseTargets()')) "Send To scans known NAS roots only after Connect Server succeeds."
Add-Check "Core.RibbonFullInvalidate" ($runtimeStatusText.Contains('ribbon.Invalidate')) "Auth/storage changes refresh enabled callbacks."
Add-Check "Validator.ButtonGetEnabledRead" ($validatorText.Contains('GetEnabled = $getEnabled')) "Packaged validator reads getEnabled."
Add-Check "Validator.ButtonGetEnabledAssert" ($validatorText.Contains('RibbonButtonGetEnabled')) "Packaged validator asserts getEnabled on required buttons."
Add-Check "Validator.CallbackGetEnabledAssert" ($validatorText.Contains('CallbackGetEnabled')) "Packaged validator asserts callback capability mapping."
Add-Check "Validator.DirectActionAssert" ($validatorText.Contains('DirectAction') -and $validatorText.Contains('callbackHasDirectAction')) "Packaged validator asserts direct ribbon actions."
Add-Check "Validator.StatusLabelAssert" ($validatorText.Contains('Get-RibbonLabelControls') -and $validatorText.Contains('StatusLabel')) "Packaged validator asserts server status labels."

$failed = @($checks | Where-Object { -not $_.Passed }).Count
$passed = $checks.Count - $failed

$lines = @()
$lines += "# Phase 6 Ribbon Generation Contract Results"
$lines += ""
$lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
$lines += "- Passed: $passed"
$lines += "- Failed: $failed"
$lines += ""
$lines += "| Check | Result | Detail |"
$lines += "|---|---|---|"
foreach ($check in $checks) {
    $result = if ($check.Passed) { "PASS" } else { "FAIL" }
    $detail = [string]$check.Detail
    $detail = $detail.Replace("|", "/")
    $lines += "| $($check.Name) | $result | $detail |"
}
[System.IO.File]::WriteAllLines($resultPath, $lines)

if ($failed -gt 0) {
    Write-Output "PHASE6_RIBBON_GENERATION_CONTRACT_FAILED"
    Write-Output "RESULTS=$resultPath"
    Write-Output "PASSED=$passed FAILED=$failed TOTAL=$($checks.Count)"
    exit 1
}

Write-Output "PHASE6_RIBBON_GENERATION_CONTRACT_OK"
Write-Output "RESULTS=$resultPath"
Write-Output "PASSED=$passed FAILED=0 TOTAL=$($checks.Count)"
