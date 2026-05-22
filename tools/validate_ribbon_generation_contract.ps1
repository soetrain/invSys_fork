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

Add-Check "Build.GetEnabledXml" ($buildText.Contains('getEnabled="RibbonRequiredCapabilityGetEnabled"')) "RequiredCapability buttons emit getEnabled."
Add-Check "Build.GetEnabledCallback" ($buildText.Contains('Public Sub RibbonRequiredCapabilityGetEnabled')) "Generated callback exists."
Add-Check "Build.ReceivingCapability" ($buildText.Contains('RequiredCapability = "RECEIVE_POST"')) "Receiving buttons declare capability."
Add-Check "Build.ShippingCapability" ($buildText.Contains('RequiredCapability = "SHIP_POST"')) "Shipping buttons declare capability."
Add-Check "Build.ProductionCapability" ($buildText.Contains('RequiredCapability = "PROD_POST"')) "Production buttons declare capability."
Add-Check "Validator.ButtonGetEnabledRead" ($validatorText.Contains('GetEnabled = $getEnabled')) "Packaged validator reads getEnabled."
Add-Check "Validator.ButtonGetEnabledAssert" ($validatorText.Contains('RibbonButtonGetEnabled')) "Packaged validator asserts getEnabled on required buttons."
Add-Check "Validator.CallbackGetEnabledAssert" ($validatorText.Contains('CallbackGetEnabled')) "Packaged validator asserts callback capability mapping."

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
