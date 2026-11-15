Param(
    [string]$RepoRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$validatorPath = Join-Path $repo "tools/validate_release1_full_chain.ps1"
$adminConsolePath = Join-Path $repo "src/Admin/Modules/modAdminConsole.bas"

$validatorText = if (Test-Path -LiteralPath $validatorPath) {
    Get-Content -LiteralPath $validatorPath -Raw
} else {
    ""
}
$liveValidatorPath = Join-Path $repo "tools/validate_phase6_live_role_workflows.ps1"
$d14PackagedPath = Join-Path $repo 'tests/tooling/Run-D14PackagedCreate.ps1'
$d14ControllerPath = Join-Path $repo 'tests/tooling/Invoke-D14PackagedGate.ps1'
$runtimeExporterPath = Join-Path $repo "tools/export-invsys-runtime-state.ps1"
$packagedActionText = $validatorText
if (Test-Path -LiteralPath $liveValidatorPath) {
    $packagedActionText += Get-Content -LiteralPath $liveValidatorPath -Raw
}
$adminConsoleText = Get-Content -LiteralPath $adminConsolePath -Raw
$d14PackagedText = if (Test-Path -LiteralPath $d14PackagedPath) {
    Get-Content -LiteralPath $d14PackagedPath -Raw
} else {
    ''
}
$d14ControllerText = if (Test-Path -LiteralPath $d14ControllerPath) {
    Get-Content -LiteralPath $d14ControllerPath -Raw
} else {
    ''
}
$runtimeExporterText = if (Test-Path -LiteralPath $runtimeExporterPath) {
    Get-Content -LiteralPath $runtimeExporterPath -Raw
} else {
    ""
}

$results = [System.Collections.Generic.List[object]]::new()
function Add-Check {
    Param(
        [string]$Name,
        [bool]$Passed,
        [string]$Detail
    )

    $results.Add([pscustomobject]@{
        Check = $Name
        Passed = $Passed
        Detail = $Detail
    })
}

Add-Check "Slice14.Validator.Exists" `
    (Test-Path -LiteralPath $validatorPath) `
    "A dedicated Release 1 validator must own the ordered chain and its combined evidence."

Add-Check "Slice14.Admin.PrimitiveEntryBoundary" `
    (($adminConsoleText -match '(?i)Public\s+Function\s+SeedDemoInventoryForAutomation') -and
     ($validatorText -match '(?i)invSys\.Admin\.xlam') -and
     ($validatorText -match '(?i)modAdminConsole\.BootstrapWarehouseLocalAdmin') -and
     ($validatorText -match '(?i)modAdminConsole\.SeedDemoInventoryForAutomation')) `
    "Fresh warehouse creation and fake inventory seeding must execute through packaged Admin primitive callbacks."

$d14Names = @(
    'D14.GeneratedSchemas.ManagedHeadersNoROW',
    'D14.Seed.UniqueKeysConditionGood',
    'D14.RoundTrip.PreservesSystemKey',
    'D14.AdminSeed.RepeatedCreatesUniqueKeysNoMigration',
    'D14.OperatorRefresh.PreservesCustomColumn',
    'D14.OperatorReopen.PreservesCustomColumn'
)
$d14NamesPresent = @($d14Names | Where-Object { -not $d14PackagedText.Contains("'$_'") }).Count -eq 0
Add-Check 'Slice14.D14PackagedRegression.Wired' `
    ($validatorText -match [regex]::Escape('Invoke-D14PackagedGate.ps1') -and
     $validatorText -match [regex]::Escape('AdminEntry.D14PackagedRegression') -and
     $validatorText -notmatch [regex]::Escape('AdminEntry.SourceIntegrationRegression') -and
     $d14ControllerText -match [regex]::Escape('ExactSixChecks') -and
     $d14NamesPresent) `
    'The ordered chain must run the six D14 assertions through current packaged callbacks.'

$orderedTokens = @(
    'GenerateFreshWarehouse',
    'SeedDemoInventoryThroughAdmin',
    'ReceiveInventory',
    'ProcessorApplyReceive',
    'RefreshAfterReceive',
    'ProductionTwoBatches',
    'ProductionConsumptionAndOutput',
    'BoxingVersionSelection',
    'ShipmentStagingAndSent',
    'ProcessorApplyShipment',
    'FinalRefresh',
    'RestartAndReconcile'
)
$lastIndex = -1
$ordered = $true
foreach ($token in $orderedTokens) {
    $nextIndex = $validatorText.IndexOf($token, $lastIndex + 1, [System.StringComparison]::Ordinal)
    if ($nextIndex -lt 0) {
        $ordered = $false
        break
    }
    $lastIndex = $nextIndex
}
Add-Check "Slice14.Chain.Ordered" $ordered `
    ("Required ordered phase tokens: " + ($orderedTokens -join " -> "))

Add-Check "Slice14.PackagedRoleActions" `
    (($packagedActionText -match '(?i)RunReceivingConfirmWritesFormActionForTest') -and
     ($packagedActionText -match '(?i)ProductionFormTwoBatchActionReportForTest') -and
     ($packagedActionText -match '(?i)RunRelease1BoxingActionForTest') -and
     ($packagedActionText -match '(?i)RunShipmentsSentFormActionForTest')) `
    "The chain must exercise the same packaged action boundaries used by operators."

$requiredAssertions = @(
    'UniqueSystemKeys',
    'NoRowHeaders',
    'HeaderPersistence',
    'ExactBalancesAndLocations',
    'EventIdentityStatusLogAndReplay',
    'NoNegativeInventory',
    'ProductionBatchState',
    'BoxingBomVersion',
    'LocksReleased',
    'OverlayPreserved',
    'RestartReconciliation',
    'CanonicalWorkbooksHidden',
    'NoDuplicatePackagesOrCallbacks',
    'RuntimeFivePackages',
    'StaticRetiredPathRatchet'
)
$missingAssertions = @($requiredAssertions | Where-Object {
    $validatorText -notmatch [regex]::Escape($_)
})
Add-Check "Slice14.RequiredAssertions" `
    ($missingAssertions.Count -eq 0) `
    $(if ($missingAssertions.Count -eq 0) {
        "All required invariant assertions are named."
    } else {
        "Missing: " + ($missingAssertions -join ", ")
    })

Add-Check "Slice14.RestartUsesSavedRuntime" `
    (($validatorText -match '(?i)\.Save\(') -and
     ($validatorText -match '(?i)\.Close\(') -and
     ($validatorText -match '(?i)Workbooks\.Open\(')) `
    "The validator must save, close, and reopen selected runtime boundaries before reconciliation."

Add-Check "Slice14.RuntimeAndStaticEvidence" `
    (($validatorText -match '(?i)export-invsys-runtime-state\.ps1') -and
     ($validatorText -match '(?i)inventory-vba-surface\.ps1') -and
     ($validatorText -match 'Review-RetiredRowWarnings\.ps1') -and
     (Test-Path -LiteralPath (Join-Path $repo 'tests/tooling/Test-Slice14RowWarningReview.ps1'))) `
    "The final gate must include read-only five-package runtime evidence and the static retired-path ratchet."

Add-Check "Slice14.RuntimeEvidence.IsolatedPackageSet" `
    (($validatorText -match "Suspend-NonTargetInvSysAddins") -and
     ($validatorText -match "Restore-SuspendedInvSysAddins") -and
     ($validatorText -match "expectedPackagePaths") -and
     ($validatorText -match "isolationExcel") -and
     ($validatorText -match "restoreExcel")) `
    "The five-package extractor must isolate and verify the intended package paths while restoring globally registered invSys add-ins."

Add-Check "Slice14.RuntimeEvidence.ExactExcelSession" `
    (($runtimeExporterText -match "ExcelHwnd") -and
     ($runtimeExporterText -match "AccessibleObjectFromWindow") -and
     ($validatorText -match "-ExcelHwnd")) `
    "The read-only runtime extractor must attach to the exact full-chain Excel session instead of an arbitrary ROT session."

Add-Check "Slice14.Evidence.RedactsRunIds" `
    ($validatorText.Contains("'(?i)RunId=[^;|\s]+'") -and
     $validatorText -match "RunId=<redacted>") `
    "Committed full-chain evidence must redact generated processor run identifiers."

Add-Check "Slice14.Evidence.RecordsD13Trace" `
    (($validatorText -match "## D13 trace") -and
     ($validatorText -match "Focused RED") -and
     ($validatorText -match "Behavioral RED") -and
     ($validatorText -match "GREEN")) `
    "Committed Slice 14 evidence must record meaningful RED and final GREEN."

Add-Check "Slice14.Harness.KeepArtifactsDiagnostic" `
    (($validatorText -match '(?i)\[switch\]\$KeepArtifacts') -and
     ($validatorText -match '(?i)-not\s+\$KeepArtifacts')) `
    "The full-chain harness must be able to preserve its disposable generated validator for a failed automation diagnosis."

$failed = @($results | Where-Object { -not $_.Passed })
foreach ($row in $results) {
    $status = if ($row.Passed) { "PASS" } else { "FAIL" }
    Write-Host ("[{0}] {1} - {2}" -f $status, $row.Check, $row.Detail)
}
Write-Host ("Slice 14 full-chain contract: {0} passed, {1} failed" -f
    ($results.Count - $failed.Count), $failed.Count)

if ($failed.Count -gt 0) {
    exit 1
}
