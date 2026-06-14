Param(
    [string]$RepoRoot = ".",
    [int]$StartAt = 1,
    [int]$EndAt = 0
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject {
    Param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
}

function Import-BasModule {
    Param(
        [object]$VbProject,
        [string]$BasPath
    )

    if (-not (Test-Path $BasPath)) {
        throw "Missing BAS module: $BasPath"
    }
    [void]$VbProject.VBComponents.Import($BasPath)
}

function New-NormalizedImportFile {
    Param([string]$SourcePath)

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-harness-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null
    $tempPath = Join-Path $tempDir ([System.IO.Path]::GetFileName($SourcePath))
    $raw = Get-Content -LiteralPath $SourcePath -Raw
    $normalized = $raw -replace "`r?`n", "`r`n"
    [System.IO.File]::WriteAllText($tempPath, $normalized, [System.Text.Encoding]::ASCII)
    return $tempPath
}

function New-NormalizedFormImportFile {
    Param([string]$FormPath)

    $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-harness-form-" + [guid]::NewGuid().ToString("N"))
    New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

    $tempFrmPath = Join-Path $tempDir ([System.IO.Path]::GetFileName($FormPath))
    $raw = Get-Content -LiteralPath $FormPath -Raw
    $normalized = $raw -replace "`r?`n", "`r`n"
    [System.IO.File]::WriteAllText($tempFrmPath, $normalized, [System.Text.Encoding]::ASCII)

    $sourceFrxPath = [System.IO.Path]::ChangeExtension($FormPath, ".frx")
    if (Test-Path -LiteralPath $sourceFrxPath) {
        $tempFrxPath = [System.IO.Path]::ChangeExtension($tempFrmPath, ".frx")
        Copy-Item -LiteralPath $sourceFrxPath -Destination $tempFrxPath -Force
    }

    return $tempFrmPath
}

function Remove-ExistingVBComponent {
    Param(
        [object]$VbProject,
        [string]$ComponentName
    )

    try {
        $existing = $VbProject.VBComponents.Item($ComponentName)
    }
    catch {
        $existing = $null
    }

    if ($null -eq $existing) { return }
    if ($existing.Type -eq 100) {
        throw "Refusing to remove document component '$ComponentName'."
    }
    [void]$VbProject.VBComponents.Remove($existing)
}

function Assert-VBComponentType {
    Param(
        [object]$VbProject,
        [string]$ComponentName,
        [int]$ExpectedType,
        [string]$Context
    )

    try {
        $component = $VbProject.VBComponents.Item($ComponentName)
    }
    catch {
        throw "$Context failed: component '$ComponentName' was not present after import."
    }

    if ($component.Type -ne $ExpectedType) {
        throw "$Context failed: component '$ComponentName' imported with type $($component.Type), expected $ExpectedType."
    }
}

function Test-FormRequiresStub {
    Param([string]$FormPath)

    $frxPath = [System.IO.Path]::ChangeExtension($FormPath, ".frx")
    return -not (Test-Path -LiteralPath $frxPath)
}

function Get-StubUserFormCode {
    Param([string]$FormPath)

    $rawLines = Get-Content -LiteralPath $FormPath
    $hasRuntimeMarker = $false
    foreach ($line in $rawLines) {
        if ($line -match "'@RuntimeStubUserFormCode") {
            $hasRuntimeMarker = $true
            break
        }
    }
    if (-not $hasRuntimeMarker) {
        return "Option Explicit"
    }

    $codeLines = New-Object System.Collections.Generic.List[string]
    $inCode = $false
    foreach ($line in $rawLines) {
        if (-not $inCode) {
            if ($line -match "'@RuntimeStubUserFormCode") {
                $inCode = $true
            }
            continue
        }
        if ($line -match '^Attribute VB_') {
            continue
        }
        [void]$codeLines.Add($line)
    }

    if ($codeLines.Count -eq 0) {
        return "Option Explicit"
    }
    return [string]::Join([Environment]::NewLine, $codeLines)
}

function Add-StubUserForm {
    Param(
        [object]$VbProject,
        [string]$FormPath
    )

    $formName = [System.IO.Path]::GetFileNameWithoutExtension($FormPath)
    Remove-ExistingVBComponent -VbProject $VbProject -ComponentName $formName
    $component = $VbProject.VBComponents.Add(3)
    $component.Name = $formName

    $captionLine = Get-Content -LiteralPath $FormPath | Where-Object { $_ -match '^\s*Caption\s*=\s*"' } | Select-Object -First 1
    if ($null -ne $captionLine) {
        $caption = [regex]::Match($captionLine, '"([^"]*)"').Groups[1].Value
        if ($caption -ne "") {
            try { $component.Designer.Caption = $caption } catch {}
        }
    }

    $module = $component.CodeModule
    if ($module.CountOfLines -gt 0) {
        $module.DeleteLines(1, $module.CountOfLines)
    }
    $stubCode = Get-StubUserFormCode -FormPath $FormPath
    $module.AddFromString($stubCode)
    Assert-VBComponentType -VbProject $VbProject -ComponentName $formName -ExpectedType 3 -Context $FormPath
}

function Import-ClassModule {
    Param(
        [object]$VbProject,
        [string]$ClassPath
    )

    if (-not (Test-Path $ClassPath)) {
        throw "Missing class module: $ClassPath"
    }

    $componentName = [System.IO.Path]::GetFileNameWithoutExtension($ClassPath)
    Remove-ExistingVBComponent -VbProject $VbProject -ComponentName $componentName
    $normalizedPath = New-NormalizedImportFile -SourcePath $ClassPath
    try {
        [void]$VbProject.VBComponents.Import($normalizedPath)
    }
    finally {
        Remove-Item -LiteralPath (Split-Path $normalizedPath -Parent) -Recurse -Force -ErrorAction SilentlyContinue
    }

    $component = $VbProject.VBComponents.Item($componentName)
    if ($component.Type -ne 2) {
        throw "Class import failed for $ClassPath; component '$componentName' resolved as type $($component.Type)."
    }
}

function Import-FormModule {
    Param(
        [object]$VbProject,
        [string]$FormPath
    )

    if (-not (Test-Path $FormPath)) {
        throw "Missing form module: $FormPath"
    }

    $componentName = [System.IO.Path]::GetFileNameWithoutExtension($FormPath)
    if (Test-FormRequiresStub -FormPath $FormPath) {
        Add-StubUserForm -VbProject $VbProject -FormPath $FormPath
        return
    }

    Remove-ExistingVBComponent -VbProject $VbProject -ComponentName $componentName
    $normalizedPath = New-NormalizedFormImportFile -FormPath $FormPath
    try {
        [void]$VbProject.VBComponents.Import($normalizedPath)
        Assert-VBComponentType -VbProject $VbProject -ComponentName $componentName -ExpectedType 3 -Context $FormPath
    }
    finally {
        Remove-Item -LiteralPath (Split-Path $normalizedPath -Parent) -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Run-TestFunction {
    Param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$FunctionName
    )

    $fullMacro = "'$WorkbookName'!$FunctionName"
    try {
        $result = $Excel.Run($fullMacro)
    }
    catch {
        throw "Excel.Run failed for $fullMacro :: $($_.Exception.Message)"
    }
    if ($null -eq $result) { return 0 }
    return [int]$result
}

function Format-ProgressBarText {
    Param(
        [int]$Current,
        [int]$Total,
        [int]$Width = 24
    )

    if ($Total -le 0) {
        return ("." * $Width)
    }
    $completed = [int][Math]::Floor(($Current / [double]$Total) * $Width)
    if ($completed -lt 0) { $completed = 0 }
    if ($completed -gt $Width) { $completed = $Width }
    return ("#" * $completed) + ("." * ($Width - $completed))
}

function Write-HarnessStatus {
    Param(
        [string]$Phase,
        [int]$Current,
        [int]$Total,
        [string]$Detail = "",
        [datetime]$StartedAt = (Get-Date)
    )

    $percent = 0
    if ($Total -gt 0) {
        $percent = [int][Math]::Floor(($Current / [double]$Total) * 100)
        if ($percent -gt 100) { $percent = 100 }
    }

    $elapsed = (Get-Date) - $StartedAt
    $bar = Format-ProgressBarText -Current $Current -Total $Total
    $line = "[{0}] {1,3}% {2}/{3} {4} {5}" -f $bar, $percent, $Current, $Total, $Phase, $Detail
    if ($elapsed.TotalSeconds -ge 1) {
        $line = "$line elapsed=$([int]$elapsed.TotalSeconds)s"
    }
    Write-Host $line

    $activity = "Phase 6 Excel validation"
    $status = "$Phase $Current/$Total"
    if ($Detail -ne "") { $status = "$status - $Detail" }
    Write-Progress -Activity $activity -Status $status -PercentComplete $percent
}

function Write-TestResultsFile {
    Param(
        [string]$ResultPath,
        [object[]]$TestRows,
        [int]$StartAt = 1,
        [int]$EndAt = 0,
        [int]$TotalAvailable = 0,
        [bool]$Complete = $false
    )

    $passedCount = @($TestRows | Where-Object { $_.Passed }).Count
    $failedCount = $TestRows.Count - $passedCount

    $lines = @()
    $lines += "# Phase 6 VBA Test Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Passed: $passedCount"
    $lines += "- Failed: $failedCount"
    if ($TotalAvailable -gt 0) {
        $lines += "- Range: $StartAt-$EndAt of $TotalAvailable"
    }
    if (-not $Complete) {
        $lines += "- Status: PARTIAL"
    }
    $lines += ""
    $lines += "| Test | Result |"
    $lines += "|---|---|"
    foreach ($r in $TestRows) {
        $detail = if ($r.Passed) { "PASS" } elseif ([string]::IsNullOrWhiteSpace($r.Error)) { "FAIL" } else { "FAIL - $($r.Error)" }
        $lines += "| $($r.TestName) | $detail |"
    }
    [System.IO.File]::WriteAllLines($ResultPath, $lines)
}

function Add-BootstrapModule {
    Param([object]$Workbook)
    $comp = $Workbook.VBProject.VBComponents.Add(1)
    $comp.Name = "modHarnessBootstrap"
    $comp.CodeModule.AddFromString("Public Function HarnessPing() As Long: HarnessPing = 1: End Function")
    return $comp
}

function Add-TestWrappers {
    Param(
        [object]$BootstrapComponent,
        [string[]]$TargetFunctions
    )

    $cm = $BootstrapComponent.CodeModule
    $wrappers = @()
    for ($i = 0; $i -lt $TargetFunctions.Count; $i++) {
        $fn = $TargetFunctions[$i]
        $moduleName = $fn.Split(".")[0]
        $wrapper = "RunT" + ($i + 1)
        $errCell = "A" + ($i + 1)
        $line = @"
Public Function $wrapper() As Long
On Error GoTo ErrHandler
ThisWorkbook.Worksheets(1).Range("$errCell").Value = ""
On Error Resume Next
Application.Run("$moduleName.ClearLastTestFailure")
On Error GoTo ErrHandler
$wrapper = Application.Run("$fn")
If $wrapper = 0 Then
    On Error Resume Next
    ThisWorkbook.Worksheets(1).Range("$errCell").Value = Application.Run("$moduleName.GetLastTestFailure")
    On Error GoTo ErrHandler
End If
Exit Function
ErrHandler:
ThisWorkbook.Worksheets(1).Range("$errCell").Value = Err.Description
$wrapper = 0
End Function
"@
        $cm.AddFromString($line)
        $wrappers += $wrapper
    }
    return ,$wrappers
}

$repo = (Resolve-Path $RepoRoot).Path
$fixtures = Join-Path $repo "tests/fixtures"
$harnessStamp = Get-Date -Format "yyyyMMdd_HHmmss_fff"
$harnessPath = Join-Path $fixtures "Phase6_Inventory.Domain_Harness_$harnessStamp.xlsm"
$resultPath = Join-Path $repo "tests/unit/phase6_test_results.md"

$excel = $null
$harness = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $modulePaths = @(
        (Join-Path $repo "src/Core/Modules/modConfigDefaults.bas"),
        (Join-Path $repo "src/Core/Modules/modDeploymentPaths.bas"),
        (Join-Path $repo "src/Core/Modules/modWarehouseBootstrap.bas"),
        (Join-Path $repo "src/Core/Modules/modWarehouseRetire.bas"),
        (Join-Path $repo "src/Core/Modules/modRuntimeWorkbooks.bas"),
        (Join-Path $repo "src/Core/Modules/modNasConnection.bas"),
        (Join-Path $repo "src/Core/Modules/modRibbonRuntimeStatus.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleWorkbookSurfaces.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleUiAccess.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas"),
        (Join-Path $repo "src/Core/Modules/modOperatorReadModel.bas"),
        (Join-Path $repo "src/Core/Modules/modPerfLog.bas"),
        (Join-Path $repo "src/Core/Modules/modDiagnostics.bas"),
        (Join-Path $repo "src/Core/Modules/modInventoryDomainBridge.bas"),
        (Join-Path $repo "src/Core/Modules/modWarehouseSync.bas"),
        (Join-Path $repo "src/Core/Modules/modLockManager.bas"),
        (Join-Path $repo "src/Core/Modules/modProcessor.bas"),
        (Join-Path $repo "src/Core/Modules/modConfig.bas"),
        (Join-Path $repo "src/Core/Modules/modAuth.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventorySchema.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryPublisher.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryBridgeApi.bas"),
        (Join-Path $repo "src/InventoryDomain/Modules/modInventoryApply.bas"),
        (Join-Path $repo "src/Receiving/Modules/modReceivingInit.bas"),
        (Join-Path $repo "src/Shipping/Modules/modShippingEventCreator.bas"),
        (Join-Path $repo "src/Production/Modules/modProductionEventCreator.bas"),
        (Join-Path $repo "src/Admin/Modules/modAddinsPublish.bas"),
        (Join-Path $repo "src/Admin/Modules/modAdminWorkbookTarget.bas"),
        (Join-Path $repo "src/Admin/Modules/modAdminConsole.bas"),
        (Join-Path $repo "tests/unit/TestPhase2Helpers.bas"),
        (Join-Path $repo "tests/unit/TestStub_modTS_Received.bas"),
        (Join-Path $repo "tests/unit/TestAddinsPublish.bas"),
        (Join-Path $repo "tests/unit/TestWarehouseBootstrap.bas"),
        (Join-Path $repo "tests/unit/test_RetireMigrateSpec.bas"),
        (Join-Path $repo "tests/unit/TestWarehouseRetireReAuth.bas"),
        (Join-Path $repo "tests/unit/TestWarehouseRetireArchive.bas"),
        (Join-Path $repo "tests/unit/TestWarehouseRetireMigration.bas"),
        (Join-Path $repo "tests/unit/TestWarehouseRetireLifecycle.bas"),
        (Join-Path $repo "tests/unit/TestReceivingReadiness.bas"),
        (Join-Path $repo "tests/unit/TestPhase6CoreSurfaces.bas"),
        (Join-Path $repo "tests/unit/TestPhase6RoleSurfaces.bas")
    )

    $formPaths = @(
        (Join-Path $repo "src/Core/Forms/frmWarehouseConnection.frm"),
        (Join-Path $repo "src/Core/Forms/frmSignIn.frm"),
        (Join-Path $repo "src/Admin/Forms/frmReAuthGate.frm")
    )

    $classPaths = @(
        (Join-Path $repo "src/Core/ClassModules/WarehouseTarget.cls"),
        (Join-Path $repo "src/Receiving/ClassModules/cAppEvents.cls")
    )

    $allTests = @(
        "TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_ReadsWarehouseIdFromConfig",
        "TestPhase6CoreSurfaces.TestNasGetCurrentTarget_ReturnsDeepCopy",
        "TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_RequiresStationInboxRejectsBlankStation",
        "TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_AllowsRoamingBlankStationWithoutInboxRequirement",
        "TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_TwoStationsHaveIndependentInboxRoots",
        "TestPhase6CoreSurfaces.TestNasScanRoot_ReturnsPathStringsWithoutWarehouseInference",
        "TestPhase6CoreSurfaces.TestNasScanRoot_RejectsMismatchedConfigAuthPair",
        "TestPhase6CoreSurfaces.TestNasResolveRememberedTarget_UnreachableFailsClosed",
        "TestPhase6CoreSurfaces.TestNasResolveRememberedTarget_ReachableRecomputesCachedHints",
        "TestPhase6CoreSurfaces.TestNasFallbackPolicy_RoleRejectsFallbackAdminAccepts",
        "TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_SignsInAndStatusOk",
        "TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_AcceptsResetPinForUserId",
        "TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_RejectsDisplayNameAsUserId",
        "TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_RejectsMismatchedTargetWarehouse",
        "TestPhase6CoreSurfaces.TestAuthCapabilityScope_AllowsSelectedRuntimeFolderAlias",
        "TestPhase6CoreSurfaces.TestAuthFailedCredential_DoesNotReplaceSignedInUser",
        "TestPhase6CoreSurfaces.TestAuthCorrectCredentialWithoutCapability_ReturnsNoCapabilities",
        "TestPhase6CoreSurfaces.TestRuntimeStatusUserLabel_UnsignedShowsNotSignedIn",
        "TestPhase6CoreSurfaces.TestRuntimeStatusUserLabel_TracksAuthSignIn",
        "TestPhase6CoreSurfaces.TestRoleWriteCurrent_RejectsUnsignedUser",
        "TestPhase6CoreSurfaces.TestRoleWriteCurrent_RejectsMissingCapability",
        "TestPhase6CoreSurfaces.TestRoleWriteCurrent_RejectsFallbackTarget",
        "TestPhase6CoreSurfaces.TestRoleWriteCurrent_AllowsSignedInReceivePost",
        "TestPhase6CoreSurfaces.TestAuthSignOut_ClearsUserButKeepsWarehouseTarget",
        "TestPhase6CoreSurfaces.TestAuthCanPerform_SignedOutFailsClosedWithLoadedAuth",
        "TestPhase6CoreSurfaces.TestAuthTtlExpiry_FailsClosedForIsSignedInAndCanPerform",
        "TestAddinsPublish.TestVerifyAddinsPublished_AllPresent",
        "TestAddinsPublish.TestVerifyAddinsPublished_OneMissingLogsDiagnostic",
        "TestAddinsPublish.TestVerifyAddinsPublished_ZeroByteFileLogsDiagnostic",
        "TestAddinsPublish.TestPublishAddins_IdempotentRepublishWritesManifest",
        "TestWarehouseBootstrap.TestValidateWarehouseSpec_TrimsFieldsAndAllowsBlankSharePoint",
        "TestWarehouseBootstrap.TestValidateWarehouseSpec_RejectsEmptyWarehouseId",
        "TestWarehouseBootstrap.TestValidateWarehouseSpec_RejectsWarehouseIdWithSpaces",
        "TestWarehouseBootstrap.TestValidateWarehouseSpec_AllowsWarehouseIdWithHyphenAndUnderscore",
        "TestWarehouseBootstrap.TestValidateWarehouseSpec_RejectsWarehouseIdWithOtherSpecialCharacters",
        "TestWarehouseBootstrap.TestWarehouseIdExists_LocalFolderExists",
        "TestWarehouseBootstrap.TestWarehouseIdExists_SharePointArtifactExists",
        "TestWarehouseBootstrap.TestWarehouseIdExists_NeitherLocalNorSharePointExists",
        "TestWarehouseBootstrap.TestWarehouseIdExists_SharePointUnavailableReturnsFalseAndLogsSkip",
        "TestWarehouseBootstrap.TestBootstrapWarehouseLocal_CreatesBootableLocalRuntime",
        "TestWarehouseBootstrap.TestBootstrapWarehouseLocal_FailureRollsBackPartialFolders",
        "TestWarehouseBootstrap.TestPublishInitialArtifacts_PublishSuccess",
        "TestWarehouseBootstrap.TestPublishInitialArtifacts_SharePointUnavailableReturnsFalse",
        "TestWarehouseBootstrap.TestPublishInitialArtifacts_RepeatedPublishIsIdempotent",
        "test_RetireMigrateSpec.TestValidateRetireMigrateSpec_TrimsAndAcceptsArchiveOnly",
        "test_RetireMigrateSpec.TestValidateRetireMigrateSpec_RejectsEmptySourceWarehouseId",
        "test_RetireMigrateSpec.TestValidateRetireMigrateSpec_RejectsMissingTargetForMigrate",
        "test_RetireMigrateSpec.TestValidateRetireMigrateSpec_RejectsEqualSourceAndTarget",
        "test_RetireMigrateSpec.TestValidateRetireMigrateSpec_RejectsUnconfirmedWriteOperation",
        "test_RetireMigrateSpec.TestValidateRetireMigrateSpec_RejectsInvalidArchiveDestPath",
        "TestWarehouseRetireReAuth.TestValidateUserCredential_SucceedsWithCorrectPasswordAndRole",
        "TestWarehouseRetireReAuth.TestReAuthGate_WrongPassword_ShowsInlineErrorAndDoesNotAuthenticate",
        "TestWarehouseRetireReAuth.TestReAuthGate_ThreeFailures_LocksOutAndLogs",
        "TestWarehouseRetireReAuth.TestReAuthGate_Cancel_LeavesUnauthenticatedWithoutLog",
        "TestWarehouseRetireArchive.TestWriteArchivePackage_SuccessCreatesAtomicArchive",
        "TestWarehouseRetireArchive.TestWriteArchivePackage_PartialFailureRollsBackTempArchive",
        "TestWarehouseRetireArchive.TestWriteArchivePackage_AuthExportMasksPinHash",
        "TestWarehouseRetireMigration.TestMigrateInventoryToTarget_SuccessAppendsInventoryAndTracesSource",
        "TestWarehouseRetireMigration.TestMigrateInventoryToTarget_RejectsMissingArchiveManifest",
        "TestWarehouseRetireMigration.TestMigrateInventoryToTarget_RejectsMissingTargetWarehouse",
        "TestWarehouseRetireMigration.TestMigrateInventoryToTarget_DoesNotCopyAuthIdentities",
        "TestWarehouseRetireMigration.TestMigrateInventoryToTarget_PreservesTargetConfigIdentity",
        "TestWarehouseRetireLifecycle.TestRetireSourceWarehouse_WritesRetirementMarker",
        "TestWarehouseRetireLifecycle.TestRetireSourceWarehouse_WritesValidTombstoneJson",
        "TestWarehouseRetireLifecycle.TestRetireSourceWarehouse_SharePointUnavailableDoesNotBlockRetirement",
        "TestWarehouseRetireLifecycle.TestDeleteLocalRuntime_RejectsWithoutTombstone",
        "TestWarehouseRetireLifecycle.TestDeleteLocalRuntime_RejectsWithoutConfirmation",
        "TestReceivingReadiness.TestCheckReceivingReadiness_AllReady_ReturnsReady",
        "TestReceivingReadiness.TestCheckReceivingReadiness_AllReady_WhenCapabilityStationWildcard",
        "TestReceivingReadiness.TestCheckReceivingReadiness_SnapshotOk_WhenAuthMissingCapability",
        "TestReceivingReadiness.TestCheckReceivingReadiness_SnapshotStale_ReturnsStale",
        "TestReceivingReadiness.TestCheckReceivingReadiness_SnapshotMissing_ReturnsMissing",
        "TestReceivingReadiness.TestCheckReceivingReadiness_SnapshotUnreadable_ReturnsUnreadable",
        "TestReceivingReadiness.TestCheckReceivingReadiness_AuthOk_WhenSnapshotMissing",
        "TestReceivingReadiness.TestCheckReceivingReadiness_AuthNoUser_ReturnsNoUser",
        "TestReceivingReadiness.TestCheckReceivingReadiness_AuthMissingCapability_ReturnsMissingCapability",
        "TestReceivingReadiness.TestCheckReceivingReadiness_AuthInactive_ReturnsInactive",
        "TestReceivingReadiness.TestCheckReceivingReadiness_RuntimeOk_WhenSnapshotMissingAndNoUser",
        "TestReceivingReadiness.TestCheckReceivingReadiness_RuntimeMissingTables_ReturnsMissingTables",
        "TestReceivingReadiness.TestEnsureReceivingSurface_BlankWorkbookWithConfigLoaded_DoesNotApplyReadiness",
        "TestReceivingReadiness.TestCheckReceivingReadiness_RuntimePathUnresolved_ReturnsPathUnresolved",
        "TestPhase6CoreSurfaces.TestOpenOrCreateConfigWorkbookRuntime_CreatesCanonicalWorkbook",
        "TestPhase6CoreSurfaces.TestLoadConfig_AutoBootstrapsCanonicalWorkbook",
        "TestPhase6CoreSurfaces.TestLoadConfig_BlankContextAutoBootstrapsDefaultRuntimeWorkbook",
        "TestPhase6CoreSurfaces.TestEnsureStationBootstrap_CreatesLocalConfigAndInbox",
        "TestPhase6CoreSurfaces.TestLoadConfig_QuarantinesContaminatedConfigSheet",
        "TestPhase6CoreSurfaces.TestLoadAuth_AutoBootstrapsCanonicalWorkbook",
        "TestPhase6CoreSurfaces.TestLoadAuth_BootstrapGrantsCurrentOperatorCapabilities",
        "TestPhase6CoreSurfaces.TestResolveInventoryWorkbookBridge_PrefersCanonicalWorkbookOverOperatorSurface",
        "TestPhase6CoreSurfaces.TestEnsureInventoryManagementSurface_RemovesDomainArtifacts",
        "TestPhase6CoreSurfaces.TestOpenOrCreateConfigWorkbookRuntime_PrunesUnexpectedSheets",
        "TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_UpdatesReadModelAndMetadata",
        "TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSharePoint_UpdatesReadModelAndMetadata",
        "TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSharePoint_StaleSnapshotMarksReadModelStale",
        "TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromCache_PreservesLocalStagingAndLogs",
        "TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_AddsRowsWhenInvSysStartsEmpty",
        "TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_AppliesCatalogMetadataForZeroQtyRows",
        "TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_NormalizesLegacyLocationSummary",
        "TestPhase6CoreSurfaces.TestRefreshInventoryReadModel_MissingSnapshotMarksStaleWithoutMutatingReceivingTally",
        "TestPhase6CoreSurfaces.TestRefreshInventoryReadModel_MissingSharePointSnapshotMarksCachedWithoutMutatingLocalTables",
        "TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_StaleSharePointSnapshotShowsVisibleMetadataWithoutMutatingLocalTables",
        "TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_MissingSnapshotDoesNotBlockQueueAndRefresh",
        "TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_FullRuntimeCloseReopenReloadsCanonicalWorkbooks",
        "TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_ReopenRefreshPreservesLocalTables",
        "TestPhase6CoreSurfaces.TestReceivingSetupUi_ForceRefreshesRegisteredWorkbook",
        "TestPhase6CoreSurfaces.TestInventoryPublisher_PublishesSnapshotForOpenInventoryWorkbook",
        "TestPhase6CoreSurfaces.TestLanSharedSnapshot_TwoSavedOperatorWorkbooksRefreshWithoutCrossContamination",
        "TestPhase6CoreSurfaces.TestLanTwoStationProcessorRun_RespectsLockAndPreservesOperatorWorkbooks",
        "TestPhase6CoreSurfaces.TestProcessor_DiscoversClosedConfiguredStationInboxWorkbook",
        "TestPhase6CoreSurfaces.TestSavedShippingWorkbook_RefreshPreservesStagingAndLogs",
        "TestPhase6CoreSurfaces.TestSavedShippingWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs",
        "TestPhase6CoreSurfaces.TestShippingEventCreator_QueuesSignedInCurrentTargetEvent",
        "TestPhase6CoreSurfaces.TestSavedProductionWorkbook_RefreshPreservesStagingAndLogs",
        "TestPhase6CoreSurfaces.TestSavedProductionWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs",
        "TestPhase6CoreSurfaces.TestProductionEventCreator_QueuesSignedInCurrentTargetEvent",
        "TestPhase6CoreSurfaces.TestSavedAdminWorkbook_ReopenRefreshReissuePreservesAudit",
        "TestPhase6CoreSurfaces.TestApplyReceive_RebuildsDeletedProjectionTablesInCanonicalWorkbook",
        "TestPhase6RoleSurfaces.TestEnsureInventoryManagementSurface_RemovesDuplicateAliasColumns",
        "TestPhase6RoleSurfaces.TestEnsureReceivingWorkbookSurface_CreatesExpectedTables",
        "TestPhase6RoleSurfaces.TestEnsureReceivingWorkbookSurface_RecreatesDeletedArtifacts",
        "TestPhase6RoleSurfaces.TestEnsureShippingWorkbookSurface_CreatesExpectedTables",
        "TestPhase6RoleSurfaces.TestEnsureShippingWorkbookSurface_RecreatesDeletedArtifacts",
        "TestPhase6RoleSurfaces.TestEnsureProductionWorkbookSurface_CreatesExpectedTables",
        "TestPhase6RoleSurfaces.TestEnsureProductionWorkbookSurface_RecreatesDeletedArtifacts",
        "TestPhase6RoleSurfaces.TestEnsureAdminWorkbookSurface_CreatesExpectedTables",
        "TestPhase6RoleSurfaces.TestResolveAdminTargetWorkbook_PrefersActiveVisibleWorkbook",
        "TestPhase6RoleSurfaces.TestResolveAdminTargetWorkbook_ExplicitWorkbookWinsOverActiveWorkbook",
        "TestPhase6RoleSurfaces.TestOpenUserManagement_WithoutWorkbookArgTargetsActiveWorkbook",
        "TestPhase6RoleSurfaces.TestOpenAdminConsole_WithoutRuntime_DoesNotCreateDefaultWarehouse"
    )

    $totalAvailableTests = $allTests.Count
    if ($EndAt -le 0) { $EndAt = $totalAvailableTests }
    if ($StartAt -lt 1) { throw "StartAt must be >= 1." }
    if ($EndAt -lt $StartAt) { throw "EndAt must be >= StartAt." }
    if ($EndAt -gt $totalAvailableTests) { throw "EndAt $EndAt exceeds available test count $totalAvailableTests." }

    if ($StartAt -ne 1 -or $EndAt -ne $totalAvailableTests) {
        $rangeSuffix = "{0:D3}_{1:D3}" -f $StartAt, $EndAt
        $resultPath = Join-Path $repo "tests/unit/phase6_test_results_$rangeSuffix.md"
    }

    $selectedTests = @()
    for ($testIndex = $StartAt - 1; $testIndex -le $EndAt - 1; $testIndex++) {
        $selectedTests += $allTests[$testIndex]
    }
    $allTests = $selectedTests

    $scriptStart = Get-Date
    Write-HarnessStatus -Phase "Preparing harness" -Current 0 -Total $allTests.Count -Detail "Excel workbook/import setup; selected $StartAt-$EndAt of $totalAvailableTests" -StartedAt $scriptStart

    $harness = $excel.Workbooks.Add()
    $bootstrap = Add-BootstrapModule -Workbook $harness
    $vbProject = $harness.VBProject
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")

    foreach ($c in $classPaths) {
        Import-ClassModule -VbProject $vbProject -ClassPath $c
    }
    foreach ($f in $formPaths) {
        Import-FormModule -VbProject $vbProject -FormPath $f
    }
    foreach ($m in $modulePaths) {
        Import-BasModule -VbProject $vbProject -BasPath $m
    }
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")

    $wrapperNames = Add-TestWrappers -BootstrapComponent $bootstrap -TargetFunctions $allTests
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    $harness.SaveAs($harnessPath, 52)
    Write-HarnessStatus -Phase "Running tests" -Current 0 -Total $allTests.Count -Detail "starting" -StartedAt $scriptStart

    $testRows = @()
    for ($i = 0; $i -lt $allTests.Count; $i++) {
        $name = $allTests[$i]
        $wrapperName = $wrapperNames[$i]
        $testNumber = $i + 1
        $absoluteTestNumber = ($StartAt + $i)
        $testLabel = "[global $absoluteTestNumber/$totalAvailableTests] $name"
        Write-HarnessStatus -Phase "Starting test" -Current $testNumber -Total $allTests.Count -Detail $testLabel -StartedAt $scriptStart
        Write-Progress -Activity "Phase 6 Excel validation" `
            -Status ("Running selected {0}/{1}, global {2}/{3}: {4}" -f $testNumber, $allTests.Count, $absoluteTestNumber, $totalAvailableTests, $name) `
            -PercentComplete ([int][Math]::Floor(($i / [double]$allTests.Count) * 100))
        $passed = Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName $wrapperName
        $errorText = [string]$harness.Worksheets.Item(1).Range("A$($i + 1)").Value2
        $testRows += [pscustomobject]@{
            TestName = $name
            Passed   = ($passed -eq 1)
            Error    = $errorText
        }
        $resultText = if ($passed -eq 1) { "PASS" } elseif ([string]::IsNullOrWhiteSpace($errorText)) { "FAIL" } else { "FAIL - $errorText" }
        Write-TestResultsFile -ResultPath $resultPath -TestRows $testRows -StartAt $StartAt -EndAt $EndAt -TotalAvailable $totalAvailableTests -Complete $false
        Write-HarnessStatus -Phase "Completed test" -Current $testNumber -Total $allTests.Count -Detail "$resultText $testLabel" -StartedAt $scriptStart
    }
    Write-Progress -Activity "Phase 6 Excel validation" -Completed

    $passedCount = @($testRows | Where-Object { $_.Passed }).Count
    $failedCount = $testRows.Count - $passedCount

    Write-TestResultsFile -ResultPath $resultPath -TestRows $testRows -StartAt $StartAt -EndAt $EndAt -TotalAvailable $totalAvailableTests -Complete $true

    Write-Output "PHASE6_VALIDATION_OK"
    Write-Output "HARNESS=$harnessPath"
    Write-Output "RESULTS=$resultPath"
    Write-Output "PASSED=$passedCount FAILED=$failedCount TOTAL=$($testRows.Count)"
    Write-Output "RANGE=$StartAt-$EndAt AVAILABLE=$totalAvailableTests"
}
finally {
    if ($null -ne $harness) {
        try { $harness.Close($true) } catch {}
        Release-ComObject $harness
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
}
