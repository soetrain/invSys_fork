Param(
    [string]$RepoRoot = "."
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
    [void]$VbProject.VBComponents.Import($FormPath)
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
        $wrapper = "RunT" + ($i + 1)
        $errCell = "A" + ($i + 1)
        $line = @"
Public Function $wrapper() As Long
On Error GoTo ErrHandler
ThisWorkbook.Worksheets(1).Range("$errCell").Value = ""
$wrapper = Application.Run("$fn")
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
        (Join-Path $repo "src/Core/Modules/modWarehouseBootstrap.bas"),
        (Join-Path $repo "src/Core/Modules/modWarehouseRetire.bas"),
        (Join-Path $repo "src/Core/Modules/modRuntimeWorkbooks.bas"),
        (Join-Path $repo "src/Core/Modules/modRoleWorkbookSurfaces.bas"),
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
        (Join-Path $repo "src/Admin/Forms/frmReAuthGate.frm")
    )

    $classPaths = @(
        (Join-Path $repo "src/Receiving/ClassModules/cAppEvents.cls")
    )

    $allTests = @(
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
        "TestPhase6CoreSurfaces.TestSavedProductionWorkbook_RefreshPreservesStagingAndLogs",
        "TestPhase6CoreSurfaces.TestSavedProductionWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs",
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

    $harness = $excel.Workbooks.Add()
    $bootstrap = Add-BootstrapModule -Workbook $harness
    $vbProject = $harness.VBProject
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")

    foreach ($m in $modulePaths) {
        Import-BasModule -VbProject $vbProject -BasPath $m
        [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    }
    foreach ($c in $classPaths) {
        Import-ClassModule -VbProject $vbProject -ClassPath $c
        [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    }
    foreach ($f in $formPaths) {
        Import-FormModule -VbProject $vbProject -FormPath $f
        [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    }

    $wrapperNames = Add-TestWrappers -BootstrapComponent $bootstrap -TargetFunctions $allTests
    [void](Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName "HarnessPing")
    $harness.SaveAs($harnessPath, 52)

    $testRows = @()
    for ($i = 0; $i -lt $allTests.Count; $i++) {
        $name = $allTests[$i]
        $wrapperName = $wrapperNames[$i]
        $passed = Run-TestFunction -Excel $excel -WorkbookName $harness.Name -FunctionName $wrapperName
        $errorText = [string]$harness.Worksheets.Item(1).Range("A$($i + 1)").Value2
        $testRows += [pscustomobject]@{
            TestName = $name
            Passed   = ($passed -eq 1)
            Error    = $errorText
        }
    }

    $passedCount = @($testRows | Where-Object { $_.Passed }).Count
    $failedCount = $testRows.Count - $passedCount

    $lines = @()
    $lines += "# Phase 6 VBA Test Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Passed: $passedCount"
    $lines += "- Failed: $failedCount"
    $lines += ""
    $lines += "| Test | Result |"
    $lines += "|---|---|"
    foreach ($r in $testRows) {
        $detail = if ($r.Passed) { "PASS" } elseif ([string]::IsNullOrWhiteSpace($r.Error)) { "FAIL" } else { "FAIL - $($r.Error)" }
        $lines += "| $($r.TestName) | $detail |"
    }
    [System.IO.File]::WriteAllLines($resultPath, $lines)

    Write-Output "PHASE6_VALIDATION_OK"
    Write-Output "HARNESS=$harnessPath"
    Write-Output "RESULTS=$resultPath"
    Write-Output "PASSED=$passedCount FAILED=$failedCount TOTAL=$($testRows.Count)"
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
