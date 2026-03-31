Param(
    [string]$RepoRoot = "."
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Add-ResultRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Check,
        [bool]$Passed,
        [string]$Detail
    )

    $Rows.Add([pscustomobject]@{
        Check  = $Check
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

function Invoke-RepoScript {
    param(
        [string]$RepoRootPath,
        [string]$ScriptRelativePath
    )

    $scriptPath = Join-Path $RepoRootPath $ScriptRelativePath
    if (-not (Test-Path -LiteralPath $scriptPath)) {
        throw "Missing script: $scriptPath"
    }

    $output = & powershell -ExecutionPolicy Bypass -File $scriptPath -RepoRoot $RepoRootPath 2>&1
    $exitCode = $LASTEXITCODE
    return [pscustomobject]@{
        Script   = $ScriptRelativePath
        ExitCode = $exitCode
        Output   = @($output | ForEach-Object { [string]$_ })
    }
}

function Find-MarkdownRow {
    param(
        [string]$ResultPath,
        [string]$Identifier
    )

    if (-not (Test-Path -LiteralPath $ResultPath)) { return "" }
    foreach ($line in [System.IO.File]::ReadAllLines($ResultPath)) {
        if ($line.StartsWith("| $Identifier |")) {
            return $line
        }
    }
    return ""
}

function Add-RowFromMarkdown {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Check,
        [string]$ResultPath,
        [string]$Identifier,
        [string]$Prefix
    )

    $line = Find-MarkdownRow -ResultPath $ResultPath -Identifier $Identifier
    if ([string]::IsNullOrWhiteSpace($line)) {
        Add-ResultRow -Rows $Rows -Check $Check -Passed $false -Detail "$Prefix missing row: $Identifier"
        return
    }

    $passed = ($line -match "\|\s*PASS(?:\s|$)")
    $detail = "$Prefix $line"
    Add-ResultRow -Rows $Rows -Check $Check -Passed $passed -Detail $detail
}

function Write-Results {
    param(
        [string]$ResultPath,
        [System.Collections.Generic.List[object]]$Rows
    )

    $passedCount = @($Rows | Where-Object { $_.Passed }).Count
    $failedCount = $Rows.Count - $passedCount

    $lines = @()
    $lines += "# Phase 6 LAN + WAN Proving Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Passed: $passedCount"
    $lines += "- Failed: $failedCount"
    $lines += "- Sources: tests/unit/phase5_test_results.md; tests/unit/phase6_test_results.md; tests/unit/phase5_hq_boundary_results.md; tests/unit/phase6_packaged_wan_hq_results.md"
    $lines += "- Scope note: proof rows below are the targeted Phase 6 LAN + WAN and central aggregation evidence items."
    $lines += ""
    $lines += "| Check | Result | Detail |"
    $lines += "|---|---|---|"
    foreach ($row in $Rows) {
        $result = if ($row.Passed) { "PASS" } else { "FAIL" }
        $detail = ([string]$row.Detail).Replace("|", "/")
        $lines += "| $($row.Check) | $result | $detail |"
    }
    [System.IO.File]::WriteAllLines($ResultPath, $lines)

    return [pscustomobject]@{
        Passed = $passedCount
        Failed = $failedCount
        Total  = $Rows.Count
    }
}

$repo = (Resolve-Path $RepoRoot).Path
$phase5ResultPath = Join-Path $repo "tests/unit/phase5_test_results.md"
$phase6ResultPath = Join-Path $repo "tests/unit/phase6_test_results.md"
$hqBoundaryResultPath = Join-Path $repo "tests/unit/phase5_hq_boundary_results.md"
$packagedResultPath = Join-Path $repo "tests/unit/phase6_packaged_wan_hq_results.md"
$resultPath = Join-Path $repo "tests/unit/phase6_lan_wan_proving_results.md"
$rows = New-Object 'System.Collections.Generic.List[object]'

$phase5Run = Invoke-RepoScript -RepoRootPath $repo -ScriptRelativePath "tools/run_phase5_excel_validation.ps1"
$phase6Run = Invoke-RepoScript -RepoRootPath $repo -ScriptRelativePath "tools/run_phase6_excel_validation.ps1"
$hqBoundaryRun = Invoke-RepoScript -RepoRootPath $repo -ScriptRelativePath "tools/run_phase5_hq_boundary_validation.ps1"
$packagedRun = Invoke-RepoScript -RepoRootPath $repo -ScriptRelativePath "tools/validate_phase6_packaged_wan_hq.ps1"

Add-ResultRow -Rows $rows -Check "Runner.Phase5Validation" -Passed ($phase5Run.ExitCode -eq 0) -Detail (($phase5Run.Output -join " ; ").Trim())
Add-ResultRow -Rows $rows -Check "Runner.Phase6Validation" -Passed ($phase6Run.ExitCode -eq 0) -Detail (($phase6Run.Output -join " ; ").Trim())
Add-ResultRow -Rows $rows -Check "Runner.HqBoundaryValidation" -Passed ($hqBoundaryRun.ExitCode -eq 0) -Detail (($hqBoundaryRun.Output -join " ; ").Trim())
Add-ResultRow -Rows $rows -Check "Runner.PackagedWanHqValidation" -Passed ($packagedRun.ExitCode -eq 0) -Detail (($packagedRun.Output -join " ; ").Trim())

Add-RowFromMarkdown -Rows $rows -Check "WAN.PublishOnline" -ResultPath $phase5ResultPath -Identifier "TestPhase5Sync.TestWanPublish_OnlineCopy_PublishesLocalArtifactsToSharePoint" -Prefix "phase5"
Add-RowFromMarkdown -Rows $rows -Check "WAN.PublishOfflineNonBlocking" -ResultPath $phase5ResultPath -Identifier "TestPhase5Sync.TestWanPublish_OfflineFailure_DoesNotBlockLocalProcessing" -Prefix "phase5"
Add-RowFromMarkdown -Rows $rows -Check "WAN.PublishSafeRerun" -ResultPath $phase5ResultPath -Identifier "TestPhase5Sync.TestWanPublish_SafeRerun_ReplacesPublishedArtifacts" -Prefix "phase5"
Add-RowFromMarkdown -Rows $rows -Check "WAN.PublishInterruptedRestore" -ResultPath $phase5ResultPath -Identifier "TestPhase5Sync.TestWanPublish_InterruptedReplacement_RestoresPriorArtifactAndAllowsCleanRerun" -Prefix "phase5"
Add-RowFromMarkdown -Rows $rows -Check "WAN.DelayedPublicationCatchup" -ResultPath $phase5ResultPath -Identifier "TestPhase5Sync.TestDelayedPublicationRecovery_PreservesLocalOutboxAndGlobalCatchup" -Prefix "phase5"
Add-RowFromMarkdown -Rows $rows -Check "WAN.UnreadablePublishedSnapshotFallback" -ResultPath $phase5ResultPath -Identifier "TestPhase5Sync.TestHqAggregation_SkipsUnreadablePublishedSnapshotAndRetainsLastGoodData" -Prefix "phase5"
Add-RowFromMarkdown -Rows $rows -Check "WAN.MixedWarehouseInterruptionCatchup" -ResultPath $phase5ResultPath -Identifier "TestPhase5Sync.TestHqAggregation_MixedWarehouseInterruption_RetainsLastGoodAndCatchesUp" -Prefix "phase5"

Add-RowFromMarkdown -Rows $rows -Check "LAN.SharedSnapshotTwoOperators" -ResultPath $phase6ResultPath -Identifier "TestPhase6CoreSurfaces.TestLanSharedSnapshot_TwoSavedOperatorWorkbooksRefreshWithoutCrossContamination" -Prefix "phase6"
Add-RowFromMarkdown -Rows $rows -Check "LAN.TwoStationProcessorLockBoundary" -ResultPath $phase6ResultPath -Identifier "TestPhase6CoreSurfaces.TestLanTwoStationProcessorRun_RespectsLockAndPreservesOperatorWorkbooks" -Prefix "phase6"
Add-RowFromMarkdown -Rows $rows -Check "Operator.SharePointStaleMetadataVisible" -ResultPath $phase6ResultPath -Identifier "TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSharePoint_StaleSnapshotMarksReadModelStale" -Prefix "phase6"
Add-RowFromMarkdown -Rows $rows -Check "Operator.SavedWorkbookSharePointStaleVisibleNonDestructive" -ResultPath $phase6ResultPath -Identifier "TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_StaleSharePointSnapshotShowsVisibleMetadataWithoutMutatingLocalTables" -Prefix "phase6"
Add-RowFromMarkdown -Rows $rows -Check "Operator.MissingSharePointSnapshotCachedNonDestructive" -ResultPath $phase6ResultPath -Identifier "TestPhase6CoreSurfaces.TestRefreshInventoryReadModel_MissingSharePointSnapshotMarksCachedWithoutMutatingLocalTables" -Prefix "phase6"
Add-RowFromMarkdown -Rows $rows -Check "Operator.MissingSnapshotDoesNotBlockPosting" -ResultPath $phase6ResultPath -Identifier "TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_MissingSnapshotDoesNotBlockQueueAndRefresh" -Prefix "phase6"

Add-RowFromMarkdown -Rows $rows -Check "HQ.StaggeredWarehouseUpdates" -ResultPath $phase5ResultPath -Identifier "TestPhase5Sync.TestHqAggregation_RebuildsGlobalSnapshotAfterStaggeredWarehouseUpdates" -Prefix "phase5"
Add-RowFromMarkdown -Rows $rows -Check "HQ.RepeatedRunsRemainStable" -ResultPath $phase5ResultPath -Identifier "TestPhase5Sync.TestHqAggregation_RepeatedRunsRemainStableForWH1AndWH2Fixtures" -Prefix "phase5"
Add-RowFromMarkdown -Rows $rows -Check "HQ.AdvisoryOnlyVisibility" -ResultPath $phase5ResultPath -Identifier "TestPhase5Sync.TestHqAggregation_GlobalSnapshotStatusIsAdvisoryOnly" -Prefix "phase5"
Add-RowFromMarkdown -Rows $rows -Check "HQ.RealPublishedArtifactsBoundaryInitial" -ResultPath $hqBoundaryResultPath -Identifier "Aggregate.Initial" -Prefix "hq-boundary"
Add-RowFromMarkdown -Rows $rows -Check "HQ.RealPublishedArtifactsBoundaryCatchup" -ResultPath $hqBoundaryResultPath -Identifier "Aggregate.Catchup" -Prefix "hq-boundary"
Add-RowFromMarkdown -Rows $rows -Check "HQ.PackagedPublishedArtifactsInitial" -ResultPath $packagedResultPath -Identifier "Aggregate.Initial" -Prefix "packaged"
Add-RowFromMarkdown -Rows $rows -Check "HQ.PackagedPublishedArtifactsCatchup" -ResultPath $packagedResultPath -Identifier "Aggregate.Catchup" -Prefix "packaged"

$summary = Write-Results -ResultPath $resultPath -Rows $rows

Write-Output "PHASE6_LAN_WAN_PROVING_OK"
Write-Output "RESULTS=$resultPath"
Write-Output "PASSED=$($summary.Passed) FAILED=$($summary.Failed) TOTAL=$($summary.Total)"
