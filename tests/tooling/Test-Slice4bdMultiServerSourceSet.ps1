[CmdletBinding()]
param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$adminPath = Join-Path $repo "src\Admin\Modules\modAdmin.bas"
$consolePath = Join-Path $repo "src\Admin\Modules\modAdminConsole.bas"
$aggregatorPath = Join-Path $repo "src\Core\Modules\modHqAggregator.bas"
$formPath = Join-Path $repo "src\Admin\Forms\frmAggregationSources.frm"

$adminText = Get-Content -LiteralPath $adminPath -Raw
$consoleText = Get-Content -LiteralPath $consolePath -Raw
$aggregatorText = Get-Content -LiteralPath $aggregatorPath -Raw
$formText = if (Test-Path -LiteralPath $formPath) { Get-Content -LiteralPath $formPath -Raw } else { "" }

$checks = @(
    [pscustomobject]@{ Check = "Slice4bd.SourceSet.FormExists"; Passed = Test-Path -LiteralPath $formPath; Detail = "Admin needs the session-scoped Aggregation Sources form." },
    [pscustomobject]@{ Check = "Slice4bd.SourceSet.PublicHandler"; Passed = $adminText -match '(?s)Public\s+Sub\s+Admin_AggregateGlobalSnapshot_Click\s*\(\s*\).*?frmAggregationSources\.Show'; Detail = "The existing public Aggregator callback must open the source-set form." },
    [pscustomobject]@{ Check = "Slice4bd.SourceSet.ReadOnlyDiscovery"; Passed = ($consoleText -match 'Public\s+Function\s+DiscoverAggregationSourcesForAdmin') -and ($consoleText -match 'Public\s+Function\s+ConnectAggregationServerForAdmin'); Detail = "Admin needs discover/connect services that do not select the operational target." },
    [pscustomobject]@{ Check = "Slice4bd.SourceSet.ExplicitAggregation"; Passed = ($consoleText -match 'Public\s+Function\s+RunHQAggregationFromSourceSet') -and ($aggregatorText -match 'Public\s+Function\s+GenerateGlobalSnapshotFromFiles'); Detail = "Selected published snapshots must be aggregated explicitly, not by scanning one current root." },
    [pscustomobject]@{ Check = "Slice4bd.SourceSet.NoSendToMutation"; Passed = ($formText -match 'ConnectNasRootWithCredentials|GetKnownWarehouseTargetRoots') -and ($formText -notmatch 'SelectWarehouseTarget'); Detail = "The form may connect/discover, but must never change Send To." },
    [pscustomobject]@{ Check = "Slice4bd.SourceSet.VisibleSourceState"; Passed = ($formText -match 'Selected Sources') -and ($formText -match 'Rejected|Skipped'); Detail = "The operator must see selected and rejected/skipped source state." }
    [pscustomobject]@{ Check = "Slice4bd.SourceSet.ConnectedFirstWorkflow"; Passed = ($formText -match 'DiscoverConnectedRoots') -and ($formText -match 'Add Server') -and ($formText -match 'mBtnAddServer'); Detail = "Already connected NAS roots must discover automatically; credentials are an explicit Add Server path." }
)

$checks | Format-Table -AutoSize
$failed = @($checks | Where-Object { -not $_.Passed })
Write-Host ("Slice 4bd multi-server source set: {0} passed, {1} failed" -f ($checks.Count - $failed.Count), $failed.Count)
if ($failed.Count -gt 0) { exit 1 }
