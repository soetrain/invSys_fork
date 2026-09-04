[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$buildPath = Join-Path $repo "tools\build-xlam.ps1"
$adminPath = Join-Path $repo "src\Admin\Modules\modAdmin.bas"
$consolePath = Join-Path $repo "src\Admin\Modules\modAdminConsole.bas"
$formPath = Join-Path $repo "src\Admin\Forms\frmAggregationSources.frm"

$buildText = Get-Content -Raw -LiteralPath $buildPath
$adminText = Get-Content -Raw -LiteralPath $adminPath
$consoleText = Get-Content -Raw -LiteralPath $consolePath
$formText = Get-Content -Raw -LiteralPath $formPath

$checks = @(
    [pscustomobject]@{
        Check = "Slice4bd.Ribbon.AdminAggregatorCommand"
        Passed = $buildText -match '(?s)Id\s*=\s*"btnAdminHqAggregator".*?Label\s*=\s*"Aggregate Global Snapshot".*?Macro\s*=\s*"modAdmin\.Admin_AggregateGlobalSnapshot_Click".*?RequiredCapability\s*=\s*"ADMIN_MAINT"'
        Detail = "Admin must expose an ADMIN_MAINT-gated Aggregate Global Snapshot command."
    },
    [pscustomobject]@{
        Check = "Slice4bd.PublicAdminHandler"
        Passed = $adminText -match '(?s)Public\s+Sub\s+Admin_AggregateGlobalSnapshot_Click\s*\(\s*\).*?frmAggregationSources\.Show'
        Detail = "The public Ribbon callback must open the explicit source-set action."
    },
    [pscustomobject]@{
        Check = "Slice4bd.AdminAuthorityAndAdvisoryBoundary"
        Passed = ($consoleText -match '(?s)Public\s+Function\s+RunHQAggregationFromSourceSet.*?EnsureAdminContext.*?RequireAdminMaintenance.*?GenerateGlobalSnapshotFromFiles.*?AppendAuditEntry') -and
                  ($formText -match 'RunHQAggregationFromSourceSet') -and
                  ($consoleText -match 'ADVISORY')
        Detail = "The source-set action must enforce Admin context, aggregate only explicit sources, audit the attempt, and retain advisory wording."
    }
)

$checks | Format-Table -AutoSize
$failed = @($checks | Where-Object { -not $_.Passed })
Write-Host ("Slice 4bd Admin Aggregator command: {0} passed, {1} failed" -f ($checks.Count - $failed.Count), $failed.Count)
if ($failed.Count -gt 0) { exit 1 }
