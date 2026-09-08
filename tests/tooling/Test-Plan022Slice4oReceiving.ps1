Param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path $RepoRoot).Path
$adminForm = Get-Content (Join-Path $repo "src/Admin/Forms/frmSeedInventory.frm") -Raw
$adminModule = Get-Content (Join-Path $repo "src/Admin/Modules/modAdmin.bas") -Raw
$receivingForm = Get-Content (Join-Path $repo "src/Receiving/Forms/frmReceiving.frm") -Raw
$receivingNavigation = Get-Content (Join-Path $repo "src/Receiving/Modules/modReceivingNavigation.bas") -Raw
$receivingModule = Get-Content (Join-Path $repo "src/Receiving/Modules/modTS_Received.bas") -Raw
$surface = Get-Content (Join-Path $repo "src/Core/Modules/modRoleWorkbookSurfaces.bas") -Raw

$checks = @(
    [pscustomobject]@{
        Name = "DemoInventory.CloseIsSilent"
        Passed = ($adminForm -notmatch 'mBtnCancel|btnCancel') -and
            ($adminModule -notmatch 'Seed inventory cancelled\.')
        Contract = "Demo Inventory has no redundant Cancel button and a window close does not emit a misleading cancellation dialog."
    },
    [pscustomobject]@{
        Name = "Receiving.ConditionIsEstablishedAtReceipt"
        Passed = ($receivingForm -match 'cboCondition') -and
            ($receivingForm -match 'Condition') -and
            ($surface -match 'ReceivedTally.+Condition') -and
            ($surface -match 'AggregateReceived.+Condition')
        Contract = "Receiving captures line Condition and persists it through both staging projections."
    },
    [pscustomobject]@{
        Name = "Receiving.ReturnsIsOperational"
        Passed = ($receivingForm -match 'tabReturns') -and
            ($receivingForm -match 'modReceivingNavigation.ApplyTab\(Me, mTabs.Value\)') -and
            ($receivingNavigation -match 'Add Disposition') -and
            ($receivingNavigation -match 'Disposition reason') -and
            ($receivingForm -match 'RETURN,DUMP') -and
            ($receivingModule -match 'RunReceivingReturnsTabContractForTest')
        Contract = "Receiving exposes operational RETURN/DUMP disposition actions through a public testable form action boundary."
    },
    [pscustomobject]@{
        Name = "Receiving.RefreshRebuildsAggregate"
        Passed = ($receivingForm -match 'RebuildAggregationForWorkbook') -and
            ($receivingModule -match 'BuildReceivingAggregateGroupKey')
        Contract = "Refresh rebuilds the complete grouped Aggregate Received projection from Received Tally."
    },
    [pscustomobject]@{
        Name = "Receiving.ViewerRemainsReadOnly"
        Passed = ($receivingForm -match 'ViewerReadOnly=True')
        Contract = "The Receiving form contract explicitly keeps Condition editing out of Inventory Viewer."
    }
)

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$resultPath = Join-Path $repo "tests/integration/plan022_slice4o_receiving_results.md"
$lines = @(
    "# Plan 022 Slice 4o Receiving Contract Results",
    "",
    "- Passed: $passed",
    "- Failed: $failed",
    "",
    "| Check | Result | Contract |",
    "|---|---|---|"
)
foreach ($check in $checks) {
    $result = if ($check.Passed) { "PASS" } else { "FAIL" }
    $lines += "| $($check.Name) | $result | $($check.Contract) |"
}
[System.IO.File]::WriteAllLines($resultPath, $lines)
$lines -join [Environment]::NewLine
if ($failed -gt 0) { exit 1 }
