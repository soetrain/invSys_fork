[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$receivingFormText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Receiving/Forms/frmReceiving.frm")
$receivingServiceText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Receiving/Modules/modReceivingPostingService.bas")
$receivingControllerText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Receiving/Modules/modReceivingActivityAction.bas")
$productionFormText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Production/Forms/frmProduction.frm")
$productionText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Production/Modules/mProduction.bas")
$readModelText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Core/Modules/modOperatorReadModel.bas")
$processorText = Get-Content -Raw -LiteralPath (Join-Path $repo "src/Core/Modules/modProcessor.bas")

function Procedure-Text {
    param(
        [string]$Text,
        [string]$Name
    )

    [regex]::Match(
        $Text,
        "(?ms)^(?:Public|Private) (?:Function|Sub) $([regex]::Escape($Name))\b.*?^End (?:Function|Sub)"
    ).Value
}

$receivingCallback = Procedure-Text -Text $receivingFormText -Name "mBtnConfirm_Click"
$receivingAction = Procedure-Text -Text $receivingServiceText -Name "ExecuteConfirmWrites"
$productionCallback = Procedure-Text -Text $productionFormText -Name "mBtnManagerApplyOutput_Click"
$productionFormAction = Procedure-Text -Text $productionFormText -Name "CompleteProductionRun"
$productionAction = Procedure-Text -Text $productionText -Name "CompleteProductionRunAfterCheckInForOutput"
$sharedRuntime = Procedure-Text -Text $readModelText -Name "RunBatchAndRefreshOperatorWorkbook"

$processorCall = $sharedRuntime.IndexOf("modProcessor.RunBatch", [System.StringComparison]::OrdinalIgnoreCase)
$duplicatePublishCall = $sharedRuntime.IndexOf("PublishInventorySnapshotBridge", [System.StringComparison]::OrdinalIgnoreCase)

$checks = @(
    [pscustomobject]@{
        Check = "Receiving.Persistence.FormSummary"
        Passed = ($receivingCallback -match 'modReceivingActivityAction\.ConfirmWrites') -and
            ($receivingControllerText -match 'modReceivingPostingService\.ExecuteConfirmWrites') -and
            ($receivingCallback -match 'ShowStatus\s+statusMessage') -and
            ($receivingAction -match 'Persistence summary:') -and
            ($receivingAction -match 'receiving inbox batch saved') -and
            ($receivingAction -match 'processor durability saves retained')
        Contract = "The real Confirm Writes/Dispositions callback reports its batched inbox and processor persistence once in Receiving txtStatus."
    },
    [pscustomobject]@{
        Check = "Production.Persistence.FormSummary"
        Passed = ($productionCallback -match 'CompleteProductionRun') -and
            ($productionFormAction -match 'CompleteProductionRunAfterCheckInForOutputResult') -and
            ($productionFormAction -match 'ShowStatus') -and
            ($productionAction -match 'Persistence summary:') -and
            ($productionAction -match 'Production inbox events saved') -and
            ($productionAction -match 'processor durability saves retained')
        Contract = "The real Complete Run callback reports Production event and processor persistence once in Production txtStatus."
    },
    [pscustomobject]@{
        Check = "Production.Persistence.QuietBoundary"
        Passed = ($productionAction -match 'modUiQuiet\.BeginQuietUi') -and
            ($productionAction -match 'modUiQuiet\.EndQuietUi') -and
            ($productionAction -match 'ExecuteProductionSession')
        Contract = "Production Complete Run keeps one quiet-UI boundary around queue, processor, and refresh persistence."
    },
    [pscustomobject]@{
        Check = "Operations.Persistence.SingleSnapshotOwner"
        Passed = ($processorCall -ge 0) -and
            (($duplicatePublishCall -lt 0) -or ($duplicatePublishCall -lt $processorCall)) -and
            ($processorText -match 'EventPersistenceSaves=') -and
            ($processorText -match 'GenerateWarehouseSnapshot')
        Contract = "The processor remains the snapshot and durability owner; shared Receiving/Production refresh does not publish a second snapshot."
    }
)

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$resultPath = Join-Path $repo "tests/integration/plan022_slice4v_receiving_production_persistence_feedback_results.md"
$lines = @(
    "# Plan 022 Slice 4v Receiving and Production Persistence Feedback Results",
    "",
    "- Passed: $passed",
    "- Failed: $failed",
    "",
    "| Check | Result | Contract |",
    "|---|---|---|"
)
foreach ($check in $checks) {
    $lines += "| $($check.Check) | $(if ($check.Passed) { 'PASS' } else { 'FAIL' }) | $($check.Contract) |"
}
Set-Content -LiteralPath $resultPath -Value $lines -Encoding UTF8

$checks | Format-Table Check, Passed -AutoSize
Write-Host ("Plan 022 Slice 4v Receiving/Production persistence feedback: {0} passed, {1} failed" -f $passed, $failed)
if ($failed -gt 0) { exit 1 }
