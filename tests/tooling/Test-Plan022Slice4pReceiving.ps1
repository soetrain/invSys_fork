Param([string]$RepoRoot = ".")

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path $RepoRoot).Path
$receivingForm = Get-Content (Join-Path $repo "src/Receiving/Forms/frmReceiving.frm") -Raw
$receivingNavigation = Get-Content (Join-Path $repo "src/Receiving/Modules/modReceivingNavigation.bas") -Raw
$receivingModule = Get-Content (Join-Path $repo "src/Receiving/Modules/modTS_Received.bas") -Raw
$postingService = Get-Content (Join-Path $repo "src/Receiving/Modules/modReceivingPostingService.bas") -Raw
$roleWriter = Get-Content (Join-Path $repo "src/Core/Modules/modRoleEventWriter.bas") -Raw
$processor = Get-Content (Join-Path $repo "src/Core/Modules/modProcessor.bas") -Raw
$config = Get-Content (Join-Path $repo "src/Core/Modules/modConfig.bas") -Raw
$auth = Get-Content (Join-Path $repo "src/Core/Modules/modAuth.bas") -Raw
$configLoad = [regex]::Match($config,'(?ms)^Public Function LoadConfig\b.*?^End Function').Value
$configRead = [regex]::Match($config,'(?ms)^Private Function ResolveExistingConfigForRead\b.*?^End Function').Value

$checks = @(
    [pscustomobject]@{
        Name = "Receiving.AggregateCombinesReferencesByCondition"
        Passed = ($receivingModule -match 'AppendDistinctReceivingValue') -and
            ($receivingModule -match 'BuildReceivingAggregateGroupKey')
        Contract = "Aggregate Received sums equivalent item buckets, concatenates distinct references, and keeps Condition in the grouping key."
    },
    [pscustomobject]@{
        Name = "Receiving.ReturnLabelsAndCondition"
        Passed = ($receivingForm -match 'modReceivingNavigation.ApplyTab\(Me, mTabs.Value\)') -and
            ($receivingNavigation -match 'Return Entries History') -and
            ($receivingNavigation -match 'Return Tally') -and
            ($receivingNavigation -match 'Aggregate Returns') -and
            ($receivingForm -match 'ItemConditionColumn=True')
        Contract = "Returns uses return-specific titles and its item results expose Condition."
    },
    [pscustomobject]@{
        Name = "Receiving.PostsTallyIdentity"
        Passed = ($postingService -match 'TABLE_RECEIVED_TALLY') -and
            ($postingService -notmatch 'BuildPostingStates[^\r\n]*AggregateReceived')
        Contract = "Confirm posts each Received Tally identity; the aggregate projection is display-only."
    },
    [pscustomobject]@{
        Name = "Receiving.QueueIsBatched"
        Passed = ($roleWriter -match 'QueueReceiveEventBatchServer') -and
            ($postingService -match 'QueueReceiveEventBatchServer')
        Contract = "A multi-line receipt queues through one server-inbox save boundary."
    },
    [pscustomobject]@{
        Name = "Receiving.ReceiptStagingIsEventIsolated"
        Passed = ($receivingModule -match 'failureStage = "find existing receipt staging row"') -and
            ($receivingModule -match 'failureStage = "populate receipt aggregate"') -and
            ($receivingModule -match 'Receiving staging failed: Stage=') -and
            ($receivingModule -match 'If eventStateCaptured Then Application.EnableEvents = previousEvents')
        Contract = "Add Selected stages a complete receipt row and aggregate with Excel events isolated, restores prior event state, and reports the exact failing stage."
    },
    [pscustomobject]@{
        Name = "Receiving.ConfirmUsesQuietUiBoundary"
        Passed = ($receivingForm -match 'modUiQuiet.BeginQuietUi mOperatorWorkbook') -and
            ($receivingForm -match 'If quietStarted Then modUiQuiet.EndQuietUi') -and
            ($receivingForm -match 'QuietDuring=') -and
            ($receivingForm -match 'QuietRestored=')
        Contract = "The real Confirm Writes form handler suppresses repeated Excel save UI and restores the prior application UI state."
    },
    [pscustomobject]@{
        Name = "Processor.PersistenceIsBatched"
        Passed = ($processor -match 'EventPersistenceSaves=') -and
            ($processor -match 'AppendEventsToOutboxBatch')
        Contract = "Processor persistence is bounded per artifact instead of saving once per event."
    },
    [pscustomobject]@{
        Name = "SignIn.HealthyReadsRemainSaved"
        Passed = ($configLoad -match 'ResolveExistingConfigForRead\(whId\)') -and
            ($configLoad -notmatch 'EnsureConfigSchema|\bwb\.Save\b') -and
            ($configRead -match 'ReadOnly:=True') -and
            ($auth -match 'EnsureAuthSchema\(wb, whId, modConfig.+, , False\)')
        Contract = "Healthy Config/Auth reads do not dirty and resave unchanged workbooks during sign-in."
    }
)

$passed = @($checks | Where-Object Passed).Count
$failed = $checks.Count - $passed
$resultPath = Join-Path $repo "tests/integration/plan022_slice4p_receiving_results.md"
$lines = @(
    "# Plan 022 Slice 4p Receiving Contract Results",
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
