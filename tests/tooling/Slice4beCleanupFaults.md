# Slice 4be-A harness cleanup fault proof

This is a test-harness defect correction, not a product/replay contract change. A workbook-close exception previously escaped to a shared empty catch, skipping Excel.Quit and silently omitting the cleanup failure. Quit now has an independent finally stage, with typed failed checks for workbook-cleanup and Quit-stage exceptions. No raw exception content is included in those checks.

## Focused proof

Run `Test-Slice4beCleanupFaults.ps1` with explicit `-HarnessPath` and a fresh, nonexistent `-EvidenceRoot`. It parses the supplied harness, extracts the actual bounded cleanup block and actual Check function, then executes them against managed stand-ins. It never executes the full harness, creates COM, invokes actual Excel, changes cursor/input state, records a trace, or force-stops a process.

Eight cases cover empty and two-workbook success, diagnostics enabled/disabled, first/last workbook-close failure, Quit failure, and combined failures. Fifty-six assertions cover exactly one Quit attempt, typed failure records, lifecycle-marker order, unchanged Close(false)/stop-on-first-error/no-retry behavior, non-leaking failure output and failure persistence through JSON serialization. The latter proves fixture serialization, not the full harness report/exit pipeline.

Evidence files are `summary.json`, `checks.json`, and `cases.json` in each new root; retain stdout, stderr and actual child exit status separately. Parse/boundary/setup failures are not acceptable behavioral RED.

## Results recorded November 4, 2026

- **RED first:** test-only commit `4f1b4cafb073ca200797795544fd55063cbfb288`, against the unchanged harness from `44557070b2f1c3d7d744e1f8040315a837933a8f`: 40 PASS / 16 FAIL, child exit 1.
- **GREEN after fix:** same test bytes, 56 PASS / 0 FAIL, child exit 0, on cloud PowerShell 7.6.9. Normal Windows PowerShell 5.1 confirmation is pending.
- **Windows RED retained:** the [Windows report](https://github.com/soetrain/invSys_fork/pull/8#issuecomment-6640328284) confirms preparation, 40 PASS / 16 exact expected FAIL and exit 1. GREEN never started because the supervisor wrapped the JSON array in another array.
- **JSON compatibility correction:** direct assignment replaces array-subexpression capture in the fixture's serialization assertion and in the continuation supervisor. Cloud simulation of the reported Windows array-output behavior first reproduced 55 PASS / 1 false FAIL in `CloseAndQuitThrow.FailureSurvivesReportSerialization`; the correction restores 56/56. It changes neither the cleanup implementation nor the 56 assertions. Existing Windows RED evidence remains valid; Windows GREEN still requires one first dispatch with the corrected fixture.
- **Source regression:** existing PR7 shutdown-variant suite, including transformation compatibility and preserved final cleanup calls: 30 PASS / 0 FAIL.
- **Known validation limitation:** `Test-LauncherCleanupEvidence.ps1` stops because `Wait-ReusableAutomationExit` is unavailable under cloud PowerShell 7.6.9, identically on baseline and candidate. This is not reported as GREEN or attributed to this patch.
- **Static maintenance:** the candidate scanner/schema validation passes: 6 packages, 391 components, 6,742 procedures, 1,246 scanner candidates, 1,248 reviewed candidates and 27 oversized-module ratchets. No VBA, callback, layout, build or package bytes change in this patch.
- **Frozen Windows inputs:** an independent read-only hash check at 2026-11-04T19:17:23Z found 535/535 unchanged and Excel count zero. The candidate exists in a separate source checkout, not the frozen runtime directories.

## Boundaries

Workbook enumeration still stops at its first failure; this patch does not add close retries or release more COM references. Existing settings/trace hooks remain in their original order; faults inside those hooks are not newly covered. Final COM release, GC and finalizer behavior after this block are untouched.

This unit proof does not establish native COM lifetime, packaged callback acceptance, compile/layout/live-role/restart/full Release 1 acceptance, or a shortened shutdown. No XLAM was rebuilt or run. The earlier OwnedTableTraversal experiment cannot settle other retained-reference hypotheses, and the observed Excel kernel CPU does not identify a cause. Excel/capture remain paused; the cursor policy block is not bypassed. Inventory GREEN, early B0, 4be-A closure and the shutdown cause remain unresolved.
