# Excel shutdown: isolated table-reference diagnostic

Date: 2026-11-03 Pacific. This is diagnostic tooling, not a product fix or release qualification. The Excel shutdown cause, Inventory GREEN, and Slice 4be-A acceptance remain unresolved.

## Evidence and hypothesis

The retained unassisted diagnostic under `reports/runtime/excel-shutdown-timing/7a9733a34d404c1487dd6b7a936e6c6a` measured Quit at 5.8042 ms, application RCW release at 537.5711 ms, GC collection at 40.0194 ms, and finalizer wait at 334012.7779 ms. Native Excel exit occurred just before the finalizer wait returned. This localizes the measured delay; it does not establish its cause.

The selected hypothesis is that the worker's `Table($Workbook,$Name)` lookup retains intermediate COM references or enumerators. The opt-in `OwnedTableTraversal` mode replaces only that helper in a generated temporary worker, using indexed acquisition and one `ReleaseComObject` per acquired intermediate reference. The borrowed workbook and returned matching table are not released. There is no new GC, Quit, process kill, retry, settings override, or package rebuild.

`Baseline` remains the default and leaves the original helper byte-for-byte unchanged. Both modes retain the original final cleanup calls and original helper-directory binding. The variant is labeled `DiagnosticHelperLifetime`, not observation-only. Both remain `FullAcceptance=false`; any external assisted cleanup disqualifies the run even if the controller's own `AssistedTeardown=false` field remains present.

## Source cohort and compatibility repair

- Diagnostic parent: `97716fa2d3ca69e5fbff63f6a8859235a5576927`.
- Current worker/helper source: `44557070b2f1c3d7d744e1f8040315a837933a8f`, `work/d8-inv-edit-receipts`.
- Current worker SHA256: `5D2D2192F2DAE05F67310A42AF9FA039ABE5E05EFCBB6613EC3C251D066EA4F3`.
- Frozen package candidate: `invSys_fork/deploy/validation-inventory-uom-02`, not `validation-inventory-config-02`.

The old instrument rejected the current worker before dispatch because the current source separates Collect and WaitForPendingFinalizers with an existing inventory trace marker. The adapter now accepts exactly the reviewed combined layout or the reviewed split layout, preserving that marker. Unknown intervening work, duplicate GC calls, changed Table source, and malformed workers refuse before Excel dispatch. This initial refusal is a tooling compatibility failure, not a meaningful packaged product RED.

The retained historical workload had 50 PASS / 28 FAIL. Its worker hash and four helper-file hashes differ from the current source cohort, so its timing is contextual evidence only. A fresh baseline and variant must use identical current inputs; never compare the variant alone against the historical timing and call it a controlled result.

## Source verification

Initial eight-check tooling regression: 4 PASS / 4 FAIL, with the split layout and absent variant failing as expected. After implementation, the expanded suite passes 30/30 on PowerShell 7.6.9 and Windows PowerShell 5.1. The existing timing suite passes 14/14 on Windows PowerShell 5.1, including independent native exit observation of a synthetic PowerShell child, not Excel.

Checks cover both accepted cleanup layouts, refusal on drift, original cleanup call counts, unchanged baseline, only-Table-region transformation, case-insensitive first match, missing/empty lookup behavior, reference-transfer boundaries, and nested cleanup on acquisition failure. Managed stand-ins cannot prove native COM reference behavior. Controller checks additionally cover parsing, default baseline selection, transformation ordering, variant labels, and diagnostic-file preservation reporting.

Native source-test logs are retained under `invSys_fork/reports/runtime/perplexity-shutdown-isolation-source/`: `variants-ps51.txt` and `timing-ps51.txt`. Excel count was zero before and after these tests. All five frozen XLAM hashes still matched the retained candidate.

## Runtime gate

Codex is the normal-Windows-session test executor only. Perplexity owns implementation and interpretation; Codex must return failures without modifying code, packages, fixtures, security settings, or cleanup policy. Runtime requires the separately issued, commit-pinned request with exact package/input hashes, commands, deadline, evidence paths, and cleanup rules.

Only one Excel controller may run. Baseline and variant are sequential, never parallel. Freeze all worker helpers, diagnostic scripts, monitor scripts, and XLAMs from preflight through final preservation checks. Require fresh desktop/cursor proof, complete check identities/outcomes, all four timings, exact native process ownership/exit evidence, restored settings, preserved hashes, and delayed native Excel event audit. A missing or assisted result is incomplete, not fast GREEN.

If the variant changes any assertion identity, order, or outcome, or the baseline unexpectedly differs from the reviewed workload, stop and return the evidence. No broader Inventory GREEN, B0, live-role, restart/binding, full Release 1, or operator acceptance is claimed by this experiment. Preserve existing valid evidence and the approved D18-REPLAY-01 contract.
