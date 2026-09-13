# Slice 4be.3 persisted Events publication and source-read evidence

Last verified 2026-09-13. The full Release 1 / Slice 4be goal remains active.
The Events artifact and grouped Viewer are not implemented by this checkpoint.
The discovered canonical source-write breach is corrected and regression-validated
in `deploy/validation-publication-source-read`; broader release work remains below.

## Contract and source-read correction

D18's persisted-publication refinement names the Core-owned Events JSON artifact,
whole-group/source coverage structure and existing exact-byte hash convention.
It was recorded in Architecture v4.11, Plan 022 and controls before implementation.
This implements the approved publication separation under semantic inheritance;
it does not change canonical authority, weaken the 5,000-complete-group rule,
or approve partial source coverage. Reader-only clipping cannot satisfy it.

The existing public Admin Generate Inventory Snapshot command pre-opened
canonical Inventory through `ResolveInventoryWorkbookBridge`. Its resolver
can ensure schema and save the workbook. The new test proves Inventory bytes
change during that ostensibly read-only publication source path. Auth, Config,
Outbox, operator workbook and source-copy bytes remain unchanged; the pending
Auth provisioning decision is not implicated or changed.

Admin now delegates its optional source directly to Core's snapshot command.
Core's scoped `cSnapshotInventorySource` borrows a supplied/already-open workbook,
or opens the existing captured-runtime canonical file read-only with events and
macros disabled for the open. It performs no create, schema ensure, repair or
save. It closes without saving only its own transient source, restoring Excel's
prior event/security settings. The existing permission gate and canonical writer
resolver remain unchanged for their other callers. No new cross-XLAM workbook
contract, Application.Run site, form or role-state mutation is introduced.

## D13 evidence

The new `-CheckViewerPublication` mode retains the grouped Viewer suite, then
calls the real public Admin snapshot command. A read-only copy of the generated
published volume fixture supplies the existing event-table read boundary. The
ordinary canonical source still resolves and the actual publication command runs;
no canonical inventory rows are fabricated. The test inspects the persisted
Events artifact independently of Viewer rendering. Its synthetic source has
5,001 groups / 5,003 lines, including the repeated-key/unlike-unit boundary group.

- Initial attempt: **0 PASS / 1 harness failure**, in the earlier Viewer action,
  before publication. No behavioral publication RED is claimed. No bounded
  Application-1000 Excel fault was found. The recorded residual Excel process
  had no workbooks or add-ins during read-only inspection; its ID/start time
  were checked before test cleanup. No runtime cause or fix is inferred.
- Unchanged retry: **8 PASS / 19 FAIL**, including a final hash-probe exception.
  Admin succeeds and the nine Events-artifact assertions fail because no artifact
  exists. The file-hash probe initially disallowed Excel's existing writable
  handle; the test now uses read-only FileAccess with FileShare.ReadWrite.
- Clean combined RED: **8 PASS / 19 FAIL**, without a harness exception. The
  source-preservation assertion now runs and fails. Earlier grouped checks retain
  their **7 PASS / 9 FAIL** scope.
- Publication-only diagnosis: **1 PASS / 10 FAIL**. Per-source Boolean evidence
  isolates canonical Inventory as the only changed file. This mode does not run
  Viewer actions and is prohibited from claiming acceptance GREEN.
- Expanded source-scope baseline: **1 PASS / 12 FAIL**. The actual Inventory
  source is writable, remains open, and changes bytes. These three failures plus
  the successful Admin command are the protecting source-read subset.
- Corrected candidate, complete combined run: **11 PASS / 18 FAIL**. The four
  source-read checks are **4/4 GREEN**; all earlier grouped checks retain 7/9,
  and the nine unimplemented Events-artifact checks remain RED. No failure is
  relabelled or omitted to claim complete Events publication.

The diagnostic-only path prepares the same generated volume fixture without
running the unrelated Viewer action matrix. It keeps separate report naming.
All fixture copies and source hashes remain local; reports expose fixed source
categories and Booleans, not row values, credentials or physical paths.

## Candidate gates

All five candidate packages compile/cold start. Compiled comparison with
`validation-viewer-detail-context` finds only two changed components
(`modWarehouseSync`, `modAdminConsole`) and the new Core class; every other
component hash matches exactly. Static maintenance is regenerated and all three
JSON contracts validate: **202 components / 5,666 procedures / 126,192 lines;
8 literal / 45 unresolved Application.Run sites; 189 duplicate-body candidates**.
All 28 oversized-module ratchets hold. The new class is 63 lines and its two
procedures are at most 39 lines. Admin shrinks; `modWarehouseSync` does not grow.
Detail remains **34/34**, Refresh failure **16/16**, and Settings **187/187**,
without a retry. Populated Viewer regression passes and packaged smoke is
**86/86**. The current default Event Detail capture was inspected and its
default/larger/native maximize/restore geometry remains GREEN. No form layout
changes were made. The unchanged-package full-chain retry passes **31/31**, with
ordered live-role child **48/48** and successful parent/child process completion.
The maintained chain report is dated **2026-09-13 11:27:38**. Full Receiving is
**854/854**, preserving every prior check identity and GREEN with no duplicates.
All **165 prior + 5 candidate package pins** and **16 protected source pins**
match with Excel closed. All 19 changed PowerShell scripts parse, changed evidence
links resolve, and diffs/status were reviewed. Unrelated handoff 067 and critique
023 remain untouched and unstaged. Accepted deployment and operational NAS
workbooks are untouched. Automated captures do not establish human acceptance.

The first full-chain attempt waited in its **Create Warehouse source integration**
child, before ordered live-role execution. Read-only native/VBE inspection found
`User-defined type not defined` in the generated harness's `modWarehouseSync`:
its explicit import list omitted `cSnapshotInventorySource`. This was not an
XLAM compile failure or behavioral RED. After identifying only generated test
workbooks, the waiting runners were stopped and those workbooks closed without
saving; the verified empty residual Excel process was then cleaned up.

The 16 explicit source-harness import lists now include the scoped class and
its missing `modTrainingWire`/`WarehouseTarget` dependencies. All 16 scripts parse.
A retry and isolated source-harness run then failed before executing the macro.
An explicit compile probe located `VERSION 1.0 CLASS` as code: direct Excel import
had treated the new LF-only class export as a standard module. The class export
now uses CRLF, matching VBA's direct-import format; the XLAM builder had already
normalized its input, explaining the earlier packaged compile success. The
unchanged candidate packages were not rebuilt. The repaired standalone Create
Warehouse integration passes **15/15**. Other legacy source harnesses have import
and parser verification only, not newly claimed lifecycle GREEN.

The next chain passed its four Admin/source-entry checks, then the ordered
live-role child stopped at **32 PASS / 1 harness exception**. Its last passing
check was `InventoryDomain.ProjectionRecovery.Delete`; the failed COM call
reported `0x800706BE`. Application-1000 evidence for that owned Excel process
records `ntdll.dll / c0000028`. The root report is **4 PASS / 1 harness exception**;
it is not a completed live-role or chain gate. The subsequent residual Excel
process was verified to contain zero workbooks before cleanup. Root cause is
unresolved; no runtime repair or passing result is inferred from the native fault.
The unchanged-package retry subsequently passes 31/31 with the child 48/48.
That successful retry does not establish the earlier native fault's root cause.

## Reproduction and ignored artifacts

Run Excel validators serially, with Excel closed before each isolated run:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-publication-source-read -CheckViewerPublication -Phase RED
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-publication-source-read -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CheckReceivingLocalActivity -CheckReceivingLifecycleActivity -CheckReceivingNavigationActivity -CheckReceivingSurfaceCoverage -CheckReceivingLauncherDenial -CaptureEvidence
```

`-Phase RED` is intentional while the 18 publication/paging assertions remain
unimplemented. The correction's four check identities are:
`ActualAdminSnapshotCommandSucceeded`, `TransientInventorySourceReadOnly`,
`TransientInventorySourceReleased`, and `SourceCopyAndCanonicalWorkbookBytesUnchanged`,
each prefixed with `ViewerPublication.`. The narrower `-ViewerPublicationOnly
-Phase RED` is diagnostic, not the complete acceptance run.

Ignored root `reports/runtime/` contains:

- `viewer-publication-initial-red.log`, `viewer-publication-setup-faults.json`,
  `viewer-publication-residual-inspection.json`.
- `viewer-publication-retry-red.log`, `viewer-publication-clean-red.log`,
  `viewer-publication-source-diagnostic.log`, `viewer-publication-source-scope-red.log`.
- `publication-source-read-build.log`, `publication-source-read-compile.log`,
  `publication-source-read-compiled.json`, `publication-source-read-focused.log`,
  `publication-source-read-component-comparison.json`,
  `publication-source-read-static.log`, `publication-source-read-static-ratchets.json`.
- Per-run JSON and `publication-source-preservation.json` are under
  `slice4be-viewer-groups/<run-id>/`; diagnostic results are explicitly prefixed
  `diagnostic-publication-`.
- Exact source-scope RED: `64b9778ad4644d59ab6be9d1bba5c3b1/diagnostic-publication-red.json`;
  corrected combined result: `04df0ade13244badb0143768bfa1d49a/red.json`.
- `publication-source-read-detail.log`, `publication-source-read-refresh.log`,
  `publication-source-read-settings.log`, `publication-source-read-viewer.log`,
  `publication-source-read-smoke.log`, `publication-source-read-check-preservation.json`.
  The new detail capture/report is in `slice4be-viewer-detail/5dea972b608b4fe28545b481d1f404d4/`.
- `publication-source-read-fullchain.log` (interrupted source harness),
  `publication-source-read-fullchain-retry.log` (3 PASS / 1 harness exception),
  `publication-source-read-chain-window-state.json`,
  `publication-source-read-chain-compile-location.json`,
  `publication-source-read-chain-project-classification.json`,
  `publication-source-read-chain-empty-recovery.json`.
- `publication-source-read-create-harness.log`,
  `publication-source-read-create-compile-probe.json`,
  `publication-source-read-create-harness-green.log` (15/15),
  `publication-source-read-static-final.log`.
- `publication-source-read-fullchain-final.log`,
  `publication-source-read-fullchain-native-failure.md`,
  `publication-source-read-chain-native-faults.json`,
  `publication-source-read-chain-native-residual.json` preserve the native failure.
- `publication-source-read-fullchain-native-retry.log` records the successful
  complete retry; [full-chain evidence](slice14_results.md) retains sanitized checks.
- `publication-source-read-receiving.log`, `publication-source-read-receiving-green.json`,
  `publication-source-read-receiving-preservation.json` record the fresh 854/854
  run and exact check comparison. The harness writes `launcher-denial-green.json`
  for this full flag set, not the older `navigation-green.json`.
- `publication-source-read-preservation.log`, `publication-source-read-preservation.json`,
  `publication-source-read-script-validation.json`,
  `publication-source-read-import-validation.json` record final preservation and
  parser/import checks.

## Remaining release scope

Implement and validate atomic Events publication, integrity/schema rejection,
all expected owner sources and coverage, the 5,000-group bound and 100-record
paging. Replace Shipping authority reads only when their accepted current-state
data has a published source. Preserve missing/stale/unknown-zone semantics and
all exact-key detail lines. Comprehensive Operations/Admin activity, recording,
How-To/Diagnostic/Compare both, guide import/export, physical NAS/multi-station
and fresh human acceptance remain required. No operational rollout is made here.
