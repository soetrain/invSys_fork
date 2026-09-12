# Plan 022 Slice 4be.1 Receiving worksheet activity

Last verified: 2026-09-12. This is an implementation checkpoint, not completion
of Slice 4be or human Release 1 acceptance.

## Contract and candidate

Architecture v4.11 D18's Receiving worksheet confirmation clarification, Plan
022's worksheet D13 entry and controls v1.87 were committed in documentation
commit `9b70705` before implementation. Catalog 7 registers
`RECEIVING_WORKSHEET_CONFIRM` under `RECEIVING_WORKFLOW`, preserving catalogs
1-6. The native `btnConfirmWrites` handler captures its worksheet/workbook and
trusted context and rechecks context after optional observation before entering
the existing posting owner. Programmatic macro calls remain unattributed.

The candidate changes Operations' Receiving entry/activity adapter and Core's
catalog/source-reference validation. Its five packages are isolated in
`deploy/validation-receiving-worksheet-activity`. Accepted deployment and NAS
workbooks are untouched. Confirmation does not assert Domain application:
owner-returned source identities and Submitted/Unknown states retain their
meaning. Optional observation failure must remain visible without preventing
or repeating the authorized business command.

## Verified D13 evidence

- Minimal native input calibration: **8/8**.
- Original pinned package: **121 PASS / 12 behavioral FAIL**, with both native
  entries and exact shape caller proven, empty-staging owner rejection confirmed,
  and no harness exception. Missing worksheet records/correlation/registered
  metadata are the expected RED.
- Candidate: **133/133 GREEN** for the same focused checks, including native
  entry in sole-visible-sheet and saved/reopened cases, and no activity from
  the separate programmatic macro call.
- Full existing activity regression: **845/845 GREEN**. This is separate from
  native worksheet submission scenarios and does not replace them.
- Five-package build, explicit VBE compile of all five projects and Operations
  cold-start dependency check: **PASS**.
- Source/tooling checks: control surface **6/6**, Operations cutover **14/14**,
  Receiving stabilization **10/10**, Tool contracts **62/62**.

The Receiving source assertion now checks the form wrapper and its shared
action's single owner call separately. Its prior whole-file count included the
distinct compatibility route; correcting that assumption changes no runtime
behavior. The packaged 845-check regression passed before this test correction.

Ignored local evidence under `reports/runtime/slice4be-receiving-activity/`:

- RED: `native-surface-7be919f9abae41169e3f4c9536a8bd14/worksheet-activity/diagnostic-surface-red.json`.
- GREEN: `native-surface-6d1528afc4cd40fcb9f681d007d445a7/worksheet-activity/diagnostic-surface-green.json`.
- `worksheet-activity-red.log`, `worksheet-activity-green.log`,
  `worksheet-full-regression.log`, `worksheet-build.log`, `worksheet-compile.log`.
- `worksheet-package-hashes.json` pins the new five-package set. All 25 pins
  across the original, rebuilt, finalized, compiled and new candidate sets match.

The full prior 845-check baseline is preserved separately at
`reports/runtime/slice4be-worksheet-activity/activity-baseline-845.json`.
Reports contain fixed check identities and pass/fail evidence; raw runtime
records, fixture values, credentials and screenshots are not committed.

## Expanded scenarios and remaining gates

`Slice4beReceivingWorksheetScenarios.ps1` extends the native test to actual
two-event Applied, Pending, UnknownSubmission, StoreFailure and OlderPolicy
cases, with independent inbox/Domain evidence and unknown-column preservation.
Initial attempts failed in fixture preparation before submission and are not
behavioral RED. The diagnostic localized the exception to the native helper's
indexed workbook-window lookup despite an open non-add-in fixture with one
window. Retaining the exact fixture window before activation resolves that
harness failure; no runtime workaround was introduced.

The expanded original-package run finishes normally at **175 PASS / 35 expected
FAIL** across 210 checks, with zero harness exceptions. All five native callers,
independent submission/Domain evidence, unknown headers, pending identities and
custom values, Config bytes and unrelated workbook preservation pass. Its failures
are missing worksheet activity/correlation/reference/read checks and missing
tracking-failure notices. Exact RED:
`reports/runtime/slice4be-receiving-activity/native-surface-6d27c789b4e0452cbf23a05d1d858338/worksheet-activity/diagnostic-surface-red.json`;
log `worksheet-scenarios-captured-window-red.log` in the parent activity directory.
The unchanged candidate then passes **210/210 GREEN**, retaining every RED check
identity with no duplicates. Exact GREEN:
`reports/runtime/slice4be-receiving-activity/native-surface-41842fa702b1402f904f6d09cad3123c/worksheet-activity/diagnostic-surface-green.json`;
log `worksheet-scenarios-green.log`. Applied and Pending worksheet captures were
visually inspected: the intended control and synthetic staging are visible and
the marked input lands on Confirm Writes. These captures precede submission;
outcomes are independently asserted by the test, not inferred from images.
Tracking notices are asserted at the existing notification boundary; this is
automated evidence, not human operator acceptance.

Final static generation passes all JSON contracts: components175, procedures5503
(+2), source lines123685 (+57 for the new registered route and shared adapter),
literal Application.Run8, unresolved dynamic calls45, duplicate candidates195,
scanner candidates1097 and reviewed candidates1099. No module was added and no
oversized cap increased across28 modules; `modTS_Received` shrank1575 ->1572.
Package smoke is **86/86 GREEN**; live-role workflows are **48/48 GREEN**. Ignored logs are under
`reports/runtime/slice4be-worksheet-activity/`.

Viewer regression passes its existing public Operations action, reuse, export,
published-event labels, refresh, rolling date filters, remembered range and
read-only/snapshot-byte checks. Layout validation passes three requested sizes
across five pages, zero geometric out-of-bounds/interactive overlaps, and native
minimize/restore/maximize/restore. The default screenshot was inspected: the
second line of the Inventory Check **Committed / Used** heading appears clipped.
This is an observation in the empty layout-validation form, not yet a proven
regression in the normal launched Production form. Recheck that normal surface
before claiming visible acceptance; geometry-only PASS is not text-legibility
proof. No Production layout implementation changed in this checkpoint.

The first candidate full-chain attempt stopped after four passing Admin/Seed
checks with an RPC failure `0x800706BE`, before Receiving evidence. Windows
recorded an Excel native exception `c0000028` in `ntdll.dll`; no business defect
or runtime repair is inferred. The later empty Excel instance had zero workbooks
or add-ins. During the original-package comparison its recovery-retention prompt
was inspected, **Yes, I want to view these files later** was selected and visually
verified, and the instance exited without discarding recovered files. The
unmodified pinned original package completed the full chain **30/30**. After all
Excel processes exited, the unchanged candidate's clean rerun also completed
**30/30**, with normal exit and Excel closed. This is a standard full-chain GREEN,
not proof that the earlier native crash was fixed or its cause identified.
Failure/comparison logs and sanitized report copies are retained in the same
ignored evidence directory. No speculative runtime or validator change was made.

Windows also recorded `c0000005` in `combase.dll` during the successful chain's
time window. Its 30 passing checks do not resolve Excel native stability. The
first candidate public-launcher run subsequently stopped after Receiving passed,
with another RPC failure and `c0000028` in `ntdll.dll`, leaving Production and
Shipping unverified in that run. The pinned original package passes the same
standard public-launcher gate **3/3**. The clean candidate launcher rerun fails
again after Receiving, at the fixed progress step **invoke Production batch-scale
contract**. The standard isolated candidate Production/restart gate subsequently
passes **2/2**, without reduced flags or diagnostic mutations. Thus the combined
launcher gate remains unresolved despite isolated Production GREEN. All failure
and comparison reports remain separate under the ignored directory. A separate
five-package rebuild from identical source now passes the combined launcher
gate **3/3**, build, all five compiles and cold-start dependency resolution.
It is `deploy/validation-receiving-worksheet-rebuild`; its pins and exported
source hashes are `rebuild-package-hashes.json` and `rebuild-source-report.json`
in the ignored evidence directory. The four changed runtime source files still
match their recorded hashes. The rebuild has distinct binary hashes; this is
an observed successful artifact comparison, not proof of an identified build
defect or native-crash repair. Its focused worksheet checks now pass **210/210**
at `native-surface-7fcd0a43c21b4fa98eca950d40cf0974/worksheet-activity/diagnostic-surface-green.json`
under the activity evidence directory. All30 package pins remain unchanged.
The first candidate's complete845 result is preserved as
`first-candidate-full-845.json` in the worksheet evidence directory. Remaining
rebuilt-package regressions must pass separately before adopting that set for
acceptance. The rebuilt full activity suite passes **845/845**, retaining every
prior check identity with no duplicates (`rebuilt-full-845.json`). Rebuilt smoke
also passes **86/86** (`rebuilt-smoke.log`). Its first live-role attempt ends at
**39 PASS / 1 harness exception**, at Production **Complete Run**, with RPC
`0x800706BE` and Windows `c0000028` in `ntdll.dll`. The remaining empty Excel
instance's recovery-retention dialog was inspected; **Yes, I want to view these
files later** was selected, visually verified and confirmed. Recovered files
were preserved. This failure is retained in `rebuilt-live-role.log` and
`rebuilt-live-role-first-failure.md`; it is not a business RED or passing gate.
The unchanged pinned original then passes the same gate **48/48**
(`original-live-role-comparison.log`). The clean unchanged rebuild then passes
**48/48** (`rebuilt-live-role-clean.log`). Neither passing comparison identifies
or repairs the native cause. Remaining gates run serially against this same set.
The rebuilt ordered full chain now passes **30/30**, including restart,
reconciliation, exact balances/identities, five-package extraction and the
static warning ratchet (`rebuilt-full-chain.log`). Windows nevertheless records
`c0000005` in `combase.dll` during this run's time window. Business/check success
does not establish native stability. Viewer also passes its existing public
launch/reuse/export/filter/read-only contract (`rebuilt-viewer.log`). Rebuilt
layout geometry and native window checks pass (`rebuilt-layout.md`); actual
rendered dimensions remain subject to the current desktop bounds. The earlier
header-legibility finding and normal populated-form review remain open.
The normal full reusable Production/restart gate passes **2/2**, without reduced
flags, debugger or diagnostic mutations (`rebuilt-production.log` and
`rebuilt-production/production-reusable-production.md`). These package-specific
results preserve the required regression scope; unresolved native crash records
and human acceptance are not cleared by them.

## Supplemental native guards

The same rebuilt set passes **201/201** for Denied, SignedOut, SwitchWorkbook,
SignOutDuringTracking and CloseDuringTracking. Adding a second generated,
signed-in warehouse during optional tracking yields **215/215**. All actions
use the actual worksheet button. Unsaved seams observe owner-entry count and
exact workbook object, and deliver one interruption at the optional-tracking
return. They do not replace the posting owner or change deployed packages.
The tests prove no owner retry/redirection, rejected-action authority and staging
preservation, unknown-column preservation, exact references on successful
completion, and no invented result or attribution to the new context.

Initial guard reports under the activity evidence directory:

- `native-surface-47bd74395e554e61991ec3d2a95ff8af/worksheet-activity/diagnostic-surface-green.json` (201).
- `native-surface-d7d17f13d0474aa4adec7d5952813682/worksheet-activity/diagnostic-surface-green.json` (215).

Run the guard suite with Excel closed against the isolated rebuilt package:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-receiving-worksheet-rebuild -Phase GREEN `
  -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity `
  -CheckReceivingStagingActivity -CheckReceivingLocalActivity -CheckReceivingLifecycleActivity `
  -CheckReceivingNavigationActivity -CheckReceivingSurfaceCoverage -ReceivingSurfaceOnly `
  -CheckReceivingNativeSurface -CheckReceivingWorksheetActivity -CheckReceivingWorksheetGuards
```

The separate210 scenario suite replaces the final guard switch with
`-CheckReceivingWorksheetScenarios`; neither focused mode replaces full845.

The strengthened target assertion reads the actual current target's
exact warehouse ID, runtime root and signed-in actor after the real target and
authentication APIs return success. This is supplemental protection of D18,
not a runtime change or retrospective substitute for the recorded 175/35 RED.
Its first rerun stopped before native entry at **94 PASS / 1 harness exception**:
the fixture workbook could not obtain the foreground window. No input was sent
without the existing exact-window guard. `worksheet-exact-target-guards-green.log`
retains that setup failure; it is not behavioral RED or proof of the stronger
assertion. The earlier215 result applies to its original assertion scope.
The unchanged strengthened rerun then passes **215/215**, retaining every check
identity with no duplicates. Exact report:
`native-surface-0e344a9320cc467c8b15fe67fe567a23/worksheet-activity/diagnostic-surface-green.json`;
log `worksheet-exact-target-guards-retry.log`. The second target's warehouse ID,
runtime root and signed-in actor are now directly verified. No native input
helper or runtime implementation was changed to obtain this result.

At this checkpoint, Excel is closed and all30 package pins plus all four runtime
source pins match. Static maintenance and its JSON contracts pass; runtime
components/procedures are byte-for-byte equal in the scanner data to the committed
candidate (175/5503), and all28 oversized-module limits hold. The new Shipping
[source map](plan022_slice4be_shipping_coverage.md) was checked against24 named
controls and25 actual event handlers; this is discovery, not activity GREEN.

Remaining worksheet coverage assessment, native stability and visible
operator acceptance remain open. The
broader Operations/Admin coverage, Event Tracking Settings, comprehensive
Viewer publication/detail, recording/conclusions, guides and comparison remain
required under 4be.1-4be.6. No narrower completion is claimed here.
