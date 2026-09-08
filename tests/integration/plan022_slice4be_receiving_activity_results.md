# Slice 4be.1 Receiving activity candidate

Last verified: 2026-09-07. Architecture v4.11 D18 and Plan 022 govern this
implementation checkpoint. Comprehensive control coverage, Events publication,
Settings policy editing and both Action Path presentations remain incomplete.
Candidate XLAMs are in `deploy/validation-activity`; accepted `deploy/current`
and operational warehouses are unchanged.

## Behavior and authority

The real Receiving Receipts-tab Confirm Writes handler captures its session and
workbook, records an attempt, calls the existing posting owner once, and records
that owner's outcome and every related source EventId. The service still owns
validation, queue submission, processing/refresh and staging cleanup. Core's
activity boundary observes and validates; it does not perform business commands.
Direct posting is not an operator-control event. Receiving Add and Confirm
Dispositions activity coverage remain pending.

Catalog 2 adds `RECEIVING_CONFIRM_WRITES`. CONFIRMED and PENDING both retain
Unknown Domain effect. Submitted references mean confirmed queue acceptance;
Unknown retains possibly submitted identities after uncertain failures. Exact
WarehouseId/SourceKind/EventId/SubmissionState references reject extra fields,
duplicates, invalid identity text and unsupported source/state/catalog values.
An aggregate successful batch is never per-event application proof.

Supported catalog-1 records and policies retain their definitions. An older
policy does not implicitly enable a new control. Stale sessions reject the
business action independently of tracking. An unavailable activity store adds
Tracking unavailable while otherwise authorized Receiving still completes.

The focused fixture also exposed a Core setup ownership defect: implicit Config
resolution did not report newly opened workbooks to the existing cleanup path.
Admin Seed's station-inbox setup left Config open; Excel subsequently marked it
dirty, so optional policy reads correctly refused that pre-existing dirty book.
The resolver now records ownership before opening and closes only its own book.
Pre-existing Config and unknown columns survive. Fresh read-only policy reads
do not save the file; unrelated dirty open Config remains unavailable/preserved.

## Test-first evidence

- Initial unchanged foundation RED: **30 PASS / 8 FAIL**; see
  [initial Receiving RED](plan022_slice4be_receiving_activity_red_results.md).
- Expanded source-reference/stale-session/store-failure RED before runtime
  implementation: **54 PASS / 12 FAIL**, no harness exceptions.
- Setup ownership RED before its fix: **61 PASS / 10 FAIL**. The new implicit
  ownership assertion failed; existing-open/extra-column guards passed. The
  remaining nine failures came from unavailable observations/valid record read.
- First combined GREEN after the ownership fix: **123 PASS / 0 FAIL**, retaining
  all 70 foundation checks and adding 53 Receiving checks.
- Expanded delivery, compatibility, redaction and visible-capture run:
  **133 PASS / 0 FAIL**. Catalog-1 records remain readable; catalog-1 policy
  does not enable the new control. Identical completion retries preserve every
  byte; conflicting source references are rejected. Independent SHA-256 and
  private-input exclusion checks protect both Receiving attempt/result pairs.

Actual Add handlers generate two distinct keys/events for each disposable case.
Applied proves both exact keys/quantities in inbox, applied events and inventory
log, then empty staging. Pending proves both submitted but neither applied nor
logged, with staging retained. Stale proves no submission and visible rejection.
Store failure proves business completion and a visible tracking warning without
a durable local fallback. All cases protect captured binding, quiet UI, other
workbook file bytes/sentinel content and the extra staging column header.

Synthetic wire/delivery checks supplement the real form tests; they are not
captured user actions or Diagnostic conclusions. Runtime JSON, fixture values,
screenshots and raw logs stay ignored. A COM fixture exception in an intermediate
run was excluded from behavioral RED; the verified empty owned recovery Excel
session was closed before a fresh serial retry. After the successful full-chain
gate, a verified empty recovery Excel was closed normally and the identified
completed validator process with disconnected automation windows was stopped.
The next focused run started only after its Excel-isolation guard passed.

## Gates

| Gate | Candidate result |
|---|---|
| Cold start and explicit compile | Five packages PASS |
| Packaged smoke | 81/81 PASS |
| Live-role | 48/48 PASS |
| Full Release 1 chain | 30/30 PASS, including restart/reconciliation and five-package runtime evidence |
| Viewer | PASS; displayed-list export, event labels/reservation exclusion, refreshed date filters/preferences and snapshot byte preservation |
| Production layout/window states | PASS; three size requests across five pages, minimize/restore/maximize/restore, no overlap/out-of-bounds; all three captures inspected. Smaller request clamps to approved minimum |
| Packaged launcher reuse | 3/3 PASS; no eligible workbook initially open; Receiving/Production/Shipping provision and reuse their saved role workbook/form |
| Reusable Production/restart | 2/2 aggregate checks PASS; worksheet/lifecycle/run/Chai scenarios and a fresh Excel restart reuse the same saved workbook and released Recipe |
| Visible Receiving evidence | Four actual form captures inspected: applied, pending, stale and unavailable store; staging and completion/warning text visible. Settings save/tracking warning also inspected. Human acceptance remains pending |
| Static maintenance | 1,077 candidates; 192 duplicate groups; 45 unresolved dynamic calls; eight literal Application.Run targets; unchanged |
| Module growth | All 28 previous oversized ratchets respected; Receiving form 1,260 -> 1,259 lines; Config remains 1,615 |

Source regressions pass: behavior locks 13, Receiving stabilization 10,
persistence feedback 4, Receiving 8, dispositions 6, launcher contracts 24,
Production Run List layout 7, Production layout 8, full-chain contract 13,
Batch Note/Viewer 5, UOM 7 and variable quantity modes 6. Existing source checks
now follow form -> observation controller -> the same posting owner. The old
Receiving sign-in source check was corrected to the already accepted D5
read-only resolver; it no longer requires retired schema repair on reads.

## Reproduction

With Excel closed, run validators serially:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build-xlam.ps1 -OutputRoot deploy/validation-activity -Apply
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-PackagedVbaCompile.ps1 -DeployRoot deploy/validation-activity
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/validation-activity -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CaptureEvidence
```

The RED counts describe earlier candidate revisions. Running current packages
with `-Phase RED` does not recreate those revisions. The ignored ownership RED
is preserved as `reports/runtime/slice4be-receiving-activity/ownership-red.json`.
Do not rebuild candidates while an Excel validator is still running.

## Candidate package hashes

SHA-256 for the candidate used by these gates:

| Package | SHA-256 |
|---|---|
| invSys.Core.xlam | 3211070f1eaeaf2dde9dc171671f07e4933d223e9b490eede639534b30f30a6f |
| invSys.Inventory.Domain.xlam | 506fbb4f11251e16cc0357b3544af5fe3958661428488621538c4e74a300b301 |
| invSys.Designs.Domain.xlam | 0afa4f4248b1487a6b1ce29425898fe760390688d3e5f3235f2d44fd1250960d |
| invSys.Operations.xlam | 4338d2f6da02f882c16eb6db5388ba097fbbb130ae05211ec15009f5aa9867de |
| invSys.Admin.xlam | 9928e3dfc71ea6914b8e6c17752edcc4ea04983b49cb2b3f973c100edf2d5f0e |
