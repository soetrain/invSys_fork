# Slice 4be.1 Receiving activity RED

Last verified: 2026-09-07. Architecture v4.11 D18 and Plan 022 govern this
test-first checkpoint. No runtime VBA, form implementation or XLAM was changed.
The existing activity foundation remains limited to Settings and UOM retrieval;
this checkpoint does not complete Receiving coverage or Slice 4be.

## Reproduction and observed result

With Excel closed, run the unchanged foundation candidate:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/validation-activity -Phase RED -CheckReceivingActivity
```

Result: **30 PASS / 8 FAIL**, 38 checks, zero harness exceptions. All 18 D5
checks remain GREEN. Twelve new Receiving business/binding guards pass. The
eight failed assertions cover four requirements in each of two scenarios:
attempt/result presence, stable ActivityId correlation, every exact source
EventId, and truthful data-effect evidence. They reflect absent Receiving
observations, not eight independent runtime defects.

The five candidate package hashes match the unchanged `03f7f20` foundation
[package evidence](plan022_slice4be_activity_foundation_results.md). Neither
`deploy/current` nor an operational warehouse was changed. Local result JSON
under `reports/runtime/slice4be-receiving-activity/red.json` contains check names
and booleans only and remains ignored.

## Protecting operator path and independent evidence

The test reuses the packaged D5 Admin Generate Warehouse fixture and accepted
Admin Seed automation boundary. Two calls to the real Receiving Add handler
create two different receipt entities and event identities at the owning
staging boundary. The test does not invent or import inventory keys.

An unsaved test seam retains that actual `frmReceiving` instance, activates a
different saved workbook, and invokes the existing
`TestRunConfirmWritesActionForWorkbook` -> `mBtnConfirm_Click` handler. Runtime
handler implementation is untouched.

| Scenario | Independent proof before checking missing activity |
|---|---|
| Applied | Both staged EventIds exist in the inbox, `tblAppliedEvents`, and `tblInventoryLog`; each log line retains its exact generated System_Key and quantity; confirmation succeeds and staging clears. |
| Pending | Real confirmation queues both events; a test-only Core bridge gate withholds processor/refresh execution. Both inbox rows exist, neither event is in the Domain applied/log tables, confirmation fails and both staging rows remain. |

Both cases preserve the captured Receiving workbook, quiet-UI entry/restoration,
the other workbook's file bytes and sentinel content, and the unknown staging
column header. The staging-column assertion does not claim preservation of
values in rows intentionally cleared by successful confirmation.

Source tables are read independently only inside generated disposable fixtures.
This is test evidence, not a new Viewer authority-read path. Activity assertions
require source references by exact EventId; a pending result cannot claim
COMPLETED/APPLIED or a known Domain data effect. The normal activity result may
retain Unknown until the owner supplies trustworthy per-event application
evidence. Future published diagnostics still require all terminal source events;
this RED does not prove that later publication/evaluation requirement.

## Harness diagnosis excluded from RED

- The initial unrelated-workbook hash read used incompatible file sharing;
  the corrected read opens a read-only stream allowing Excel's existing handle.
- Admin-generated inboxes use the configured inbox directory. The test now
  resolves the path through the existing Core resolver and verifies it stays
  inside the generated runtime instead of assuming a root-level inbox.
- Editing a referenced VBA project after creating the form reset its globals.
  Fault instrumentation is now installed before fixture/session/form creation,
  then toggled through a test-only primitive without editing live code.
  The failed dialog and empty recovery Excel process were verified as owned
  test state and closed. The final run ended with no Excel process remaining.
- Excel changed the unrelated workbook's Saved flag from True to False in both
  cases. Its saved bytes, sentinel cell, sheet count and empty table collection
  were unchanged. Separate binding/UI/content checks replace the misleading
  combined Saved-flag assertion; no runtime repair or new architectural
  interpretation was introduced.

## Next implementation and remaining gates

Add the Receiving owner-supplied result/source envelope and actual-handler
observation boundary under D18. Keep source references distinct from application
proof, preserve session/workbook binding, and keep optional tracking failure
independent from business command success. Extend strict source-reference,
denial/rejection, stale-session, repeated-action and tracking-failure tests
before their corresponding changes. Confirm Dispositions and other controls
remain separate pending coverage.

No new GREEN, compile, layout, static-maintenance, broad role/chain regression,
deployment or human acceptance is claimed here. Those gates remain mandatory
for the implementation candidate. PowerShell parsing and diff checks pass for
this test-only checkpoint; the prior foundation gate record remains historical.
