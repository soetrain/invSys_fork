# Slice 4be.1 Receiving staging and disposition activity

Architecture v4.11 D18 governs this extension of the approved shared activity
contract. This Receiving staging/disposition candidate checkpoint is GREEN.
Comprehensive 4be.1 coverage and Release 1/human acceptance remain incomplete.

## Focused RED, 2026-09-07

Unchanged runtime code `1689bff` and the five package hashes recorded in
[Receiving confirmation evidence](plan022_slice4be_receiving_activity_results.md)
were used for both runs. Tests invoke actual packaged `mBtnAdd_Click` and
`mBtnConfirm_Click` handlers using unsaved fixture-only access seams.

- Initial new coverage: **206 PASS / 15 FAIL**. Missing Add Selected,
  Add Disposition and Confirm Dispositions attempt/outcome records.
- Expanded rejection and protected-staging failure coverage: **210 PASS /
  35 FAIL**. No harness errors; all previous 190 checks remain GREEN.
- Independent guards prove two actual Add operations, three subsequently
  submitted exact event identities (including one direct-service staging call),
  all receipt or RETURN/DUMP event applications once, captured workbook binding,
  preservation of the unrelated workbook and extra staging header, and no
  user-control attribution for the direct staging service call.
- Invalid quantity and a protected staging sheet exercise the actual Add handler
  before runtime edits. Both preserve the entire staged row data and show the
  existing cause. Expected observation effects differ: pre-validation is
  Unchanged, while service failure remains Unknown without inferring rollback.

Ignored evidence: `reports/runtime/slice4be-receiving-activity/staging-red.json`
and `staging-outcomes-red.json`. Reports contain check names and booleans only;
raw activity, fixture values and screenshots remain outside committed evidence.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/validation-activity -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CaptureEvidence
```

## Candidate GREEN

First GREEN is **245/245**. The final expanded focused run is **262/262**, with
all previous 190 checks retained. No runtime edits separate those GREEN runs.
Additional checks prove readable catalog-3 records, valid catalog-2 policies
without implicit collection of any new control, preserved staging after session
renewal, and successful Add with a visible warning when optional tracking fails.
The final submission cases each apply four exact source events once, including
one Add performed while tracking was unavailable and one direct service staging
call. Neither receives invented Add activity; Confirm still references all four.

The actual form Add boundary records REQUESTED, performs the existing validation
and staging call once, and records STAGED, REJECTED or FAILED. The new small
Receiving UI helper contains the extracted input validation and isolated finish
notification. Direct staging services are unchanged and create no activity.
Confirm Dispositions uses the existing posting owner and its exact source IDs.
Catalog 3 preserves older definitions; observations never imply Domain application.

Eleven actual form captures were inspected: receipt/disposition staged, rejected,
protected-sheet failure and tracking-unavailable states, both completed
confirmations, and stale Add rejection. The existing status retains the business
cause and shows tracking failure alongside successful staging. These automated
captures are evidence, not human user acceptance. They remain ignored under
`reports/runtime/slice4be-receiving-activity/coverage-*.png`; final boolean results
are preserved in `staging-full-green.json`.

| Gate | Current candidate result |
|---|---|
| Five-package build, cold start and explicit compile | PASS |
| Focused actual-handler and foundation suite | 262/262 PASS |
| Static maintenance | 1,077 candidates, 192 duplicate groups, 45 unresolved dynamic calls and eight literal Application.Run targets; unchanged |
| Code growth | All 28 oversized-module ratchets hold; Receiving form 1,259 -> 1,235 lines; one 39-line UI helper and three procedures added overall |
| Receiving source regressions | Behavior locks 13, stabilization 10, persistence feedback 4, Receiving 4o 5 and 4p 8, ListBox 7: all PASS |
| Packaged smoke | 81/81 PASS |
| Live-role | 48/48 PASS |
| Full Release 1 chain | 30/30 PASS, including restart/reconciliation and five-package runtime evidence |
| Viewer | PASS: actual displayed-list export, accepted event labels/reservation exclusion, refreshed date filters/preferences and snapshot byte preservation |
| Production layout/window states | Three size requests across five pages and minimize/restore/maximize/restore PASS; all three captures inspected. Minimum request clamps to the default minimum |
| Packaged launcher reuse | 3/3 PASS from NoEligible: Receiving, Production and Shipping open/reuse their saved operator workbook/form; repeated 3/3 with cooperative dialog-observer shutdown |
| Reusable Production/restart | Two independent full runs each 2/2 PASS, including fresh Excel processes, with cooperative dialog-observer shutdown; no focused-only switches |

No accepted deployment or operational workbook was rebuilt. The candidate is
`deploy/validation-activity`; all checkpoint gates above pass on its recorded hashes.
Both original failed attempts are retained as
ignored `production-harness-rpc*.md` reports under
`reports/runtime/slice4be-staging-production`. The second progress marker was the
Production output-regulation handler; the first stopped earlier. Windows reports
native Excel failures in ntdll.dll (c0000028), which identifies the interruption
but does not establish its cause or rule out a regression. A separate detached
checkout at pre-change `f813f2f` builds into `deploy/validation-staging-baseline`
for comparison using the same reusable Production validator. All five baseline
packages compiled, but its validator also failed with the same native/RPC error,
at the released Process edit/export step. This proves the interruption is not
unique to the new runtime; it does not establish a cause or satisfy the gate.

The validator now signals its dialog observer to stop and waits for a terminal
job state instead of forcibly stopping it during a possible UI Automation call.
An abnormal observer remains a failed gate. With unchanged runtime packages and
workflow assertions, two independent cooperative runs each pass reusable
Production and clean restart (2/2). Excel closes normally, all five candidate
hashes remain unchanged, and neither run records an Excel native crash event.
Revised-harness launcher checks also pass 3/3. This supports the harness
hypothesis without establishing the native crash's cause; the failed candidate
and pre-change runs remain part of the evidence. Launcher/restart source checks
pass 24/24 and 6/6 respectively. No workflow assertion was weakened or skipped.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_plan022_packaged_launchers.ps1 -DeployRoot deploy/validation-activity -OutputDirectory reports/runtime/slice4be-staging-production-cooperative -CallbackFilter Production -WorkbookState ProductionReusable
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_plan022_packaged_launchers.ps1 -DeployRoot deploy/validation-activity -OutputDirectory reports/runtime/slice4be-staging-launchers-cooperative -WorkbookState NoEligible
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_plan022_packaged_launchers.ps1 -DeployRoot deploy/validation-activity -OutputDirectory reports/runtime/slice4be-staging-production-cooperative-repeat -CallbackFilter Production -WorkbookState ProductionReusable
```

The respective ignored reports are `production-reusable-production.md`,
`packaged-launcher-noeligible.md` and `production-reusable-production.md`.
The pre-change comparison and its five-package hash record remain under
`reports/runtime/slice4be-staging-production-baseline`.

The packaged smoke verifier left its own disconnected automation process after
81/81; it was identified from the isolated run and stopped before the next gate.
Its report heading also hard-coded deploy/current despite the actual candidate
invocation. The report-only correction now uses the validated path through the
existing redactor; parsing, diff review and a direct redactor check passed. This
metadata correction changes no runtime contract and needs no D13 behavioral RED.

## Candidate hashes

| Package | SHA-256 |
|---|---|
| invSys.Core.xlam | 6e827aa70e353e6ea6fc55d6675a592aa6c4ecd60132ca50ac55a0ff962ebf36 |
| invSys.Inventory.Domain.xlam | 6632ad147ade59d4f97a53eb6640474d0fe407cadffe0e591a994191ffc01481 |
| invSys.Designs.Domain.xlam | 66815dfb1cd970e7761235f09331b2501ccf0243fe7842428aa350efc7d3d8c0 |
| invSys.Operations.xlam | 4bc12d48bbbb3d7359efd63d4af23ea330045f657ccfff7d7f9c7161517a8f0b |
| invSys.Admin.xlam | d9ab9e28026452178ef2958d654a22f82f97e1e3d7f8ab7ed6b6d541f442a787 |
