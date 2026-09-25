# Production lifecycle observation checkpoint

Last verified: 2026-09-25 UTC. Slice4be.1 under Architecture v4.11 D18,
Plan022 and the maintained controls catalog. **Partial acceptance; not Slice4be completion.**

## Contract and implementation

The normative specification, plan and control catalog were updated before runtime
edits. Catalog13 adds the six Process/Recipe Save Draft, Release and Obsolete
controls; versions1-12 retain their definitions. Actual click handlers reuse the
captured-context and permission boundary. A per-form guard prevents overlapping
designer commands; context is rechecked after confirmation/yielding and before
refresh. Shared submissions and imports do not create additional click records.

Operations-local typed facts retain the exact generated event ID and whether the
owning writer reached a write attempt. Core exposes that existing fact through
an optional primitive output; business schemas are unchanged. Source references
are bound to the captured warehouse and registered Designs source. CONFIRMED
means the command finished, with Unknown Domain effect. No report parsing,
processor count or current definition status establishes event application.
The separate Designs Action Path evaluation contract is **not implemented here**.

## Packaged D13 evidence

Frozen predecessor: `deploy/validation-production-header`. Candidate:
`deploy/validation-production-lifecycle`. Tests install unsaved adapters and invoke
the six actual packaged form handlers; source/reference assertions supplement them.

- Initial observation RED:60 PASS/105 expected FAIL. Expanded pre-write RED:
  76 PASS/113 expected FAIL. Owning actions, exact source events and unknown user
  columns passed before implementation.
- Final valid RED: **139 PASS/155 expected FAIL**, all294 check identities,
  UTC22:47:51--22:50:21. Report:
  `reports/runtime/slice4be-production-designer/d1af022ef12b471394502476d73e6182/red.json`.
- GREEN: **294/294**, UTC22:51:51--22:54:32. Report:
  `reports/runtime/slice4be-production-designer/dc3ed0ec5e7c4bbba644d5373a6188dc/green.json`.
  Each successful action produces distinct correlated REQUESTED/CONFIRMED records
  with its exact owning Designs ID. Entered data/credentials are absent, unknown
  operator columns survive, all68 earlier definitions remain, invalid reference
  envelopes are rejected, denial/validation preserve drafts and submit no event,
  and stale target/session actions stop without redirected activity.

Both final runs close Excel normally and restore settings without changing the
five frozen packages. Controllers:
`production-lifecycle-controller/81fc8407c7764f9b947d38b7f9795c34` and
`production-lifecycle-controller/3c8cf5d04d094d719d017491b9433c95` under reports/runtime.

Earlier invalid attempts remain recorded: the first probe named a nonexistent
Recipe query and failed instrumented compile; it was corrected to GetRecipeGraph.
A later reference-case array expression was corrected before the final RED.
That intermediate run also encountered VBA80010007 in the closed-workbook fixture;
verified disposable Excel termination allowed the original controller to restore
settings. Its attempted native dismissal was not verified. These are harness
failures, not meaningful RED or normal-shutdown evidence. Closed-workbook coverage
requires a corrected independent test; it is not silently treated as passing.

## Subsequent packaged failure and safety tests

This extension changes test tooling only, under the existing D18 lifecycle contract.
Runtime source and both frozen package sets remain unchanged. Unsaved instrumentation
injects a failure at one real boundary; it never replaces a form handler, writer,
processor or activity record. Each failed queue is isolated beneath the disposable
fixture so a later case cannot silently process it.

All six handlers now exercise four owner-failure boundaries: before the write
attempt (FAILED, no unused source ID), after the actual append but before acknowledgment
(FAILED, exact Unknown reference), after acknowledged submission but before processing
(PENDING, exact Submitted reference), and after actual processing but before refresh
(PENDING, exact Submitted reference). The last case deliberately proves that an applied
source does not imply that the user command finished. Activity remains Unknown effect,
with no entered values, paths or raw injected error text.

The failure extension records **RED211 PASS/275 expected FAIL -> GREEN486/486**,
retaining all294 preceding identities and adding192 assertions for24 failure cases.
Both runs have five instrumented compiles, normal unassisted closure, restored
settings, unchanged packages and zero Excel Application failures. Private receipt:
`reports/runtime/production-lifecycle-fault-verification.json`.

The further safety extension reaches **RED316 PASS/299 expected FAIL -> GREEN615/615**,
retaining all486 preceding identities. Both final runs have no harness failure,
five instrumented compiles, normal delayed unassisted closure, restored settings,
unchanged package/test bytes and zero Excel Application failures. It covers
tracking-off, unavailable activity storage, nested actual-handler
invocation and closed captured workbooks. Original tracking policy is restored;
prior activity bytes and unknown workbook columns remain protected. Re-entrancy is
observed at the real writer, so duplicate submissions are explicit failures instead
of requiring a readable publication after an already-failed nested command.

Final RED: UTC00:01:28--00:06:20 September25,
`slice4be-production-designer/c4a8ba88cf1f4b89b9e4840221ea1371/red.json`.
Final GREEN: UTC00:06:37--00:12:12,
`slice4be-production-designer/3adbf45b4da744c6bd4bbfbc22b10f4d/green.json`.
Controllers are `production-lifecycle-controller/c0f1a59131c4452cb8210183f718314f`
and `production-lifecycle-controller/dbc26d6b84a1411fbaeaf82129959f12` under
reports/runtime. Receipt: `production-lifecycle-safety-verification.json`.
The615 checks do not establish either new lifecycle Action Path method or native
confirmation/cancellation. These remain separate acceptance requirements.

Three intermediate attempts are retained as invalid/incomplete evidence: the first
queries publication after the predecessor's duplicate action and encounters an
unavailable source; the next hashes an open Excel file with incompatible sharing;
the third stops in VBA80010007 and requires verified disposable Excel termination.
The original controller restores settings in all three. The corrected closed-book
fixture pins saved bytes before reopening and activates a surviving decoy window
before constructing the form. Only the captured workbook closes; the form's host
window survives, and the decoy must never replace the binding. The final predecessor
reaches all six closed-book actions without that dialog. These fixture corrections
do not establish a broader cause for Excel/RPC or desktop failures.

Final static evidence regenerates after the fixture corrections at
`reports/runtime/production-lifecycle-safety-final-static`:
all six metrics remain exactly unchanged, three report schemas pass, and282 tooling
PowerShell files parse. All28 module caps hold with the **existing** Plan022 +45-line
draft-control exception for frmProduction (11700 +45 ceiling; current11743). No new
exception or package rebuild is introduced. Earlier build/layout/regression records
retain their exact frozen-package scope; failed chain/reusable gates stay open.

## Maintenance and prior package evidence

All five packages build and compile, including cold-start Operations dependency
resolution. Layout gates pass8/8 and7/7. The first layout invocation omitted the
required explicit RepoRoot and failed in its default parameter; the corrected
invocation passes, with no layout/runtime edit.

Final static reports: `reports/runtime/production-lifecycle-final-static`. Counts are
253 components/6058 procedures/133154 lines, nine literal/45 unresolved dynamic
calls and192 normalized duplicate groups. All28 oversized limits hold. Plan022
records the single reviewed literal-only ControlIds duplicate exception; no
scanner-reported code was deleted. Current-emission test expectations advance
12->13 while historical catalog assertions remain unchanged.
All three report schemas and280 PowerShell parses pass. Verification receipts:
`production-lifecycle-focused-verification.json`,
`production-lifecycle-final-static-verification.json` and
`production-lifecycle-passing-regressions-verification.json` under reports/runtime.

The supplemental bulk-import source probe initially reports7 PASS/one obsolete
helper-name assertion FAIL. Aligning that assertion with the shared
SubmitDesignerAction(True, PROCESS_SAVE) boundary restores8/8; no additional
runtime change. The legacy Slice4x source probe reports8/10: its old caption and
Viewer-module text predicates also fail on the frozen HEAD source. Its private
result is retained and the tracked historical report restored. Neither source
probe replaces actual packaged workflow evidence.

## Broader regression scope and failures retained

Current-candidate smoke passes86/86, UTC23:07:31--23:07:51, restoring settings and
its tracked report with five package hashes preserved. Its existing helper may
terminate Excel, so this is not unassisted-closure evidence. Headless Settings
passes202/202, UTC23:08:36--23:12:58, restoring settings and package hashes with
normal closure. This includes current catalog/policy, detail-profile, personal
preference, restart and Operations-without-Admin routes; no captures were requested.

The earlier six draft controls and their original recording/publication/diagnostic
routes retain **390/390**, UTC23:15:26--23:25:12, with exact prior check identities.
This includes their closed-workbook guards, real released-Process Recipe validation,
immutable observations/journal, Viewer selection and negative terminal conditions.
It does not cover the six new lifecycle controls' Action Paths. A residual hidden
Excel process was observed at23:25:07, then exited normally before the planned
recovery command found it. No termination or COM reattachment occurred. The
provisional commentary calling this a shutdown failure is superseded by that
terminal evidence. Settings/pins restore and the Application audit is clear.
Receipt: `reports/runtime/production-lifecycle-designer-verification.json`.

Two current-candidate chain attempts and one frozen-predecessor comparison all
stop at **Delete and rebuild canonical inventory projections**:
`modProcessor.RunBatchReportForAutomation`, HRESULT800706BE. Each records
chain5 PASS/one harness FAIL, live-role32 PASS/one harness FAIL and Create
Warehouse15/15. Disposable recovery Excel processes required termination; the
original controllers then restored settings and three tracked reports. Candidate,
predecessor and test-source pins were verified at that comparison checkpoint.
The identical predecessor failure narrows attribution but does not establish a
cause or accept the candidate. Receipt:
`reports/runtime/production-lifecycle-chain-failures-verification.json`.

Reusable Production also stops before its aggregate result, at
`mProduction.RunProductionBatchScaleContractTest`, HRESULT800706BE,
UTC23:05:44--23:06:11. Settings restore and Excel closes under the existing
potentially assisted cleanup helper. Its67 prior Boolean observations are not
re-established. Do not substitute focused294/294 or smoke86/86 for either gate.

Read-only cursor probes, including23:42:36 UTC September24, still return error5;
no native capture is accepted for this candidate. The user suggests Group Policy
as a possibility and reports no changes there. A scoped read-only registry check
finds machine inactivity timeout0 and no configured usual user screensaver-policy
or machine/user RDP timeout-policy values. This does not establish the full effective
policy or whether either desktop is unlocked. No Windows policy or permission setting
was changed; the cause remains unproven. At00:09:10 UTC September25, the current
session reports Active while UOI_IO for the automation desktop is False. The
query and read-only input-desktop open succeed; both observed desktop names are
Default. This confirms a current input-availability problem without proving a
lock, policy cause or need for elevation. Cursor access still returns error5 at
00:19:18 UTC. The user was asked to confirm the host
desktop is visible inside unminimized RDP with the client unlocked.

Pending: native cancellation; original recordings,
published Events and both How-To/Diagnostic methods; complete Designs applied,
awaiting and incomplete evidence; visible operator acceptance. The failed regression
and chain gates above remain open. The
[remaining acceptance checklist](plan022_slice4be_remaining_acceptance.md) continues
to govern the full release outcome.
