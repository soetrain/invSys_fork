# Production lifecycle observation checkpoint

Last verified: 2026-09-24 UTC. Slice4be.1 under Architecture v4.11 D18,
Plan022 and controls1.249. **Partial implementation; not Slice4be acceptance.**

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

## Maintenance and remaining acceptance

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

The read-only desktop probe at23:07:30 again returns cursor error5; no native
capture is accepted for this candidate. Group Policy and the earlier desktop
failure's cause remain unproven. The user has been asked whether both desktops
are unlocked. No policy or permission setting was changed.

Pending: native cancellation; pending/uncertain/error owner routes; tracking-off,
tracking failure and re-entrancy; closed captured workbook; original recordings,
published Events and both How-To/Diagnostic methods; complete Designs applied,
awaiting and incomplete evidence; visible operator acceptance. The failed regression
and chain gates above remain open. The
[remaining acceptance checklist](plan022_slice4be_remaining_acceptance.md) continues
to govern the full release outcome.
