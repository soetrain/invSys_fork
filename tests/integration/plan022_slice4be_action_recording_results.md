# Slice 4be.4 recording lifecycle

Last verified 2026-09-13. Architecture v4.11 D18 and Plan022 govern the explicit
actor/warehouse sequence, immutable observations and advisory conclusions.
Slice4be and full Release1 acceptance remain open.

## Test-first evidence

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-events-maintenance -Phase RED -CheckActionRecording
```

The frozen five-package candidate from source checkpoint `1d6a2ad` records
**34 PASS / 16 FAIL**, 50 unique check identities, exit1 as expected for RED.
All 30 existing published-reader checks retain their exact identities and pass;
comparison uses the preceding `events-maintenance-filters.log` result. No
`Harness.Exception` occurs.

The extension opens Viewer through `modInventoryViewer.OpenInventoryViewer`,
enters Events through its existing handler, locates recording buttons by their
approved captions and delivers CommandButton.Value to the actual event binding.
Missing product controls return MISSING; a missing Viewer/test seam raises a
harness error. The status probe reads `lblRecordingStatus` without refreshing or
creating a record. Instrumentation is installed only in unsaved test projects.

The actual Admin capture checkbox and Save Tracking Policy handlers configure
the disposable warehouse. Repeated Save Value actions enter
`frmAdminSettings.mBtnSaveConfig_Click`, changing synthetic BatchSize values.
Each action must produce exactly one REQUESTED and one COMPLETED record with
the same ActivityId and registered ADMIN_SETTINGS_SAVE_VALUE identity before
its sequence assertions run. Missing ordinary activity is a fixture failure,
not recording RED. No synthetic owner outcome or sequence is supplied.

| Observation | Result | Meaning |
|---|---|---|
| Prior published-reader checks | 30/30 PASS | Existing public Viewer behavior retained within this test's scope. |
| Recording buttons, disabled explanation and disabled Start | 5 FAIL | The approved surface is absent. |
| Start/counter, first/second ordinals and distinct repeated occurrence | 5 FAIL | Valid ordinary actions have blank SequenceId and Ordinal0. |
| Stop control and stopped-without-conclusion status | 2 FAIL | No lifecycle transition exists. |
| Second Start, new sequence, Cancel and cancelled status | 4 FAIL | No new-run/cancel lifecycle exists. |
| Ordinary activity before/after attempted recording and earlier byte retention | 4 PASS | Existing behavior is preserved; these passes do not prove Stop/Cancel worked. |

## Original test-only checkpoint: preservation and limits

No VBA, form, package, schema implementation, accepted deployment or NAS workbook
changes are made. Admin-generated disposable fixture configuration changes are
intentional. The harness removes those fixtures and restores its saved personal
settings after the run. All five frozen candidate package hashes remain equal
to their pre-run pins. Excel is closed. No Excel Application Error1000 appears
between creation of this run's log and the post-run observation; this does not
resolve the previously recorded intermittent native faults.

The architecture, Plan022 and controls name the already approved recording
surface consistently. This refinement changes no business authority, capability,
activity catalog identity, collection rule or accepted behavior. No new build,
compile, static-maintenance, role/full-chain or visible acceptance is claimed for
this test-only checkpoint. The candidate's preceding gates retain their recorded
scope in the Events filter evidence.

This first lifecycle RED does not yet protect saved Action Path schema/hash,
atomic incremental persistence, 1MiB overflow, 256-action closure, tracking
failure, mid-sequence policy/context invalidation, restart, all Operations roles,
exact multi-event submissions or evaluation. Add those focused cases before
their implementation. Then implement actual Viewer handlers and headless Core
recording, preserve these 50 checks, and complete all D18/D13 gates. How-To,
Diagnostic, Compare both, library/version/import/export and human acceptance
remain required; Stop alone must never stand in for a conclusion.

## Recorded-run foundation in validation

The next frozen-candidate run extends the original 50 identities to **64**:
**35 PASS / 29 FAIL**, with no harness exception. It adds durable Start before
any action, distinct run/record identities, incremental hash-linked observations,
complete immutable closing evidence, policy/Viewer-close/sign-out interruption,
storage obstruction and ordinary Viewer-user eligibility. The further action
boundary run completes 256 real Admin Save Value handlers and records **36 PASS
/ 32 FAIL**, **68** identities. Action257 still completes as ordinary activity;
the missing sequence, partial closing record and visible counter are RED.

The new Core recorder exposes only primitive Control/Status/ContextClosed
boundaries. Internal services own the captured target, activity ordinals and
Start/Observation/Close journal. The target is copied at Start; no role workbook
or form is shared. Activity bodies retain their owning identities, captions,
outcomes and exact source references. The journal stores original decoded bodies
with hash-linked immutable versions and generated filenames; its schema does
not accept paths, input values or arbitrary added fields. Core remains headless.

Operations owns three actual Viewer buttons through one typed event binding and
a status/counter. The Events recording row leaves Inventory list space intact.
Start succeeds only after durable persistence. Stop never asserts a conclusion;
Cancel preserves work. Known Viewer close, authentication invalidation and policy
update boundaries close incomplete evidence. A failed journal write leaves
ordinary owner work available. Policy validation exposes its existing capture
and eligibility flags directly, avoiding a full editor projection on every action.
The 16 explicit Core source-import harnesses include the four recorder modules.

The first build stops explicit compile at a procedure/argument name collision in
the new recorder; this is an implementation failure, not behavioral RED. After
the naming correction, `deploy/validation-recording-compiled` passes five-package
explicit compile, Operations cold-start references and **64/64** focused checks.
All five package hashes remain unchanged and Excel closes. This establishes the
tested foundation, not full recording or Release1 acceptance.

Review identifies one new duplicate pair, ContextClosed/PolicyChanging. Both
typed callers and the guard are inspected and consolidated into one internal
context-scoped interruption procedure. No accepted dead code is removed. The
cleaned source candidate `deploy/validation-recording-limits` again compiles all
five packages with cold-start reference validation. Its extended run passes
**77/77**: all prior 50/64/68 check identities remain GREEN, plus four recording
geometry checks and five supplemental storage-bound checks. The latter verify
an exact 1 MiB record, integrity readback, idempotent bytes, explicit one-byte
overflow rejection and no partial publication. These five supplemental checks
are not claimed as separate pre-implementation behavioral RED.

The same candidate passes full chain **31/31**, live-role **48/48**, Create
Warehouse **15/15**, Viewer filters/layout **59/59** and Settings **187/187**.
All 59 preceding filter identities remain GREEN. Full-chain cleanup preserves
the five package hashes and restores the three original generated reports.
An Excel combase/c0000005 fault occurs during that passing full-chain run;
clean native reliability remains unproven. A read-only desktop probe finds no
foreground window or readable input desktop, so visible acceptance is open.

Publication attempt one stops before assertions with
an RPC failure in PublishProjectionFixtureForTest; attempt two passes 63 checks
then stops at PublishViewerGroupsForTest during the locked-destination case.
These are harness/native failures, not meaningful behavioral RED. The serial
comparison then passes **82/82** on the preserved maintenance candidate and
**82/82** on the unchanged recording candidate, each exiting zero with Excel
closed. Exact identity comparison retains every preceding publication check.
No publication implementation was changed to obtain that result; the two prior
failures remain evidence of unresolved reliability, not a proven product cause.
Both failed runs' recovery children were identified before guarded cleanup;
Excel closure was subsequently verified. The second cleanup did not reach its
post-close installed-file hash comparison, so that byte-preservation claim is
not made. No arbitrary Excel process was terminated.

Final preservation verifies 190 historical/publication package pins and 50
additional reader/paging/recording candidate pins, 15 unchanged protected source
checks and the one previously reviewed Shipping visibility-only difference.
Excel is closed. No Application Error1000 Excel fault is observed from the start
of the Settings run through the two completed publication comparison runs;
this bounded observation does not resolve earlier native failures.
All 19 changed PowerShell files parse, all 16 Core import lists contain each
required recorder module once, 42 maintained Markdown links resolve, and code
and applicable documentation diffs pass whitespace checks.

The final candidate's focused and publication commands are:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-recording-limits -Phase GREEN -CheckActionRecording -CheckRecordingLimits -CheckRecordingStorageBounds
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-recording-limits -Phase GREEN -CheckViewerPublication
```

Static maintenance returns to **189** duplicate groups, **45** unresolved and
**9** literal Application.Run calls. All **28** existing oversized components
are compared individually to the committed baseline and none grows. All **6**
new modules and **28** new procedures meet 1000/200-line limits. Totals are
215 components/5755 procedures, 1137 scanner/1139 reviewed candidates. The sole
new candidate identity is the required cRecordingButton.mButton_Click dynamic
root; it is not new unreachable code or a maintenance exception.

Still required: full-journal current-policy reader and missing-Close/Excel
restart proof, corruption and missing-link handling through the eventual reader,
recording across real Operations tasks and exact multi-event submissions,
late-result/session isolation, failure/retry evaluation, How-To/Diagnostic/Compare
both, library/version/import/export, broader regressions and visible/human
acceptance. The existing record reader validates individual entries; do not
claim a completed full-chain journal reader or completed diagnostic evidence.

## Ignored raw evidence

- `reports/runtime/action-recording-red.log`
- `reports/runtime/slice4be-viewer-published-read/d3e82e13f0f04c5aa9a6f0881f0775a1/red.json`
- `reports/runtime/events-maintenance-package-pins.json`
- `reports/runtime/events-maintenance-filters.log`
- `reports/runtime/action-recording-journal-red.log`
- `reports/runtime/action-recording-limits-red.log`
- `reports/runtime/recording-compiled-focused.log`
- `reports/runtime/recording-compiled-source.json`
- `reports/runtime/recording-compiled-package-pins.json`
- `reports/runtime/recording-limits-compiled.json`
- `reports/runtime/recording-limits-module-ratchets.json`
- `reports/runtime/recording-limits-size-checks.json`
- `reports/runtime/recording-limits-focused.log`
- `reports/runtime/recording-prior-check-retention.json`
- `reports/runtime/recording-limits-chain-summary.json`
- `reports/runtime/recording-regression-filters.log`
- `reports/runtime/recording-regression-settings.log`
- `reports/runtime/recording-regression-publication.log`
- `reports/runtime/recording-regression-publication-retry.log`
- `reports/runtime/recording-publication-baseline-comparison.log`
- `reports/runtime/recording-publication-controlled-comparison.log`
- `reports/runtime/recording-final-preservation.json`

The maintained report contains only check identities, aggregate results and
technical scope. Runtime fixture identities and row values are not copied here.
