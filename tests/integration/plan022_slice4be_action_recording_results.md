# Slice 4be.4 recording lifecycle

Last verified 2026-09-14. Architecture v4.11 D18 and Plan022 govern the explicit
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

## Saved-run reader: test-first entry

The frozen `17d788c` recording candidate runs `-Phase RED -CheckRecordingReader`:
**71 PASS / 18 FAIL**, **89** unique identities, retaining all **68** applicable
foundation checks GREEN and no harness exception. The fixture records two real
Admin Save Value actions through Viewer Start/Stop, verifies their distinct
ordinals and six-entry hash chain, then uses the actual Action Paths button and
library controls. Missing library controls are observed product failures.

Reader failures cover open/reuse/selection, ordered original actions, read-only
evidence, Search, stopped-without-conclusion status, missing Close, a missing
middle entry, damaged hash, a rehashed broken previous link, omitted closing
observations, repeated ordinals, restoration, current visibility and sign-out.
The three new passes establish absence of publication/Shipping authority calls,
training-byte preservation and no library left after Viewer close; they do not
prove a working reader. Fault injection changes only the generated run and
restores every original journal byte in finally blocks.

The first `deploy/validation-recording-reader` candidate builds and explicitly
compiles all five packages with Operations cold-start references. Core owns
journal enumeration/validation and permitted primitive projections; Operations
owns the Events-only library launcher and captured read form. Focused GREEN is
**89/89**, retaining every RED identity, with normal Excel closure and all ten
foundation/reader package pins unchanged. No Excel Application Error1000 appears
in this focused run's bounded observation window; earlier native faults remain
unresolved. All 19 changed PowerShell files parse and all 16 import harnesses
contain the two new Core modules exactly once.

Static maintenance records218 components/5777 procedures,1150 scanner/1152
reviewed candidates,45 unresolved/9 literal Application.Run calls. All28
existing oversized modules pass individual comparisons and the three new
modules remain below1000 lines. Duplicate groups grow189->193: one copied
MakeControl builder and three short event-handler bodies. Nine further
candidates are required form/callback dynamic roots. The duplicate findings
remain unresolved; no maintenance exception or completed maintenance gate is
claimed. Resolve the builder duplication and review the callback bodies before
accepting the reader checkpoint.

Still required for this reader: its own full layout/native geometry and visible
inspection, additional warehouse/package/identity corruption and stale-context
cases, current-policy failure/older-catalog cases, actual Excel interruption and
restart, plus relevant broader regressions on the reader candidate. Foundation
release results above apply to the preceding candidate. All-role recording,
source-result correlation, conclusions and How-To/Diagnostic/Compare both remain
required. This is an initial focused GREEN checkpoint, not a completed reader
or Slice4be/Release1 acceptance claim.

Reader evidence (ignored): `recording-reader-red.log`,
`recording-reader-red-summary.json`, `recording-reader-build.log`,
`recording-reader-compile.log`, `recording-reader-compiled.json`,
`recording-reader-package-pins.json`, `recording-reader-green.log` and
`recording-reader-static.log` and `recording-reader-checkpoint-summary.json`,
all under `reports/runtime/`.

## Reader integrity and maintenance follow-up

The final owner-provenance candidate explicitly compiles all five packages and
passes **117/117** focused checks, retaining all77 foundation and89 initial-reader
identities. All104 expanded identities are GREEN with the corrected expectations
documented below. This includes four library geometry sizes, actual Close/reopen,
foreign-warehouse/schema rejection, owning result-build mismatch, known older
catalogs, opaque build/package warnings,256 actions and1MiB boundaries.
Filters/layout pass **59/59**, retaining every prior identity.

Static validation confirms218 components/5776 procedures,1149 scanner/1151
reviewed candidates,192 duplicates (only the three explicit exceptions below),
45 unresolved/9 literal calls and all28 individual oversized-module limits.
The copied builder group is removed. Publication's first attempt stops0/1 at
PublishProjectionFixtureForTest with RPC unavailable, before product assertions.
The exact recovery child is verified empty and closed; all three loaded add-in
files retain their hashes. The unchanged candidate's publication retry passes
**82/82** and Event Detail passes **34/34**. Full-chain validation passes **31/31**,
including live roles **48/48** and Create Warehouse **15/15**, with five candidate
hashes unchanged, Excel closed and the three original generated reports restored.
The chain observation window records one combase.dll/c0000005 Excel fault.
Both native failures remain unresolved despite the passing assertions.

Final preservation checks verify190 earlier package pins plus65 recording,
reader, paging and related candidate pins, including the rejected candidate.
Fifteen protected source files remain exact; the one previously reviewed
Shipping visibility-only change retains its documented scope. All19 reader
procedures fit the200-line limit (maximum69). Both changed PowerShell scripts
parse and repository whitespace checks pass. No accepted deployment is replaced.

Reproduce the focused candidate gate:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-recording-owner-provenance -Phase GREEN -CheckRecordingReader -CheckRecordingLimits -CheckRecordingStorageBounds
```

**Owner-build correction:** The first combined candidate fails58/15 and never
reaches the new reader fixture. Fourteen failures lose accepted recording checks;
one is the consequent missing-journal fixture exception. Source inspection
locates the error in the newly added model comparison: Activity.MakeBody reads
Admin/Operations BuildIdentity, while the journal header reads Core. D18 original
owner provenance requires those distinct identities. This candidate is not
GREEN and supplies no new acceptance. Excel closes normally.

Remove only the incorrect build-equality comparison; keep common package-set
and catalog validation. The stable ObservationBuildMismatch check now changes
only outcomes, protecting the existing owning attempt/result equality instead
of imposing equality with Core. The earlier99/5 run therefore supplies four
valid RED failures (two common-provenance checks and two truthful difference
warnings), not five. Its proposed Core/role-build failure is invalid and is
explicitly excluded. The corrected build test is a regression for existing
ReadRun behavior. Original observations and package build IDs are never rewritten.

The expanded packaged test covers common package-set/catalog provenance,
owning attempt/result builds, foreign warehouse/schema and release warnings;
four library bounds/non-overlap sizes; and changed-target recovery. Rehashed
faults retain a calibrated six-entry hash chain, so a hash failure cannot stand
in for semantic rejection. The frozen reader candidate is the RED target.
These cases enforce the existing D18 identity, provenance and context contract.

The first expanded run completes99/5 across104 identities. Its two common
package-set/catalog failures are valid RED; the proposed Core/role build equality
is excluded as explained above. Two warning failures were based
on an invalid age inference: the build tool generates random GUID identities
and writes PackageSetVersion as the compatibility label R1-5. Neither orders
releases. Those two original failures are not claimed as meaningful RED for
chronological age. Architecture, Plan022 and controls now clarify the inherited
truthful-provenance rule: a lower supported catalog can say Older release;
otherwise differing package/build identity says Different release/build with
relative age unavailable.

The corrected warning run is again99/5; its initially claimed five valid failures
are qualified by the owner-build correction above. Every89 prior reader GREEN
identity is retained. Historical check IDs
OlderPackage/OlderBuild remain stable, but their expectations now require the
truthful difference/unknown-age label; their original-evidence checks also remain.
The final model rejects package-set/catalog disagreement between a journal and
its observations, preserving per-owner build identity. The reader distinguishes
known older catalogs from opaque
identity differences. The final focused GREEN and static results above supersede
the failed candidate. Added lower-catalog
and actual Close/reopen checks are supplemental regressions, not new behavioral
RED claims.

The duplicated MakeControl routine is replaced by one declarative control table
inside library initialization, retaining the same controls, geometry and event
bindings. This is refactoring under the existing89/89 GREEN; the expanded
geometry and action checks protect the resulting candidate.

**Explicit maintenance exception RDR-UI-THUNKS-01:** retain only the following
three LOW-confidence duplicate groups, each a single-statement event body:

- `duplicate:862c21be3babc752:mBtnClose_Click+mClose_Click` (Me.Hide).
- `duplicate:9546f46f4520c584:UserForm_Layout+UserForm_Layout` (guarded layout).
- `duplicate:c6fd55022a9ee927:mRefresh_Click+mSearch_Change` (RefreshPaths).

These are distinct required MSForms event entrypoints; the action/layout logic
already has one shared implementation. Removing or indirectly rerouting events
solely to change a scanner count would obscure actual operator entry. This
bounded exception follows AGENTS.md's explicit-exception rule and explains the
three groups above the preceding189 baseline; the expected ceiling is192 after
builder cleanup. It grants no general duplication, module-growth, dynamic-call
or architectural exception. Other new groups remain failures. Packaged action
and geometry checks, compile and individual module limits still apply. Raw
generated scanner findings remain visible and are not rewritten as deletions.

Ignored follow-up evidence under reports/runtime: recording-reader-integrity-red.log,
recording-reader-provenance-red.log, recording-provenance-red-summary.json,
recording-provenance-build.log, recording-provenance-compile.log,
recording-provenance-compiled.json, recording-provenance-package-pins.json,
recording-provenance-focused.log and recording-provenance-static.log.
Final candidate evidence uses the recording-owner-provenance prefix for build,
compile, compiled source, package pins, focused, filters, publication/retry,
detail and chain reports. recording-owner-provenance-prior-checks.json and
recording-owner-provenance-maintenance-checks.json retain exact comparisons;
recording-owner-publication-recovery-closure.json retains guarded recovery proof.

## Foundation raw evidence

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
