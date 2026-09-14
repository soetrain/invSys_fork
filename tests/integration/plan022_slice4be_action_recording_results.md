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

## Late-result sequence isolation test entry

Under D18's immutable selected-run and re-entrancy boundaries, an old action's
result may not terminate a newer recording. The new packaged isolation test
creates both runs through actual Viewer Start/Stop controls and Admin Save Value.
It replays only the old handler's recorded COMPLETED outcome, first under the
same policy and then after actual Settings policy saves. It does not fabricate
a business result or replay a task. Same-policy replay must remain idempotent;
an old-policy result must be rejected without changing either journal, current
recording status, activity bytes or configuration. The new run must still stop
normally through its actual control.

The current FinishAction calls global Interrupt on policy mismatch, so the
expected behavioral RED is closure of the new run when the old result is
rejected. The isolated frozen target is the five-package owner-provenance
candidate. All117 prior focused identities remain required. No runtime change
or isolation GREEN is claimed at this checkpoint.

The frozen candidate completes **127 PASS / 3 FAIL**,130 unique checks, retaining
all117 prior GREEN identities and no harness failure. Same-policy replay passes
all six controls; changed-policy rejection preserves activity/config bytes but
fails ReplayPreservesBothJournals, CurrentRunRemainsActive and
CurrentRunStopsNormally. Those are meaningful behavioral RED: the stale result's
global policy interruption appends an Incomplete Close to the unrelated new run.
Five candidate package hashes remain unchanged and Excel closes normally. No
Excel Application Error1000 is observed during this test window; the preceding
native faults remain unresolved. Foreground/input-desktop probes are unavailable.

Next implement an internal owning-sequence guard for FinishAction interruption,
then rerun the exact130-check gate and applicable packaged/regression/maintenance
gates. Preserve policy rejection and same-policy idempotence; do not suppress
current-run tracking failures globally. Actual cold interruption/restart,
Operations sequences, conclusions and comparison remain required.
Ignored evidence: `recording-isolation-red.log`, `recording-isolation-red-exit.json`,
`recording-isolation-red-summary.json` and `recording-isolation-desktop.json` under
`reports/runtime/`.

The expanded pre-implementation matrix adds stale append-failure and result-
exception cases, plus the corresponding owning-sequence positive controls.
A conflicting generated activity file exercises Append=False; a test-only
exception at the resolved owning-action boundary exercises the actual
FinishAction error handler. All outcomes come from real Admin Save Value actions.
Restore exact fixture bytes and clear the exception selector; journal
observations and business outcomes are never fabricated. An owning
failure must still append an Incomplete Close, so suppressing all interruption
cannot satisfy the test. The first extension stops75PASS/1harness failure because
reusing the Settings value produces a legitimate UNCHANGED outcome, unsuitable
for this COMPLETED fixture. That run supplies no new RED. The corrected matrix
uses distinct values and runs before any runtime change. That corrected-value
run completes94/11: six valid failures reproduce stale policy/append interruption,
but five assertions belong to a CVErr-based exception fixture that did not reach
the intended error branch. Those five failures are excluded from behavioral RED.
Replace coercion with a calibrated Err.Raise after action/context resolution in
the disposable unsaved Core project; require the actual sanitized result-error
response before counting exception isolation. The package files remain unchanged.

The first explicit-error probe stops68/1 because its case-sensitive marker does
not match VBE's `Action` identifier spelling; no exception case ran. A read-only
package probe verifies exactly one case-insensitive resolved-action marker.
The calibrated matrix then completes **96PASS/9FAIL** across105 unique checks,
without a harness failure. All nine failures are the three stale-result symptoms
under policy, append-conflict and exception cases. Both active-sequence failure
controls pass every check, including immutable prior entries and Incomplete Close.

The implementation adds internal `modRecordingSession.InterruptAction`, checking
the action's captured Context and exact SequenceId before delegating interruption.
FinishAction's three failure paths use it. Global authentication/context/policy
boundaries retain their existing behavior. This enforces D18 ownership without
adding UI, write authority or a cross-package contract. The new isolated
`deploy/validation-recording-isolation` candidate builds and explicitly compiles
all five packages, including Operations cold-start dependency validation. Compile
initially deferred while build Excel exited; a subsequent closed-process check
allowed compile of the same candidate, without rebuilding.

Static generation reports218 components/5777 procedures,1149 unchanged scanner
candidate identities and1151 reviewed candidates,192 duplicate groups,45 unresolved/
9 literal calls. Every28 oversized source-file limit is preserved; the sole added
procedure is6 lines. Existing RDR-UI-THUNKS-01 scope is unchanged. The expanded154-
check focused gate and full chain remain required.

The first combined candidate run reaches109PASS/1harness failure. All37 isolation
checks pass, including the nine corrected stale-result assertions and owning-
failure controls; reader entry then loses its packaged RPC connection
(0x800706BE). Reader/storage gates and the chain do not complete. This is partial
GREEN evidence, not a passing154-check gate. Its exact owned child has zero
workbooks. Normal Quit is requested; the helper refuses forced termination while
a native window remains, and a subsequent process check verifies normal exit.
No forced stop occurs and all five candidate hashes are preserved. The helper's
guarded exit1 is retained; do not claim its post-Quit project-hash loop completed.
No Excel Application Error1000 is observed at the subsequent query, which does
not resolve the RPC/recovery failure. An unchanged candidate/test retry is running
before the full chain; native-clean and visible acceptance remain unproven.

The unchanged retry completes **154/154 GREEN**, exit0, preserving every check
identity from the prior117-check reader baseline,130-check original isolation
run and105-check calibrated matrix. The compiled-source comparison covers211
components in each candidate; only modActivity and modRecordingSession differ.
This proves the corrected interruption scope while retaining owning failures,
idempotence, journal/policy/context reads, layouts and256-action/1MiB boundaries.
The same candidate completes the full Release1 chain **31/31**, live-role
**48/48**, and Create Warehouse **15/15**, exit0. All five candidate hashes are
unchanged and the three tracked fixture reports are restored exactly. No Excel
Application Error1000 was observed during that chain window. The first combined
RPC failure is not erased by the successful retry.

The chain leaves one residual Excel instance. Read-only inspection identified
zero workbooks and three loaded add-in projects; normal Quit was requested and
all three add-in files remained unchanged afterward. The process persists with
a visible native dialog whose text is unavailable through automation. No forced
termination occurred. Its relationship to registration restoration is suspected,
not proven. Null COM properties during shutdown are not evidence of empty state.
Excel closure, clean native reliability and visible human acceptance remain open;
no further build or Excel job is started while the residual remains.

A read-only hash snapshot verifies **260 package pins**,15 protected source files
unchanged and one previously reviewed Shipping visibility-only difference. This
snapshot records ExcelClosed=False and does not claim clean shutdown. The initial
additional-pin script incorrectly wrapped ConvertFrom-Json arrays; correcting
that PowerShell enumeration made the70 additional pin checks pass without changing
packages. This was a verifier setup error, not product RED.

Final ignored evidence under `reports/runtime/`: `recording-isolation-green-summary.json`,
`recording-isolation-focused-retry.log`, `recording-isolation-chain-summary.json`,
`recording-isolation-chain.md`, `recording-isolation-chain-live.md`,
`recording-isolation-chain-create.md`, `recording-isolation-maintenance-checks.json`,
`recording-isolation-compiled-diff.json`, `recording-isolation-chain-residual-closure.json`,
`recording-isolation-chain-residual-window-classes.json`,
`recording-isolation-historical-preservation.json` and
`recording-isolation-final-preservation.json`.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-recording-owner-provenance -Phase RED -CheckRecordingReader -CheckRecordingLimits -CheckRecordingStorageBounds -CheckRecordingIsolation
```

## Foundation raw evidence

### Actual interrupted Excel restart: first execution and calibration

`-CheckRecordingRestart` retains the reader foundation and adds12 checks through
the existing Viewer Start/Stop, Action Paths selection and Admin Save Value
handlers. It creates a real durable Start plus two owning observations, verifies
the exact initial Excel HWND/process and all five candidate project paths, checks
every non-package workbook belongs to the generated temporary fixture and is
saved, then deliberately terminates only that held process. No Close record is
deleted and no runtime reset substitutes for interruption. A fresh Excel process
loads the same package files; credentials and copied test drivers stay in memory.

The expected contract is Interrupted, never Conclusion observed or a resumed
recorder. Reads preserve all fixture files. A later ordinary action has no
sequence, and a new explicitly started/stopped run gets a distinct identity while
preserving the old evidence. Package hashes must remain unchanged. Three changed
PowerShell parsers and whitespace checks pass; runtime execution/compile of the
fresh-process drivers remains unverified. No behavioral RED or GREEN is claimed.
The existing Excel-open guard prevents running while the chain's residual dialog
remains. This adds D13 evidence for existing D18; it changes no runtime contract.

Run separately from the unchanged154-check limits/storage/isolation gate:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-recording-isolation -Phase RED -CheckRecordingRestart
```

Failure to set up the fixture, prove process ownership, restart Excel or compile
the disposable drivers is a harness failure, never meaningful product RED.

The previous chain's residual dialog is now identified through MSAA as Document
Recovery. The exact process identity was reverified, **Yes, I want to view these
files later** was selected and its checked state verified, then OK allowed normal
exit. No forced stop or recovery-file deletion was performed. This explains that
shutdown obstruction; the earlier RPC/native failures remain unresolved. See
`recording-isolation-chain-dialog-options.json` and
`recording-isolation-chain-final-closure.json` under ignored runtime evidence.

The first cold-restart run completes **111PASS/1harness failure**, exit1. All108
preceding recording/reader checks and the first three restart checks pass: a real
unclosed journal is durable, deliberate interruption preserves it, and a different
Excel process loads the same five package paths. Fresh-process Viewer opening
then encounters a VBA compiler dialog: **Error accessing file. Network connection
may have been lost.** The selected code pane is cOperationsAnchorManager, line75;
that observation does not establish a root cause or actual network loss.

After acknowledging that exact compiler error, the verified test project remains
in break mode. Its enabled Reset command is used only to end this failed test;
the harness records its error and Excel exits normally. No reset substituted for
the earlier actual interruption, and no later restarted-reader checks passed in
this attempt. All five candidate file hashes remain unchanged. This is a harness
failure, not behavioral RED or full cold-restart GREEN.

The next controlled run adds one calibration check: open the pristine packaged
Viewer through its existing action wrapper before installing fresh-process test
drivers. Close it, install only those unsaved drivers, then continue all original
12 restart checks. This distinguishes pristine launch from instrumented launch
without changing runtime packages. Thirteen unique restart checks and PowerShell
syntax are verified. Ignored first-run evidence:
`recording-restart-first.log`, `recording-restart-first-exit.json`,
`recording-restart-dialog-classification.json` and
`recording-restart-compiler-position.json`.

The calibrated run also ends **111PASS/1harness failure**, exit1, preserving the
same108 earlier checks and three interruption/process checks. The pristine
`modInventoryViewer.RunInventoryViewerActionForTest` call encounters the same
compiler message **before any fresh-process test drivers are installed**. Thus
those edits are not a necessary trigger in this scenario. This does not prove a
general cold-start defect: the verified scenario uses the same PowerShell
controller after intentionally terminating the preceding Excel process.

After acknowledgement, the COM attachment supplies no usable code pane or debugger
state; no null value is treated as proof of an empty or reset process. The exact
failed fresh test process is then terminated, releasing the waiting harness.
Its final RPC-unavailable message follows that deliberate cleanup and is not a
new independent crash finding. Both runs are terminal and Excel is closed.
The original interruption and failed-test cleanup are distinct actions; no
post-reset/cleanup observation is promoted to restart acceptance.

The closed-Excel preservation check passes260 package pins,15 protected source
files and the existing reviewed Shipping visibility-only difference. Runtime
source and packages remain unchanged. Full cold-restart, clean native and visible
acceptance remain open. Next isolate the controller boundary: start the reader
in a separate fresh PowerShell process with in-memory fixture input, then prove
pristine Viewer launch before adding drivers. Do not infer network loss from the
compiler wording or patch runtime behavior without a protecting behavioral test.

Additional ignored evidence: `recording-restart-calibrated.log`,
`recording-restart-calibrated-exit.json`, `recording-restart-pristine-dialog.json`,
`recording-restart-calibrated-cleanup.json`, `recording-restart-first-summary.json`
and `recording-restart-final-preservation.json`.

### Separate reader controller and standalone pristine startup

The reader now runs in `Slice4beRecordingRestartWorker.ps1`, a distinct PowerShell
process verified by its actual parent/process identities. Fixture input travels
only through redirected stdin; command-line arguments contain the worker path,
not fixture data. Worker output contains fixed stages and boolean check records;
raw exceptions and input are not emitted. `Test-Slice4beRecordingWorkerProtocol.ps1`
passes6/6: valid/malformed/missing/extra/escaped-root input, output redaction, and
no Excel or fixture creation. These are transport tests, not product acceptance.

`Slice4beRecordingFixture.ps1` shares the existing14 helper bodies unchanged;
source comparison verifies that extraction, five PowerShell parsers pass, and
all13 previous restart check identities remain. The fresh-controller assertion
adds one check. No runtime VBA, schema, control, architecture or package changes
are made. The parent skips its final Viewer close only when its Excel reference
has already been cleared after the proven interruption.

The full run completes **112PASS/2harness flags**, exit1. All108 earlier checks,
the original interruption/process checks and the distinct-controller check pass.
Pristine Viewer launch in the worker reaches the same compiler file-access
dialog before driver installation. The worker and enclosing harness both flag
that one failure; they are not two independent product defects. Acknowledgement
does not release the invocation, so only the exact verified failed worker Excel
process is terminated. Both controllers then exit and Excel is closed. A fresh
reader controller alone does not resolve the failure; the original creator
controller is still alive during this comparison.

A separate `Test-Slice4bePristineViewer.ps1` uses the established Admin Generate
Warehouse/auth fixture functions, without running their surrounding instrumentation
or editing VBA. In its own new Excel process, with no interruption within this
testcase, the existing Viewer action wrapper opens and reuses the same form.
It passes **6/6**, exit0, preserves all fixture files and five package files, and
Excel closes normally. This disproves a universal pristine-launch failure under
that tested setup. It does not isolate the cause: this control both lacks a live
old creator and performs warehouse bootstrap in the fresh Excel process.

Next use a neutral coordinator so the creator controller fully exits before a
fresh reader consumes the same saved recording fixture, with no repeated bootstrap
or VBA edits before pristine launch. Keep fixture transfer private/in memory and
restore registry/fixture ownership across both processes. Full interrupted-reader,
native reliability, Operations sequences, conclusions and guide comparison remain
open; do not substitute the standalone6/6 for them.

Ignored evidence: `recording-controller-static.json`,
`recording-fresh-controller.log`, `recording-fresh-controller-exit.json`,
`recording-fresh-controller-processes.json`, `recording-fresh-controller-dialog.json`,
`recording-fresh-controller-cleanup.json`, `pristine-viewer-first.log` and
`pristine-viewer-first-exit.json` under `reports/runtime/`.
Final preservation verifies260 package pins,15 protected source files and the
existing reviewed Shipping visibility-only change with Excel closed. All six
changed PowerShell files parse,83 local document links resolve, and every111
previously passing restart-attempt identity remains GREEN in the fresh-controller
run. See `recording-controller-final-preservation.json` and
`recording-controller-final-checks.json`. No new runtime build, static-baseline
regeneration, full-chain rerun or human acceptance is claimed for this test-only
checkpoint.

### Creator exit before the fresh recording reader (2026-09-14)

`Test-Slice4beRecordingControllerExit.ps1` coordinates two sequential controllers
without opening Excel itself. The creator retains the108-check foundation, starts
a real recording through Viewer, records an actual Admin Save Value, and deliberately
interrupts only its verified disposable Excel process. It transfers the existing
fixture through a current-user-only named pipe and exits completely. The fresh
reader verifies that exit, receives the fixture through private stdin, and opens
the same five unchanged packages and saved warehouse without repeated bootstrap
or VBA edits before the pristine Viewer callback. Fixture input remains in memory;
only fixed stages and boolean checks reach output. The coordinator restores
settings and removes its disposable fixture only after both controllers and Excel
are terminal. A live process defers cleanup explicitly.

The transfer calibration initially passes4/5: Unicode text is not preserved by
the console's implicit input decoding. Explicit UTF-8 stream reads/writes correct
the transport; the unchanged test passes5/5, including the large Unicode payload,
current-user ACL, creator exit, no public payload output and no Excel creation.
Worker protocol/redaction checks pass6/6. These are harness checks, not D13 product
RED/GREEN. The first coordinator launch fails before Excel/checks because it passes
an absolute package path to the existing relative-path harness; forwarding the
original relative argument corrects that setup error.

The second coordinator run passes **123/123**, exit0, with normal final Excel
closure. Every112 prior passing check identity is retained, with no duplicates;
the108-check foundation and all15 restart checks pass. This includes verified
creator exit, pristine Viewer launch, no resumed recorder, selection of the real
unclosed journal as Interrupted rather than concluded, original attempt/outcome
visibility, read-only fixture preservation, ordinary work outside that sequence,
a distinct new explicit recording, preserved interrupted evidence and package
hashes. All14 existing shared fixture helper bodies remain unchanged.

This establishes the focused interrupted-reader gate under the tested process
lifecycle. It does not explain every earlier compiler/native failure or prove
general native reliability. No runtime, package, schema or architectural contract
changes are made; no new build/static/full-chain or human UAT result is claimed.
Full Operations/Admin multi-submission and multi-event sequences, deferred source
outcomes, conclusions, How-To/Diagnostic/Compare, versioned guides and visible
NAS/operator acceptance remain required. Continue with actual Operations handlers
inside one recording and exact multi-event references through published outcomes.

Command: `powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beRecordingControllerExit.ps1 -RepoRoot . -DeployRoot deploy/validation-recording-isolation -Phase RED`.
The phase argument names the evidence run;123/123 is GREEN on the unchanged
candidate, and neither setup error is a behavioral RED.

Ignored evidence under `reports/runtime/`: `recording-controller-exit.log` (setup
failure), `recording-controller-exit-second.log` (123/123),
`slice4be-viewer-published-read/6738f2b85dce440c96e55bdc51a7d685/controller-exit-checks.json`,
`recording-controller-exit-check-retention.json` and
`recording-controller-exit-preservation.json`. Final preservation verifies260 package
pins,15 protected source files and the existing reviewed Shipping visibility-only
change, with Excel closed. All seven changed PowerShell files parse.

### Admin/Operations sequences and deferred source application (2026-09-14)

The new `-CheckRecordingOperations` gate retains the108-check reader foundation.
Its generated warehouse uses Admin Generate Warehouse and Seed; all unsaved test
driver/fault edits precede authentication and live forms. The Receiving launcher
enters the actual generated Ribbon callback through an IRibbonControl fixture.
Add and Confirm enter the same handlers on that launcher-owned form, without
rebinding it when the unrelated workbook is active. Admin Save Value begins and
ends the shared sequence. No activity, event ID or inventory key is fabricated.

The scenario records nine actions,18 original observations and20 journal entries.
Two first-submission receipts remain pending while two new receipts are staged.
The next Confirm retains all four source references, including the two earlier
identities; the library must display all six references across those observations.
Only automatic processing/refresh is withheld. The ordinary queue persists, and
the later owning Admin processor applies four distinct events. Publication has
zero of those applied groups beforehand and four afterward, preserving each exact
System_Key. A dedicated disposable Admin workbook receives console audit output.
Original journal bytes and unknown staging columns are protected. Evaluation may
append separate derived records under D18; the journal preservation assertion does
not prohibit those new records.

The focused expected RED is the missing actual library **Evaluate** action in
both pending and applied views. Selection/read evidence must not claim a conclusion
without an expectation. This presence test alone cannot establish evaluator GREEN:
before runtime implementation, refine the expectation/evaluation surface and wire
under approved D18 and add positive, awaiting, failed/cancelled/incomplete, stale,
restricted and corrupt-evidence cases, plus immutable derived-result checks.
How-To/Diagnostic/Compare and versioned guides remain required.

Calibration history is retained separately from behavioral RED:

- First run:110PASS/1harness failure at Add. The probe reused a reference while
  expecting a new row; the owning staging service merges matching receipt
  combinations. Distinct references and actual Refresh readiness calibrate the
  fixture without changing that service or generating replacement identities.
- Calibrated run:130PASS/3FAIL. Both missing Evaluate checks fail after successful
  real submissions/application. The third assertion combines byte preservation
  with Saved=True and requires further separation; it is not proof of a data write.
- Capture-enabled run:19PASS/1harness failure in an existing Viewer foreground
  screenshot guard, before the new scenario. No new visible acceptance is claimed.
- Expanded preservation run:134PASS/3FAIL. Both Evaluate REDs remain; every staged
  diagnostic preserves the unrelated file bytes, one-sheet/one-cell shape and
  sentinel. Saved is False by the first-submission checkpoint; this probe does not
  separate its two Add actions from Confirm. Historical GREEN
  Receiving cases already record the same flag transition while asserting bytes.
  The corrected check protects file/content and leaves Saved diagnostic; it never
  forces the workbook clean. A standalone calculation-mode control was rejected
  by automatic approval review with only "blocked by policy"; that causal
  hypothesis remains unverified and supplies no justification for runtime changes.
- The first corrected rerun stops120PASS/1harness failure at Admin publication with
  RPC unavailable. A remaining Excel instance is reliably bound by window/process
  identity and has zero workbooks plus Core/Admin/Operations add-in projects.
  Normal Quit preserves all three loaded add-in files, then the exact process
  terminates. No force termination or recovery-file deletion occurs. This failure
  is neither evaluator RED nor a clean final gate, and remains a native-reliability
  limitation even if a subsequent unchanged-candidate run succeeds.

Command: `powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-recording-isolation -Phase RED -CheckRecordingOperations`.
The final clean rerun reports **135PASS/2FAIL**, exit1 as expected for RED, with
normal Excel closure. Only `RecordingOperations.Pending.EvaluateActionAvailable`
and `RecordingOperations.Applied.EvaluateActionAvailable` fail. Every108 foundation
identity and all130 earlier passing identities remain GREEN. Both repeated-reference
display checks and all file/content preservation checks pass. The earlier native
failure is not resolved by this successful fixture run. Three changed scripts parse.

Runtime remains `99a69aa`; no package, architecture or runtime source changes are
made. The generated driver changes are not saved to XLAMs. No new build, static
baseline, full-chain or human acceptance is claimed for this test checkpoint.

Ignored evidence under `reports/runtime/`: `recording-operations-first.log`,
`recording-operations-calibrated.log`, `recording-operations-preservation.log`,
`recording-operations-preservation-retry.log`, `recording-operations-final-red.log`,
the per-run `recording-operations-other-state.json`,
`recording-operations-residual-quit.json` and
`recording-operations-residual-terminal.json`, `recording-operations-final-red-retry.log`
and `recording-operations-final-checks.json`.

### Earlier foundation artifacts

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
