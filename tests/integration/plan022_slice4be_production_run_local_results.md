# Slice4be Production Run local observations

Architecture v4.11 D18 specifies nine catalog24 controls under approved semantic
inheritance: Scale, Clear, Load, Loader/Manager Refresh, List/Tree Apply and Tree
Expand/Collapse. Owner is PRODUCTION_RUN_LOCAL. Local completion is distinct from
inventory application; these controls carry no source-event references. D15's
experimental Tree status and D14's exact-key/unknown-column rules remain binding.

This is test-first work, not completed implementation or acceptance. Runtime
remains catalog23/118 IDs and55/68 constructed Production buttons. The separate
full reusable native crash remains unresolved; see
`plan022_slice4be_production_assignment_results.md` for its bounded diagnostics
and the thirty verified shared gates on the frozen Assignment candidate.

## Protecting test and candidate

`tests/tooling/Test-Slice4beProductionRunPresentation.ps1` runs the isolated
ConfigCommands harness against frozen, unpromoted
`deploy/validation-production-assignment-01`. Its unsaved adapters in
`Slice4beProductionRunPresentation.ps1` call the original packaged
`mBtnRunTreeExpandAll_Click` / `mBtnRunTreeCollapseAll_Click` handlers, not replacement
handlers. They stage synthetic local palette rows without inventory writes.
The opt-in gate is independent of the existing ConfigCommands selections.

The initial baseline covers only this presentation pair: expanded/collapsed/
process-collapsed/empty trees, repeated actions, unchanged palette contents,
fixed metadata, exact paired observations, redaction/integrity/terminal semantics,
loading/busy suppression, stale target/session/sign-out/closed workbook, saved
authority, unknown local values/formula and older-record preservation. All nine
controls, the complete optional-policy/failure/yield matrix and paired Action Path
publication/reader evidence remain required before implementation acceptance.
No runtime source, package build, deployment or promotion has changed here.

## Verified presentation-pair RED

Controller `reports/runtime/production-run-presentation-controller/397ec444067d43b68e61488b307aaea5`;
result `reports/runtime/slice4be-production-run-presentation/92792238dab347198548a78248cc4a3c/red.json`.
2026-10-01 06:07:45.6116044--06:09:33.3276574 UTC: **98 PASS /134 FAIL /232 unique checks**.
The134 failures are exactly two absent catalog definitions,112 absent observation/
terminal facts, four missing loading/busy guards and16 missing captured-context
mutation/refusal guards. Every existing tree shape, palette value and status check
passes, including preserved collapsed Process groups and empty/repeated actions.
All42 prior shared checks remain GREEN in exact order; the other56 passes establish
the preserved owner/preservation behavior. No assertion is removed or relaxed.

`verification.json` records five instrumented package compiles, canonical frozen
package pins, preserved packages/settings, normal unassisted closure, no remaining
Excel and zero delayed Application1000/1001/1002 failures. This is meaningful RED
for the two actual handlers only. It neither protects all nine controls nor
completes the optional-policy/failure/yield/recording matrix. No screenshot or
human-acceptance claim is supplied by this baseline.

Static evidence `reports/runtime/run-presentation-static-01`:280 components,
6150 procedures,135006 lines,9 literal/45 unresolved dynamic calls,190 duplicate
bodies,28 non-growing module caps. Three schemas validate and356 PowerShell
scripts parse. Runtime metrics/source are unchanged. The two unrelated document
hashes remain pinned and preserved. Desktop monitoring through this baseline
reported no new Win32 error5; it does not guarantee future unlocked access.

## Excluded fixture attempts

- Controller `reports/runtime/production-run-presentation-controller/3025ae9e856b437591c7e54d6c9f96cc`,
  result `reports/runtime/slice4be-production-run-presentation/a894e4e652ad42068a8ecedb0f978fe1/red.json`,
  2026-10-01 06:00:28.5397603--06:01:51.9510085 UTC:39 PASS/1 harness FAIL.
  An incorrectly parenthesized PowerShell argument array passed one argument to
  the two-argument catalog Definition probe. This is not behavioral RED. Corrected
  test call; runtime unchanged. Five compiles, normal closure, settings/package
  preservation and zero delayed Excel Application1000/1001/1002 failures.
- Controller `reports/runtime/production-run-presentation-controller/b65cd6db964141ce92b580ce709f20fd`,
  result `reports/runtime/slice4be-production-run-presentation/5f31a81995d34b2985a0d0cd1a99df1c/red.json`,
  2026-10-01 06:02:12.2937350--06:05:04.5627583 UTC:39 PASS/2 FAIL.
  One missing-metadata assertion preceded fixture Stage error91. The owned Visual
  Basic dialog was inspected through allowlisted error/button facts and dismissed
  using End; cleanup was assisted, so this is excluded from protecting RED.
  Settings/packages were preserved, Excel closed and delayed Application audit
  found zero failures. Error91 is not desktop Win32 error5. The next adapter returns
  only a fixed failing setup stage and numeric error, rather than leaving a dialog.
- Controller `reports/runtime/production-run-presentation-controller/c14544f577b04218a052da0d819bee85`,
  2026-10-01 06:05:53.3664289--06:07:14.8565239 UTC:39 PASS/2 FAIL.
  The bounded fixture diagnostic identifies TreeState/error91. In the adapter,
  `EnsureRunTreeState:` was parsed as a VBA label rather than a no-argument call,
  so the dictionary was never initialized. Separating the call onto its own line
  corrects only the fixture. This diagnostic is excluded from RED; five compiles,
  unassisted closure, preservation and delayed zero-failure audit pass. Do not
  repeat the colon form for no-argument procedure calls in VBA adapters.

## Seven-control reusable baseline

The seven-control reusable baseline is protected through
`Test-Slice4beProductionRunLocal.ps1`, `Slice4beProductionRunLocalProbe.ps1` and
`Slice4beProductionRunLocalActivity.ps1`. Admin Generate Warehouse/Seed supplies
real inventory; actual Process/Recipe Save/Release supplies released definitions.
The tests invoke Load, Scale, Clear, Loader/Manager Refresh and List/Tree Apply
unchanged. They cover scale limits, reset of prior allocation/actual-output/note
state, false-load clearing, allocation quantity precedence, single-row selection,
palette refill, input/location/requirement refusals and captured-context guards.
The completed-node scale refusal uses explicitly synthetic local state; it is
not evidence of completing a Process or applying inventory events. Worksheet
branches, multiple-key bucket expansion, full policy/fault/yield and paired-path
coverage remain required. No runtime implementation is changed.

First seven-control attempt is excluded: controller
`reports/runtime/production-run-local-controller/4350b735ccc14405bb566769dc57c56a`,
result `reports/runtime/slice4be-production-run-local/013b5feff05f4bf6aea1e35f96ea75d6/red.json`,
2026-10-01 06:17:08.7487641--06:25:26.0073287 UTC. Its460 unique checks contain
131 PASS/329 FAIL, including two harness/cleanup failures. All33 measured owner
behavior cases pass before the failure; this partial observation is not protecting
RED. Invoking a retained form reference during repeated closed-workbook cases
raised VBA80010007 and RPC800706BE. Windows Application1000 at06:24:02.3256648 UTC
records c0000005, with1001 at06:24:13.8035447. These are not desktop Win32 error5.
The owned End action and closing three verified, saved temporary-fixture recovery
workbooks without saving are assisted cleanup. Settings/packages were preserved
and Excel exited; final canonical-authority preservation was not reached/proven.

The corrected fixture follows the established Assignment/worksheet native-lifetime
pattern: show the actual form, measure visibility and exact workbook closure,
invoke the handler through an outer trapped adapter only if the form survives,
and record dismissal as no handler invocation. Independently retain the typed
closed-binding check, in-memory owner state, disk bytes and activity counts.
Cache fixture identities before disposal and release the old reference before
opening the next workbook. `-ClosedDiagnostic` exercises only that boundary and
cannot substitute for the complete seven-control baseline. No native-lifetime
or broader reusable-crash repair is claimed.

The narrow closure diagnostic, controller
`reports/runtime/production-run-local-controller/ed6efd93ba594c09a73003d33b06603d`,
result `reports/runtime/slice4be-production-run-local-closed/155320e9e9334bf798c7ee4925ac0f2a/red.json`,
2026-10-01 06:27:29.7044899--06:29:44.9533729 UTC, has84 PASS/5 expected FAIL/89
unique checks. SCALE, LOADER_REFRESH and ALLOCATE survived closure and entered
their actual handlers without adapter errors; the other four forms were dismissed
and their handlers were not invoked. Owner mutation/refusal guards account for
the five failures. Five compiles, preservation, unassisted closure and delayed
zero Excel Application failures pass. This was a fixture diagnostic only.

The full baseline additionally selects the actual List/Tree Run page before each
closure. Controller
`reports/runtime/production-run-local-controller/2b5df5badbba45019b981d281d216d23`,
result `reports/runtime/slice4be-production-run-local/4e9de4535352471483dbb279664df653/red.json`,
2026-10-01 06:31:20.2375897--06:36:51.0284396 UTC: **171 PASS/332 expected FAIL/503
unique checks**. All33 existing owner cases and42 prior shared GREEN identities
in exact order pass. Failures are seven absent metadata definitions,264 absent
observation/terminal checks,14 missing loading/busy guards,40 changed-context
mutation/refusal checks and seven surviving-form closed-context guard checks.
The two Target allocation cases already preserve state through existing owner
validation; their missing explicit context refusal still fails independently.

All seven native closure receipts prove visible Run pages, exact captured-book
closure and preserved decoy. LOAD, CLEAR, LOADER_REFRESH and ALLOCATE survived
and entered their handlers with adapter error0. SCALE, MANAGER_REFRESH and
TREE_ALLOCATE were dismissed, so no handler invocation is claimed for them.
Native lifetime varies between runs; dismissal is never counted as a handler
test. All saved authority/operator bytes, unknown values/formula and older
activity records remain unchanged. Five compiles, canonical frozen package pins,
settings/package preservation, normal unassisted closure and delayed zero Excel
Application1000/1001/1002 failures are verified in the controller's
`verification.json`. Runtime source/builds remain unchanged. This protects the
seven reusable branches, not worksheet branches or the full nine-control gate.

Static `reports/runtime/run-local-baseline-static-01` retains280 components,
6150 procedures,135006 lines,9 literal/45 unresolved dynamic calls,190 duplicate
bodies and28 non-growing caps; three schemas and359 script parses pass.

## Supplemental Core contract RED

`Test-Slice4beProductionRunLocal.ps1 -ContractOnly` runs
`Slice4beProductionRunContract.ps1` separately from the operator-handler baselines.
Controller `reports/runtime/production-run-local-controller/67b171a992eb4881a7de972256be7bf8`,
result `reports/runtime/slice4be-production-run-contract/36394b1a48134c2c9e374678ed0e6c4c/red.json`,
2026-10-01 06:40:14.2539766--06:41:32.7523882 UTC: **492 PASS/272 expected FAIL/764
unique checks**, retaining all42 shared checks GREEN in exact order. All nine
controls are checked for catalog23 exclusion, catalog24 extension and preservation
of118 existing definitions, enabled/disabled defaults, fourteen candidate outcome
codes, exact terminal classification, wrong owner/catalog rejection and rejection
of Inventory/Designs references in both Submitted and Unknown states.

Failures are119 catalog24 extension/definition-preservation assertions,18 absent
defaults,45 absent supported outcome definitions, nine absent positive terminals
and81 unsupported outcomes currently accepted by the generic empty-reference
decoder. This does not mean that118 existing catalog23 definitions regressed:
catalog24 does not yet exist. Five instrumented compiles, frozen pins, settings/
package preservation, unassisted closure and delayed zero Excel Application
failures pass. This supplements actual handlers; it supplies no human acceptance.

## Optional-policy and denial RED

`Test-Slice4beProductionRunLocal.ps1 -PolicyOnly` uses the original nine handlers,
real Admin-seeded/released reusable fixture, and the existing synthetic local
Tree presentation fixture. It checks denied-user entry, disabled collection,
catalog23 policy compatibility, Navigation default-off, unavailable activity
storage, suppression of programmatic setup events and saved/local preservation.
Only disposable fixture paths are moved/restored to simulate unavailable storage.

Initial controller `reports/runtime/production-run-local-controller/2a94d6e91a4a48c98738cf89c1075b6a`,
result `reports/runtime/slice4be-production-run-policy/c4b13a28a6dc4ecd9dd46763e87f5c45/red.json`,
2026-10-01 06:42:47.6516635--06:45:27.4519967 UTC:173 PASS/36 expected FAIL/209
unique checks. These failures cover27 denied-user protection/record checks and
nine unavailable-tracking notices. Five compiles, preservation, normal closure
and delayed zero Excel audit pass. State equality alone cannot prove absence of
owner reads, so the next run adds nine explicit owner-entry counters.

Strengthened controller
`reports/runtime/production-run-local-controller/8996b1ce03934f45b06dcbfe9d00d002`,
result `reports/runtime/slice4be-production-run-policy/d8f8d3ad9f0848af8d75a87a1f04c7a2/red.json`,
2026-10-01 06:46:13.9329190--06:48:53.6727950 UTC: **173 PASS/45 expected FAIL/218
unique checks**. All209 prior identities/results retain their exact order; the
nine additional failures prove denied actions enter the existing reusable owner
or Tree presentation routine. All29 authorized optional-policy action cases
preserve their existing results, including the populated Tree shape/palette.
Programmatic setup produces no observations. Forty-two shared checks, five
compiles, canonical package pins, saved authority/operator bytes, unknown values/
formula, older records and other-warehouse activity preservation all pass.
Unassisted closure and delayed zero Excel Application1000/1001/1002 failures are
verified. No handler or runtime implementation was replaced or changed.

All361 tooling scripts parse. Runtime source and the validated three-schema
static baseline remain unchanged at280 components/6150 procedures/135006 lines,
9 literal/45 unresolved calls,190 duplicate bodies and28 non-growing caps.

## Owner-exception and nested-click RED

`Test-Slice4beProductionRunLocal.ps1 -FaultOnly` installs unsaved boundary hooks
through `Slice4beProductionRunFault.ps1`; all nine original handlers remain
intact. Nine exception cases and nine one-shot nested actual clicks reach their
intended owner routines. Exception hooks retain existing propagation/handled
failure behavior and local effects: Load has already cleared the reusable owner;
Loader Refresh has already refreshed recipe lists. No rollback is asserted.
Guard state is measured before any adapter reset, avoiding a test-created cleanup
that could hide a stuck runtime guard.

Controller `reports/runtime/production-run-local-controller/47316acdc3d74b7b8474c773dd9e48f3`,
result `reports/runtime/slice4be-production-run-fault/b3e401ded2304f9a92b5d54d15b5d2cc/red.json`,
2026-10-01 06:54:31.7257636--06:56:18.9449328 UTC: **127 PASS/118 expected FAIL/245
unique checks**. Failures are108 absent observation/terminal facts, nine repeated
owner entries and the nested Load result disturbed by reentry. All18 boundaries,
existing exception propagation/partial effects and guard-restoration measurements
pass. The nested Load failure protects the approved nesting guard; it does not
authorize changing the ordinary load algorithm.

All42 prior shared checks retain exact GREEN order. Five compiles, canonical
package pins, saved authority/operator bytes, unknown values/formula, older
records and other-warehouse activity preservation pass. Cleanup is unassisted;
delayed Excel Application1000/1001/1002 failures are zero. Runtime/source metrics
remain unchanged;362 scripts parse. This is a bounded owner-exception/nesting
baseline. Real-read failures/yields, worksheet branches, exact stock-bucket
expansion and independent paired paths remain open.

## Real-read sign-out RED

`Test-Slice4beProductionRunLocal.ps1 -YieldOnly` uses
`Slice4beProductionRunYield.ps1` and the common fault fixture. The first attempt,
controller `reports/runtime/production-run-local-controller/1e0e4adb8a47437ebb7f98f7a7249fdd`,
result `reports/runtime/slice4be-production-run-yield/653ab2e07ef3466c8597eb362e2c60b6/red.json`,
2026-10-01 06:59:40.1183981--06:59:49.5575550 UTC, is excluded:1 PASS/1 setup FAIL.
The adapter incorrectly assumed the referenced Core primitive bridge module was
present in the Operations VBProject. Subscript-out-of-range is neither product
RED nor desktop error5. There is no compile evidence for that attempt. Settings,
canonical package pins, unassisted closure and delayed zero Excel audit pass.

Corrected probes run at the existing Operations owner call sites immediately
after real reads return. They preserve the payload, capture the local state at
that boundary in memory, then sign out. Observed subsequent reads cover the
reusable owner's eight Inventory entity-read sites, released Recipe validation,
Recipe graph/Process payload reads and the form's Process/Recipe list reads.
No reverse dependency or new runtime bridge is introduced.

Controller `reports/runtime/production-run-local-controller/1c67e19b1b1049d7b1424dc85eb5c085`,
result `reports/runtime/slice4be-production-run-yield/5252226e68ea428a862e590ccea21203/red.json`,
2026-10-01 07:01:23.2907183--07:04:02.3303002 UTC: **97 PASS/49 expected FAIL/146
unique checks**. All11 cases prove an available owning read returned before
sign-out: four Load boundaries, Scale's Inventory read, three Loader Refresh
boundaries, Manager Refresh and both allocation actions. All11 currently continue
observed reads/change the captured local projection, lack the required refusal
and lack an original-context attempt without a misattributed outcome (44 FAIL).
Five cases also mutate owner state beyond its captured boundary (49 FAIL total).
Earlier local clearing/scaling/loading effects are retained rather than rolled
back. All guard checks occur before fixture reset. No unhandled errors occur.

All42 shared checks retain exact GREEN order; the common fixture's five final
preservation checks pass. Five compiles, canonical pins, settings/package
preservation, unassisted closure and delayed zero Excel Application failures pass.
Runtime/static metrics remain unchanged;363 scripts parse. This covers the
specified read-return boundaries, not every possible source failure or yield.

## Remaining gates

Worksheet-branch tests must distinguish quiet helper return from successful owner
completion: `mProduction.LoadRecipeChooser` and `BtnClearRecipeChooser` are Subs,
and the scale branch subsequently scales/prepares local tables. The former reads
the released Design/BOM bridge, while the reusable Run loader reads released
Recipe graphs and Process versions. This source audit does not establish a
runtime compatibility defect or authorize changing either algorithm. Protect
D14 unknown columns around generated-table rebuild and D15's no-fallback rule
through actual handlers before proposing any repair. A missing fixture or modal
test setup failure is not behavioral RED.

Preserve every established presentation baseline identity through GREEN. Add the
other seven owner behaviors and complete
context/policy/failure coverage; then require packaged XLAM, compile, layout,
static maintenance, live roles, full Release1 chain, reusable and independent
Action Path evidence. Agent inspection is not human acceptance.
