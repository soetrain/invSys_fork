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

## Multi-key stock allocation RED

`Test-Slice4beProductionRunLocal.ps1 -StockOnly` extends the common fault fixture
with `Slice4beProductionRunStock.ps1`. It preserves the original List/Tree Apply
handlers. Admin Seed supplies the first entity; the owning receiving writer and
processor create a distinct second `System_Key` at the same location before
measured local planning. Identities and quantities remain in memory.

Two initial attempts are excluded before any stock action. Both have39 PASS/one
harness FAIL, five compiles, canonical package/settings preservation, unassisted
closure and delayed zero Excel Application1000/1001/1002 events:

- Controller `production-run-local-controller/da6671d81b6f4d2eaa63e65d48c2bc48`,
  result `slice4be-production-run-stock/8071bcc215d043a6a11dd4a059060c05/red.json`,
  2026-10-01 07:12:09.5355313--07:13:28.7362455 UTC: SeedShape setup failure.
- Controller `production-run-local-controller/59621de216a34cf992c567db5f383b47`,
  result `slice4be-production-run-stock/429f40a2cf744ef1969eedebceac284e/red.json`,
  2026-10-01 07:17:31.2161879--07:18:53.9716061 UTC: narrower diagnostics identify
  the fixed seed-quantity assumption. These paths are beneath `reports/runtime/`.

The corrected fixture derives its sizes from the available quantity returned by
the owning Inventory read. With each entity quantity q, requirement2q+2, inputs
q+1 and75% each force two-key expansion;2q+1 exceeds stock while remaining below
the requirement. Zero must clear both allocations. This changes test setup only;
it neither changes Seed behavior nor establishes a Seed quantity defect.

Corrected controller `production-run-local-controller/f3d5506fd2f9440980992912bd18333a`,
result `slice4be-production-run-stock/c9f30ef7f70b40679117770f40c2f2df/red.json`,
2026-10-01 07:19:56.5153890--07:21:43.9162408 UTC: **82 PASS/56 expected FAIL/138
unique checks**. All eight actual-handler allocation cases pass: quantity takes
precedence over percentage, percentage-only input expands across both original
keys, zero clears both, and over-stock refusal preserves the earlier plan/inputs.
Exact identities and on-hand quantities remain unchanged; no submission occurs.
All failures are missing observation-pair, context, redaction and terminal facts,
including explicit exclusion of both exact keys from the observation payload.

All42 prior shared GREEN identities retain exact order. Five compiles, canonical
package pins, saved authority/operator bytes, unknown values/formula, older
records and other-warehouse activity preservation pass. Closure is unassisted;
delayed Excel Application1000/1001/1002 failures are zero. Runtime source and the
validated static baseline remain unchanged; all365 repository PowerShell scripts
returned by `rg --files -g '*.ps1'` parse. This is packaged handler evidence,
not visible operator or human acceptance. Preserve all138 checks through GREEN.

## Worksheet allocation RED

`Test-Slice4beProductionRunLocal.ps1 -WorksheetOnly` installs
`Slice4beProductionRunWorksheet.ps1` into the unsaved packaged form. Admin Seed
provides the exact inventory identity; a disposable operator Production table
stages that identity with reordered managed headers, custom data and a formula.
The real List/Tree Apply handlers run on their respective pages while a decoy
workbook is active. Inactive-page inputs deliberately differ. No legacy recipe
fallback or canonical inventory write is introduced.

Controller `reports/runtime/production-run-local-controller/96682e2fc5d0454e8f6dc760d6cbfbc6`,
result `reports/runtime/slice4be-production-run-worksheet/62829354876041d491e2146f799dc441/red.json`,
2026-10-01 07:27:40.5323304--07:29:19.7323917 UTC: **156 PASS/116 expected FAIL/272
unique checks**. Six ordinary cases per handler pass: quantity-source priority,
percentage-source priority, quantity-only inference, zero, over100% refusal after
input mirroring, and wrong-location refusal that clears an earlier allocation.
Checks retain normalized-header writes, List/Tree synchronization and overrides.
All16 cases preserve custom columns/formula, exact identity and canonical source.

Two unavailable-surface cases per handler expose a presentation contradiction: a missing
palette table or QUANTITY column still produces an allocation-updated success
message. FAILED owner facts are already required, but changing that message needs
Architecture D18 decision RUN-UI-01 because the contract also preserves existing
operator messages. These four assertions are a pending proposal, not approved RED;
they do not authorize rollback or repair of missing surfaces. The other112 failures
are missing observation/context/redaction/terminal facts. Tests do not pin the
accidental local UI changes in unavailable-surface cases as required behavior.

Forty-two prior shared GREEN checks retain exact order. Five compiles, canonical
package pins, saved authority/operator bytes, unknown values/formula, older
records and other-warehouse preservation pass. Unassisted closure and delayed zero
Excel Application failures are verified. Runtime/static metrics remain unchanged;
366 repository PowerShell scripts parse. Preserve the268 approved checks through
GREEN; the remaining four presentation assertions depend on RUN-UI-01 approval.
Worksheet load/clear/refresh branches, remaining read failures and independent
paired paths still require evidence. This gate is not visible or human acceptance.

## Headless Core catalog24 GREEN

Unpromoted `deploy/validation-production-run-catalog-01` implements only the Core
part of the approved Run contract. `modProductionRunCodes` supplies the nine
definitions and fixed outcomes. `modActivityCatalog`, `modActivityReferences` and
`modEvaluationMatches` route those definitions, reject unsupported/nonempty source
references and require the exact declared positive outcome for CommandCompleted.
No Run form handler or workflow algorithm changed. RUN-UI-01 is not implemented.

Controller `reports/runtime/production-run-local-controller/bf38070b9ca7440f96602c37e72506ac`,
result `reports/runtime/slice4be-production-run-contract/76fa7339038647e79d9611808e1ca7e0/green.json`,
2026-10-01 07:33:45.7306014--07:35:08.2965816 UTC: **764/764 GREEN**, retaining the
492 PASS/272 FAIL RED's exact ordered identities. All118 catalog23 definitions are
unchanged in catalog24; defaults, outcomes, references and terminal cases pass.
Forty-two shared checks, five instrumented compiles, candidate pins, preservation,
normal closure and delayed zero Excel Application failures pass.

Build `reports/runtime/run-catalog-build-01`,2026-10-01
07:32:41.1624265--07:33:20.9863254 UTC, passes five-package generation and independent
cold-start compile. The settings snapshot precedes the build and is retained only
in memory; restoration, frozen candidate preservation, unassisted cleanup and
delayed zero Excel audit pass. Compiled comparison allows exactly three modified
Core components and the one new codes module;270 prior components are unchanged
ignoring identifier casing, with string literals compared exactly and none lost.

Older worksheet Core regression: controller
`reports/runtime/process-worksheet-activity-controller/fcd587a3bf6d4246b65b7d0daa2330ac`,
result `reports/runtime/slice4be-process-worksheet-activity/19b9d234eb0248869cf045e6867e318b/green.json`,
2026-10-01 07:36:48.0738644--07:38:10.3113951 UTC: **188/188 GREEN**, exact prior
order, five compiles, preservation, normal closure and delayed zero Excel audit.

`reports/runtime/run-catalog-static-01` has281 components/6154 procedures/135103
lines: growth1/4/97 for the catalog feature. Calls remain9 literal/45 unresolved,
duplicates remain190, and all28 oversized caps do not grow. Three schemas and365
tooling PowerShell parses pass. Fifteen current-output assertions advance their
exact catalog expectation23->24, with policy count118->127; historical catalog
fixtures retain their original versions. Broader workflow regression is still
required. The candidate has127 definitions, but Run observation coverage does not
increase until the handlers are integrated and their actual-handler gates pass.

Settings regression: controller
`reports/runtime/run-catalog-regression/settings-3685675ac2f9490298e056b7f547a584`,
result `reports/runtime/slice4be-tracking-settings/6059a972ea45422a8d3830ec92ee40d5/green.json`,
2026-10-01 07:38:32.1263206--07:42:41.3075299 UTC: **202/202 GREEN**, retaining exact
prior order. The editor loads catalog24/127 definitions; policy/detail/preference,
Admin close and Operations Settings captured-context behavior pass. Five compiles,
candidate pins, settings/config preservation, unassisted closure and delayed zero
Excel audit are verified. No visible human acceptance is inferred.

Ingredients Assignment regression: controller
`reports/runtime/production-assignment-controller/c4b84dccf3f3482a9b7360f8f1c7fc81`,
result `reports/runtime/slice4be-production-assignment/c7d8afe521404b90b0060186ab1f4899/green.json`,
2026-10-01 07:43:44.3667698--07:46:35.2303245 UTC: **349/349 GREEN**, preserving
the exact prior order. Five compiles, original actions/observations, context guards,
candidate pins, saved authority/operator bytes, older records, settings restoration,
normal closure and delayed zero Excel audit pass. Its expanded789 companion and
independent paths, plus remaining shared/full-release gates, are not rerun here;
the recorded frozen catalog23 acceptance remains separate from this candidate.

## Worksheet Clear and Refresh owner baseline

On the catalog24 candidate, `-WorksheetOwnerOnly` exercises both real Refresh
handlers and Clear with their local surfaces present or unavailable. Controller
`reports/runtime/production-run-local-controller/a6db581725d841118533048336893cc5`,
result `reports/runtime/slice4be-production-run-worksheet-owner/555fc5547190439ebd46d7bb0f62e5da/red.json`,
2026-10-01 08:07:08.5045527--08:08:50.3581007 UTC: **84 PASS/38 FAIL/122 checks**.
Thirty-six failures are missing observations; two expose Clear's existing owner
binding defect. All42 prior shared GREEN checks retain exact order; five compiles,
canonical candidate pins, saved authority/operator bytes, unknown values/formulas,
older records, settings restoration, normal cleanup and delayed zero Excel
Application failures pass. Runtime source and static metrics are unchanged;
367 repository PowerShell scripts parse.

The fixture uses the supported Core Production surface and an Admin-seeded exact
System_Key. It stages only the owned unsaved operator workbook, activates a decoy,
and temporarily renames local surfaces. Both Refresh buttons preserve the real
LOCAL owner's Boolean result: a missing invSys table returns False; ordinary
success does not prove freshness and can retain cached data. Custom columns,
formula and exact key survive all six cases; canonical entity values and saved
authority remain unchanged. Ordinary Clear deletes generated chooser/palette
tables and clears chooser/output/check contents as before.

**Captured-owner defect:** with the operator's Production sheet unavailable,
`BtnClearRecipeChooser` resolves an `OtherAddin` worksheet; ordinary Clear resolves
`Captured`. The read-only probe observes the actual resolved Worksheet before
cleanup. `ResolveBoundProductionWorkbook(requiredSheet)` rejects a bound workbook
missing that sheet; `ResolveProductionWorkbook` searches elsewhere, then
`SheetExists` can fall back to `ThisWorkbook`. This violates the existing captured
workbook invariant and D18's missing-owner-surface rule. The two failing checks
are `CLEAR.Missing.RealOwnerResultAndNotificationCount` and
`CLEAR.Missing.OwnerUsesCapturedWorkbook` under `RunWorksheetOwner`. The eventual
guard must prevent entering another workbook/add-in owner; a quiet Sub return is
not successful completion. No broad resolver rewrite or message exception is
approved by this diagnosis. RUN-UI-01 remains pending and separate.

Only three notification calls are captured in memory by unsaved adapters; their
owning work and handlers run normally. Notification counts are automated evidence,
not native modal or visible operator acceptance. Raw reports and workbook names
stay in memory; the diagnostic outputs only fixed target categories.

Three incomplete fixture attempts are excluded: controller `18ffd54c0ce74cf2ad4590f9e314cf17`
failed an exact notification seam before handler execution; `9e26b032a3c747509d838844ba125c21`
and `3d761ac58b904d878e30c6a070df3b64` stopped during Clear staging because default
row insertion would shift neighboring tables. Using the surface's reserved blank
row with `AlwaysInsert:=False` corrected setup. All had normal preservation/cleanup
and delayed zero Excel audits; the latter two compiled all five packages.
The completed120-check diagnostic, controller `4bd5235e4ea14784a94f7361dc7c6d0f`,
result `slice4be-production-run-worksheet-owner/94e693a7ddca45f5a96acd62caff272a/red.json`,
was83 PASS/37 FAIL; the final122 retains those identities and adds the two explicit
Clear owner-target checks. Its five compiles, preservation, normal cleanup and
delayed audit also pass.

## Clear captured-owner guard

The populated-decoy extension retains the122 checks and adds14. Controller
`reports/runtime/production-run-local-controller/3f53e34bff944cbb9fb3c93b07159d1e`,
result `reports/runtime/slice4be-production-run-worksheet-owner/4898f6aa7bf8446fa243d612e14c14bc/red.json`,
2026-10-01 08:12:23.6588556--08:14:06.2194191 UTC: **89 PASS/47 FAIL/136 checks**.
Clear resolves `OtherWorkbook` when a populated Production sheet exists on the
active decoy, then clears its chooser contents, including its custom value and
formula. This proves an actual cross-workbook mutation, not just a misleading
status. With no decoy surface it resolves `OtherAddin`. Five binding/preservation
failures and42 missing-observation failures supply RED. All earlier PASS results,
five compiles, package/settings/saved authority preservation, unassisted closure
and delayed zero Excel audit pass.

Unpromoted `deploy/validation-production-run-binding-01` adds
`modProductionRunBinding.BindWorksheetOwner`: require the original workbook object
to be in Application.Workbooks, not an add-in, and contain the Production sheet;
only then bind it to the existing owner. Worksheet Clear calls this guard before
entering `BtnClearRecipeChooser`. Reusable Clear, ordinary cleanup, current
messages, other handlers and the general fallback resolver are unchanged. This
implements the existing captured-workbook and missing-owner requirements; no new
architectural exception or RUN-UI-01 approval is inferred. Missing-sheet Clear
returns before the owner and its success notification/status. FAILED tracking
still needs the separate Run observation integration.

Controller `reports/runtime/production-run-local-controller/b2757a112c8341fbbffd1c69ac40001f`,
result `reports/runtime/slice4be-production-run-worksheet-owner/43b3eb71e5a9416989140c22be9c78ab/green.json`,
2026-10-01 08:15:47.2063526--08:17:35.5809681 UTC: **94 PASS/42 FAIL/136 checks**.
All five binding checks change RED->GREEN, all89 earlier PASS checks remain GREEN,
and all136 identities/order are retained. Both unavailable-sheet cases report
`NotEntered`; ordinary Clear remains `Captured`. Decoy contents/custom formula
are preserved. Five compiles, package pins, saved/local/settings/older-record
preservation, normal cleanup and delayed zero Excel audit pass. Despite the
driver's `green.json` filename, **the overall suite remains RED** for42 missing
observations. No Run observation coverage or human acceptance is gained here.

Build `reports/runtime/run-binding-build-01`,2026-10-01
08:14:45.3863653--08:15:23.7254859 UTC, passes five-package generation and independent
cold-start compile. Prebuild settings capture/restoration and the prior catalog24
candidate's package preservation pass. Compiled comparison identifies one changed
form and one new Production module;273 other components are preserved, comparing
string literals exactly and ignoring identifier case, with none removed. The
delayed build Excel audit is zero; it was performed after the subsequent owner
gate had already started, rather than before starting that gate.

`reports/runtime/run-binding-static-01` records282 components/6155 procedures/
135128 lines: growth1/1/25. Dynamic calls remain9 literal/45 unresolved; duplicate
body candidates remain190. All28 oversized-module caps are non-growing, all three
schemas validate,366 tooling scripts and367 total repository PowerShell scripts
parse. `frmProduction` does not grow.

Seven-control reusable regression on the binding candidate: controller
`reports/runtime/production-run-local-controller/58003afa48104caa95c47ba1e8ffe7e8`,
result `reports/runtime/slice4be-production-run-local/c3eda01b7f964be485fa3ab1c5bfa2b9/red.json`,
2026-10-01 08:18:20.1918838--08:23:50.5998490 UTC: **178 PASS/325 FAIL/503 checks**.
All503 identities/order and all171 prior PASS checks are preserved. The only seven
changed results are FixedMetadata checks, now GREEN because catalog24 is present;
the binding guard does not claim those catalog changes. All33 existing owner
cases pass. Seven native-close receipts retain four surviving/invoked handlers
and three dismissed surfaces. Five compiles, canonical candidate pins, settings/
saved authority/operator bytes/older records, unassisted cleanup and delayed zero
Excel audit pass. This remains a RED observation/guard suite; it does not replace
the separate unresolved full reusable171/two-aggregate/replay37 acceptance gates.

## Worksheet Apply Scale baseline and binding guard

`-WorksheetScaleOnly` invokes the actual Apply Scale handler in its non-reusable
branch, with Admin-seeded identity, supported Core staging surfaces and the real
released reusable Recipe selector entry. After Seed, the actual Admin Settings
save handler explicitly enables `DesignsEnabled`; the fixture verifies it before
the action gate. No Domain return is replaced and no legacy recipe is imported.

Controller `reports/runtime/production-run-local-controller/a0bf13495d7b48d7af190d3c906d556c`,
result `reports/runtime/slice4be-production-run-worksheet-scale/d4512cf1dd82450cb31c843c662675bc/red.json`,
2026-10-01 08:36:15.1025891--08:37:56.6499880 UTC: **120 PASS/65 FAIL/185 checks**.
Nine cases cover no selection, nonnumeric scale, below/above bounds, unavailable
Design/BOM at ordinary/minimum/maximum scale, missing captured Production sheet,
and that missing sheet with an active decoy Production sheet. Sixty-three failures
are missing observations; two prove owner retargeting to `OtherAddin` and
`OtherWorkbook`. The decoy's custom value/formula remains unchanged in this failure
fixture; unlike Clear's earlier defect, this run does not prove cross-workbook
mutation. All42 shared GREEN checks, five compiles, canonical candidate pins,
saved authority/operator bytes/unknown values/formulas/older records/settings,
normal unassisted cleanup and delayed zero Excel audit pass.

The owner queries `tblDesigns` through the released Design/BOM bridge, while the
selected reusable Recipe is stored separately. Its real read returns no staging
workbook. The existing Sub then returns to the caller, which still scales the
existing named worksheet quantity columns at valid percentages. The probe retains
those partial effects and unknown columns/exact identity. Validation refusals
occur before owner entry, and all nine cases preserve canonical source values.
There is no legacy fallback with the feature enabled. These facts require FAILED
observations for owner-read failures; they neither prove a successful worksheet
load nor authorize changing its algorithm or routing Apply Scale to a different
owner. Positive Design/BOM rebuild and Designs-disabled branch coverage remain
open, including D14 preservation through rebuild.

Unpromoted `deploy/validation-production-run-binding-02` changes only the worksheet
scale helper's binding line to use `modProductionRunBinding.BindWorksheetOwner`
before `LoadRecipeChooser`. This enforces the existing captured-workbook/D18 rule.
Selection/scale validation and prior tree-state clearing keep their original
order; no rollback is claimed. Reusable Scale and existing owner-read/partial
scaling algorithms are unchanged. No new operator wording is introduced.

Controller `reports/runtime/production-run-local-controller/2b040a5f54db4fb3bef42eac843779ba`,
result `reports/runtime/slice4be-production-run-worksheet-scale/8953dc1ae5c04a0d88f969d5250e8fc6/green.json`,
2026-10-01 08:40:28.0542668--08:42:18.6581495 UTC: **122 PASS/63 FAIL/185 checks**.
Both retargeting checks become GREEN (`NotEntered`), all120 prior PASS checks and
all185 identities/order are retained. Five compiles, preservation, normal cleanup
and delayed zero Excel audit pass. The file's GREEN phase name does not make the
overall suite GREEN:63 missing-observation failures remain.

Build `reports/runtime/run-binding-build-02`,2026-10-01
08:39:20.5352629--08:39:59.1074967 UTC, passes five-package generation and independent
cold-start compile with prebuild settings capture/restoration, prior candidate
preservation and unassisted closure. Delayed zero Excel audit completed before
the next gate. Compiled comparison finds only `frmProduction` changed among275
components; all274 others and all string literals are preserved, ignoring
identifier case. `reports/runtime/run-binding-static-02` retains282 components/
6155 procedures/135128 lines,9 literal/45 unresolved calls and190 duplicate
candidates, with all28 caps non-growing. Three schemas,367 tooling scripts and368
repository PowerShell scripts validate/parse.

Two fixture attempts are excluded. `1a9bdd92b7f74459b2c054fd269dd99b`
(`slice4be-production-run-worksheet-scale/0d8ec9f2970d4f71af9a4b342a6fff97/red.json`)
stopped before handler execution at an identifier-case-sensitive probe seam.
`f0473c6186f24a40b64d97b7d9ab766a`
(`slice4be-production-run-worksheet-scale/48c3a4b004d448a5a27275c3cef24340/red.json`)
did not explicitly enable Designs; the value probe raised VBA9 while locating
expected generated staging, and cleanup raised VBA91 after End reset adapter state.
Two owned-dialog End actions assisted recovery. Its five compiles and restored
settings/packages are retained only as diagnostic facts; it is not product RED,
normal-cleanup or acceptance evidence. Both excluded runs have delayed zero Excel
Application audits. No desktop Win32 error5 occurred during these attempts.

Clear/Refresh regression on binding02: controller
`reports/runtime/production-run-local-controller/8f792b2f6cfd430b846ea9008f049da5`,
result `reports/runtime/slice4be-production-run-worksheet-owner/98c50450fb3242e0916aac920fb58fd3/green.json`,
2026-10-01 08:43:04.9983186--08:44:52.4996202 UTC: **94 PASS/42 FAIL/136 checks**,
all prior identities/order/results retained, including the five binding GREEN
checks. Five compiles, candidate pins, saved/local/settings/older-record
preservation, unassisted cleanup and delayed zero Excel audit pass. The42 missing
observations remain open.

Seven-control reusable regression on binding02: controller
`reports/runtime/production-run-local-controller/0b97b78a08fe49ba9c0882853153337b`,
result `reports/runtime/slice4be-production-run-local/1828ae7211fb4539ab8727af724f0728/red.json`,
2026-10-01 08:46:00.8565643--08:51:31.4310662 UTC: **178 PASS/325 FAIL/503 checks**.
All503 identities/order/results match binding01, including all33 owner cases and
all171 historical PASS checks. Seven native-close receipts retain four invoked
handlers and three dismissed surfaces. All42 shared checks, five compiles,
canonical binding02 pins, preservation, normal unassisted closure and delayed
zero Excel Application audit pass. The325 tracking/guard failures remain open;
this is not the separate full reusable171/two-aggregate/replay37 acceptance gate.

## Remaining gates

Preserve the seven worksheet-owner cases and the five binding GREEN checks through
Run handler integration, together with the185 Scale checks and two binding GREEN
checks. Positive worksheet Scale tests must distinguish quiet helper return
from successful owner completion: `mProduction.LoadRecipeChooser` is a Sub,
and the scale branch subsequently scales/prepares local tables. Actual-handler
evidence confirms that helper queries the Design/BOM bridge (`tblDesigns`) for a
selection supplied by the released reusable Recipe selector. D15 requires only
released Process/Recipe projections with Designs enabled; D18's instruction to
preserve the non-reusable branch must not be interpreted as waiving that rule.
This authority mismatch needs explicit reconciliation before owner-routing work.
Do not create a matching legacy Design merely to make the modern Recipe fixture
pass. Record a normative clarification/decision and synchronize Plan022/Controls
before any changed contract; until then preserve the guarded branch and report
its observed failure honestly. D14 preservation through a successful rebuild is
still unproven. Missing fixtures and modal setup failures are not behavioral RED.

Load Recipe itself always calls the released reusable loader; the worksheet
reload helper is reached by non-reusable Apply Scale. Do not invent a separate
worksheet Load-button branch from the helper's name.

Preserve every established presentation baseline identity through GREEN. Add the
other seven owner behaviors and complete
context/policy/failure coverage; then require packaged XLAM, compile, layout,
static maintenance, live roles, full Release1 chain, reusable and independent
Action Path evidence. Agent inspection is not human acceptance.
