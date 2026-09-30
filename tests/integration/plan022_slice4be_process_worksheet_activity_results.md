# Slice4be Process worksheet observations — test-first baseline

Authority: Architecture v4.11 D18's Process worksheet discovered-control
refinement, committed in docs `ba3c206` before these tests; Plan022 and Controls
1.346 synchronized. D14/D15 retain local-save/import behavior and authority.
Catalog22's Core and Operations observations are implemented and focused595 GREEN;
worksheet107, picker115 and independent paths139 pass. Broader regression and
Release1 acceptance remain pending. The original frozen baseline is
`deploy/validation-process-worksheet-picker`, whose scoped gates are complete in
`plan022_slice4be_process_worksheet_picker_results.md`.

## Focused RED

Controller:
`reports/runtime/process-worksheet-activity-controller/d1b17c9d89774c1fad17fdbfb36266bc`.
Result:
`reports/runtime/slice4be-process-worksheet-activity/55e84e2d17b1447cad76687175837592/red.json`.
Run: 2026-09-30,19:21:29.0507595–19:22:56.7300898 UTC.
**62 PASS / 158 expected FAIL / 220 checks**, exact expected failure sequence.
All42 shared prior GREEN checks and five instrumented package compiles pass.

The110 catalog failures cover the missing catalog22 extension, preservation of
all106 earlier control definitions within that absent version, and three missing
metadata definitions. The48 other failures arise from six absent attempt/result
pairs: two Send actions, Add, rejected Retrieve, single-table Retrieve and
two-table Retrieve. Their correlation, context, outcome, reference, redaction,
integrity and terminal assertions cannot pass without those records; this is not
evidence of corrupt records or leaked values. Actual handlers run successfully.

Unchanged packaged behavior passes: Send saves only the captured workbook despite
an active decoy; Add appends its pair and saves; invalid Retrieve leaves both
tables; valid single retrieval removes only its selected table; multi-area
retrieval removes both selected tables. Local actions preserve canonical bytes.
Successful imports preserve Auth/Config bytes and the exact six Inventory
business tables. Unrelated workbook values remain intact. Queue-return observation
captures actual IDs/states in memory: zero for local/rejected actions, one for
single retrieval, two for multi retrieval. No fabricated identity, status-text
parsing, Domain handler replacement or source-value log is used. The unsaved
adapter calls the real form handlers and observes the existing submission return;
it does not intercept an error handler's Err reads.

Settings/package pins are preserved, Excel closes without forced termination,
and the delayed Application1000/1001/1002 audit is clear. Desktop probes report
no error5. One operator capture is directly reviewed and hashed in verification:
the populated Process Designer shows three DRAFT definitions and the actual
"Retrieved 2 selected Process table(s) as DRAFT." status. It proves the visible
business result, not tracking, human acceptance or a completed Action Path.

No runtime VBA, package, catalog registration or static runtime metric changes.
The three changed/new PowerShell test files parse; all347 PowerShell files parse.
Generated fixture/report/image data remains ignored. This is meaningful initial
RED, not a compile/setup failure or an acceptance claim.

## Extended partial-result and guard RED

Controller `reports/runtime/process-worksheet-activity-controller/5e5d13b95ab2498dbb4322400e20d706`,
result `reports/runtime/slice4be-process-worksheet-activity/8c2df34b83594a73aedab379eea2d2d7/red.json`,
2026-09-30,19:29:22.9170273–19:33:04.8811071 UTC:
**119 PASS / 214 expected FAIL / 333 checks**. Every original220 check retains its
exact result and relative order, including all42 shared GREEN. Five compiles,
package/settings preservation, normal closure and delayed Excel audit pass.

Three faults are injected only into the unsaved fixture package, with each fault
verified to occur exactly once. The real handlers and real Designs queue writes
still execute. UncertainSecond suppresses the second actual write's acknowledgment:
one table remains, with exact owner states Submitted/Unknown. RemoveSecond returns
failure before the second removal: one table remains, with Submitted/Submitted.
SaveSecond follows the real second lo.Delete, then takes the existing failure
return before wb.Save: no tables remain in memory, the workbook is unsaved, and
both references are Submitted. This simulates a save refusal; no real disk fault
is claimed. The tests do not fabricate identities, parse report text or use
unhandled Domain Err.Raise/VBE interruption. Later processing may catch up the
earlier queued event; each action's expected references still come only from its
own submission returns. Missing FAILED pairs account for24 new failures.

The32 other new failures prove current worksheet/save guards are absent for all
three actions under Loading, Nested, changed Session and changed Target, and for
Send/Add while SignedOut. Retrieve submits in the first four contexts; signed-out
Retrieve does not submit. Inventory business state remains exact. These are
behavioral failures under D18, not compile or setup failures. No runtime fix yet.

## Earlier closed-workbook fixture failures — excluded

Broader capability/policy tests and form-draft preservation checks are written,
but the full sequence has not reached valid completion. Do not count those test
definitions or partial output as additional accepted coverage.

- Excluded controller `80ed43dbd7464c2c96f1cd26a229604a` under the same controller
  parent,19:36:50.9240298–19:43:57.1351238 UTC, stops at
  TestProductionDesigner.WorksheetActivityGuard after workbook closure. The owned
  Excel host shows a VBA automation dialog,0x80010007, with Continue disabled.
  VBE End/Reset is not used. After verifying the exact owned PID/start identity,
  only that disposable host is forcibly terminated at19:43:43.6633634 UTC so the
  live controller can restore settings. The subsequent RPC error and harness
  failure are not product RED. Package/settings restoration and delayed audit
  pass, but assisted termination excludes the attempt.
- Isolated `-ClosedDiagnostic` controller `0e95014cd1c547b488fbc29e99b5a1a9`,
  19:46:46.9664771–19:48:03.0987213 UTC, result
  `reports/runtime/slice4be-process-worksheet-activity/33dc04141f654c84aff029a08438f634`,
  shows the actual form before closing its captured workbook. An outer adapter
  trap prevents an unavailable form reference from becoming an unhandled VBE
  error. The form adapter is entered and returns normally, outer error0, workbook
  bytes preserved and no activity. Shared42 and five compiles pass; normal
  closure, preservation and delayed audit pass. This is diagnostic-only, not a
  full guard gate or proof of a causal repair.
- Excluded controller `a16832f3a7c3457c9dbf6a493297d435`,
  19:49:03.5524662–19:52:45.0175255 UTC, uses the visible-form setup in the longer
  sequence. The first closed-workbook guard cannot enter its form adapter. The
  defensive entry assertion stops the test as a fixture failure. It closes
  normally without another dialog or forced termination; helper/package/settings
  pins and delayed audit pass. This does not establish that the actual operator
  handler failed, or that the form remained reachable after workbook closure.

All three delayed Application1000/1001/1002 audits have zero Excel events.
Desktop probes remain healthy, with no error5. The HRESULT/fixture failure must
not be confused with the user's conditional desktop-error5 stop instruction.
No native-crash repair is claimed. All347 PowerShell files parse.

## Expanded RED and calibrated closure boundary

Controller `4fd3b34ad15644d4913299cf8a50d6b3`,20:04:35.9981473 through
20:08:14.9290077 UTC on2026-09-30, adds fixed boundary measurements to the
longer sequence. Result `010bd90f9f504e62a8f84c166823c7b5` proves that closing
the captured workbook removes its visible Production window, decreases open
workbooks from4 to3, and leaves the decoy open. Forcing the old form reference
then returns80010007 before adapter entry. This attempt remains excluded; it
closes normally, preserves packages/settings, and has zero delayed Excel events.
The observations resolve the fixture ambiguity, not a runtime defect or its cause.

The test now measures native visibility before invoking a closed-binding handler.
If workbook shutdown dismisses the visible surface, it checks no save, submission
or activity and supplements those facts with the existing typed binding guard.
It explicitly records that no handler was invoked. If the surface remains visible,
the actual form handler must be entered normally and preserve the draft, saved
workbook and activity. The test never treats a disconnected private reference as
an operator click. All adapters remain confined to unsaved disposable packages.

The subsequent full RED is verified:

- Controller `reports/runtime/process-worksheet-activity-controller/2369e088c595488d92511a9ca1b81c0d`.
- Result `reports/runtime/slice4be-process-worksheet-activity/32dc95a4f4e146b595f30f8a6010fee1/red.json`.
- 2026-09-30,20:09:19.5531516 through20:14:41.7703628 UTC.
- 185 PASS /256 expected FAIL /441 checks. Every prior333 result and relative
  order is retained, including all42 shared GREEN checks; five packages compile.
- 108 new checks add66 passing assertions and42 expected failures:10 form-draft
  guard failures,29 permission-denial failures, and3 missing tracking notices.
  The29 include24 dependent assertions for three absent DENIED pairs, three
  worksheet/draft preservation failures and two actual save violations.
- Actual capability loss preserves the captured context in all three cases.
  Send/Add still edit and save; Retrieve changes its form draft but does not save
  the workbook or submit. No failure is a compile/setup substitute for RED.
- Off/older/unavailable tracking permits all three business actions and saves;
  the older policy still reads existing controls. Unavailable notices are absent.
- Closed Send dismisses the native surface; Add/Retrieve stay visible and enter
  their actual handlers with error0. All corresponding preservation checks pass.
  Send's dismissed branch does not claim a click-handler invocation.
- Prior activity, Inventory business state and final Auth/Config bytes remain
  unchanged. Helper/package pins and settings restore; Excel closes normally.
  Delayed Application1000/1001/1002 audit has zero Excel events. No forced cleanup.

`verification.json` in the controller records exact subset/failure-sequence and
preservation checks. The populated Process form capture was reviewed again:
three DRAFT entries, mixed-UOM requirements, output and the two-table retrieval
status are visible. It proves the existing form result, not new activity logging,
human acceptance or a closure screenshot. No runtime implementation has changed.
Desktop monitors report zero error5; the native/COM distinction remains explicit.

## Core catalog/reference GREEN; Operations still RED

The supplemental wire-contract test exercises allowed/unsupported outcomes,
empty versus required references, exact Designs source kind, per-reference
Submitted/Unknown rules, malformed/duplicate identities, wrong warehouse,
missing/extra fields and mixed failure references. It supplements the packaged
form-action tests; its synthetic reference values are schema fixtures only.

- RED controller `c22f97c0363d46f595905d6f7e43d021`, result
  `71ec9c0331c848e0afcca60bd18ed3e0`:147 PASS/38 expected FAIL/185,
  20:19:39.2571491 through20:20:57.0720521 UTC on2026-09-30.
- First Core candidate `deploy/validation-process-worksheet-catalog` passes the
  exact185 checks, controller `5882116b23804cb59cf02309cb19c331`, result
  `c5ef57ec18cb4956aea1b2927a7a39fa`,20:22:56.2666048 through20:24:19.8988031 UTC.
  Static review finds one new duplicate control-definition body; this candidate
  is retained as intermediate evidence, not the final maintenance result.
- Final Core candidate `deploy/validation-process-worksheet-catalog-final` uses
  `modProductionControlCatalog.Command` for shared record fields. Recipe Order
  retains its exact mapping; worksheet captions follow its ordered ControlIds.
  This consolidates record construction and removes the duplicate-body increase.
- Final GREEN controller `744cdf30128e422fb05d9880c7468784`, result
  `8ea0ea156a29416c952e4feb5b751a56`:188/188,20:32:18.3894537 through
  20:33:38.5685170 UTC. Exact prior185 order and all42 shared checks remain GREEN;
  three new assertions protect every field of Recipe Order's version17 definitions.

Controller paths are under `reports/runtime/process-worksheet-activity-controller`;
result paths are under `reports/runtime/slice4be-process-worksheet-activity`.
All three runs compile five instrumented packages, close normally, preserve
packages/settings and have zero delayed Excel Application1000/1001/1002 events.
`reports/runtime/process-worksheet-catalog-final-build` records packaged build,
five compiles, Operations cold start, preservation and delayed audit. Of267
compiled components,261 are unchanged against the picker candidate. Four existing
Core modules change: modActivityCatalog, modActivityReferences, modEvaluationMatches
and modProductionRecipeOrderCodes. Two new Core modules hold worksheet codes and
shared command fields. No Operations runtime component changes in this checkpoint.
Static evidence is `reports/runtime/process-worksheet-catalog-final-static`:
274 components,6120 procedures,134573 lines,9 literal and45 unresolved Application.Run
sites,190 duplicate-body groups,28 non-growing oversized caps. No exception added.

The broader actual-handler RED on the intermediate Core candidate is controller
`95d3be4b762647368230b201ee566a86`, result `26824f27c5b649f0838f61c2771e50a5`,
20:24:35.4178185 through20:30:26.8539744 UTC. It verifies442 PASS/149 expected
FAIL/591: every prior441 identity/order remains, all prior GREEN pass, and110
catalog assertions turn GREEN. All143 wire checks pass. A new seven-check case
signs out immediately after the first real queue return. Exactly one Submitted
reference is observed, session loss is real, and the other target is untouched.
Three expected failures prove a second queue call is attempted, worksheet state
is removed/saved after session loss, and no original REQUESTED record exists yet.
The existing worksheet algorithm may finish the first owning save despite sign-out;
the missing context guard must stop subsequent worksheet removal and submissions.
No fabricated event ID or unhandled VBA error is injected. Final inventory,
Auth/Config and prior activity preservation pass, with normal closure, package/
settings pins, five compiles and a zero-event delayed audit. Desktop error5 stays0.

Current test source adds the three Recipe Order checks to the broad gate, making
its next expected total594; this is not a claim that the full594 have run yet.
Catalog22 now defines109 controls while the three worksheet handlers still emit
no new observations. Production's observed-button coverage remains45/68 until the
handlers and required recording/Action Path gates pass. No release acceptance,
promotion, or completed worksheet slice is claimed.

## Operations focused GREEN and closure-test mapping

Frozen candidate: `deploy/validation-process-worksheet-activity-03`.
Controller `60dfa8b42ffd41ad8e601a5ac281781b`, result
`d9f95fbea535463bb9f05d5fe32d5fcb`, under the same activity controller/result parents,
runs2026-09-30,20:52:59.7781547 through21:01:41.0444761 UTC: **595/595 GREEN**.
All574 non-closure checks from broadRED591 retain exact identities and relative
order; Core188 and all42 shared checks pass. The595 total also includes the
three new Recipe Order metadata checks and the18 stable closure assertions below.
Five instrumented compiles, normal focused closure, package/settings preservation
and delayed Application1000/1001/1002 audit pass with zero Excel events.
`verification.json` in the controller records the subset, closure mapping, compiled
change scope and capture hash. No desktop error5 occurs.

The implementation adds typed `cProductionWorksheetAction`,
`cProductionWorksheetDraft` and `modProductionWorksheetActions` in Operations.
The actual three click handlers delegate to guarded coordination. Initial stale/
loading/nested/permission loss cannot edit, save or submit. Every subsequent read,
submission and removal boundary rechecks the captured context; post-queue sign-out
stops before a second queue call or removal/save. The existing public owning
Process save supplies typed per-event facts, including mixed Submitted/Unknown
states. The class serializes each through Core with the captured warehouse,
retains earlier references on partial failure, and makes optional tracking failure
visible without blocking authorized work. All-before-write validation, draft
restoration on rejection, deterministic imports and existing status/save behavior
remain. Unexpected Add service errors retain their original propagation behavior.
Known worksheet validation refusals are distinguished from exceptions through
optional local service outputs. No canonical schema or authority changes.

Build evidence: `reports/runtime/process-worksheet-activity-build-03` records
five compiles, Operations cold start, preservation and zero delayed Excel events.
Of270 compiled components,264 remain unchanged against the final Core candidate;
only the form, worksheet service and column helper change, plus the three new
typed components. Static evidence `reports/runtime/process-worksheet-activity-static-03`
has277 components/6137 procedures/134745 lines,9 literal and45 unresolved
Application.Run sites,190 duplicate-body groups and28 non-growing oversized caps.
All three schemas validate. The form shrinks11685 to11599 lines and worksheet
service1331 to1325; no new size exception. Existing identity-formatting behavior
moves to the column helper. The populated Process form capture is reviewed:
three DRAFT entries, mixed-UOM rows and successful two-table retrieval are visible.
It does not establish Event Detail/Action Path or human acceptance.

Two implementation failures remain explicit:

- The first candidate (`validation-process-worksheet-activity`, build record
  `process-worksheet-activity-build`) fails Operations compile at the moved
  formatting call: the existing local alternativePairCount hides the function
  of the same name. Renaming the local count fixes the ambiguity. Compile failure
  is not D13 RED; the prior591 behavioral baseline remains authoritative.
- Candidate02 (`validation-process-worksheet-activity-02`) compiles but returns
  531 PASS/64 FAIL/595 in controller `48b04458cab6488da58b64741ea62c7e`, result
  `45d47578c0044ba8b02c4d8efae28dbf`,20:43:07.9628185 through20:50:28.6361682 UTC.
  The reviewed form reports missing definition JSON; valid retrieval never reaches
  its owner. Worksheet reader outputs passed directly to class members do not
  populate the snapshot. Candidate03 retains outputs in local strings and assigns
  the snapshot explicitly. Only modProductionWorksheetActions changes between
  the compiled02 and03 packages. Failed02 closes normally, preserves packages/
  settings and has zero delayed Excel events; its capture is failure evidence.
  Neither failed candidate is acceptance or a native-crash repair.

The provisional594 total assumed one native lifetime. Candidate02 exposes595
because closed Send remains visible, whereas the final run dismisses it.
The test now emits six invariant checks for each closed-book action, with fixed
metadata preserving the actual branch. This replaces17 prior branch-dependent
assertions with18 stable assertions without weakening their protected behaviors:

| Prior assertion(s) | Stable assertion/evidence |
|---|---|
| NativeSurfaceDismissed; visible-before/closed-book/decoy fixture facts | SurfaceLifetimeEstablished plus measured VisibleAfter/HandlerInvoked; a dismissed branch never invokes a disconnected reference |
| NoUnhandledError, ActualHandlerEntered, FormDraftUnchanged | UserActionProtected requires all three when the form remains visible; metadata records each fact; dismissed surfaces are explicitly marked not invoked |
| BindingGuardRejectsClosedBook (previously dismissed Send only) | BindingGuardRejectsClosedBook now runs for all three actions |
| NoWorkbookSave, NoSubmission, NoActivityOrRetarget | Same assertions for all three, with activity counts captured before workbook closure so internal shutdown is covered too |

Final native facts: Send is dismissed and not invoked; Add/Retrieve remain visible,
enter their handlers normally, return error0 and preserve drafts. No prior
non-closure check disappears or changes order. The test never substitutes a
disconnected-object exception for an actual user-handler result.

The existing worksheet-header regression also passes exact107/107 on this candidate:
controller `reports/runtime/process-worksheet-headers-controller/e667ec75af14439d8d8bea89a004aaf6`,
result `reports/runtime/slice4be-process-worksheet-headers/faf3c1810c304048848adf470acf60ac/green.json`,
21:03:53.0900220 through21:06:43.0529513 UTC. All previous107 identities/order,
five compiles, package/settings preservation, normal cleanup and delayed audit
pass. Normalized/custom columns, mixed-UOM imports, confirmed-only removal and
unchanged Auth/Config and Inventory business state remain protected.

The picker regression retains its exact115 prior checks on the same candidate:
controller `reports/runtime/process-worksheet-picker-controller/7a14c9f42d61468e8017ab29426b6a2d`,
result `reports/runtime/slice4be-process-worksheet-picker/e4a3bbcae0114dcb91e2dd1691077691/green.json`,
21:07:43.1911562 through21:09:37.5695177 UTC. Five compiles, package/settings
preservation, normal cleanup and delayed Excel audit pass. Package hashes were
also rechecked after closure. No human acceptance or deployment promotion.

Catalog22 defines109 controls and the three actual worksheet buttons now produce
their observations, bringing observed Production buttons to48/68. Twenty buttons
and30 nonbutton handlers still need coverage review. This is a focused implementation
checkpoint; the full scoped release gates remain open.

## Independent worksheet recordings and Action Paths

The same frozen candidate passes139/139 in a separate recording/publication/UI gate,
21:17:07.3871286 through21:24:11.3278393 UTC on2026-09-30. Controller:
`reports/runtime/process-worksheet-activity-controller/d09bcceecbac47baafaeccfc88195d39`.
Result: `reports/runtime/slice4be-process-worksheet-paths/abf89c5585984810832cb1bac6f65fff/green.json`.
All42 shared prior GREEN identities remain. Five instrumented compiles, package/
settings preservation, normal cleanup and delayed Excel Application audit pass.
No runtime or architectural change accompanies these additional acceptance tests;
the preceding focused behavioral RED/GREEN remains the implementation evidence.

Two independent recordings invoke the original Send, Add, Send and Retrieve handlers.
Each saves two local tables and retrieves both through real owning submissions.
The test checks the added fifth item/SKU pair, exact original request/result pairs,
ordinals, two distinct Designs event references per retrieval, saved local state,
Auth/Config bytes and six Inventory business tables. Admin publication retains all
original records and both complete applied Designs events for each run.

The actual Event Detail distinguishes selected REQUESTED from CONFIRMED, retains
the source identity and shows Unknown effect. Actual expectation-editor/Evaluate
actions establish CommandCompleted for local Send/Add, but SourceEventsApplied
remains Incomplete with empty references. The full four-action sequence concludes
as CommandCompleted in order; a separate SourceEventsApplied evaluation requires
both exact Designs references, Applied owner states, matching line counts/hashes
and no invented inventory keys.

Actual guide creation and save retain the first recording's provenance and explicit
four-step intent. A reader without ACTION_PATH_MAINT pairs that guide with the
second recording and evaluates its exact four occurrences, with zero extras.
How-To, Diagnostic and Compare preserve the same guide/run/evaluation evidence,
original journals, saved authority and custom workbook values. Read-only viewing
does not save the operator workbook. Only the guide's CommandCompleted conclusion
is displayed in the paired capture; it explicitly does not assert application.

Six worksheet captures are reviewed: populated editor, Event Detail, How-To,
Diagnostic, Compare and scrolled conclusion. They show four DRAFT definitions,
the successful two-table import, all four authored instructions, the separate
observed run, both submitted references and four matched occurrences. Their hashes
are in the controller verification record; images remain ignored runtime evidence.
No human acceptance or deployment promotion is inferred.

Static evidence was regenerated in `reports/runtime/process-worksheet-paths-static`:
277 components/6137 procedures/134745 lines,9/45 dynamic calls,190 duplicate groups,
28 non-growing caps, three valid schemas and347 parsed tooling scripts. Runtime
source is unchanged.

The shared harness also retains the exact90 prior Close-path identities/order on
this candidate,21:24:51.4711739 through21:28:59.3442753 UTC. Controller:
`reports/runtime/production-close-paths-controller/587ad91882424be0bc1b63998216a00f`;
result `reports/runtime/slice4be-production-close-paths/d1c5fd3095524b0cbe6875304c035631/green.json`.
Five compiles, normal cleanup, package/settings preservation and delayed audit
pass. Further scoped regressions follow.

## Scoped regression progress

The same `validation-process-worksheet-activity-03` packages retain:

- Packaged smoke86/86 in exact prior order,21:29:46.9850404 through21:30:08.8498235
  UTC. Controller: `reports/runtime/process-worksheet-activity-regression/smoke-6eee0876b82a435aaea78ef84cdcaa54`.
  Initial and final shutdown facts in
  `reports/runtime/packaged-smoke-closure/e7835111f54a4371a3bd44f8b38138d2`
  prove unassisted exits with no termination requested. Packages, settings and
  tracked report bytes are preserved; the delayed Excel audit is clean.
- Layout's exact prior report: three requested sizes, five pages, zero overlaps
  or out-of-bounds controls and passing native window actions,21:30:57.7529481
  through21:31:12.5000012 UTC. Controller:
  `reports/runtime/process-worksheet-activity-regression/layout-8a14c261a69641bd855d8f2b2f00c91a`.
  Normal cleanup, package/settings preservation and delayed audit pass. Three
  reviewed captures show empty Run List layouts, not populated workflow acceptance.
  The minimum request remains clamped to the same1110x800 actual size as before.
- Ordered full chain32/32, live roles48/48 and warehouse creation15/15, all in exact
  prior order,21:31:44.8780854 through21:36:57.2754022 UTC. Controller:
  `reports/runtime/process-worksheet-activity-regression/chain-eaef72bd22cb42f0ba87140ce1a0067a`.
  The normative creation/seed/Receiving/Production/Boxing/Shipping/restart sequence
  completes. Normal cleanup, package/settings and tracked-report restoration,
  and the delayed Excel audit pass.

Full reusable/restart is verified below; remaining shared observation regressions are open.
No runtime edit, rebuild, deployment promotion or human acceptance accompanies
these regression runs.

### Excluded full-reusable attempt and bounded diagnosis

The first uninstrumented full-reusable attempt fails before reusable assertions:
controller `reports/runtime/production-restart-diagnostic/de95d3fa46e14415bd5df54401101a47`,
21:37:31.5556553 through21:37:54.9514544 UTC. The original initial
`mProduction.RunProductionBatchScaleContractTest` callback returns RPC0x800706BE;
Windows records Excel `ntdll.dll/c0000028` at21:37:50.4828873 UTC
(14:37:50.483 Pacific), plus its error-reporting event. This matches the earlier
picker candidate's initial failure signature/location; root cause is unresolved.
The two HARNESS/HARNESS_CLEANUP rows are failures, not behavioral D13 RED or release
acceptance. Excel exits through the crash; final workbook closure is incomplete.
No termination is requested. Parent settings, packages and original validator are
preserved. This is not a desktop error5.

A disposable diagnostic retains the original setup and initial callbacks, then
stops immediately after batch scale. Generation proves original statements intact,
credentials inherited only in memory and no parse errors. It first attaches the
existing redacted NativeExceptionObserver before initial Excel visibility, then
runs separately without an observer:

- `reports/runtime/initial-batchscale-diagnostic/6cf31b1e54de44b78be83385bc7f71bd`,
  21:40:31.6362474 through21:40:54.5446317 UTC: batch scale passes; observer reaches
  Attached/Ready/Exited0 with no captured fault.
- `reports/runtime/initial-batchscale-diagnostic/2ec9102a07b9461eba90e417b9e84c75`,
  21:41:43.3089437 through21:42:05.8498794 UTC: fresh unobserved batch scale passes.

Both diagnostics close normally, preserve settings/packages/validator and have
zero delayed Excel Application failures. Neither runs the full reusable workflow
or changes runtime source; neither establishes a crash repair or full acceptance.
After those bounded comparisons, one full uninstrumented validation with independent
restart replay completes in controller
`reports/runtime/production-restart-diagnostic/cf56200b22f44bdaa683c4b15bf2f131`.
It passes both aggregates and retains all171 Boolean observations in exact prior
order,21:42:29.4928306 through21:53:31.6611378 UTC. Generation preserves original
workflow statements and inherits the creator credential only in memory; no
debugger or VBA instrumentation is installed. Initial/restart-process shutdowns
are normal, with2613/2542ms release waits and no termination requests. Package,
settings, validator and helper hashes are preserved; the delayed Excel audit is clean.

The independent replay at
`reports/runtime/production-restart-replay/eecfe96c873340e189a42ec0f67ca0c1`
retains all37 prior checks in order,21:52:05.6462167 through21:53:31.6243568 UTC.
It reuses the same creator credential through an in-memory SecureString and uses
the public workbench/restart probes to recreate two draft tables from the retained
final fixture state. It does not claim a pristine pre-restart fixture. Both replay
processes close normally, preserve settings/packages and have no delayed Excel
events. The successful full/replay observations satisfy this regression checkpoint;
the first failed attempt stays excluded and no native-crash repair is claimed.

## Current-catalog regression fixture maintenance

The first Close/query rerun retains all312 identities but has307 PASS/5 FAIL,
2026-09-30,21:54:59.5450574 through21:58:29.8264812 UTC. Controller
`reports/runtime/production-close-controller/c4be577d190f4e5eb8c98675dad03dd0`;
result `reports/runtime/slice4be-production-close/ebadd12df483479fb348c69a60943ae9/green.json`.
Four current-record assertions still require catalog21. The declared catalog20
policy fixture removes only Close, retaining the three catalog22 controls; the
existing policy validator correctly rejects that inconsistent fixture. These
are stale test expectations, not behavioral RED or a reason to change runtime.
Five compiles, normal cleanup, settings/package preservation and delayed audit
pass. The failed run stays excluded from acceptance.

Under the already approved catalog22 contract, current-record expectations in
Close and twelve other regression fixtures advance21 to22; the current Settings
editor expects109 controls instead of106. The Close historical fixture now keeps
exactly its version20 IDs. Explicit historical definition/terminal tests remain
unchanged. Check identities and other assertions are retained. This is test
maintenance, with no runtime/package or normative contract change; the affected
broader suites still require their own packaged GREEN evidence.

The unchanged candidate then passes312/312 in the same order as the prior gate,
21:59:22.4312537 through22:02:58.3687809 UTC. Controller
`reports/runtime/production-close-controller/1aa0bde812fd4eeb8e9d878bf2e8a3dd`;
result `reports/runtime/slice4be-production-close/2d1ca0d80b6a41c49d967f7cc7cb1ee9/green.json`.
Five compiles, all80 query checks, all42 shared checks, preserved settings/packages,
normal closure and zero delayed Excel Application failures pass. Four images are
directly reviewed and hashed in `visible-review.json`: pre-dismissal button/native
views, public reopening and separate saved-configuration/tracking-unavailable
feedback. Disposal is established by the actual-handler/window checks, not images.

Regenerated `reports/runtime/process-worksheet-regression-static` retains277
components,6137 procedures,134745 lines,9 literal/45 unresolved Application.Run
calls,190 duplicate-body groups and28 non-growing oversized caps. All three
schemas and347 PowerShell parses pass. No runtime source change, deployment,
native-crash repair, desktop-error5 or human-acceptance claim accompanies this work.

## Settings regression GREEN on catalog22

Settings retains202/202 exact prior checks and five compiles,
2026-09-30,22:03:24.7362664 through22:07:29.0830101 UTC. Controller
`reports/runtime/process-worksheet-activity-regression/settings-97fb0ea590864012a0fe727e021a7d31`;
result `reports/runtime/slice4be-tracking-settings/1f7571b0dd894279ac8deeabf29458be/green.json`.
The109-control editor, cancelled-save preservation, version checks, all three
presentation choices, saved-preference restart and Operations without Admin pass.
Restart releases all ten COM references with zero failures and exits unassisted.
Normal final closure, settings/package preservation and delayed Excel audit pass.

Settings activity completes497/497 with five compiles,
22:07:53.1855691 through22:22:32.7469702 UTC. Controller
`reports/runtime/process-worksheet-activity-regression/settingsactivity-39af05a7d7dd4dfab1b94b950016e5df`;
result `reports/runtime/slice4be-settings-activity/588ff66082c44bf0b5b1710c2f6823d0/green.json`.
All494 prior checks retain their relative order. The only additions are
`SettingsActivity.OlderPolicy.Excludes.PRODUCTION_PROCESS_WORKSHEET_SEND`,
`SettingsActivity.OlderPolicy.Excludes.PRODUCTION_PROCESS_WORKSHEET_ADD_ITEM` and
`SettingsActivity.OlderPolicy.Excludes.PRODUCTION_PROCESS_WORKSHEET_RETRIEVE`.
The existing catalog enumeration supplies these checks; the isolated verifier
requires precisely this addition set instead of accepting arbitrary extra checks.

Eight captures are directly reviewed and hashed in the result's
`visible-review.json`: tracking, detail, personal and Operations preferences,
denied reload, failed profile read, and separate owner-save/tracking-unavailable
feedback. They distinguish staged choices from effective saved choices. After
assertions finish, Excel remains alive without a main window, then exits normally
before the controller completes. No extra Quit, termination, debugger or runtime
change is used. Settings/packages restore, Excel closes, and the delayed
Application1000/1001/1002 audit is clean. No desktop error5 is observed. This
establishes these regression gates, not deployment promotion or human acceptance.

## Production lifecycle regression GREEN

The unchanged worksheet candidate retains615/615 exact ordered lifecycle checks,
2026-09-30,22:23:25.3762313 through22:29:27.7743770 UTC. Outer controller
`reports/runtime/process-worksheet-activity-regression/lifecycle-7a753ab4f08e410ba8d1314844ae3c47`;
inner `reports/runtime/production-lifecycle-controller/67747839709543f3962030d9928905b1`;
result `reports/runtime/slice4be-production-designer/8cdc858926114f3c8b46ca25bba6cd33/green.json`.
Five compiles, current-owner event correlation, guarded submissions, uncertain
failure facts, optional tracking, reentry suppression, prior records and unknown
columns retain their protecting checks. Both controllers confirm normal cleanup
and preservation; all five frozen package hashes are rechecked. The delayed
Excel Application audit is clean. No runtime edit or human acceptance is claimed.

Draft/Action Paths retains390/390 exact prior checks,
22:29:54.0641636 through22:38:46.5576510 UTC. Outer controller
`reports/runtime/process-worksheet-activity-regression/draft-a476a62a264d41278f31e577cf511e42`;
inner `reports/runtime/production-designer-controller/c246540cc5b9410a8488be874b1bb36a`;
result `reports/runtime/slice4be-production-designer/24c2e6893e384fa19825150853a08386/green.json`.
The original eight-action recording, exact owner matches, explicit authored
intent and local-only terminal classification pass. REQUESTED/rejected outcomes
and absent Domain evidence do not establish application or completion. Prior
activity/journals, saved authority and unknown workbook columns are preserved.
Five compiles, normal closure, package/settings preservation and the delayed
Excel Application audit pass. No runtime or package changes accompany this gate.

Native cancellation retains94/94 exact prior checks,
22:39:23.7178279 through22:41:14.8561910 UTC. Outer controller
`reports/runtime/process-worksheet-activity-regression/native-a0382739c0ac456582041e1760ff484c`;
inner `reports/runtime/production-lifecycle-native-controller/53897cf634d641dbab3678220cc43b5d`;
result `reports/runtime/slice4be-production-lifecycle-native/afcb537af3bf406188962cb8e831a214/green.json`.
The four original Process/Recipe Release/Obsolete dialogs default to No; actual
native No delivery preserves local drafts, owner source, captured workbook and
unknown values, producing only correlated REQUESTED/CANCELLED records without
source references. Five compiles, preservation, normal closure and delayed audit
pass. Five images are reviewed and hashed in `visible-review.json`: four original
questions and the separate Settings owner-save/tracking-unavailable message.
Images precede cancellation; native-input and state assertions prove its result.
No desktop error5, runtime change, promotion or human acceptance is claimed.

## Regulation regression GREEN

Regulation retains696/696 exact prior checks,
2026-09-30,22:41:59.0055678 through22:45:16.2641689 UTC. Controller
`reports/runtime/production-regulation-controller/bdd8a07200784ee1a3ccb622cd705caa`;
result `reports/runtime/slice4be-production-regulation/50c93044d24a49d19d8d82f255ccbc2c/green.json`.
Both Process defaults and Recipe overrides retain local editing, validation,
partial-failure, captured-context, permission and optional-tracking checks.
Five compiles, saved-authority/unknown-column/prior-record preservation, package
pins, settings restoration, normal closure and delayed Excel audit pass. Two
principal captures are reviewed and hashed in `visible-review.json`: Process
Apply remains local staged regulation, and Recipe Clear retains its original
cleared status. No runtime, package or contract change accompanies this gate.

Regulation paired paths retain102/102 exact prior checks,
22:45:49.7292902 through22:50:50.4017103 UTC. Controller
`reports/runtime/production-regulation-controller/85c8ef2b7f6e41b786f7fa30652c7ec4`;
result `reports/runtime/slice4be-production-regulation-paths/2468e9365be44aa0951dc5e9e238b325/green.json`.
Separate four-action recordings retain original REQUESTED/STAGED pairs through
Admin publication, exact Event Detail, authored expectations, independent pairing
and How-To/Diagnostic/Compare. All four steps match in order with zero extras;
CommandCompleted concludes without asserting Domain application. SourceEventsApplied
remains incomplete for these local actions. Original activity/journals, saved
authority and unknown workbook columns/bytes remain unchanged. Five compiles,
normal closure, settings/package preservation and delayed Excel audit pass.
Six principal images are reviewed and hashed in `visible-review.json`; the
scrolled conclusion explicitly shows all four matches and no application claim.

## Designer load/refresh regression GREEN

Designer reads retain621/621 exact prior checks,
2026-09-30,22:51:35.6763402 through22:54:47.9699304 UTC. Controller
`reports/runtime/production-design-read-controller/ccad9f82f6d949d19a2387725a366b0c`;
result `reports/runtime/slice4be-production-design-reads/1409f8347c6f480fb1be940495c5faa1/green.json`.
Process Refresh/View/Reuse and Recipe Refresh/Load retain their existing local
draft behavior, malformed/unavailable-read handling, partial-failure facts,
permission/context guards and optional-tracking assertions. Prior records,
saved authority and unknown workbook values remain protected. Five compiles,
normal cleanup, settings/package preservation and delayed Excel audit pass.
No runtime/package changes or broader acceptance claim accompany this regression.

Designer-read paired paths retain114/114 exact prior checks,
22:55:22.0243887 through23:01:44.8881610 UTC. Controller
`reports/runtime/production-design-read-controller/c32e812cd09e4f80bf4a1b17049b8294`;
result `reports/runtime/slice4be-production-design-read-paths/fed0ed1170c54ab8a116c02f0d6ef99e/green.json`.
Separate five-action recordings preserve their original pairs through Admin
publication, exact Event Detail, explicit guide intent and independent evaluation.
How-To, Diagnostic and Compare retain the same immutable evidence. All five
REFRESHED/PRESENTED/STAGED steps match in order, zero extras; the conclusion is
local command completion, not Domain application. Five compiles, preservation,
normal closure and delayed audit pass. Six principal captures are reviewed and
hashed in `visible-review.json`; selected Load/PRESENTED remains Info/Unchanged,
and the scrolled comparison shows all five matches and the local-only conclusion.

## Recipe structure regression GREEN

Recipe structure retains786/786 exact prior checks,
2026-09-30,23:02:32.6843099 through23:05:59.3018715 UTC. Controller
`reports/runtime/production-recipe-structure-controller/5db51a54a62445798dce08619c4e3091`;
result `reports/runtime/slice4be-production-recipe-structure/21ab47a9d9d3427cb6fe655e00d555f5/green.json`.
Five compiles, captured-context/permission and optional-tracking checks, saved
authority, prior activity and unknown workbook values remain protected. All five
frozen package hashes, restored settings, normal cleanup and the delayed Excel
Application audit pass. Two principal captures are reviewed and hashed in
`visible-review.json`: Add Process creates a local Recipe node from the selected
released Process, and Connect retains its original connection-staged status.
No runtime, package or contract changes accompany this regression. Images do not
establish Domain application or human acceptance.

Released-data regression retains55/55 exact prior checks,
23:06:30.3769936 through23:08:16.2269253 UTC. Controller
`reports/runtime/production-recipe-structure-controller/9bdb0f40392840e7b04933a415c81590`;
result `reports/runtime/slice4be-production-recipe-structure-released/0712ee27f35749daad5fb02dfc3991ac/green.json`.
The actual Update handler writes all seven validated fields against released
Process records, including changed quantity/percentage, and preserves routing,
UOM, nodes, instructions, saved Domain authority and unknown workbook columns/bytes.
Five compiles, normal closure, settings/package preservation and delayed audit
pass. The populated released-data capture is reviewed and hashed: edited required
quantity4/percentage75/LB and the original staged status are visible. Output Flow
still shows produced quantity5/yield100; the field assertions, rather than this
image alone, establish the changed local connection row. No application or human
acceptance is inferred.

Structure paired paths retain109/109 exact prior checks,
23:08:46.9957171 through23:14:26.5324083 UTC. Controller
`reports/runtime/production-recipe-structure-controller/98bf6fb1445844028c57e63906c74a1d`;
result `reports/runtime/slice4be-production-recipe-structure-paths/9c3729c9b2764f598cbfa2f11d7b5b70/green.json`.
Separate five-action recordings reach the expected local draft, retain original
REQUESTED/STAGED pairs through Admin publication and exact Event Detail, and
support explicit guide intent and an independent reader's comparison. All five
steps match in order with zero extras. CommandCompleted concludes locally;
SourceEventsApplied remains incomplete without owning application evidence.
Five compiles, immutable records/journals, saved authority, unknown workbook
values/bytes, normal closure, settings/package preservation and delayed audit pass.
Six principal images are reviewed and hashed in `visible-review.json`. Selected
Remove Process remains STAGED/Info/Unchanged; the scrolled comparison states that
Domain application is not asserted. The Recipe form retains the original Add
status after Remove; that old notice is not used as proof of the final draft.

## Recipe ordering regression GREEN

Recipe ordering retains463/463 exact prior checks,
2026-09-30,23:14:59.7114925 through23:17:48.4480919 UTC. Controller
`reports/runtime/production-recipe-order-controller/7330293af92e4b748841453f32341a81`;
result `reports/runtime/slice4be-production-recipe-order/cd7ba1f807b8408e84a2ff14ced5df27/green.json`.
Existing row order, fields following node identity, renumbering, selection,
connections/instructions, no-movement/refusal, partial failures, guards and
optional-tracking cases retain their assertions. Five compiles, saved authority,
prior activity, unknown columns, normal cleanup, settings/package preservation
and delayed Excel audit pass. Two captures are reviewed and hashed in
`visible-review.json`: Auto Order shows the reordered nodes and original updated
status; Move Down shows the moved selection but retains the earlier injected
failure notice. That second image is row-layout evidence only, not a current
success notice. Actual-handler row/event assertions establish the result.

Ordering paired paths retain93/93 exact prior checks,
23:18:17.1798904 through23:22:39.9108759 UTC. Controller
`reports/runtime/production-recipe-order-controller/b2fdb5ced1e74421b331733af711bb00`;
result `reports/runtime/slice4be-production-recipe-order-paths/0fa9ce8159dc4d429ea996abeff62bb9/green.json`.
Separate three-action recordings retain original pairs through Admin publication,
exact Event Detail, explicit guide intent and independent-reader evaluation.
How-To, Diagnostic and Compare use the same immutable evidence. All three STAGED
steps match in order with zero extras; CommandCompleted concludes locally, while
empty source references cannot establish SourceEventsApplied. Five compiles,
normal cleanup, settings/package preservation, prior records/journals, saved
authority, unknown workbook values/bytes and delayed audit pass. Six principal
captures are reviewed and hashed in `visible-review.json`; Auto Order remains
STAGED/Info/Unchanged, and the scrolled comparison explicitly excludes a Domain
application claim. No runtime/package changes or human acceptance are implied.

## Process component regression GREEN

Component editing retains795/795 exact prior checks,
2026-09-30,23:23:42.7044613 through23:26:59.6233227 UTC. Controller
`reports/runtime/production-component-controller/76e6d67291bb47d0b10fd6557782587d`;
result `reports/runtime/slice4be-production-components/db51a9d6aaa747b2bccd909c928391ff/green.json`.
Requirement/output editing, partial-failure uncertainty, captured-context and
permission guards, optional tracking and prior-record preservation retain their
protecting assertions. Five compiles, normal cleanup, saved authority, unknown
workbook values/bytes, settings/package preservation and delayed Excel audit pass.
Two principal captures are reviewed and hashed in `visible-review.json`: local
requirement/output quantity and basis edits are visible alongside the explicit
unchanged-saved-definitions status. No runtime/package/contract changes or human
acceptance accompany this regression.

Component paired paths retain142/142 exact prior checks,
23:27:29.3820554 through23:33:20.7072768 UTC. Controller
`reports/runtime/production-component-controller/bdbaf550b78e4b9f812b09867e30566c`;
result `reports/runtime/slice4be-production-components/94428c02c49646f0850f46e3f0e58490/green.json`.
Separate ten-action recordings retain original pairs through publication, exact
Event Detail, explicit guide intent and independent evaluation. All ten STAGED
steps match with zero extras; CommandCompleted concludes locally and empty source
references cannot prove SourceEventsApplied. Five compiles, normal cleanup,
preservation of saved authority, prior records/journals, unknown workbook values
and bytes, settings/package pins and delayed audit pass. Six principal images
are reviewed and hashed in `visible-review.json`: the guide shows all ten actions,
Output Remove remains STAGED/Info/Unchanged, and the scrolled lower comparison
shows matches4-10 and zero extras. The first three matches and conclusion text
are outside that viewport; assertions establish all ten matches and local-only
completion. This capture limitation is not represented as full visible acceptance.

## Next required work

Preserve the595 focused checks, documented closure mapping and exact107/115
regressions plus independent paths139, Close paths90, smoke86, layout and full
chain32/live48/Create15, full reusable171/replay37, Close/query312, Settings202,
Settings activity497 (494 retained plus three exclusions), lifecycle615,
draft/paths390, native cancellation94, Regulation696/paths102, design reads621/paths114
and Recipe structure786/released55/paths109, Recipe ordering463/paths93,
components795/paths142.
Complete the remaining shared-observation regressions on the frozen candidate.
SourceEventsApplied must use every exact owning published Designs reference;
queue acknowledgment and
CommandCompleted remain distinct from applied evidence. No package promotion or
human acceptance is implied by focused595, worksheet107, picker115 or paths139.
The current DeleteProcessWorksheetTable performs lo.Delete before wb.Save. A
post-confirmation workbook-save failure can leave a local deletion in memory;
the failure observation must retain the confirmed Designs reference and Unknown
effect, not claim restoration. D15 still prohibits removal before confirmed
Designs draft save. The normative observation section explicitly distinguishes
these cases without changing the existing save algorithm.

Remaining shared-observation regressions and their visible
operator evidence remain outstanding. No promotion or Slice4be completion is claimed.
