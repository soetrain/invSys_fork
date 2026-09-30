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

## Next required work

Preserve the595 focused checks, documented closure mapping and exact107/115
regressions plus independent paths139, Close paths90, smoke86, layout and full
chain32/live48/Create15 and full reusable171/replay37. Complete the remaining
shared-observation regressions on the frozen candidate. SourceEventsApplied
must use every exact owning published Designs reference; queue acknowledgment and
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
