# Slice 4be Production output-regulation observations

Last verified:2026-09-30 UTC. Incomplete; no operational promotion or human
acceptance. Architecture v4.11 D15/D18 specifies catalog20 under approved semantic
inheritance. Documentation commit3d7d603 precedes tests and runtime changes;
Plan022, controls1.305 and Production coverage1.48 carry the same contract.

The bounded group is Apply Regulation and Clear Override in Production Settings.
Both observe existing local Process-default/Recipe-override staging. Only STAGED
can establish CommandCompleted; no Save/Release or Domain application is implied.
Preserve existing conversion, validation, clear, refresh and status semantics.

## Packaged test-first RED

`Test-Slice4beProductionRegulation.ps1 -DeployRoot deploy/validation-production-design-reads-final -Phase RED`
invokes the actual `mBtnOutputRegulationApply_Click` and
`mBtnOutputRegulationClear_Click` through unsaved typed adapters. Its real released
Process fixture is created before saved-authority pins are captured. Both local
scopes, exact Recipe binding, unrelated outputs/nodes and unknown workbook columns
are protected. Test adapters return fixed failure codes rather than raw errors.

RED:168 PASS/528 FAIL across696 unique checks,
10:19:20.1996933--10:21:49.2485182 UTC. Controller:
`reports/runtime/production-regulation-controller/6bb456b66ff849ecaacc7c133ad354f2`.
Result: `reports/runtime/slice4be-production-regulation/5bbfd42a206d44b3834e101ed60d990b/red.json`.
Five instrumented compiles pass; there are no harness failures. Settings and
package pins restore, runtime source matchesa2aba6a, Excel closes unassisted and
the delayed Application1000/1001/1002 audit finds zero Excel failures.

Every selected local effect, other-output/node/scope preservation, exact released
binding, refresh/editor reset and numeric/disabled-bound semantic check passes.
Existing blank/nonnumeric conversion exceptions are observed without mutation;
their new handled/fixed failure observations remain RED. Partial staging is not
rolled back. Saved-authority, workbook-byte and prior-activity pins pass.

Failures cover the absent catalog20/observations, current permission and captured
context guards, loading/reentry suppression, fixed exception handling and visible
optional-tracking failure. The catalog-preservation assertions are RED because
catalog20 is absent; they require all103 prior definitions to remain exact when
that catalog exists. Missing compilation, a broken fixture or an unavailable
workbook is not counted as this behavioral RED.

## Focused implementation checkpoint

Catalog20 registers105 global controls and44/68 constructed Production controls;
24 remain unregistered. Typed Operations guards precede the unchanged local
editing algorithms. Core supplies fixed, versioned control/outcome facts and
recognizes only STAGED as local command completion. No canonical mutation or
rollback is asserted. Three existing regulation storage helpers move intact into
the typed module; the form shrinks from11685 to11681 lines.

The first candidate `deploy/validation-production-regulation` records695 PASS/
one redaction-check failure,10:29:00.3510432--10:32:25.0971451 UTC, controller
`reports/runtime/production-regulation-controller/126ab042b28f42f0a3d3e66557d973e2`.
Result: `reports/runtime/slice4be-production-regulation/7f563c11398d4b85b060fab724437db9/green.json`.
The failed identity is `Regulation.CLEAR.Recipe.Success.NoEnteredDataOrSources`.
Disposable raw records were deleted by normal cleanup, so the exact matching
field cannot be recovered. All behavior/preservation checks pass; this attempt
is retained as failed evidence.

The test's unstructured substring predicate demonstrably confuses short business
IDs with generated identifiers, hashes and timestamps, and misses JSON-escaped
sensitive text. A separate deterministic calibration records6 PASS/4 FAIL then
10/10 GREEN, under `reports/runtime/production-regulation-redaction/`:
RED `3c77c0db3d6b44ed949a0b9f4c6f6c90`, GREEN `c41ad2db6b51480fb24bd9af7ff22c4d`.
The corrected predicate decodes JSON and checks nested values. Only named
technical fields with complete valid syntax are exempt from short-ID substring
matching; sensitive canaries remain forbidden everywhere. Actual identity values,
identities in messages/guidance, nested leaks and escaped sensitive text fail.
No runtime record format changes to accommodate the test.

Final candidate: `deploy/validation-production-regulation-final`. Build evidence
`reports/runtime/production-regulation-final-build`,10:34:12.2421732--10:34:50.6839940
UTC, records five builds/compiles, Operations cold load,262 compiled components,
and changes only to the five intended components. Settings/frozen packages are
preserved; normal exit and delayed Application1000/1001/1002 audit pass.

Focused GREEN:696/696,10:35:00.0213553--10:38:29.4134107 UTC. Controller
`reports/runtime/production-regulation-controller/32dd4debc7d64d31b973b507e6e58a2f`;
result `reports/runtime/slice4be-production-regulation/0c66683fd7ab46d8b5a9070ae1d34ada/green.json`.
All696 ordered RED check identities remain. Five instrumented compiles,
saved-authority/workbook/custom-column/prior-activity preservation, settings and
five package pins pass; Excel exits unassisted with zero delayed Application
failures. Two directly reviewed principal captures show Process Apply and Recipe
Clear on Production Settings, with original status wording. Their hashed review
receipt is `visible-review.json` in the result root. Agent review is not human
acceptance; images and runtime records remain ignored.

Static evidence: `reports/runtime/production-regulation-static`,269 components,
6107 procedures,134347 lines. Literal/unresolved Application.Run remain9/45,
duplicate-body candidates remain190; all28 oversized-module caps do not grow.
Three schemas and330 PowerShell parses pass. No dead code is removed.

## Paired-path test correction

The first paired run records100 PASS/two test-assertion failures across102 checks,
10:38:41.0837532--10:43:33.0649978 UTC. Controller
`reports/runtime/production-regulation-controller/ee677e719a454adaa52b15ff14261c56`;
result `reports/runtime/slice4be-production-regulation-paths/d6ec8e554fb545689825fce1a1c09c5c/green.json`.
The single-step terminal assertions incorrectly expect the final Clear action,
although Clear/STAGED appears twice. Read-only inspection before fixture cleanup
confirms the existing evaluator correctly selects the first eligible occurrence,
ordinal2, leaving three extras: CommandCompleted concludes, SourceEventsApplied
remains incomplete/SOURCE_UNAVAILABLE. This is a test defect, not product RED.
The corrected gate authors all four expected steps and selects the fourth as
terminal; it requires all four exact observed identities and zero extras. No
runtime/evaluation contract changes. The separate four-step guide evaluation and
all three presentation methods already pass in this first execution.

Corrected paired-path GREEN:102/102,10:44:01.1727705--10:49:08.3675576 UTC,
controller `reports/runtime/production-regulation-controller/028961e6442c4db7b3211306e22219ea`;
result `reports/runtime/slice4be-production-regulation-paths/fa1d0d5165ec48dcb1b60363c0f3165f/green.json`.
All102 prior ordered identities remain; five instrumented compiles, immutable
source/journal/guide evidence, canonical files, workbook bytes/custom columns,
settings/package pins, normal closure and delayed Application audit pass.
Two independent four-action recordings use the actual Apply/Clear handlers in
both scopes; each reaches the expected local draft. Admin publication preserves
the original records. Explicit four-step expectations match all four observed
actions with zero extras: CommandCompleted concludes, SourceEventsApplied remains
incomplete. Reader-side How-To, Diagnostic and Compare both preserve the same
immutable evidence and separate guide provenance from the observed run.

Six principal captures are directly reviewed with hashes in `visible-review.json`.
Clear/STAGED detail visibly shows Info/Unchanged; How-To shows all four authored
instructions. The scrolled diagnostic conclusion shows four STAGED matches,
zero extra observations and "Command completed; Domain application not asserted".
This is agent-visible evidence, not human acceptance. Runtime source is3559e2d;
the paired-test correction did not change or rebuild the frozen packages.

## Final-candidate regressions

All roots below are under `reports/runtime/production-regulation-regression/`.
These executions use the same five frozen catalog20 package hashes. Settings
restore, tracked reports restore where applicable, and delayed
Application1000/1001/1002 audits find zero Excel failures.

| Gate | Result | Controller | UTC interval |
|---|---|---|---|
| Ordinary full reusable Production | Two aggregates and all171 prior ordered boolean observations pass | `reusablefull-300b8da304b040e7a6764a57b75cba1b` | 10:49:41.6915279--10:58:01.1642066 |
| Packaged smoke | Exact86 prior checks pass | `smoke-71dd9816d894493e972cfad2f9cf9a59` | 10:58:22.1110436--10:58:44.6051153 |
| Production layout | Exact prior geometry at three sizes/five pages; zero bounds/overlap violations | `layout-b4f12a8c7e3445feb6995753e1ee0663` | 10:59:07.2226206--10:59:21.8566631 |
| Full Release1 chain/live roles/Create Warehouse | Exact32/48/15 prior checks pass | `chain-4af55c8b76fd470f88df4ea61da2fb67` | 10:59:46.0281049--11:04:55.9037997 |

Full reusable Production uses the ordinary uninstrumented launcher gate, not the
native-observer diagnostic. Both restart and final Excel exits are unassisted,
with zero release failures or termination requests. Restart closes four packages
in reverse order; final closure completes three workbooks/four packages with none
remaining. All171 observations retain the prior full GREEN identities and order.
This establishes the current candidate's ordinary gate; it does not diagnose or
repair the earlier intermittent native failure.

Smoke retains normal Initial/Final exits, no termination requests and zero
release failures; shutdown root:
`reports/runtime/packaged-smoke-closure/4448f3661f214f4eaf5009a804690ed1`.
All three layout captures are directly reviewed with hashes in the controller's
`visible-review.json`; minimum clamping, footer reachability and scrolling retain
the accepted geometry. Agent review is not human acceptance.
The ordinary full chain retains the normative phase order, live-role and warehouse
creation results, normal closure and all prior check identities. No repeat or
native observer is needed for this catalog20 execution; the earlier native
failures remain retained without a repair claim.

Designer read regressions on the same catalog20 packages retain621/621 focused
checks and114/114 paired-path checks, with exact prior ordered identities, five
instrumented compiles each, preserved settings/packages/authority/workbook bytes,
normal closure and zero delayed Excel Application failures. Controller roots:

- `reports/runtime/production-design-read-controller/ecdfe475563b4f6ab986fbce35e30ff4`,
  11:05:24.0031181--11:08:43.0900451 UTC; result
  `reports/runtime/slice4be-production-design-reads/624b803e36864619b8605592fd9f5e48/green.json`.
- `reports/runtime/production-design-read-controller/f58a7984cdc84d0e8e74170ff208c229`,
  11:09:12.9208814--11:15:45.4737167 UTC; result
  `reports/runtime/slice4be-production-design-read-paths/271427243fee4d94925f4555814c090e/green.json`.

Six principal path captures are directly reviewed, with hashes in the result's
`visible-review.json`: Load/PRESENTED detail remains Info/Unchanged; all five
instructions and matched REFRESHED/PRESENTED/STAGED steps are visible, with zero
extras and an explicitly local conclusion. This remains agent review, not human
acceptance. No runtime edits or package rebuilds were needed.

Recipe structure regressions retain 786/786 focused checks, 55/55 released-data
checks and 109/109 paired-path checks, each with exact prior ordered identities,
five instrumented compiles, preserved settings/packages/authority/workbook bytes,
normal closure and zero delayed Excel Application failures. All intervals below
are 2026-09-30 UTC; controller roots are under
`reports/runtime/production-recipe-structure-controller/`.

| Gate | Controller | UTC interval | Result root under `reports/runtime/` |
|---|---|---|---|
| Focused 786/786 | `e648daef96c849b2ad7b11b727efdac0` | 11:18:51.2687239--11:22:25.3538702 | `slice4be-production-recipe-structure/41b1dcbcd6574bb28c01adfd7fea6594` |
| Released data 55/55 | `e50c060a50974b6fb89f8e13c477529a` | 11:22:36.8786387--11:24:25.3338116 | `slice4be-production-recipe-structure-released/5608f34135814459b917a5a070b10add` |
| Paired paths 109/109 | `8322c5484a394912a00cce861834534c` | 11:24:33.9481377--11:30:22.3694440 | `slice4be-production-recipe-structure-paths/c3785a06df274b7bbf26fc2b82891eb4` |

Six principal path captures are directly reviewed and hashed in `visible-review.json`.
Remove Process/STAGED detail remains Info/Unchanged. All five instructions are
readable; the separate observed run matches five STAGED actions with zero extras.
The scrolled conclusion explicitly states that Domain application is not asserted.
Agent review is not human acceptance; no runtime or package change was needed.

Recipe ordering regressions retain 463/463 focused and 93/93 paired-path checks,
with exact prior ordered identities, five compiles each, preservation, normal
closure and zero delayed Excel Application failures. Controller roots are under
`reports/runtime/production-recipe-order-controller/`; intervals are September30 UTC.

| Gate | Controller | UTC interval | Result root under `reports/runtime/` |
|---|---|---|---|
| Focused 463/463 | `7907566dad204d0d9697f39e16168ef3` | 11:30:30.4492520--11:33:22.1782285 | `slice4be-production-recipe-order/cca361f418624980a9d29790f0a0ceca` |
| Paired paths 93/93 | `f268c9bbba9247d388ba99fa150ba639` | 11:33:30.7917875--11:37:58.1109595 | `slice4be-production-recipe-order-paths/fa07f7bd0e374ab4b85ce32ac9d2cb86` |

Six principal captures are directly reviewed and hashed in the path result's
`visible-review.json`. Auto Order/STAGED detail remains Info/Unchanged; How-To
shows three instructions, and the separate observed run matches all three STAGED
steps with zero extras. The visible conclusion does not assert Domain application.
This is agent review, not human acceptance; runtime and frozen packages are unchanged.

Component regressions retain 795/795 focused and 142/142 paired-path checks,
with exact prior ordered identities, five compiles each, preservation, normal
closure and zero delayed Excel Application failures. Controller roots are under
`reports/runtime/production-component-controller/`; intervals are September30 UTC.

| Gate | Controller | UTC interval | Result root under `reports/runtime/` |
|---|---|---|---|
| Focused 795/795 | `28a05cdac7a641cc9e4b64af19a5e202` | 11:38:08.6320910--11:41:29.1662136 | `slice4be-production-components/d4e1f399add84cbbbddf76358cb04101` |
| Paired paths 142/142 | `3a405317b651484fb9b463f90e811a6f` | 11:41:37.9542385--11:47:35.2561154 | `slice4be-production-components/5efefea76d7040c698e448bcd28bcf39` |

Six principal captures are directly reviewed and hashed in `visible-review.json`.
Output Remove/STAGED detail remains Info/Unchanged; all ten authored instructions
are readable. The final diagnostic viewport shows matches4-10, zero extras and
all ten matched identity pairs. The conclusion heading and matches1-3 are above
that viewport; exact text assertions protect the full local-only conclusion.
This is agent review, not human acceptance. No runtime or package changes.

Instruction regressions retain 411/411 focused and 105/105 paired-path checks,
with exact prior ordered identities, five compiles each, preservation, normal
closure and zero delayed Excel Application failures. Controller roots are under
`reports/runtime/production-instruction-controller/`; intervals are September30 UTC.

| Gate | Controller | UTC interval | Result root under `reports/runtime/` |
|---|---|---|---|
| Focused 411/411 | `009b1975b166449382c28b2c83c7a37d` | 11:47:43.5194618--11:50:22.9155603 | `slice4be-production-instructions/8382f6d58bd74dc2a3c75b36c3d360aa` |
| Paired paths 105/105 | `5496d25bbe794ef7bf09fbbfedc74cd5` | 11:50:32.4392691--11:55:16.5714591 | `slice4be-production-instruction-paths/91290b803d244953b0d096ac6f15df1b` |

Six principal captures are directly reviewed and hashed in `visible-review.json`.
Selected Remove/REQUESTED detail shows Info/Unknown. How-To presents all five
instructions, and the separate observed run matches five STAGED actions with
zero extras. The scrolled comparison visibly excludes a Domain-application claim.
This is agent review, not human acceptance; runtime and packages remain unchanged.

UOM staging retains 264/264 checks and public Close retains 61/61, with exact
prior ordered identities, five compiles each, preservation, normal closure and
zero delayed Excel Application failures. Controller roots are under
`reports/runtime/production-uom-staging-controller/`; intervals are September30 UTC.

| Gate | Controller | UTC interval | Result root under `reports/runtime/` |
|---|---|---|---|
| Staging/activity 264/264 | `9fbe651e107249318c77c58e6e91852f` | 11:55:25.8822203--11:59:49.8496079 | `slice4be-production-uom-staging/8b4f4c8a744c4e0f91d90140b7329c74` |
| Public Close 61/61 | `cbd85b5df24140289e41ac0b1fde8db9` | 12:00:00.7084811--12:01:34.6437975 | `slice4be-production-uom-public-close/9ffb48e6e1144ecdb7a3ee9730d9a520` |

Two staging and three public-Close captures are directly reviewed and hashed in
their result roots. The workbench preserves operator columns/neighboring cells;
the visible reopen status says existing edits are retained and the saved catalog
is not reloaded. Public Close disposes the form and permits the approved reopen.

The first UOM paired-path attempt is incomplete:77 PASS/one harness failure,
12:01:45.4866113--12:06:53.9723721 UTC, controller
`ac00a9b3d1344f0eb1bffb16b83918ce`, result
`reports/runtime/slice4be-production-uom-paths/c1bc3ff8ed14468aad389b9c7c40d8a8/green.json`.
Compare both passes its method/evidence assertions, then visible capture fails
because the requested form is not foreground. All three bounded capture attempts
observe a different form in the same Excel process; the exact cause is unresolved.
Seven trailing preservation/header checks do not run. This is not behavioral RED
or a completed GREEN gate. Four available principal images show correct REUSED
detail, two guide steps and OPENED/REUSED diagnostic matches with zero extras;
the Compare both/conclusion captures are unavailable. Five compiles, controller
settings/package restoration, normal closure and zero delayed Excel Application
failures pass. Desktop probes report no error5.

The unchanged repeat passes84/84,12:08:00.6013895--12:12:23.5965572 UTC,
controller `da2a1a21643340188d3b16776dbb075c`, result
`reports/runtime/slice4be-production-uom-paths/2e79dbe9b16046c0bfeb32d6fa9d2ad0/green.json`.
All84 prior ordered identities, five compiles, preservation, normal closure and
zero delayed Excel Application failures pass. All six principal images are
directly reviewed and hashed: OPENED/REUSED match the two instructions with zero
extras; Compare both and the scrolled conclusion visibly exclude Domain application.
This establishes the gate for this execution but does not diagnose or repair
the earlier foreground-capture failure. No runtime, test or package change.

Draft/designer regression retains 390/390 exact prior ordered checks, including
the separate Action Path diagnostic checks, with five instrumented compiles,
preservation, normal closure and zero delayed Excel Application failures.
The outer interval is 2026-09-30 12:12:48.4840790--12:21:26.7589579 UTC.
Outer controller: `reports/runtime/production-regulation-regression/draft-d50106c4b43646eb88e3aa70f9a718f9`;
inner controller: `reports/runtime/production-designer-controller/a4c9064fedad40568b4399bd714f390f`;
result: `reports/runtime/slice4be-production-designer/219c14986f6d4b6193397371b31a0ea9/green.json`.
This gate does not request new screenshots; no new visual or human acceptance
is claimed. Runtime, tests and frozen packages remain unchanged.

Lifecycle retains 615/615 and native cancellation retains 94/94 exact prior
ordered checks. Each has five instrumented compiles, preserved settings/packages,
normal closure and zero delayed Excel Application failures. September30 UTC:

| Gate | Outer controller under `reports/runtime/production-regulation-regression/` | Inner controller under `reports/runtime/` | UTC interval | Result under `reports/runtime/` |
|---|---|---|---|---|
| Lifecycle 615/615 | `lifecycle-5aca55ecf47c432c83866ae729f8606f` | `production-lifecycle-controller/9cc2902cabd6442e9f3756364c314467` | 12:21:47.7456484--12:27:44.5793061 | `slice4be-production-designer/5699a3aeb0784e7ba6fff1e0c34f9cd4/green.json` |
| Native cancellation 94/94 | `native-afa48f07d58141478fc18eef1b00049c` | `production-lifecycle-native-controller/a6f06147b25047bdae56a661c555a086` | 12:27:58.1278385--12:29:54.7206625 | `slice4be-production-lifecycle-native/4ce129198f81433a895c1b4592283282/green.json` |

The native result has five directly reviewed images and a hashed review receipt.
All four Process/Recipe Release/Obsolete questions are readable with No focused.
The Settings image explicitly reports saved configuration and unavailable
tracking. This is agent review, not human acceptance or native-crash repair.
Runtime, tests and frozen packages remain unchanged.

Settings retains 202/202 exact prior ordered checks and five instrumented compiles,
2026-09-30 12:30:21.7877413--12:34:25.8472714 UTC. Controller:
`reports/runtime/production-regulation-regression/settings-49449ed5bfc44cffba783798d5f23ecf`;
result: `reports/runtime/slice4be-tracking-settings/afef75845a8547efaf41d9aee85e2d0c/green.json`.
The preference restart is unassisted, all ten unique COM references release with
zero failures, settings/packages are preserved, final closure is normal and
the delayed Application audit finds zero Excel failures. This gate requests no
new screenshots and does not establish human acceptance.

## Remaining work

The remaining focused regressions and run-only reusable gate remain pending on
this final candidate. Catalog19 remains unchanged
and unpromoted. Its ordinary full reusable native failure remains unresolved; see
[designer read evidence](plan022_slice4be_production_design_read_results.md).
