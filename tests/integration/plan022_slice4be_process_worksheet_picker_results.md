# Slice 4be Process worksheet picker header alignment

Last verified: 2026-09-30 UTC. Focused115 GREEN, build/compile/static, smoke86,
layout, chain32/live48/Create15, worksheet107 and Close/query312 pass. Full
reusable/restart remains unresolved. No new tracking registration or acceptance.

## Governing contract and test

Architecture v4.11 D4/D14/D15 requires the shared Core picker to update the exact
selected numbered Process item/SKU pair under normalized managed headers.
Documentation d099b48 records the clarification before runtime changes. Existing
ownership, record types, allocation and explicit save behavior remain unchanged.

The packaged test invokes the real Production process-item picker and its actual
Core CommitSelection handler. Unsaved adapters prepare Admin-generated disposable
fixtures and read expected values from the actual displayed selection. Nine cases
cover canonical pairs1/2/5, lowercase pair2, space-padded pair2, normalized pairs1/2/5,
and normalized OUTPUT pair1. Other cells, custom values/formulas, header sequence,
the active decoy workbook, saved workbook bytes and canonical authority bytes are
protected. Fixture values remain in memory; runtime captures remain ignored.

## RED

Frozen baseline: `deploy/validation-process-worksheet-headers`.
Controller `reports/runtime/process-worksheet-picker-controller/2d220813d9fd491c8ce2fa795d740974`;
result `reports/runtime/slice4be-process-worksheet-picker/32758714dac647c09667ae70cc41abf4/red.json`.
2026-09-30 18:01:08.6416716–18:02:39.6465393 UTC: **110 PASS,5 FAIL/115**.
Five instrumented compiles and all42 prior shared GREEN checks pass. All five
noncanonical INPUT cases change the item label but leave its SKU unchanged.
Canonical and OUTPUT cases and all preservation assertions pass. Normal
unassisted shutdown, package/settings preservation and delayed Excel Application
audit pass. No desktop error5 observed.

Initial controller d7782185d7a7427b9b5fe2554b9ed731/result76f18831811a4beeb3da1923145b7dc6
reported111 PASS/4 FAIL. Its direct case-only ListColumn rename did not establish
lowercase coverage. The final fixture renames through a temporary header and
verifies the exact resulting header with binary comparison before executing the
same handler. This is fixture calibration, not a runtime workaround. The final
five failures above establish the protecting RED before Core edits.

The captured picker shows a populated actual item selection before commit; the
item/SKU mismatch is proven by assertions, not a post-commit screenshot. No claim
of human acceptance or comprehensive Slice4be completion.

## Candidate and focused GREEN

Only Core `cDynItemSearch.ProcessAlternativePairNumber` changes: trim the actual
header and compare its prefix with vbTextCompare. No module/procedure/line growth,
new bridge, identity allocation or ownership/save change.
Candidate `deploy/validation-process-worksheet-picker`; build record
`reports/runtime/process-worksheet-picker-build`,18:03:32.3325656–18:04:10.7948007 UTC.
Five compiles and Operations cold start pass. Compiled-source comparison finds
only cDynItemSearch changed among265 components;264 are exactly unchanged.
Frozen worksheet-header candidate/settings preserved, normal cleanup and delayed
Excel Application audit pass.

GREEN controller `reports/runtime/process-worksheet-picker-controller/fdc4c379abe54795a3a4cf8bffbca9b3`;
result `reports/runtime/slice4be-process-worksheet-picker/467da303e9e1453086791d9f553aa915/green.json`.
18:04:30.2768494–18:06:07.7087906 UTC: **115/115**, exact RED check sequence and
all42 prior shared GREEN checks. Five instrumented compiles, preservation, normal
unassisted cleanup and delayed Excel Application audit pass. The actual populated
picker capture was directly reviewed; it is a pre-commit surface, not visual proof
of the worksheet outcome. Automated assertions prove the latter.

Static `reports/runtime/process-worksheet-picker-static`:272 components,6116
procedures,134480 lines,9 literal/45 unresolved Application.Run,190 duplicate
body candidates and28 non-growing oversized caps. All baseline metrics unchanged;
three schemas and340 PowerShell parses pass, no exception added.

## Regression status

First full reusable controller
`reports/runtime/process-worksheet-picker-regression/reusablefull-320f6c4899ab4ba290a20d9fbb6b9de9`,
18:06:41.2076434–18:07:07.5249073 UTC, is excluded: HARNESS/HARNESS_CLEANUP RED,
RPC HRESULT0x800706BE at mProduction.RunProductionBatchScaleContractTest.
Application1000 at18:07:01.9175659 UTC records ntdll.dll/c0000028/offset12d2f;
Application1001 follows. Normal workbook closure was not completed. Settings and
packages were restored/preserved, no process termination requested, Excel absent
afterward. This matches a previously observed native failure signature, but does
not establish its cause or a repair. Desktop probes stayed healthy, zero error5.
Unchanged rerun controller
`reports/runtime/process-worksheet-picker-regression/reusablefull-563e5e0d518d4f2bb34ef213bd13a004`,
18:08:14.7979385–18:19:44.7329397 UTC, is also excluded. The main-workflow aggregate
passes with the exact first162 prior boolean observations, but the remaining nine
restart observations are absent. Restart fails at mProduction.RunReusableProductionRestartActionContractTest,
HRESULT0x80020009; final closure fails. First-session normal exit is proven in2586ms,
with zero reference-release failures. Final host disappears and a separate recovery
Excel process opens a recovery copy of this run's disposable operator fixture.
No matching Application1000/1001/1002 event was found through12 seconds after the
controller finally closed; do not
infer a second native-crash cause from the COM exception alone. After verifying
PID/start identity, recovery directory and this fixture's identity, that one book
was closed without saving and Excel Quit requested at18:19:41.8379338 UTC. No forced
termination; the controller then restores settings and verifies package pins.
This assisted cleanup is not GREEN. Full reusable/restart remains unresolved;
independent gates may proceed, but this correction's full gate set is incomplete.

## Passing independent regression gates

All paths below are under `reports/runtime/process-worksheet-picker-regression/`
unless specified. These gates use the same frozen candidate, preserve package and
settings pins, restore any tracked runtime reports, close normally without
assistance, and have zero delayed Excel Application1000/1001/1002 events.

- Smoke `smoke-d5d402903113428785af8d85ee015e80`,18:20:13.6261717–18:20:34.0704333 UTC:
  **86/86**, exact prior check identities; normal shutdown evidence
  `reports/runtime/packaged-smoke-closure/2726c55ef46d4932b6701beea9e1f164`.
- Layout `layout-6b97fb47632243ee8167f09e8c1b0920`,18:21:15.2430285–18:21:29.8792958 UTC:
  exact prior geometry over three sizes/five pages, no bounds/overlap violations.
  Three directly reviewed, hashed captures show empty Run List geometry; they do
  not establish populated-workflow or human acceptance.
- Chain `chain-c14802fb90d84c56bad9afb61794301a`,18:22:18.1621949–18:27:29.9798211 UTC:
  **chain32/live48/Create15**, all exact prior ordered checks retained.
- Worksheet headers controller
  `reports/runtime/process-worksheet-headers-controller/98034027e6d8425cbc59527306c944a4`,
  result `reports/runtime/slice4be-process-worksheet-headers/0e744171009042d7bb075a0de3d2f9d9/green.json`,
  18:27:57.3143323–18:29:36.7484749 UTC: **107/107** exact prior ordered checks and
  five instrumented compiles. Three captures directly reviewed: retained custom
  values and actual missing-name rejection status for custom/normalized cases.
  The normalized-header worksheet image is blank Excel chrome and excluded.
- Close/query controller
  `reports/runtime/production-close-controller/371b78b392e24114b067608c37d62542`,
  result `reports/runtime/slice4be-production-close/f8d5acb4fa49483bb12fa308bffe6a26/green.json`,
  18:30:01.4805109–18:33:29.8571002 UTC: **312/312**, exact latest prior sequence
  including all80 inventory-query preservation checks and five instrumented
  compiles. Four captures directly reviewed/hashed: button/native images show
  pre-dismissal surfaces, public reopen shows the real form, and Settings shows
  successful save with unavailable-tracking notice. Assertions prove dismissal;
  these are not post-dismissal captures or human acceptance.

## Fresh-process restart diagnostic

The test-only diagnostic changes no production source, packaged code or normative
behavior. D13 product RED is not claimed for this tooling. Offline generation
calibration10/10 proves original workflow statements retained, observer/trace
placement only in the fresh restart, unchanged original VBA lines, fixed-label
redaction and no logger call intercepting error-handler Err reads. Calibration:
`reports/runtime/production-restart-diagnostic/af6aca6fd2a14792b83d49c09b0f24b9/calibration.json`.
All343 PowerShell files parse. The existing native observer implementation is
unchanged. The generated validator inherits its existing fixture credential in
memory; it does not copy the literal into generated source or replay metadata.

Diagnostic `reports/runtime/production-restart-diagnostic/d1b55669fc0d4632a818611b0014e82d`,
18:43:50.4702355–18:51:59.5509317 UTC, passes two aggregates and the exact171 prior
boolean observations. Only the fresh restart is observed/instrumented; all four
loaded projects compile. The trace transport rejects an unknown label, emits16
allowlisted entries, and reaches both Retrieve returns and Result.Success.
Faults-only native observation is Attached/Ready/Exited with no relevant fault,
observer exit0 and owned Excel exit00000000. First-session normal exit takes22ms;
final exit2418ms, no reference-release failures or forced termination. Final
workbooks and four packages close normally. Package/settings/helper/validator
pins and the delayed Excel Application audit pass. Desktop probes remain healthy.

The existing disposable runtime is retained for a targeted follow-up. Its ignored
fixture-context.json contains identifiers and an operator-relative path, not
credentials. It describes final state, which may contain restart mutations; it is
not a pristine checkpoint. No Auth copy or runtime data dump is created.

This proves an observed restart can complete, not a cold-run repair. The earlier
unobserved failures remain unresolved and the picker prerequisite is not complete.
Next use the retained fixture and actual public
`mProduction.RunProcessWorksheetWorkbenchContractTest` to recreate its two draft
tables, close normally, then exercise the nine restart observations without VBE
tracing. Compare native-only observation with an unobserved control as needed;
do not replace the full171 acceptance gate with this smaller diagnostic. Preserve
all frozen baselines and do not issue another undifferentiated full rerun.

## Uninstrumented full workflow and targeted replay

The picker prerequisite's scoped gate set completes on 2026-09-30 UTC. A fresh
disposable fixture passes the full uninstrumented reusable workflow, followed by
a separate uninstrumented restart replay. This supersedes the unresolved gate
status above; the earlier failures remain excluded and their cause is unproven.
No production source, package or architectural contract changes in this follow-up.
D13 product RED is not manufactured for this diagnostic tooling.

Two smaller attempts against the previously retained fixture stop at sign-in,
before worksheet work. The fixture uses a random credential held only in its
creator's memory. Re-evaluating its setup expression produces a different value;
the numeric Auth rejection is not desktop error5 or product RED. Both attempts
close normally, restore settings and preserve packages; delayed Excel audits are
clear. Evidence under `reports/runtime/production-restart-replay/`:
`4a4df065246a4644b8d086d45c9b03e6`,19:02:00.1000754–19:02:11.0747331 UTC, and
`457871c81c07441dacc536e6396a0ee2`,19:02:50.2525499–19:03:00.9925190 UTC.
Do not repeat this approach or persist the credential to enable replay.

The corrected creator passes its same random credential to the replay tool as an
in-process SecureString; it is never serialized, logged or passed on a command
line. Replay checks the exact disposable-root boundary and five frozen package
pins, runs the actual public workbench handler to recreate two outstanding tables,
then closes normally and invokes the actual restart handler in another process.
This is fresh owner-generated staging, not a pristine snapshot of an earlier run.

Generation calibration remains10/10 for observed mode
(`production-restart-diagnostic/fb9077d035094d0c96a3dc68af72e6fc`) and passes10/10
for unobserved mode (`114616ec1b9642d08f75e54a43abffdb` under the same parent).
Unobserved generation preserves original workflow statements and inserts no
debugger attachment or VBE trace. Replay's ownership/pin/dispatcher/redaction
calibration passes without opening Excel:
`production-restart-replay/4fa3b18c788e478bb574cd69c14bd32f`.
All345 PowerShell files parse. These paths are under `reports/runtime/`.

Live controller `production-restart-diagnostic/0bba8416f0794d89bd8382e5082fb6c1`,
19:04:41.4262833–19:13:37.5572090 UTC, passes two full-workflow aggregates and
**all171 observations in exact prior order**, compared with the completed header
candidate. Original workflow statements remain intact; no debugger or VBE
instrumentation runs. Initial/restart normal exits take2951/2381ms, with zero
release failures and no termination. Final fixture books and all four packages close
normally. The subsequent targeted replay
`production-restart-replay/31b7005b0d45405ea429f836ddaae5a4`,
19:12:43.2250580–19:13:37.5231706 UTC, passes **37/37 exact checks**, including
the public workbench actions, exact released recipe, captured workbook, two table
rediscovery/retrievals and no new workbook. Its two sessions exit normally in
2521/2387ms with no release failures or forced termination. Controller and replay
verification.json records preserve these distinctions.

All package, settings, helper and original-validator pins hold; delayed Excel
Application1000/1001/1002 audits report zero failures. Desktop probes remain
healthy. The smaller replay supplements the full171 gate and does not replace it.
Together with the focused115, build/compile, unchanged static ratchets, smoke86,
layout, chain32/live48/Create15, worksheet107, Close/query312 and reviewed captures
above, this completes the isolated picker correction's gates. It proves neither
a native-crash repair nor human/full Release1 acceptance. No candidate promotion
or new tracking registration. Next specify and test the three worksheet actions'
local versus Designs-submission observations under D18, including partial results.
