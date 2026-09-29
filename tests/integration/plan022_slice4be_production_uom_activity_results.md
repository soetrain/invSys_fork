# Slice 4be UOM Edit observations

Last verified: 2026-09-29 UTC. **Focused observation GREEN; broader acceptance open.**

Architecture v4.11 D18 catalog15 registers the existing **Edit UOM Catalog on
Sheet** handler as `PRODUCTION_UOM_EDIT`, owned by `PRODUCTION_UOM_STAGING`.
The typed Operations adapter rechecks current captured context and existing
PROD_POST permission, then records REQUESTED and the owning staging outcome.
OPENED/REUSED describe local workbench completion with saved authority unchanged;
they never assert catalog publication or Domain application. Fixed observations
contain no catalog values, formulas, notes or workbook paths. Catalogs1-14 remain.

## Focused D13 evidence

The actual packaged Send handler records **98 PASS / 134 behavioral FAIL** before
implementation and **232/232 GREEN** on `validation-production-uom-activity`.
Check identities match exactly, including all preceding78 staging checks. The gate
covers owner metadata, historical catalog definitions, OPENED/REUSED/reopened,
REJECTED/DENIED/FAILED, tracking disabled/unavailable, current-context refusal,
closed workbooks, loading/re-entry suppression, redaction and exact completion
classification. All five instrumented projects compile; closure is unassisted and
settings and frozen packages are preserved.

An earlier101/131 RED used an overly broad refusal-message oracle; it was tightened
before implementation. The98/134 result is the protecting baseline.

Five candidate builds/compiles and cold Operations dependency loading pass.
Among249 compiled components, six intentionally change relative to the preceding
UOM preservation candidate: Core activity catalog/completion map/new UOM codes,
Operations UOM action/staging owner and Production form. Compare exact source,
case-insensitive source and literal hashes, not XLAM container bytes.

## Action Path evidence and qualification

On the final `validation-production-uom-extent` candidate, the corrected disposable
reset now completes **84/84 unassisted**, with the exact preceding84 identities,
normal closure, settings/package preservation and zero Excel Application events.
Six principal images from that run are reviewed. Its fixed-row terminal capture
still selects REQUESTED because contributing-line order varies. The capture fixture
now resolves the exact terminal RecordId within its published group; the resulting
rerun again passes84/84 unassisted and the reviewed Event Detail image selects
REUSED, Info/Unchanged, with the retained-draft/no-reload explanation. This is a
fixture correction, not a runtime or normative contract change; no manufactured
product RED is claimed for selecting the intended capture line.

Current unassisted controller/result:
`production-uom-staging-controller/fea9b5f7c6eb45d2a1720168459a6182`;
`slice4be-production-uom-paths/fd10b376c43e48e6bd6d32f96387cba8/green.json`.
Exact-terminal capture controller/result:
`production-uom-staging-controller/76985050ea7a41d7bc69fe58808d8cbc`;
`slice4be-production-uom-paths/00411c0aa42e4c9688fa0b584f4bd0fb/green.json`.
Comparisons: `production-uom-extent-path-verification.json` and
`production-uom-extent-path-final-verification.json`.

The earlier assisted run below is retained as qualified historical evidence:

Independent guide-source and observed recordings pass **84/84** through actual
Admin publication, Viewer selection, guide authoring, reader pairing, evaluation
and How-To/Diagnostic/Compare both views. Exact original activity/journal identities
are preserved. OPENED then REUSED concludes CommandCompleted; empty source-event
references leave SourceEventsApplied incomplete. Saved authority, workbook bytes,
custom columns and immutable evidence remain unchanged.

This first path run required UI assistance confirming a verified disposable-sheet
delete dialog (`NUIDialog`). It is not unattended acceptance. The shared fixture
now suppresses alerts only around its validated disposable reset and restores the
prior value; an unassisted rerun is required. Six principal images were reviewed:
Production Settings, Event Detail, How-To, Diagnostic, Compare both and conclusion.
The first Event Detail image selects REQUESTED while showing the paired REUSED line;
the corrected fixture will select the terminal line for capture. Images/runtime
values remain ignored. These observations are not user acceptance.

Static evidence:256 components/6071 procedures/133522 lines;9 literal and45
unresolved dynamic calls,191 duplicate groups. None of28 existing oversized-module
caps grows; Production form shrinks11739 to11736 lines. Three schemas validate and
305 PowerShell files under tools/tests/tooling parse without errors.

## Remaining gates

Current-candidate Settings regression passes **202/202**, retaining every preceding
identity, normal closure, settings/packages and zero Excel Application events.
Packaged Production layout passes three sizes/five pages and native window actions;
minimum and expanded images are reviewed, and default is byte-identical to minimum.
The actual packaged Production captured-workbook close/reopen workflow passes1/1,
preserving the unrelated workbook and creating one intended role workbook, with
settings/packages preserved and no Excel Application event. Its cleanup receipt
records forced termination: this validator waits only500ms outside palette-probe
mode. It is workflow evidence, not proof of normal process shutdown or a repair
of the separate combined form-test failure.

- Settings controller/result: `production-uom-extent-regression/settings-53a3d9dcd4854a6ba9b983b7a8f98c53`; `slice4be-tracking-settings/d4d2a8b4a82946709c86f7467de68314/green.json`.
- Settings comparison: `production-uom-extent-settings-verification.json`.
- Layout: `production-uom-extent-regression/layout-2efeb4ba2ffc4c50a05524892ed605f8`.
- Captured close: `production-uom-extent-regression/capturedclosed-96e8b26ac59d46cf8192aae98e467fcb`; inspect `captured-closed/final-cleanup-observation.json` for the qualification.

The first combined visible run on `validation-production-uom-extent` stops at
**255 PASS/1 harness failure**, before completing the closed-workbook guard and
optional-tracking tail. Of the110 preservation-suite identities,107 have passed;
three common trailing checks have not run. This is not a completed GREEN gate.
A VBA dialog reports80010007 at
`TestProductionDesigner.SendUom`, the test adapter call into the form. Diagnostic
Debug/Reset intervention localizes the boundary but changes the run. Windows then
records Excel/oleaut32 nativec0000005 at20:50:17 UTC, offset0000000000033b41; recovery
termination is needed. The timing does not prove a common cause with the earlier
c0000028 failure or identify a product/harness root cause. This is not desktop
Win32 error5. Settings and packages are restored. Do not hide the failure by
reporting only earlier passing checks or treat assisted closure as acceptance.

Failed visible controller/result:
`production-uom-staging-controller/91119955725d4554911de2eea19aae16`;
`slice4be-production-uom-staging/d29842ba571541f788c60f04fe74d756/green.json`.
The controller contains sanitized dialog-location, reset and cleanup receipts.
Separate visible staging validation now completes110/110 with exact identities,
normal closure and preservation. The controlled comparison with capture disabled
completes **264/264**, including the closed-workbook and optional-tracking tail,
all earlier232 GREEN identities, normal closure and zero Excel Application events.
It uses the unchanged candidate and skips no checks; its diagnostic phase is RED,
not a fresh behavioral RED baseline. The contrast narrows investigation but does
not prove a visibility-related root cause or repair the failed visible path.
The larger visible guard failure remains open.

Comparison controller/result: `production-uom-staging-controller/c74a18a9e47746a7a228f86fefcf31c9`;
`slice4be-production-uom-staging/774c6e429e8b4de3bd5e89f327ab21a5/red.json`
(20:54:13--20:57:15 UTC). Verification: `production-uom-extent-hidden-verification.json`.

The additional blank-row preservation test exposes a separate defect: Retrieve
retains cells but Edit recreates a shortened table above the blank row. Initial
expanded result84/1 preserves all earlier78 GREEN checks. The full-extent correction
under the approved reuse decision reaches separate visible110/110 on
`validation-production-uom-extent`, following expanded RED96/14.
See [staging evidence](plan022_slice4be_production_uom_staging_results.md).

The final candidate's Instructions regression passes411/411, retaining the exact
prior check identities with normal closure, preserved settings/packages and no
Excel Application event. Controller: `production-instruction-controller/f92a5868d7304b6cb06802152dc00b89`;
result: `slice4be-production-instructions/24874fd7a68b4bfc907684d45bc2f475/green.json`;
verification: `production-uom-extent-instructions-verification.json`.

Instruction paths also pass105/105 with exact prior identities, six reviewed
principal images, normal unassisted closure, preservation and no Excel Application
event. The detail image selects REQUESTED; the paired conclusion shows all five
STAGED outcomes and CommandCompleted without asserting Domain application.
Controller: `production-instruction-controller/3f25713a2a0a4361bc706ce27db8cf8d`;
result/images: `slice4be-production-instruction-paths/e476b7d9e64a40a9a8783f34b5c14bf4`;
verification: `production-uom-extent-instruction-paths-verification.json`.

The draft/Action Path regression passes390/390, retaining exact prior identities,
normal unassisted closure, settings/packages and zero Excel Application events.
The controller finishes about61 seconds after the final report without
intervention; a post-report observation confirms Excel has exited. Controller:
`production-uom-extent-regression/draft-eab58816d05d4229b76588e88860e644`;
inner controller: `production-designer-controller/0afa3cd3559b4d53af212ad3e9fc7452`;
result: `slice4be-production-designer/f9d77d39b8be4fd2bfea52da7d353c5a/green.json`;
verification: `production-uom-extent-draft-verification.json`.

Lifecycle615/615 retains every prior identity, normal unassisted closure,
settings/packages and zero Excel Application events. Controller:
`production-uom-extent-regression/lifecycle-5edeabaa0854497ea3e9af8756b355c8`;
inner controller: `production-lifecycle-controller/33df59218b79475786666e0cb429b06b`;
result: `slice4be-production-designer/bf3a78b6db1f4c5285a73533dca7309f/green.json`;
verification: `production-uom-extent-lifecycle-verification.json`.

An opt-in UOM adapter diagnostic repeats the visible closed-workbook failure
without Debug/Reset intervention:262 PASS/1 harness failure. It calibrates a
missing test form as error91 with no form entry or activity, then records form
entry and zero adapter error for DENIED, FAILED, Target, Session and SignedOut.
The closed-workbook Send call fails with transport envelopes80131501/80020009;
the first failure receipt still lists the original Excel and a new child. The
original later exits during harness cleanup; the child prevents restoration until
terminated. This does not prove a native crash. No closed-workbook state receipt
is returned, so the probe cannot establish whether the form was entered. No Excel
Application1000/1001 event was observed during the checked window. Settings and
all five packages are restored; this is qualified diagnostic evidence, not GREEN.

Diagnostic controller/result: `production-uom-staging-controller/7497bb331a724613a8d15aeca234eabd`;
`slice4be-production-uom-staging/53bf0c953eb64aacae8ed58d9d5436bc/red.json`.
Inspect its allowlisted `uom-adapter-diagnostic.jsonl`, first-call failure facts
and controller `assisted-cleanup.json`. The diagnostic switch requires visible
activity checks and RED; attempting acceptance GREEN is rejected before Excel work.
It changes only unsaved test adapters, preserves the refusal assertions and logs
fixed case names, Boolean entry and numeric errors. A durable fixed-marker
extension addresses that evidence gap because the original instance is unavailable
after cleanup.

The durable diagnostic completes271 PASS/2 FAIL with normal unassisted closure,
all264 preceding check identities retained, five instrumented compiles, preserved
settings/packages and no Excel Application1000/1001 event. The failures are
`UomActivity.Guard.ClosedWorkbook.RefusalVisible` and the additional diagnostic
form-entry assertion. Fixed trace49--51 records ClosedWorkbook case selection,
adapter entry, then AdapterFailed -2147418105 (80010007), with no intervening
form-entry marker. No redirected activity is created. The call fails at the
retained-form dispatch boundary; this does not establish why that reference is
unavailable or prove a runtime repair. Subsequent fresh-form tracking-off and
unavailable-store cases pass. Their trace lines retain the previous case label
in this run; only49--51 are the closed-workbook attempt. The test-only label reset
is corrected for future diagnostics, without repeating the gate for that label.

Durable controller/result:
`production-uom-staging-controller/7851d87d119c4415abccbed62f800a50`;
`slice4be-production-uom-staging/da9f2da09f724d1a802103c380fdb2a8/red.json`.
Its `uom-adapter-entry.txt` contains only sequence, fixed case/stage and numeric
error fields. Verification: `production-uom-extent-adapter-diagnostic-verification.json`.
D18 requires closed binding refusal without replacement-context attribution;
the actual launcher-owned close/reopen lifecycle still needs a focused UOM proof
before any decision about the failed retained-instance message oracle. Keep the
failure open; do not treat diagnostic instrumentation as an accepted runtime change.

Native cancellation94/94 retains every prior identity, four reviewed principal
dialog images, normal unassisted closure, settings/packages and zero Excel
Application events. Controller:
`production-uom-extent-regression/native-9be0bdcf7ada4efdaaa548ed567a6902`;
result/images: `slice4be-production-lifecycle-native/0af64587bc3f4ca7a08c4c3e5a398224`;
verification: `production-uom-extent-native-verification.json`.

Packaged smoke86/86 retains every prior identity, settings/packages and the
tracked report, with no Excel Application event. Its harness can force termination
after one second and does not report whether that branch ran; normal automatic
shutdown is unobserved. Controller:
`production-uom-extent-regression/smoke-c3f670085f624961bdb4412c219ee7be`;
verification: `production-uom-extent-smoke-verification.json`.

The Settings-observation regression must run
on the final candidate. The preceding candidate's full-chain/native c0000028
blocker remains unresolved. No package promotion or Release1 acceptance is claimed.

## Traceability

All runtime paths below are beneath ignored `reports/runtime/`.

- Contract before implementation: documentation commit `13291eb`.
- Final RED controller/result: `production-uom-staging-controller/7875e27bf99d4e5d999402087ba81eca`; `slice4be-production-uom-staging/e27a45e839444857b1e644c7d36a0026/red.json`.
- GREEN controller/result: `production-uom-staging-controller/95f6458f3b984b89b2a56a38d4e3f83f`; `slice4be-production-uom-staging/93bffe93d5c7432c8e109942aea1de3b/green.json`.
- Exact check verification: `production-uom-activity-focused-verification.json`.
- Build/source: `production-uom-activity-build`; maintenance: `production-uom-activity-static`.
- Assisted paths controller/result: `production-uom-staging-controller/92807e8dd2ab4cf0b095c64d11b73411`; `slice4be-production-uom-paths/43467d5b5ca94624abc53e54a57956f2/green.json`.
- Blank-row RED controller/result: `production-uom-staging-controller/bf9c0604d0a543a28d8977e4085d1a42`; `slice4be-production-uom-staging/fe3a8086b9394ff49f11a1e040f570a4/red.json`.
- Tests: `Test-Slice4beProductionUomStaging.ps1`, `Slice4beProductionUomActivity.ps1`, shared `Slice4beProductionInstructionPaths.ps1` under `tests/tooling`.
