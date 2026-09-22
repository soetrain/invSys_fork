# Plan 022 Slice 4be.5 - direct tracked-action guide curation

Status: corrected isolated candidate passes focused **55/55** and expanded
integration **63/63**. Supporting regressions are in progress; deployment, human
acceptance and full Slice 4be completion are not claimed.
Architecture v4.11 D18's direct tracked-action curation refinement,
Plan 022 and controls were committed together in docs `49ffb47` before testing.
This implements the already approved non-recording authoring route; pending
Event Detail and D8-A proposals remain unapproved.

Protecting test: `tests/tooling/Slice4beGuideActionCuration.ps1`, invoked by the
existing packaged configuration controller with `-GuideActionCurationOnly`.
The controller installs probes into disposable copies before forms run and
explicitly compiles all five projects. The probe only observes controls or
delivers their normal handlers; multi-selection uses the actual ListBox selection
change. It never manufactures a guide or recording.

Actual Admin Settings saves create two original actions outside a recording;
ordinary Admin publication and Viewer loading provide their exact source bodies.
The tests protect independent entry, search-preserved multi-selection, same-source
reuse, exact publication binding, four picker/editor sizes, stable authored IDs,
original observations and source references, empty SourceRun, None expectation,
immutable schema-1 save/revision, source/config preservation, non-publication,
parent/child closure, sign-out, capability revocation, policy loss and target change.
Permission fixtures use the disposable Admin-generated Auth file and restore its
original bytes; no operational workbook or installed package is modified.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-published-guide-edit-notices -Phase RED -GuideDraftOnly -GuideActionCurationOnly -CheckViewerPublishedRead -CompileEvaluationProbesForTest -ViewerStartupPackageStateForTest SavedCopies
```

Initial RED is **12 PASS / 37 expected FAIL**, 49 checks, at
`reports/runtime/slice4be-viewer-published-read/b7e0e7004c774081b433ba6ab3ed91a1`,
**2026-09-22 06:40:41.2389482--06:42:18.3441713 UTC**. Four additional parent/
Viewer/sign-out/maintenance-revocation checks are then added before any runtime
change. Expanded RED is **12 PASS / 41 expected FAIL**, at
`reports/runtime/slice4be-viewer-published-read/afe5b11ae907444681cc750897d9322a`,
**2026-09-22 06:43:54.8095241--06:46:07.0742576 UTC**. All 49 earlier identities and
outcomes are retained. Every failure is the expected missing curation behavior;
all five instrumented compiles, both startup checks, the actual source fixture
and four preservation checks pass. Both runs close Excel normally without
assistance and have zero matching Application events.

`reports/runtime/verify-guide-action-curation-red.ps1` validates the expanded
typed assertion set, prior identities, terminal exit, closure, all 175 running
test-source pins and the event audit. Frozen candidate validation preserves all
55 package pins and existing 28 static size limits. Runtime evidence remains
ignored and is not committed. Required next gate: implement only the named
Core/Operations contract, build a new isolated candidate, then prove the same 53
checks GREEN plus visible evidence and all applicable existing regressions,
maintenance, live-role and full Release 1 chain gates.

The initial runtime candidate `deploy/validation-guide-action-curation` builds all
five packages and passes explicit compile/Operations cold start. Compiled scope
is exactly five existing components plus three additions, 239 to 242: Core
`modActionGuideDraft`, `modGuideModel`, new private `modCuratedGuideSource` and
primitive boundary `modGuideActionPicker`; Operations `modGuideEditor`,
`frmActionPathGuide`, `frmActionPaths`, new `frmGuideActionPicker`. The inherited
Event Detail working change is untouched and not staged.

Its first focused run gives **24 PASS / 29 FAIL**, preserving all 53 identities,
at `reports/runtime/slice4be-viewer-published-read/9e6f6d088bdb474795b9f1255ee64bff`,
**2026-09-22 06:54:42.3516298--06:57:20.0634241 UTC**. Four picker captures are
directly reviewed and hashed: layout/source/status are readable, but both valid
actions are unavailable. No editor opens. Root cause: published activity lines
include the storage ContentSha256 wrapper, whereas the existing guide observation
validator accepts the 26-field original body. The adapter incorrectly passes
the wrapper to that validator. The test also incorrectly compares the guide body
with the wrapper. Both must honor the existing distinct schemas.

D18, Plan and controls clarify that distinction in docs `b216fd9` before the
adapter correction. The test now verifies each published envelope against its
original activity file and digest, compares every original body value, and checks
that an invalid inner digest is rejected while another valid action remains
selectable. The outer publication remains valid in that disposable negative fixture;
its original bytes are restored afterward. No source format or hash rule changes.

The first follow-up attempt is a harness interruption, **22 PASS / 24 behavioral
FAIL / one harness FAIL**, 47 reached checks, root
`cdf5808dcd89430fac7cde58bd859611`, **2026-09-22 06:59:28.1241554--07:00:50.4542936
UTC**. `frmActionPaths/btnChooseGuideActions/Click` is rejected with inner
`0x80010001`; the new digest checks are not reached. First-call evidence is retained,
no owned foreground capture was available and no action is replayed. The cause is
unproven. Excel closes normally; zero matching Application events and unchanged
running sources/frozen candidate are verified. This is not behavioral RED.

The fresh-fixture retry against that same retained candidate completes **24 PASS /
31 expected FAIL**, 55 unique typed checks, root
`reports/runtime/slice4be-viewer-published-read/0c9ba553bd174cd6a0a4f3f88c65dafc`,
**2026-09-22 07:01:38.8834357--07:04:19.6056357 UTC**. All preceding 53 identities
and outcomes are unchanged; both added digest checks execute and fail. Normal
closure, zero matching Application events and unchanged test/runtime sources and
candidate hashes are verified before fixing the adapter. Summary:
`reports/runtime/guide-action-curation-envelope-red-retry-verification.json`.

Initial maintenance at `reports/runtime/guide-action-curation-static` records
249 components, 6,034 procedures, 132,470 lines, dynamic calls 9/45 and duplicate
groups 192. All 28 preceding size limits pass. The next candidate must preserve
all 55 focused identities, add visible editor evidence and complete supporting
regression gates; this first candidate is retained without overwrite.

The corrected candidate `deploy/validation-guide-action-curation-envelope` builds
all five packages and passes explicit compile/Operations cold start. It retains
all 242 compiled identities; only private `modCuratedGuideSource` differs from the
first curation candidate. Its detached-copy adapter verifies ContentSha256 before
validating and retaining the original body. No source envelope is mutated.

Focused GREEN is **55/55**, preserving every protecting RED identity, root
`reports/runtime/slice4be-viewer-published-read/a91cba903fa2443e9b23126adcbfb131`,
**2026-09-22 07:07:19.2421691--07:11:03.0440874 UTC**. Eight picker/editor captures
are directly reviewed and hashed. All controls, source notice, authored text,
original observations and Save/Cancel remain readable at minimum/default/larger/
restored sizes. Excel closes normally without assistance, sources and frozen
packages remain unchanged, and the Application event audit finds no matching
1000/1001/1002 events.

Additional tests exercise four library layouts and actual published-reader/Edit/
Save handlers for this non-recorded source. The first integration attempt reaches
50 passing checks before a read-only picker query is rejected with `0x800AC472`;
root `cb479e0047c74351831f0352aebeab06` retains first-call evidence. The published
round trip is not reached. This is a harness interruption, not behavioral RED;
no handler is replayed, runtime changed or cause inferred. Normal closure and
unchanged sources/packages are verified separately.

A fresh fixture using the existing visible-Excel harness option completes
**63/63**, root
`reports/runtime/slice4be-viewer-published-read/70fc6f19b65f4f7d875a5786e973302d`,
**2026-09-22 07:17:40.3315281--07:21:23.9654934 UTC**. Every preceding focused
identity passes. Published reading works without a selected run; Edit restores
stable step identities, and Save creates the next immutable version with empty
SourceRun, unchanged original bodies and exact predecessor hash. All twelve
library/picker/editor captures are directly reviewed and hashed. The library's
existing empty-recording unavailable notice remains visible, while Choose tracked
actions works independently. No automatic action retry is introduced. Normal
unassisted closure, zero matching Application events and unchanged test/runtime/
package pins are verified by
`reports/runtime/verify-guide-action-curation-integrated-visible.ps1`.

Corrected maintenance: **249 components, 6,035 procedures, 132,490 lines**, dynamic
calls **9/45**, duplicate groups **192**. All **28** preceding size limits pass.
Evidence stays ignored under `reports/runtime/guide-action-curation-envelope-static`.
The existing guide regression, fresh-process preference, Viewer, Boxing/Shipping,
evaluation, live-role and full Release 1 chain gates must still be verified on
this candidate. Pending D8-A and Event Detail proposals remain separate.

The first broader 351-check attempt is interrupted after **137 PASS / one harness
FAIL**, root `55b3481960b940c786693ab35cd04d75`, **2026-09-22
07:23:19.3800958--07:30:49.5751172 UTC**. The read-only probe
`frmActionPathGuide / empty control / Count / empty value` receives inner
`0x800AC472`. First-call evidence records the foreground as a terminal, without
an owned-form capture; this establishes the observed state, not the rejection's
cause. Three prior Save/library captures are directly reviewed and hashed. No
product assertion fails before interruption. Normal unassisted closure, unchanged
running source/runtime/package pins and zero matching Application events are
verified. This attempt is not a completed broader regression.

The existing opt-in read retry is extended through a separate
`-RetryGuideObservationForTest` switch for exactly two observational probe calls:
`frmActionPathGuide / empty / Count / empty`, and
`frmGuideActionPicker / lstGuideActions / Values / empty`. Both probe branches
only inspect already-loaded forms/control values. Only inner `0x800AC472` or
`0x80010001` qualify, with four total attempts and 250/500/750 ms delays. First
failure remains recorded; sanitized `readonly-observation-retries.jsonl` states
each attempt and recovery. Click, selection, toggle, write, other controls/codes/
arguments and the default flag-off path retain one execution. No runtime or
frozen package changes, automatic business retry or causal diagnosis is implied.

The actual shared Run function has offline protecting **70 PASS / 8 expected
FAIL** at `reports/runtime/form-failure-diagnostics/105cc73c7922423e85adda389c0fad98`,
then **78/78 GREEN** at
`reports/runtime/form-failure-diagnostics/4b9af27246824f26a7397d5649fa6f68`.
All prior 38 identities remain. New checks protect recovery, exact exhaustion,
excluded actions/fields/codes, first-failure retention, sanitized trace identity
and field-content exclusion. These are harness tests, not packaged runtime RED.
The full guide regression is restarted with a fresh fixture and this explicit
option; it must still run every original assertion against the frozen candidate.

That second broader attempt reaches **286 PASS / one harness FAIL**, root
`ec9f089ec0064a2795147869cd914904`, **2026-09-22
07:34:57.8181759--08:00:54.6248543 UTC**. Opening the recording library receives
inner `0x80010001`; cleanup then receives `0x800AC472`. No operator action is
replayed, no read-only retry occurs, and the cause remains unproven. All 36
captured images are directly reviewed and hashed. Normal unassisted closure,
unchanged running source/package pins and zero matching Application events are
verified. This is not a completed 351-check regression.

The source-integrity extension first stops at a fixture assumption (59 PASS /
one harness FAIL, root `37d4068bf1454883b6f70fccde842ffe`). A completed activity
appears in both Lines and Outcomes. Correcting the fixture to change both exact
representations and recompute inner/outer hashes changes no runtime behavior.

Corrected focused RED is **67 PASS / four expected FAIL**, **71 unique typed
checks**, root `d77014cd25d44290bd994fe27207c368`, **2026-09-22
08:10:06.4406656--08:14:16.6259455 UTC**. The actual picker still offers an action
with malformed result RecordId, or conflicting result CatalogVersion,
PackageSetVersion or BuildIdentity. All four valid-neighbor/source-preservation
checks pass, as do all prior 63 focused identities. Each fixture uses valid
recomputed hashes and restores exact publication bytes afterward. Normal
unassisted closure, unchanged test/runtime/package pins and zero matching
Application events are verified before any Core fix by
`reports/runtime/verify-guide-action-curation-integrity-red.ps1`.
The correction will enforce existing D18 identity and release agreement in
private Core validation; it changes no schema, controls or authority contract.
