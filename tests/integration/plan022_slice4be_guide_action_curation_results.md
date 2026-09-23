# Plan 022 Slice 4be.5 - direct tracked-action guide curation

Status: the current Settings-diagnostic candidate passes **77/77** curation
checks with **18** directly reviewed captures and normal closure. Earlier focused
55/55 and expanded 63/63 results remain below. Deployment, human acceptance and
full Slice 4be completion are not claimed.
Architecture v4.11 D18's direct tracked-action curation refinement,
Plan 022 and controls were committed together in docs `49ffb47` before testing.
This implements the already approved non-recording authoring route; pending
Event Detail and D8-A proposals remain unapproved.

Current-candidate verification, **2026-09-23 23:34:15--23:39:45 UTC**:
`deploy/validation-settings-diagnostic` retains all 77 preceding check identities.
All five instrumented projects compile before the actual handlers execute.
Direct review accepts all six sizing states of the library, picker and editor.
The library images show the fixture's explicit unavailable-recording state and
reachable direct-curation entry; they do not prove a populated recording library.
The picker shows two selected published actions; the editor separates authored
instructions/order from original observations and execution evidence. Controls
remain visible through minimum/default/larger/restored and native maximize/restore.
Source preservation, immutable revisions, parent closure, sign-out, permission/
policy/target loss and malformed-source isolation checks pass.

Excel closes immediately without assistance. The 23:39:51 UTC audit finds zero
Application failures; 299 runtime/195 test/five package hashes are unchanged.
Root: `reports/runtime/slice4be-viewer-published-read/5ef27faf08dd4ca7a8c78dcde636261b`.
Exact independent receipt:
`reports/runtime/settings-diagnostic-fixed-regression-curation-verification.json`.
This is a candidate regression result, not a new behavioral contract or human
comparison acceptance. Remaining restart/Viewer/guide and comparison gates retain
their separately recorded status.

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

The new isolated candidate `deploy/validation-guide-action-curation-integrity`
builds all five packages and passes explicit compile/cold start. All 242 compiled
component identities remain; only private `modCuratedGuideSource` changes. It
validates original RecordId as a GUID and requires attempt/result catalog,
package-set and build agreement before offering the action.

Focused GREEN is **71/71**, preserving every RED identity, root
`dbe2029e5d2f4188a588112e3640dcaf`, **2026-09-22
08:21:49.0523985--08:25:59.6998582 UTC**. All twelve library/picker/editor captures
are directly reviewed and hashed at default/minimum/larger/restored sizes.
Normal unassisted closure, unchanged running sources/frozen candidates and zero
matching Application events are verified by
`reports/runtime/verify-guide-action-curation-integrity-green.ps1`.
Maintenance has **249 components, 6,035 procedures, 132,492 lines**, dynamic calls
**9/45**, duplicate groups **192**, with all **28** preceding size limits passing.
User-file pins are preserved. Supporting regression gates still remain.

The opt-in harness readiness guard has offline **15 PASS / 17 expected FAIL**
before implementation at `reports/runtime/ready-before-dispatch/6381f50fa7964db6a76da65f2f587a0a`,
then **32/32 GREEN** at `b9bbba062b42476dae07d493a0254545`. The protecting test
loads the actual Run function and exercises read-before-dispatch ordering, eight
read bound, unavailable/False/type handling, flag-off behavior, single action
execution even on dispatch rejection, and field-value exclusion from traces.
The previous **78/78** diagnostics also remain GREEN at
`reports/runtime/form-failure-diagnostics/fe920567b4464fe7b53dc90e803aa216`.

`-WaitForExcelReadyForTest` samples the read-only Application.Ready property before
the first macro dispatch, at most eight times with 250 ms gaps. Only typed True
permits dispatch; unsupported types or exhaustion stop before the action. The
existing exact observational retries are unchanged; commands are never replayed.
`readiness-before-dispatch.jsonl` contains only fixed macro name, attempt and
status. This does not prove the cause of previous interruptions or guarantee a
future dispatch. Microsoft documents Ready as a read-only Boolean property:
<https://learn.microsoft.com/en-us/office/vba/api/excel.application.ready>.
No runtime/package changes. The full guide gate is restarted with this explicit
option and fresh fixtures against the integrity candidate; its result is pending.

The complete guide regression now passes **351/351**, retaining every prior
identity and all 46 published-edit checks, root
`ba5c5a4f5f3d4b2da6def3918080719e`, **2026-09-22
08:29:44.5889308--09:03:58.6378337 UTC**. All **52** current captures are directly
reviewed and hashed, including native maximize/restore, both presentations and
comparison, saved-result scrolling, missing/extra/rejected/no-conclusion cases,
capture-off/restriction/unavailable evidence and published editing/conflict.
Normal unassisted closure, unchanged running test/runtime/package pins and zero
matching Application events are verified by
`reports/runtime/verify-guide-action-curation-integrity-full-ready.ps1`.

The readiness trace has 1,339 property samples for 1,332 dispatches. One bounded
wait observes five Busy and two unavailable samples, then Ready on sample eight;
all other dispatches proceed after their first Ready sample. No macro call fails,
no observational retry is used and no operator action is replayed. This proves
the guard waited before dispatch in this run; it does not establish the cause
of earlier interruptions or prove a prevented failure. The initial final-audit
assumption that every sample was Ready is corrected from the retained trace,
without rerunning or altering the successful runtime evidence.
Fresh-process preference, Viewer, Boxing/Shipping, evaluation and the full Release
1 chain remain required on this candidate and are now running serially.

Fresh-process preference passes **27/27**, root
`d1b65731f506474cb0ae83565de4cc6f`, **2026-09-22
09:05:54.4938739--09:08:49.0217638 UTC**. A different verified Excel process
restores the saved Operations preference without an Admin dependency, automatic
pairing or inferred evaluation; training/config and saved probe bytes remain
unchanged. All three current captures are directly reviewed and hashed.

Viewer regression retains **94/94** prior identities, root
`8df6935681d0497285bbb8907cba731c`, **2026-09-22
09:08:50.2525609--09:11:27.0788155 UTC**. Filters, published-source integrity and
policy visibility, exact Shipping inventory key/state and read-only source access
remain protected. All three current captures are directly reviewed and hashed.
Both gates close Excel normally without assistance, preserve source/package pins
and have zero matching Application events. Verifiers are
`reports/runtime/verify-guide-action-curation-integrity-restart.ps1` and
`reports/runtime/verify-guide-action-curation-integrity-regressions.ps1 -Gates viewer`.
Boxing/Shipping, evaluation and full-chain gates remain in progress. Native
maximize/restore checks for the three curation surfaces will extend the existing
71-check route after these gates; this is missing test coverage, not product RED.

Boxing/Shipping retains **1,707 PASS / seven known pending D8-A FAIL**, all
**1,714** unique typed identities, root `a2f2ed1e08804268b368d3bb422d53f6`,
**2026-09-22 09:11:28.2058427--09:27:46.7966933 UTC**. Every prior GREEN remains
GREEN. The failure set is exactly `Shipping.Access.AuthUnavailable.<Add|Update|
Remove|Hold|Return|Stage|Send>.MissingFileNotRecreated`; no D8-A change is approved
or implemented. All **22** current captures are directly reviewed and hashed.
Normal unassisted Excel closure, unchanged running source/package pins and zero
matching Application events are verified by
`reports/runtime/verify-guide-action-curation-integrity-regressions.ps1 -Gates boxing`.

The serial queue stops after this completed gate because its outer wrapper
expects process exit 1 from a wrapper that already validated and recorded the
inner test's known-failure exit 1 but then returns normally. This is an
orchestration check error in an ignored runtime script, not a product regression or a reason to rerun the
verified matrix. Evaluation continues separately; full-chain validation follows
only after verified normal closure. The stopped queue and all evidence remain.

Evaluation completes **376/376**, retaining every preceding identity, root
`2fb2df7f872f4cd9b8d68606c0e71077`, **2026-09-22
09:28:24.3594705--09:52:52.3870677 UTC**. All **19** current captures are directly
reviewed and hashed. Actual Receiving source events progress through pending,
three-of-four applied and fully applied publication; only the last stage concludes.
Exact keys, original observations, unknown staging columns, unrelated workbook
contents and immutable results remain protected. One bounded foreground-activation
retry precedes the successful Applied/Minimum capture; it replays no workflow
action. Normal unassisted Excel closure, unchanged test/package pins and zero
matching Application events are verified by
`reports/runtime/verify-guide-action-curation-integrity-regressions.ps1 -Gates evaluation`.
The full Release 1 chain starts only after this completed verification.

Full Release 1 chain is **32/32**, live-role **48/48** and Create Warehouse
**15/15**, **2026-09-22 09:53:26.2209967--10:02:51.8506590 UTC**. Every preceding
chain/live-role identity is retained. The five frozen packages and **234** running
test/tool sources stay unchanged; all three tracked result files are restored
byte-for-byte. No matching Application 1000/1001/1002 events occur. Evidence is
`reports/runtime/guide-action-curation-integrity-chain-*`, verified by
`reports/runtime/verify-guide-action-curation-integrity-chain.ps1` before the
subsequent native-layout test-source extension.

Cleanup is **assisted**, not unattended: read-only inspection confirms one idle
instance has no workbooks and it exits without a Quit request. A later instance,
created during the isolated run, also has a typed zero workbook count and is
bound by exact native owner, process creation time and observed parent ID before
normal Quit. The Document Recovery dialog is inspected; **Yes, I want to view
these files later** is selected, captured, directly verified and confirmed.
Recovery files remain, no process is forcibly terminated and no workbook is
saved or modified by cleanup. The later instance's parent has already exited;
the evidence does not establish the cause of that recovery instance. The runner
then exits zero and Excel is closed. This does not prove unattended cleanup,
comprehensive control coverage, transfer or human acceptance.

The test-only native-layout extension passes **77/77**, preserving all **71**
preceding focused identities and adding six actual maximize/restore checks for
Action Paths, Choose tracked actions and Action Path guide. Root
`8e78bad1e2d445be8d4f60278b71b389`, **2026-09-22
10:04:18.7527167--10:08:41.7413401 UTC**. All **18** current captures are directly
reviewed and hashed; maximized captures are reviewed at original resolution.
Runtime sources and frozen XLAMs are unchanged. Excel closes normally without
assistance after the runner's bounded closure wait, and there are zero matching
Application events. Verification is `reports/runtime/verify-curation-native-layout.ps1`.
This fills existing D18 coverage; no behavioral RED is fabricated and no runtime
fix, control change or new architectural contract is introduced. The preceding
build/compile/static and complete supporting gates apply to the same package
hashes; the extension only changes its isolated focused test route.

### Operations guide comparison coverage (verified)

The new test route authors and explicitly pairs an immutable guide with the
real Receiving recording, then checks How-To, Diagnostic and Compare both at
pending, partial and fully applied publication. It retains the existing 376
evaluator assertions and adds 45 comparison assertions through the existing
packaged form handlers. The generated author receives an explicit
ACTION_PATH_MAINT fixture grant; ordinary Viewer permissions do not change.
The guide and observed recording are explicitly selected separately, even though
this fixture uses the recording from which the guide was authored. The existing
351-check guide gate retains the distinct-source/observed-run proof.

The first run, root `c88663be25874bfda5e5af349ddeac46`, is **incomplete: 242 PASS
and one harness failure**, 2026-09-22 10:19:28.0441952--10:31:03.0723888 UTC.
The eight-sample readiness guard exhausts before a guide call dispatches during
Partial; a subsequent cleanup guard also exhausts. The first exception is not
independently retained, so its full context is not inferred from the final
cleanup message. All running source pins remain unchanged, Excel closes normally
without assistance within the bounded wait, and no matching Application events
occur. Nine Pending captures are directly reviewed and hashed; five other
captures are not claimed reviewed. This is neither product RED nor acceptance.
Evidence: `reports/runtime/operations-guide-presentation-*` and that root's
`readiness-before-dispatch.jsonl` / `operator-image-review.json`.

The test-only readiness limit is now selectable from 1 through 120 samples,
default eight, with the unchanged 250 ms interval. The comparison retry explicitly
uses 40. Actual-Run offline RED is **37 PASS / 7 expected FAIL**, root
`39b9a1ce80c14d8086d3df14548c1a4e`; GREEN is **44/44**, root
`f7baa9f1dbd848c5bf9dc251adf6fa18`. All 32 prior readiness identities remain;
late readiness, bounded exhaustion and a failed single dispatch are covered.
Existing privacy/replay diagnostics remain **78/78**. Sampling never replays a
command and logs no argument values. No runtime or architectural contract changes.
Evidence: `reports/runtime/extended-readiness-{red,green,diagnostics-green}.log`
and `extended-readiness-verification.json`.

Regenerated maintenance remains **249 components / 6,035 procedures / 132,492
lines / dynamic calls 9 literal and 45 unresolved / 192 duplicate candidates**;
all 28 prior module limits and unrelated documentation pins pass. Evidence:
`reports/runtime/operations-guide-presentation-static-verification.json`.
The fresh extended-readiness packaged run passes **421/421**, retaining every
one of the 376 preceding evaluator identities and adding exactly 45 Operations
comparison checks. Root `957e171a8fa6454cb7994f6e857166d7`, **2026-09-22
10:34:24.4223736--11:06:16.3335952 UTC**. All **31** captures are directly reviewed
and hashed. The guide keeps its original authored instructions while the explicit
run/result shows Awaiting with zero, then three, applied source events, and
Concluded only when all four exact events are applied. Viewing, switching and
refresh preserve activity/training/config/publication bytes and authority-call
counters; the separate Evaluate action alone creates each saved result.

Excel closes normally without assistance within the bounded closure wait; all
test/runtime source pins and five frozen package hashes remain unchanged, with
zero matching Application events. All **1,377** readiness samples are Ready on
their first read. The 40-sample option is exercised by the offline tests, but this
live run does not prove the cause or recovery timing of the earlier exhaustion.
Verification: `reports/runtime/verify-operations-guide-presentation-extended-ready.ps1`
and `operations-guide-presentation-extended-ready-{verification,readiness-verification}.json`.
No runtime change or repeated full-chain run was needed for this test extension;
the preceding compiled packages, static limits, live roles and full-chain evidence
apply to the same frozen hashes. Comprehensive control coverage, training transfer,
unapproved decisions and human acceptance remain open.
