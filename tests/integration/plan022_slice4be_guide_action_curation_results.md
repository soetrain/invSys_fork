# Plan 022 Slice 4be.5 - direct tracked-action guide curation

Status: initial isolated implementation compiles but remains failing; source-adapter
follow-up has verified **24 PASS / 31 expected FAIL**, 55 unique typed checks.
No GREEN, deployment or acceptance is claimed. Architecture v4.11 D18's direct tracked-action curation refinement,
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
