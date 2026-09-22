# Plan 022 Slice 4be.5 - direct tracked-action guide curation

Status: focused packaged **RED verified: 12 PASS / 41 expected FAIL**, 53 unique
typed checks. No direct-curation runtime change, GREEN, deployment or acceptance
is claimed yet. Architecture v4.11 D18's direct tracked-action curation refinement,
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
