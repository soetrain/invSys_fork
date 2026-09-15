# Plan 022 Slice 4be.5 immutable guide saving

Architecture v4.11 D18's immutable guide-save refinement names Save guide,
publication wording, the fixed `Guides` storage child and its schema. Headless
Core owns current-context/capability/policy validation and atomic publication;
Operations owns the captured editor. A saved guide version never replaces its
source observations, publishes business Events or establishes a diagnostic
conclusion. This is a discovered implementation detail under approved semantic
inheritance, not a changed architectural authority.

`Slice4beGuideSave.ps1` is the protecting packaged test. It starts from the accepted
draft editor, records two real Admin Save Value actions, freezes and selects their
six-entry recording journal, and uses the actual guide controls. Its file reader
checks independently hashed saved records; it does not construct a substitute
guide or invoke an invented save result.

The test covers explicit publication wording, blank-name/empty-step rejection, first save,
distinct guide identity, exact source version/hash, authored fields and tags,
stable StepIds, unchanged observations, no inferred expectation, a second immutable
version with exact prior links, authored reorder, an occupied version filename,
whole-record oversize rejection
without draft truncation, Cancel preserving saved versions, and no business Event
publication or source/config/activity changes. Additional cases change policy
through actual Settings, revoke only the disposable author fixture's maintenance
grant, and switch to the other generated target before attempting Save. Fixture
credentials are never emitted; the generated Auth file is restored after revocation.
The recorded-source case also temporarily withholds the generated closing entry,
requires the actual guide source label to say Interrupted, then restores exact
source bytes. This protects D18's existing saved-run lifecycle rule; it is not a
new lifecycle interpretation or a reconstruction of the missing close.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-guide-draft-policy -Phase RED `
  -GuideDraftOnly -CheckGuideSave -CheckViewerPublishedRead -CompileViewerProbesForTest
```

This route retains the 67 focused draft/Viewer checks. Add `-CaptureEvidence` for
the five existing draft captures. Omit `-GuideDraftOnly` for combined recording
regression. The new Save control and persisted guide versions must be missing
behavioral RED before implementation; compile errors or broken fixtures do not
qualify. The first attempt with captures terminates at **15 PASS / one harness
FAIL**: the existing Event Detail screenshot rejects a form that is not in the
foreground, before any guide-save case runs. All five instrumented compiles pass.
This is not behavioral RED. The original report remains at
`reports/runtime/slice4be-viewer-published-read/044853132ff545a9a2f2c1dfb3f56b0a/red.json`;
terminal exit is 1 with Excel closed. A separate attempt uses the command above
without captures to establish behavioral evidence. Layout/capture acceptance
remains required independently. No guide persistence runtime is implemented;
the frozen draft candidate and its known D8-A limitations remain unchanged.

The second attempt reaches **69 PASS / 24 missing-save behavioral FAIL**, then
one harness failure reopening the next guard fixture. Its permission-restoration
sign-in creates a new session while the test retains the previous Viewer's
captured binding. The test now closes that Viewer before re-sign-in, verifies
the restored capability and exact Auth bytes, then opens Viewer for the new
session. No product binding rule is weakened. The incomplete attempt remains at
`reports/runtime/slice4be-viewer-published-read/873443f408c344e8ad9f07fb049bcee6/red.json`;
exit 1 and Excel closed are verified. A complete focused RED still requires all
guard and interrupted-source cases to run without a harness error.

The corrected attempt completes all **97 checks: 71 PASS / 26 expected FAIL**.
Every one of the 67 preceding Viewer/draft identities passes. All failures are
`GuideSave.*`: the missing Save control, publication wording and immutable records,
their associated rejection/status requirements, and the existing unclosed-source
label showing Recording instead of Interrupted. Source/config/activity preservation,
no business publication, source guards and exact interrupted-fixture restoration
pass. There is no harness exception. Exact report:
`reports/runtime/slice4be-viewer-published-read/d53b95e5ac394c9dbd22981b10fcee1c/red.json`.
Terminal exit 1 and Excel closure are verified. The verifier
`reports/runtime/guide-save-red-verification.json` proves retained check identities,
ten current/frozen package hashes, both unrelated user documents and no runtime
source changes. All three test windows have zero Application events 1000/1001/1002.
The three changed test scripts parse, and 87 local documentation links resolve.
This is the pre-implementation D13 checkpoint, not guide-save GREEN or full UAT.

After meaningful RED, implement and prove the new save boundary through the same
handlers, including appropriate packaged, compile, layout, static, role and full
chain evidence. Guide discovery/editing, guide expectations, How-To/Diagnostic/
Compare selection, direct event/action curation without a prior recording,
origin-only transfer, user capability provisioning and full
deployed/NAS/human Release 1 acceptance remain required beyond this save gate.
