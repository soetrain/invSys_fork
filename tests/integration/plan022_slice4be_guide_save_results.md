# Plan 022 Slice 4be.5 immutable guide saving

**Current focused result: 97/97 GREEN** on isolated
`deploy/validation-guide-save-identity`, retaining every protecting RED identity.
Five-package build/explicit compile/cold start, four editor layouts, static limits,
source preservation and terminal Excel closure pass. Current Save guide screenshots
and broader regression/full-chain/deployed/NAS/human acceptance remain open.

**Evaluation regression:** The unchanged Save candidate passes **376/376**,
retaining all 376 preceding evaluation identities, including five instrumented
compiles, actual Operations/Admin recording, asynchronous publication, retries,
policy/stale-evidence handling, selection/intent/version races and expectation
editor behavior. Report:
`reports/runtime/slice4be-viewer-published-read/7433add3a00443cdbc0ad241cb71222a/green.json`.
The controller exits 0 with Excel closed. All five candidate package hashes and
both unrelated user documents remain unchanged. The UTC 2026-09-15
07:11:46.3799770--07:30:16.0883292 window has zero Application events
1000/1001/1002. The local verifier initially miscounted a PowerShell 5 JSON array
wrapped in `@(...)`; correcting that verifier confirms the original 376/376
terminal report, without rerunning or changing product behavior.

All 15 diagnostic captures were directly reviewed: pending, partially applied
and fully applied, each at minimum/default/larger/restored sizes plus the source
viewport. Pending and partially applied runs retain Awaiting published result;
only all four applied terminal references show Conclusion observed. Stopped
remains capture lifecycle, and source/evaluation provenance stays visible. This
protects the shared Guides/Evaluations folder helper and preceding diagnostic
presentation. It does not supply the still-missing Save editor captures or human
comparison/full Release 1 acceptance. Raw captures/reports remain local and ignored.

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

The RED test checkpoint is **e96863d**, with synchronized specification/plan/
controls at docs **935f85c**, both pushed before runtime implementation. Historical
preservation additionally verifies 325 package hashes and 16 protected source
checks; those historical maintenance metrics are not this candidate's metrics.

Implementation is now under test in isolated `deploy/validation-guide-save`.
Core adds `modGuideModel` and `modGuideStore`; `modActionGuideDraft.SaveDraft`
revalidates the captured source/session/policy/capability, stages a separate guide
model, and appends a version without overwriting the previous record. The Operations
editor adds Save guide and permanent publication wording. `modGuideDraftSource`
renders an unclosed source as Interrupted. The fixed Guides/Evaluations child-root
check is shared through `modRecordingJournal.ChildRoot`; the evaluation store's
same folder/target/reparse behavior requires its existing regression gate.

All five explicit package compiles and Operations cold start pass. The compiled
comparison retains 229 preceding component identities, adds only the two guide
modules, and changes only the draft/source, journal/evaluation-store and guide-form
components. Compilation preserves all five candidate package hashes. Fresh static
evidence has 238 components, 5,929 procedures and 130,890 lines: growth of two
components, eight procedures and 299 lines. Duplicate groups remain 194 and literal/
unresolved Application.Run counts remain 9/45; all 28 existing module limits hold.
See ignored `guide-save-compiled-delta.json` and `guide-save-static-verification.json`.
The first GREEN attempt stops at **15 PASS / one foreground-capture harness FAIL**
before guide cases, at
`reports/runtime/slice4be-viewer-published-read/13eaadffee654bfaa3f0301dfead7e32/green.json`.
Exit 1 and Excel closure are verified. This is not save-behavior evidence. The
capture setup now checks that the callback's window handle matches the unique
visible form in the disposable Excel process, activates that form, and retains
the exact foreground check before capture. Only foreground activation may retry,
at most three times; a wrong/missing owned window still fails. The same setup is
used for the existing published Viewer capture and two optional post-save captures.
It does not alter runtime code, invoke a business control or substitute rendering.
The corrected capture attempt also stops at 15 PASS / one foreground harness
failure, at `reports/runtime/slice4be-viewer-published-read/52187c9e5c194528b0692ae9986e8d9d/green.json`.
Ownership validation succeeds but foreground acquisition does not, even after the
bounded attempts; exit 1 and Excel closure are verified. A read-only desktop check
finds an accessible input desktop with Code in the foreground. That observation
does not prove why activation failed. Visible evidence remains unproven.
The two post-save captures supplement the five existing draft captures; they do
not replace the 97 protecting behavioral checks. No persistence GREEN or completed
slice is claimed yet.

The separate behavioral attempt completes **80 PASS / 17 FAIL** across all 97
checks, at `reports/runtime/slice4be-viewer-published-read/83acdd3f4ad04dc99e6191172e88c1c5/green.json`.
The Save control, wording, blank/empty-step validation, policy/capability/target
guards, source preservation, layouts and Interrupted source label pass. First
publication and its dependent version/provenance assertions fail; the generated
fixture has no Guides folder or guide record. This is a runtime save defect,
separate from capture setup. A RED-only stage/error-number trace in the unsaved
test package copy locates `modGuideModel.Validate:58;error=0`: the cross-namespace
identity rejection. The original activity writer intentionally uses an attempt's
ActivityId as its REQUESTED RecordId. All five generated attempts inspected have
that relationship; no IDs or record values were emitted. The trace run terminates
at the same 80 PASS / 17 FAIL, exit 1 with Excel closed, at
`reports/runtime/slice4be-viewer-published-read/2f2ce18b6f1947f3bb53557e440e6890/red.json`.
The normative specification, plan and controls clarify preservation of that
existing source relationship before the validator correction. Record uniqueness
and distinct guide/Step identities remain enforced. Diagnostic instrumentation
is removed; its ignored copy is `guide-save-diagnostic-trace-probe.ps1`.

The corrected candidate is `deploy/validation-guide-save-identity`. Its static
baseline has 130,892 lines (301 above the draft baseline), with the other reviewed
metrics unchanged and all 28 module limits retained. All five packages build and
explicitly compile, Operations cold start passes, and compiled scope remains the
same five changed/two new components. The first corrected run reaches 73 PASS,
including first publication, then a harness exception at PowerShell's COM custom-
property lookup. That incomplete attempt is retained at
`reports/runtime/slice4be-viewer-published-read/07c871d315c74393882c75deb3ca8865/green.json`,
exit 1 with Excel closed. The expected publisher identity now comes from the built
XLAM's `docProps/custom.xml`, independently of the runtime saver. The same behavioral
test is running again; full GREEN remains pending. Capture calibration on an empty disposable
form can acquire foreground by both title and process; this does not prove product
capture. Two earlier calibration attempts fail their window-identity setup before
activation, and are preserved separately. Their correction puts the native null
class-name argument inside C# instead of PowerShell string conversion. No product
control, operational workbook or original training record is changed by calibration.

The corrected metadata test completes **97/97 GREEN**, at
`reports/runtime/slice4be-viewer-published-read/3256a696eb2c4508877324c350316c31/green.json`.
Both immutable guide versions, exact preceding links, independent package/source
provenance, authored Unicode/text/tags, stable step IDs, original observation order,
no inferred expectation, oversize/conflict rejection, Cancel, current permission/
policy/target guards and exact source preservation pass through actual controls.
All four editor layout checks include the new Save button and publication label.
Terminal exit is 0 with Excel closed. This run intentionally has no screenshots.

`reports/runtime/guide-save-green-verification.json` verifies all 97 protecting
identities, 20 current/frozen package hashes, both unrelated user documents, the
reviewed compiled scope/static metrics and all 28 module limits. Ten audited
build/compile/test windows contain no Application events 1000/1001/1002; build or
compile cleanup intervals extend to the next verified no-Excel preflight where
the immediate exit still observed Excel. No native-recovery repair is claimed.
The shared evaluation-folder helper still requires the broader evaluation gate,
and all preceding role/Viewer/recording/full-chain gates remain required for this
candidate. Focused saving is proven; full guide discovery/editing/presentation/
comparison/transfer and complete Slice 4be acceptance are not yet achieved.

After meaningful RED, implement and prove the new save boundary through the same
handlers, including appropriate packaged, compile, layout, static, role and full
chain evidence. Guide discovery/editing, guide expectations, How-To/Diagnostic/
Compare selection, direct event/action curation without a prior recording,
origin-only transfer, user capability provisioning and full
deployed/NAS/human Release 1 acceptance remain required beyond this save gate.
