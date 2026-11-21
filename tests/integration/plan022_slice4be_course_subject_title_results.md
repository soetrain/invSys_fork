# Slice 4be-A: Action Path course Subject and Title

Architecture v4.11 D18-COURSE-01 was approved on 2026-11-21. Subject groups
authored guides; the existing guide `Name` is displayed as Title. Raw recording
journals remain evidence without course labels. Plan 022 and Controls 1.793
are synchronized. D18-RECORDING-LIST-01 remains a separate unapproved proposal.

## Focused behavioral RED, 2026-11-21

The existing packaged guide Save route used the five unchanged candidate XLAMs
in `deploy/validation-action-search-final-20261120`. It compiled all five
instrumented packages, created a real source recording through Viewer/Admin
handlers, opened the real Action Path guide editor, saved a guide and exercised
the retained published reader and expectation editor. Result: **171 PASS / 3
behavioral FAIL**. The exact failures were:

- `GuideCourse.SubjectFieldAvailable`: the current editor has no Subject field.
- `GuideCourse.TitleLabel`: it still says Guide name.
- `GuideCourse.SavedSubjectAndTitle`: the saved guide is schema 1 and has no
  Subject, while the existing Name contains the authored title text.

All existing GuideSave, GuideLibrary and GuideExpectation identities in this
run passed. Excel exited normally and the five source XLAM hashes remained
unchanged. Local ignored evidence:
`reports/runtime/slice4be-viewer-published-read/d97309824da84ef8b5ca83c6e3d8dfc1/red.json`.
The first dispatch command stopped before Excel at a documented controller
preflight because SavedCopies requires the isolated expectation/compiled route;
it was not product RED. The corrected packaged run used:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-action-search-final-20261120 -Phase RED -GuideDraftOnly -CheckGuideSave -CheckGuideExpectation -CheckViewerPublishedRead -CompileEvaluationProbesForTest -ViewerStartupPackageStateForTest SavedCopies -WaitForExcelReadyForTest -ExcelReadyReadLimitForTest 40
```

## Packaged continuation

The first schema-3 build passed its Subject/Title editor and save checks but
exposed retained fixture and larger-layout failures. After those corrections,
a behavioral list/search RED on the pre-library-change package reported
**174 PASS / 3 FAIL**: missing four-column course display, missing Subject
search and the known pre-fix larger editor layout. The list and search
failures were exact named behavior checks; all retained library checks passed.
Ignored evidence: `reports/runtime/slice4be-viewer-published-read/a4c0d3e56da04016818fb7a8d57d1e34/red.json`.

The rebuilt five-XLAM candidate compiled and passed **177/177** through the
packaged Viewer, guide editor, Save, Published guides and expectation handlers.
This includes Subject/Title persistence, current-policy search, minimum,
default and larger layouts, exact version selection and read-only source
preservation. Ignored evidence:
`reports/runtime/slice4be-viewer-published-read/403efe35e8c74dd08aced8ae708f7a94/green.json`.
After the transfer-origin correction, the final five-XLAM candidate repeated
this route at **177/177**, with natural Excel closure:
`reports/runtime/slice4be-viewer-published-read/5f5958c325bf4b0da3fae3471d9ab17f/green.json`.

The first guide-transfer run stopped before transfer checks because its older
guide-authoring fixture did not fill the new required Subject; it was not
transfer RED. After fixture correction, an imported schema-3 editor omitted
origin wording (68 PASS / 1 FAIL). The Core provenance check was corrected;
the rebuilt candidate passed **69/69** through actual Export, Import, Edit and
re-export handlers, including Subject/Title round trip, exact origin binding,
policy refusal and source-byte preservation. Ignored evidence:
`reports/runtime/slice4be-guide-transfer/c5c5312c11a24f7aaff09c24b2306e92/green.json`.

The real picker/editor route passed **61/61**: incremental Subject typing was
responsive with no Core write per character, focus exit flushed once, dirty
Cancel/X and stale authority did not publish, and two distinct Subjects sorted
ahead of creation order in Published guides. Ignored evidence:
`reports/runtime/slice4be-viewer-published-read/09d60fbc615341c8b79384346e673381/green.json`.
Direct tracked-action guide creation passed **70/70** on the current candidate,
including an explicit schema-three assertion. All tests ended with Excel
closed naturally.

Published-guide Edit passed **60/60** on the same candidate. The actual editor
reopened the exact Subject and Title, changed both in an immutable schema-3
successor, retained the predecessor hash and original observations, then
reopened the new Subject and Title. Ignored evidence:
`reports/runtime/slice4be-viewer-published-read/578fd5fc6f8d484c88ab2624fef6b29f/green.json`.

The static inventory at `reports/runtime/action-course-static-20261121`
reports 391 components and 6767 procedures. Versus the 2026-11-20 scan,
duplicate candidates remain 193, same-project late bindings remain 7,
unresolved candidates remain 290 and warnings remain 30. Twenty changed
PowerShell tooling files parse with zero errors; `git diff --check` is clean.

## Open gates

Prove legacy schema 1/2 read/edit without rewriting historical guide bytes,
run broader Release 1 regressions, and obtain visible NAS course usability
evidence. This is not guide-course acceptance, Inventory GREEN or 4be-A
completion.
