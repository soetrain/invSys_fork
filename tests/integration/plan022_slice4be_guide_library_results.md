# Plan 022 Slice 4be.5 published-guide discovery and reading

**Focused reader GREEN: 132/132**, retaining every protecting RED identity.
Report:
`reports/runtime/slice4be-viewer-published-read/7391802610464be895be0c9556212180/green.json`.
Terminal exit is 0 and Excel closes. All 25 current/frozen package hashes and
both unrelated documents are unchanged. The build, compile and GREEN windows
have zero Application events 1000/1001/1002; the build's immediate cleanup was
still pending, so its audit extends through the subsequent no-Excel compile
preflight. The standard compile harness retains its existing owned-empty-process
cleanup fallback; this is not proof that no forced cleanup was available.
At that gate, reader/Save captures were pending. Broader recording/evaluation/
role/full-chain acceptance remains open. No accepted deployment has changed.

**Visible focused gate: 135/135**, retaining all 132 preceding identities plus
both actual Save-version captures and the published-reader capture. Report:
`reports/runtime/slice4be-viewer-published-read/150f1d31df9a4cac89ae744233b77f22/green.json`.
All three images were directly reviewed: publication wording and version labels
fit, both revisions are listed, and authored instructions/Unicode text remain
separate from original attempt/result observations. Terminal exit is 0 with
Excel closed. Twenty-five package hashes and both unrelated documents remain
unchanged. The UTC 2026-09-15 07:53:03.2197121--07:59:01.6108670 window has zero
Application events 1000/1001/1002. Refreshed static evidence at
`reports/runtime/guide-library-capture-static` retains the metrics/limits below.
This proves current candidate Save/reader screenshots, not the cause of the
earlier Event Detail foreground failures or human acceptance. Serial recording,
detail, Viewer/Shipping state, Boxing/Shipping, evaluation and full-chain gates
were started serially. Completed regression results follow; D8-A is unapproved.

**Completed reader regressions (2026-09-15):** Combined recording/guide authoring/
Save/reader passes **170/170**, retaining all 105 prior combined identities;
Event Detail passes **34/34**, retaining all 34; Viewer/filter/Shipping state
passes **94/94**, retaining all 94. Exact reports, respectively:

- `reports/runtime/slice4be-viewer-published-read/bba8c1423b9e489f822641c862d8c95e/green.json`
- `reports/runtime/slice4be-viewer-detail/26784e6f223946db9191ac9574ad14e7/green.json`
- `reports/runtime/slice4be-viewer-published-read/ae7285756ad545899c84930281ed06a0/green.json`

Each controller exits 0 with Excel closed and no Application events
1000/1001/1002 in its measured window. The combined windows span UTC
08:00:06.7586898--08:11:34.4582673. Five reader candidate package hashes and both
unrelated documents are preserved. Verification is recorded in ignored
`reports/runtime/guide-library-first-regressions-verification.json`.
These results do not establish evaluation or full-chain outcomes; log counts
alone are not terminal evidence. All six detail/Viewer regression captures were
directly reviewed: Event Detail default/maximized/restored, activity captions,
event filters and Shipping held/pending status. The long Coverage field is
visibly clipped at default/restored width and readable when maximized. Record
this as an open visible-readability issue, not complete default-width acceptance;
the screenshots do not establish the cause of older foreground failures.

**Boxing/Shipping regression:** **1,707 PASS / seven known D8-A FAIL**, retaining
all 1,714 preceding identities and all 18 owner-return observations. Report:
`reports/runtime/slice4be-shipping-activity/29cca59b1d8e497e8d6a764ce5fd947c/boxing-activity-shipping-recording-green.json`.
The seven failures are AuthUnavailable MissingFileNotRecreated for Shipping Add,
Update, Remove, Hold, Return, Stage and Send. Terminal exit is 1 as expected for
those failures. Its immediate ExcelClosed snapshot is false; the same owned
Excel process subsequently exits and the serial evaluation preflight verifies
no Excel before starting its own instance. The native-event audit therefore
extends through that preflight: UTC 08:11:34.4846768--08:28:32.5190437, zero
Application events 1000/1001/1002. Five candidate hashes and both unrelated
documents remain unchanged. Ignored verification:
`reports/runtime/guide-library-boxing-verification.json`. This is regression
preservation, not D8-A approval or full-chain acceptance. Boxing captures from
this run still require direct review; evaluation and full chain remain pending.

**Candidate implementation:** `deploy/validation-guide-library`
builds all five packages and passes explicit compile plus Operations cold start.
The compiled comparison against Save changes only `modActionPathRead` and
`frmActionPaths`, adding `modGuideLibraryRead`, private `modTrainingReadContext`
and `frmActionPathLibrary` (231 to 234 compiled components). Core shares the
existing recording context/policy guard, validates full predecessor chains and
filters retained content under current policy. Operations owns the reusable
reader and exact version/hash selection. No guide record/schema changes.

Static maintenance completes at 241 components / 5,951 procedures / 131,237
lines: growth of three components, 22 procedures and 345 lines. Duplicate groups
decrease from 194 to 193; literal/unresolved dynamic calls stay 9/45. All 28
existing module limits hold. Focused behavioral and capture gates are recorded
above; broader regression/full-release acceptance is not yet established.

**Focused RED: 102 PASS / 30 expected missing-reader FAIL; 132 total.**
All 97 preceding Save/draft/Viewer GREEN identities remain passing. No harness
failure occurs. This preceding unchanged-candidate gate authorized implementation.
Architecture v4.11 D18's
published-guide reader refinement names the Operations entry/reader and exact
version selection contract. It implements the approved searchable guide library
under semantic inheritance. No guide schema or architectural exception changes.

The protecting `Slice4beGuideLibrary.ps1` test runs inside the existing actual
Save guide fixture. Both revisions come from packaged author controls, with
different guide names and reversed authored steps. Original observations remain
unchanged. The test-only control probe reads list values and dispatches real
controls; it does not create a reader, guide or successful service response.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-guide-save-identity -Phase RED `
  -GuideDraftOnly -CheckGuideLibrary -CheckViewerPublishedRead -CompileViewerProbesForTest
```

This retains the preceding 97 Save/draft/Viewer identities. New cases cover
entry/reuse, both exact version/hash keys, authored and original observation
ordering, provenance, locked panes, exact selection through Refresh, name/tag/ID
search, four layouts, current-policy restrictions, corruption of a selected
version/predecessor, a missing predecessor, recovery after byte restoration,
Close/Viewer ownership, target invalidation, empty-folder reads, source-byte
preservation and no business publication. A separate call runs while the fixture
remains signed in after its maintenance grant is revoked, protecting ordinary
reader access independently from guide editing. The existing author fixture
restores the exact Auth bytes and session after that case.

Only generated disposable guide files are temporarily corrupted or withheld;
each is restored in `finally`. No operational workbook is replaced. Add
`-CaptureEvidence` for a visible reader capture alongside the existing draft/Save
captures. Automated geometry does not establish human acceptance.
The test-only `-CaptureGuideEvidence` option enables the two actual Save captures
and published-reader capture while retaining the existing ownership/foreground
checks. It leaves unrelated screenshot stages disabled. The dedicated visible
result above and preceding 132-check behavioral result remain separately scoped.

Expected behavioral RED is missing entry and reader behavior on the unchanged
Save candidate. Broken Save fixtures, compile failures and harness exceptions
are not accepted RED. The unchanged candidate produces the verified RED above,
including missing ordinary-reader access while the maintenance grant is revoked.
The five new passing checks protect unchanged Config/all warehouse bytes,
Shipping authority-call count, source/activity bytes and business publication.
Report:
`reports/runtime/slice4be-viewer-published-read/2c441f1f447849d88e6e3d283d3e092c/red.json`.
The controller exits 1 as expected, with Excel closed. Twenty current/frozen
package hashes and both unrelated user documents are unchanged. The UTC
2026-09-15 07:30:59.4584702--07:36:25.3513424 window contains zero Application
events 1000/1001/1002. Four changed PowerShell scripts parse; local document links
and diffs are checked. Raw reports remain ignored. That RED checkpoint did not
claim runtime reader implementation or GREEN. Deployment, NAS and full Release 1
acceptance remain open. Guide editing,
direct event/action curation, both presentations/comparison, expectations,
transfer, provisioning and the remaining release gates remain required.
