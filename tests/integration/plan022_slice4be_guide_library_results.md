# Plan 022 Slice 4be.5 published-guide discovery and reading

**Focused RED: 102 PASS / 30 expected missing-reader FAIL; 132 total.**
All 97 preceding Save/draft/Viewer GREEN identities remain passing. No harness
failure occurs. Reader runtime implementation remains pending. Architecture v4.11 D18's
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
and diffs are checked. Raw reports remain ignored. No runtime reader changes,
GREEN, deployment, NAS or full Release 1 acceptance are claimed. Guide editing,
direct event/action curation, both presentations/comparison, expectations,
transfer, provisioning and the remaining release gates remain required.
