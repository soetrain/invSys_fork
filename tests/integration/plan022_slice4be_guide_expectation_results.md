# Plan 022 Slice 4be.5 guide expectation authoring

**Focused RED: 136 PASS / 30 expected FAIL; 166 total.** Every preceding 132
reader/Save/draft/Viewer identity remains passing, with no duplicate checks or
harness exception. The unchanged reader candidate lacks the actual guide entry,
guide-specific shared-editor scope, staged summary, cancellation/invalidation and
saved expectation behavior. Four new preservation checks pass independently.
Report:
`reports/runtime/slice4be-viewer-published-read/1515cc3719c745afa21fa5452e43c594/red.json`.
Terminal exit is 1 as expected, with Excel closed. All 25 frozen package hashes
and both unrelated documents are unchanged. UTC 2026-09-15
08:52:30.7577605--09:00:26.0713116 has zero Application events 1000/1001/1002.
Ignored verification: `reports/runtime/guide-expectation-red-verification.json`.
This establishes behavioral RED before runtime changes; implementation and GREEN
remain next. The reader candidate's separate full-chain failure is not this RED.

Architecture v4.11 D18's
guide-expectation authoring refinement names Expected conclusion and the guide
summary, reusing the existing expectation editor with an explicit guide-draft
scope. Use for this guide stages intent; Save guide publishes that definition
with an immutable version. There is no new record schema or inferred success.

The packaged `Slice4beGuideExpectation.ps1` test creates a real two-action Admin
recording and explicitly authors its captured expectation through the existing
Viewer control. It then creates a guide through the existing author entry.
This lets the new test distinguish an initially empty guide definition from
the recording's already populated expectation. The generic disposable probe
delivers real controls, including Boolean retry selection; it constructs neither
an editor nor an expectation model. Instrumented packages must compile first.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-guide-library -Phase RED `
  -GuideDraftOnly -CheckGuideExpectation -CheckViewerPublishedRead -CompileViewerProbesForTest
```

This route retains the 132 reader/Save/draft/Viewer checks and adds guide-scope
wording, initial None, editor reuse, registered repeated actions, default/explicit
retry, exact expected IDs and terminal selection, Cancel before and after staging,
four editor layouts, unchanged run-analysis intent and safe scope switching.
An explicitly started/stopped empty recording must leave pending guide intent
intact; its new journal is fixture activity, never guide-triggered workflow replay.
Two actual Save clicks must persist distinct immutable expectation definitions
with exact prior links, surviving expected IDs and unchanged original observations
and instruction IDs. Cancelling the guide clears its pending expectation editor;
fresh drafts cannot inherit intent. Actual policy change, same-session maintenance
revocation and target change must invalidate pending guide intent. Generated Auth
bytes and source files are restored and verified, without emitting credentials.

The existing captured-expectation and guide Save fixtures are required setup.
Their failure or a compile/harness exception is not meaningful RED. Expected RED
is missing guide expectation entry, scope, staging and persistence on the unchanged
reader candidate. The completed run above confirms that expected RED. No new
runtime implementation or GREEN is claimed yet. Execution began after the
preceding serial regression runner was terminal and Excel closure was verified.

Adding `-CaptureGuideEvidence` retains the existing ownership/foreground checks
and captures the four expectation-editor sizes, staged guide summary and second
published version, alongside the existing Save/reader captures. These optional
captures are prepared, not yet executed or accepted.

Guide-bound evaluation, How-To/Diagnostic/Compare and preference application,
closed-guide editing, direct event curation, import/export, remaining control
coverage and full Release 1 acceptance remain required. This test entry does not
replace them with recording-only analysis or redefine completion.
