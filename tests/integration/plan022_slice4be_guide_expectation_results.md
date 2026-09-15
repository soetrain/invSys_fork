# Plan 022 Slice 4be.5 guide expectation authoring

**Focused implementation GREEN: 166/166.** The isolated
`deploy/validation-guide-expectation` candidate implements the approved explicit
guide-draft expectation scope in Core and Operations. All 166 protecting check
identities pass, including the preceding 132 reader/Save/draft/Viewer checks.
Report:
`reports/runtime/slice4be-viewer-published-read/ff91f4fc88104eeda788a6cbe13c2811/green.json`.
The controller exits 0 at UTC 2026-09-15 09:23:55.4808153 with immediate Excel
closure false; a subsequent process inspection verifies normal closure at
09:24:45.4286630. No forced termination or additional Quit was used.

The seven runtime files add the guide entry/summary, explicit shared-editor
scope and independent staged definition. New guides start at None. Use stages
only; Save deep-copies the definition into an immutable version. Cancel retains
previously staged intent, while guide closure, policy/capability/context loss
clear pending intent. Starting/stopping a recording and switching between guide
and run editors preserve their distinct scopes. Exact prior-version links,
expected-step identities, instruction identities and source observations pass.

The five isolated packages build and compile, and Operations cold start passes.
Compiled-source comparison retains all 234 components and changes exactly the
seven expected components, comparing case-insensitive VBA code and exact string
literals. The build's immediate cleanup check was false; the following compile
preflight verified no Excel. The compile harness retains its existing guarded
owned-empty-process fallback, so this is not proof that forced cleanup is absent
from that harness. Accepted deployment, installed packages and NAS are untouched.

The first combined capture attempt ends with **81 PASS / one harness FAIL** at
the existing first Save-guide image: `SaveVisibleWindow` reports **Requested form
is not in the foreground**. It produces no PNG and is not product RED. Report:
`reports/runtime/slice4be-viewer-published-read/a2e948883dfd467ea71e54dc458765aa/green.json`.
Exit 1, Excel closed; its UTC 09:09:28.5284413--09:12:52.5528332 window has zero
Application events 1000/1001/1002. The test-only diagnostic now records intended
and foreground root process/window/class identities and AppActivate's result on
that failure, without captions or workbook values. Existing ownership checks,
activation attempts, foreground requirement and retry limit remain unchanged.
PowerShell parsing and standalone compilation of the native helper pass.
The separate diagnostic capture gate also ends **81 PASS / one harness FAIL**
at the same first image, with no PNG. Report:
`reports/runtime/slice4be-viewer-published-read/fa883641b2e6485faaa55661ab440d53/green.json`.
All three observations report successful AppActivate but a different foreground
root: the intended Excel `ThunderDFrame` remains behind a VS Code
`Chrome_WidgetWin_1` window. The observation identifies the competing window,
not the cause of activation failing to take effect. Exit 1, Excel closed; UTC
09:25:01.4855933--09:28:27.1978119 has zero Application events 1000/1001/1002.
No product change, foreground bypass or repeat without a new diagnostic is
justified by this result. Visible acceptance remains open.

Refreshed static evidence is **241 components / 5,958 procedures / 131,327
lines**; duplicate groups decrease **193 to 192**, dynamic calls remain
**9 literal / 45 unresolved**, and all **28** existing module limits pass.
The reviewed backlog contains 1,186 candidates versus 1,184 scanner candidates;
no candidate is deleted automatically. All **30** current/frozen package hashes
and both unrelated user-document hashes are preserved. Four PowerShell scripts
parse, 89 local links resolve, and source scope remains exactly seven runtime
files plus the test diagnostic. The build, compile and focused behavior windows
each have zero Application events 1000/1001/1002, with the build window extended
through compile's Excel-closed preflight and behavior through observed closure.
Ignored verifiers:
`reports/runtime/guide-expectation-behavior-green-verification.json`,
`reports/runtime/guide-expectation-implementation-checkpoint-verification.json`,
`reports/runtime/guide-expectation-visible-green-capture-failure-verification.json`.
These establish focused acceptance only; broader candidate regressions remain
required, including the combined recording/guide/expectation route below.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-guide-expectation -Phase GREEN `
  -GuideDraftOnly -CheckGuideExpectation -CheckViewerPublishedRead -CompileViewerProbesForTest
```

Omit `-GuideDraftOnly` for the combined recording regression. This retains the
preceding 170 combined checks and adds the 34 guide-expectation checks; its
result must be verified separately before broader acceptance is claimed.

The following RED is retained as the test-first antecedent, not current runtime
status. Full-chain Receiving failure, guide-bound evaluation, both presentations
and Compare, closed-guide editing, direct curation, transfer and release UAT
remain open.

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
This established behavioral RED before the implementation and GREEN above.
The reader candidate's separate full-chain failure is not this RED.

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
reader candidate. The completed RED confirms that expected failure. Execution began after the
preceding serial regression runner was terminal and Excel closure was verified.

Adding `-CaptureGuideEvidence` retains the existing ownership/foreground checks
and captures the four expectation-editor sizes, staged guide summary and second
published version, alongside the existing Save/reader captures. These optional
captures remain subject to the separate visible gate described above.

Guide-bound evaluation, How-To/Diagnostic/Compare and preference application,
closed-guide editing, direct event curation, import/export, remaining control
coverage and full Release 1 acceptance remain required. This test entry does not
replace them with recording-only analysis or redefine completion.
