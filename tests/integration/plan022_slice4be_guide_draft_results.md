# Plan 022 Slice 4be.5 guide draft entry

## Contract and scope

Architecture v4.11 D18's guide draft-entry refinement and synchronized Plan 022 /
controls name the actual Create guide control and Operations guide editor. Core
validates the originating Viewer/session/warehouse, maintenance capability and
exact source run/version. Authored steps and text remain separate from immutable
observed controls/outcomes. Editing and cancellation do not save training records,
publish events or rewrite source evidence. The original observation order remains
independent of guide step order and omissions.

This test entry advances authored How-To guides. It does not accept guide persistence,
version/publication/search, How-To/Diagnostic/Compare, export/import, comprehensive
coverage or Release 1. No runtime implementation or accepted package is changed.
The frozen packaged baseline is code **9bfac38**, isolated
`deploy/validation-event-detail-labels`; five package hashes are retained in ignored
`reports/runtime/guide-draft-baseline-package-pins.json`.

## D13 protecting route

`Slice4beGuideDraft.ps1` installs a disposable facade before forms are loaded. It
only inspects actual controls or delivers their normal Click/Change actions. It
does not construct a guide model, return invented source evidence or impersonate
an owning handler. Admin-generated disposable fixtures explicitly grant the author
ACTION_PATH_MAINT; bootstrap Admin permission alone does not imply that capability.
The ordinary reader retains its existing capabilities without guide maintenance.

The test records two real Admin Settings Save Value actions, stops through the
normal recording control, validates the six-entry immutable journal and selects
the run through the actual Action Paths library. Those prerequisites must succeed
before missing Create guide/editor behavior is treated as product RED. The focused
route preserves the packaged published Viewer prerequisites; the combined route
also retains the broader recording checks.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-event-detail-labels -Phase RED `
  -GuideDraftOnly -CheckViewerPublishedRead -CompileViewerProbesForTest
```

Use `-CheckGuideDraft` in place of `-GuideDraftOnly` for combined recording
regression. The focused switch is not evidence that the broader suite ran.

## Observed RED and fixture corrections

The first two attempts each retain 76 passing existing checks but stop with a
harness error before guide entry. The first omitted `modAuth.CanPerform`'s required
user argument; the second revealed the missing explicit fixture maintenance grant.
Neither is meaningful RED. Both terminate with Excel closed; they remain preserved
under `reports/runtime/guide-draft-entry-red` and `guide-draft-entry-red-retry`
prefixes. Production authorization and runtime code are unchanged.

The corrected focused route establishes **41 PASS / 14 expected FAIL**. All failures
are missing guide entry/editor behavior; all 38 published Viewer prerequisites and
three source-preservation checks pass. Exact report:
`reports/runtime/slice4be-viewer-published-read/d48dcefd9a2f424994c3cb28d4bec87e/red.json`.
All five instrumented projects compile before forms, and Excel closes normally.

The expanded route retains those checks and adds authored wording/tags, stable
StepIds distinct from ActivityIds, actual Move up/Move down/Remove actions, retained
step instructions, unchanged original observation order, default/minimum/larger/
restored layout and clearing of stale source evidence. It establishes **41 PASS /
26 expected FAIL**, all 26 failures confined to `GuideDraft.*`. All 55 preceding
focused identities and all 38 baseline Viewer GREENs are retained. Exact report:
`reports/runtime/slice4be-viewer-published-read/6d8f37ac9c3745aabaa35e72be267914/red.json`.
Ignored `reports/runtime/guide-draft-complete-red` prefixes name this expanded draft
test; they do not mean that the full guide contract or Slice 4be is complete.

All five instrumented packages compile before forms in the expanded run. Runtime
source is unchanged, all five frozen package hashes match, both unrelated user
documents are preserved and Excel is closed. All four attempts' verified windows
contain no Application events 1000/1001/1002. Final verification is ignored
`reports/runtime/guide-draft-red-verification.json`. This is test-entry evidence,
not a repair of the earlier full-chain RPC/recovery behavior.

## Remaining implementation and acceptance

With expanded behavioral RED verified, implement the captured Core draft
boundary and Operations editor through the protected handlers. Keep save/version,
published guide discovery, both presentations, exact guide/run comparison and
validated origin-only transfer in scope, with their own protecting tests before
implementation. The prior label checkpoint retains the full-chain results and
assisted recovery limitations; this test-only entry claims no new build, runtime
static baseline, deployed/NAS behavior or human acceptance.
