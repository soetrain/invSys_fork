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
coverage or Release 1. Accepted packages remain unchanged; the isolated runtime
implementation and its verification are in progress as described below.
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

## Runtime implementation in progress

The isolated candidate adds headless Core `modGuideDraftSource` and
`modActionGuideDraft`, plus Operations `modGuideEditor` and `frmActionPathGuide`.
The existing `frmActionPaths` provides the actual Create guide entry. Core owns
memory-only authored text and stable step identities; the form owns the captured
editor and presents original observations separately. No save/publication or
presentation-switch behavior is accepted by this draft gate.

Two explicit compilation attempts caught mismatched `Next` variables in the new
form. These are source defects, not behavioral RED. Corrected isolated
`deploy/validation-guide-draft-ready` passes cold-start dependency validation and
all five explicit package compiles. Its 229 compiled components retain every prior
component except the intentionally changed `frmActionPaths`, with exactly four
new components. The first behavioral run retains 42 passing checks but has 25
guide failures: the permitted selected run does not enable Create guide. A
test-only guard-stage trace confirms error 5 at policy serialization. Excel
projects the validated saved catalog version as Double; the training serializer
accepts Integer/Long. The new guide boundary now normalizes that validated field
before hashing. The serializer itself is unchanged. The trace is removed from
the protecting test and preserved in ignored `guide-draft-guard-trace-probe.ps1`.
The corrected `deploy/validation-guide-draft-policy` candidate passes all five
explicit compiles and the same compiled-source scope comparison. Its focused
packaged gate is **72/72 GREEN**, retaining all 67 protecting identities and adding
five successful actual editor captures. Exact report:
`reports/runtime/slice4be-viewer-published-read/f433215eeb664b55a7753e670fa64876/green.json`.
Original failed candidates and reports remain preserved.

All five captures were directly reviewed: reordered steps, minimum, default,
larger and restored sizes. Labels, editable authored fields, immutable observations,
step controls, staging status and Cancel fit visibly without overlap. This is
automated operator-surface evidence, not human acceptance. The terminal test exit
is 0 with Excel closed. The verifier confirms all ten current/frozen package pins,
both unrelated user documents and 11 audited build/compile/test windows without
Application events 1000/1001/1002. Build cleanup intervals are included through the
next no-Excel compile preflight; immediate build exit did still see Excel.
See ignored `reports/runtime/guide-draft-focused-verification.json`.

The first static review records 236 components, 5,921 procedures and 130,589 lines: growth
of four components, 40 procedures and 491 lines for the authored guide feature.
All 28 existing module growth limits hold. Literal/unresolved `Application.Run`
counts remain 9/45. Scanner duplicate groups increase from 191 to 194. The
explicit maintenance exception is limited to these three reviewed groups:

- `041b258526c03dfd`: the three actual Move up/Move down/Remove handlers call one
  shared operation with different command literals. Normalization removes those
  literals; the actions are distinct, and their implementation is already shared.
- `b0153d52048d1613`: the two `CloseLibrary` procedures validate their respective
  typed library ownership before closing a different editor.
- `e9b4f06801108607`: the two `CloseEditor` procedures release and unload their
  respective typed form and Core draft. A generic late-bound editor dispatcher
  would weaken direct typed calls; a new interface solely for these small lifecycle
  routines would add more machinery than it removes.

These exceptions retain explicit form ownership and existing lifecycle behavior.
They do not authorize future duplicate growth, dynamic calls, deletion of
scanner candidates or an architectural exception. Packaged reuse/cancel/context
checks and the existing expectation regression protect the lifecycle decision;
successful current-candidate results are still required. The regenerated ignored
baseline is `reports/runtime/guide-draft-static` (1,177 reviewed candidates).
The regenerated final baseline, `reports/runtime/guide-draft-policy-static`, adds
the two-line catalog normalization: 130,591 lines, or 493 over the prior accepted
feature baseline. Other metrics and all 28 existing limits remain as reviewed.

## Remaining implementation and acceptance

Combined recording/guide passes **105/105** at
`reports/runtime/slice4be-viewer-published-read/36be87ce9d234bc2b98b08dac5d95e33/green.json`.
Event Detail passes **34/34** at
`reports/runtime/slice4be-viewer-detail/35fa03bb661b4f4e8fe01193045d9d38/green.json`.
Both terminate at exit 0 with Excel closed. Viewer/filter/Shipping state,
Boxing/Shipping, evaluation and full Release 1 chain/live-role regressions are
running serially against the corrected isolated candidate. Their results remain
pending; this is a focused GREEN checkpoint, not completed slice acceptance. Keep save/version,
published guide discovery, both presentations, exact guide/run comparison and
validated origin-only transfer in scope, with their own protecting tests before
implementation. The prior label checkpoint retains its full-chain results and
assisted recovery limitations; this candidate establishes no native-recovery
repair, deployed/NAS behavior or human acceptance. Admin's current user-role form
does not expose ACTION_PATH_MAINT among its six capability choices; the actual
user provisioning route remains a required test-first extension, independently
of the disposable fixture grant used here.
