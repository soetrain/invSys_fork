# Slice 4be.1 Receiving launcher denial observations

Architecture v4.11 D18's launcher-denial clarification governs this change.
Architecture, Plan 022 and controls v1.80 were synchronized and pushed as docs
**f918687** before runtime implementation. This inherits comprehensive activity,
owner-fact and captured-context rules; D12 capability gating remains binding.
The comprehensive 4be Goal and human Release 1 acceptance remain incomplete.

## Contract and implementation

The actual generated Operations Receiving callback still enters
`modTS_Received.ShowReceivingForm True`. Its typed
`modReceivingActivityAction.BeginOpen` records REQUESTED, then invokes the same
Core cached RECEIVE_POST guard exactly once before launch owner work. Rejection
records RECEIVE_OPEN_DENIED / Blocked / Unchanged with empty source references.
Fixed explanation: "Receiving form launch was not authorized." Advisory:
"Review Receiving permissions before reopening."

The build metadata explicitly declares this action-owned guard and validates
the exact Receiving callback/action/capability tuple before omitting its outer
guard. getEnabled capability mapping and every other generated guard remain.
No authorization ownership, canonical schema or business operation changes.

Repeated dispatches have distinct correlated attempts/results. Polling,
direct guard calls and pre-sign-in dispatches produce no warehouse user activity.
An interrupted session retains only the original-context attempt and an
incomplete-tracking notice. Optional storage failure preserves the existing
denial and shows Tracking unavailable without provisioning, opening or retrying.
Direct compatibility launch retains its existing non-click semantics.

## Test-first evidence

Protecting entry: `Test-Slice4beConfigCommands.ps1` with the actual generated
Ribbon callback and unsaved Office.IRibbonControl seam. Core instrumentation
counts the existing guard/auth decision and intercepts only its notification;
it does not replace authorization. Fixtures enter through Admin Generate/Seed.

- Unchanged candidate `deploy/validation-receiving-navigation-identity`:
  first focused **108 PASS / 19 FAIL**, expanded **112 PASS / 21 FAIL**.
  No harness exception. Missing denial records, optional failure notice and
  original-context interrupted attempt are the intended failures.
- New candidate `deploy/validation-receiving-launcher-denial`:
  focused **133/133 GREEN**. All earlier focused identities retained.
- Native-dialog variant (`-CaptureDenialDialogs`): **137/137 GREEN**, including
  all 133 original checks and four actual visibility/capture checks. Both
  process-owned modal captures were inspected: existing RECEIVE_POST denial
  and "Tracking unavailable: the training record could not be saved."
  The observer captures only the fixed dialogs and dismisses only a sole OK.
- Existing guard/notice, no provisioning/form opening, staged System_Key and
  unknown values, authority bytes, unrelated workbook/activation, polling/direct
  guard and signed-out exclusions pass before and after implementation.
- Reports contain check identities and booleans, never credential values,
  operational rows, arbitrary error/audit payloads or recorded user inputs.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-receiving-launcher-denial -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CheckReceivingLocalActivity -CheckReceivingLifecycleActivity -CheckReceivingNavigationActivity -CheckReceivingLauncherDenial -ReceivingLauncherDenialOnly
```

For reproducible RED, use the preserved navigation/identity DeployRoot and
`-Phase RED`. The full regression command additionally uses
`-CheckReceivingSurfaceCoverage -CaptureEvidence` and omits the diagnostic-only
switch. Current focused logs/results are under ignored
`reports/runtime/slice4be-receiving-activity/`; release-gate copies are under
`reports/runtime/slice4be-launcher-denial/`.

## Maintenance evidence

Regenerated source evidence: 175 components, 5,501 procedures, 1,097 scanner
candidates, 195 duplicate groups, 45 unresolved dynamic calls and eight literal
Application.Run targets. The one new procedure is the typed BeginOpen adapter.
Comparison against the saved pre-change baseline preserves all 28 existing
oversized-module caps and all component/candidate/duplicate/dynamic-call counts.
No new growth exception is requested and no code is deleted from scanner advice.

The historical Slice 12 cleanup validator reports 7/13 against both the prior
and current baselines, with the same six failing identities: registry hash,
component/procedure/candidate/duplicate budgets and old module-growth review.
Those frozen Slice 4x expectations are not rewritten to claim GREEN. This is a
pre-existing limitation, not a new regression; current D18 ratchets and recorded
exceptions remain explicit. Count/identity comparison is retained in ignored
`legacy-maintenance-comparison.json` and `static-comparison.json`.

Source checks pass: R1 control surface 6/6, Operations cutover 14/14, Receiving
stabilization 10/10 and tool contracts 62/62. Two initial source-check invocations
needed explicit `-RepoRoot .`; no product change was made for that harness issue.

## Release gates

Five-package build, five explicit VBE compiles and cold-start dependencies pass.

| Gate | Observed result |
|---|---|
| Full activity regression | 845/845; all prior 771 and surface-discovery 105 check identities retained, no duplicates |
| Packaged smoke | 86/86 |
| Live role workflows | 48/48 |
| Ordered Release 1 chain, restart/reconciliation | 30/30 |
| Viewer | PASS |
| Production layout | Three sizes, five visible pages and native transitions PASS; all three inspected PNGs match prior accepted captures byte for byte |
| Public Operations launchers | 3/3 |
| Receiving visible evidence | Normal form inspected; both native fixed denial/tracking notices captured and inspected |
| Full reusable Production | UNRESOLVED: candidate twice crashed; prior candidate comparison 2/2 |

VBA source extracted during the compile check differs from the prior candidate
in exactly four components: Core `modReceivingActivityCodes`; Operations
`modReceivingActivityAction`, `modTS_Received`, `modRibbonGenerated`.
Production and all other extracted source remain identical. This limits the
observed change surface; it does not excuse the unsuccessful Production gate.

Candidate full Production runs failed at reusable-surface and batch-scale
stages respectively, with RPC 0x800706BE after Excel exited. Both corresponding
Windows application events report EXCEL.EXE / ntdll.dll / c0000028. The same
full harness against the preserved navigation/identity candidate passes 2/2,
including clean restart. Earlier staging/lifecycle records contain similar
unresolved native faults; no root cause is inferred from that history.
There is no Production-harness VBA instrumentation and no runtime caller of
MouseScroll.EnableMouseScroll. These investigated leads supply no explanation.
Controlled package-isolation diagnostics each pass 2/2, including restart:
candidate Core with the other four prior packages, and candidate Operations
with the other four prior packages. These preloaded diagnostic combinations
are not release candidates or a cold-start dependency acceptance claim. They
do not establish a root cause.

A separate clean rebuild, `deploy/validation-receiving-denial-rebuild`, passes
cold-start dependency and five compile checks. All 168 extracted VBA components
match the original denial candidate. Its first full Production run nevertheless
crashes at batch scale with the same native fault. The planned second run and
broader gate replay were correctly not started after this failure.

A further diagnostic against the original denial files invalidates Operations
compilation with an unsaved comment insertion/removal, asserts exact source
preservation, and explicitly recompiles before invoking the same full harness.
It also crashes at batch scale. This does not support Operations recompilation
alone as a remedy. All ten original/rebuilt package hashes remain unchanged
after these diagnostics; Excel is closed at the final check (2026-09-08).

Next bounded investigation: test candidate Core and Operations together with
the prior unchanged Admin/Domain packages, recording actual loaded dependency
paths before interpreting that diagnostic. Do not accept a mixed diagnostic
set, repeat unchanged retries or modify Production speculatively. The original
and rebuilt candidates remain preserved and unaccepted at this gate. None of
the native failures is a meaningful D13 behavioral RED or permission to waive
the required regression. This is an implementation checkpoint, not a completed
4be.1 slice or Release 1 acceptance.

Evidence directory: ignored `reports/runtime/slice4be-launcher-denial/`,
including `activity-results.json`, `native-denial-results.json`, gate result
copies, `production/`, `production-retry/`, `production-baseline/`, redacted
crash summaries, `source-differences.json`, `layout-comparison.json`,
`isolation-summary.json`, `rebuild/` and `recompile-probe/`.
Original candidate's inspected dialog/form PNGs and hashes are preserved under
ignored `reports/runtime/slice4be-launcher-denial/receiving-visible/`.

The tested candidate hashes (SHA-256) are:

| Package | Hash |
|---|---|
| Admin | 5eddf176f6131ecb0e2dff1e7e3244a3600cafe4a89a78485fc6c683cdd17196 |
| Core | 42568735fc7090f5da367a840ad3e6270d805db7a17934717f76b71f15a519e5 |
| Designs Domain | 5319a37a82d5a3cd179addd09ae6ee6c2dd337d684c1c423a918b61b88f5312d |
| Inventory Domain | 113a32388e359662a17ee01c762f9312ab9b35424726c51c08fbc9272c786ef5 |
| Operations | 4759bd058816f4c19b1e0298dcdd0ce9fe5f00b57a38822cf035cae75db48f07 |

Native worksheet Confirm invocation remains independently unproved as documented
in [surface discovery](plan022_slice4be_receiving_surface_results.md). Other
Operations/Admin coverage, publication, Event Tracking Settings, comprehensive
Viewer, recording/conclusions, guide management and both comparison presentations
remain incomplete. Agent evidence does not replace physical multi-station/NAS or
human user acceptance. Accepted deployment and operational workbooks are unchanged.
