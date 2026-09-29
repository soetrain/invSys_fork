# Slice 4be projection-boundary diagnostic

Last verified: 2026-09-29 UTC. Focused traced and untraced controls pass35/35;
full Release1 chain/live roles/Create Warehouse pass32/48/15. No runtime correction or new
architectural contract is claimed.

## Governing contract and diagnostic scope

Architecture v4.11's Inventory projection contract requires derived tables to be
rebuildable by the processor from `tblInventoryLog` and `tblAppliedEvents`, while
preserving event-carried `System_Key` and canonical authority. D5, D12, D13 and
Plan022 remain binding. The preceding candidate's chain failed during
`modProcessor.RunBatchReportForAutomation`, after projection deletion, with
0x800706BE and an Excel Application Error. That harness crash was not behavioral
RED and its cause remains unresolved.

`Test-Slice4beProjectionLiveControl.ps1` already preserves the ordered live-role
workflow through projection recovery and then stops before later chain phases.
It now accepts an explicit package-pin file and optional `-TraceBoundaries`.
The latter installs37 fixed stage markers into unsaved Core/Inventory Domain
test packages before forms, compiles all four loaded workflow packages, and arms
logging only at projection deletion. It changes no business arguments, expected
values, callback handlers or persistence behavior. All five candidate files stay
pinned and are checked again after the run; settings and tracked reports retain
their existing restoration guards. The untraced route installs no adapters.

The helper `ProjectionBoundaryTrace.ps1` observes processor context, inventory
resolution, locks, schema creation, projection rewriting and publication. Its
trace contains stage constants only: no credentials, entered values, operational
rows or settings. These are diagnostic-tooling changes, not a product RED/GREEN
correction. No Core/Domain runtime source or package is rebuilt or saved.

## Focused outcomes

Both controls use `deploy/validation-production-paths` and the verified five
package pins from the previous native-cancellation controller. The protecting
projection balance and authoritative-log assertions remain unchanged.

| Control | Evidence |
| --- | --- |
| Traced |35/35, four explicit instrumented compiles,37 installed stages;77 trace entries all belong to the allowlist. The processor call returns, the cut is reached and Excel closes normally. |
| Untraced |35/35 with the exact same check identities; the processor call returns and Excel closes normally. No instrumentation is installed. |

Both controllers exit0, report zero Application failure events, restore settings,
preserve all five package hashes and leave the tracked live-role report unchanged.
The traced audit completes16:14:01 UTC. Passing both controls shows that trace
instrumentation is not required for this candidate to complete the focused case
in the current session conditions. It does not establish why the earlier crash
occurred or accept the full chain.

Private results:

- `reports/runtime/projection-live-control/60003ca1f6474cdd9b557c03cbe7ebf7/result.json`
- `reports/runtime/projection-live-control/663d6bc715d841a380ce84cb776447c9/result.json`
- `reports/runtime/projection-boundary-pair-verification.json`
- `reports/runtime/projection-boundary-trace-redaction-verification.json`

Commands: run `Test-Slice4beProjectionLiveControl.ps1` with
`-DeployRoot deploy/validation-production-paths -Cut AfterProjection` and
`-PackagePinsPath reports/runtime/production-lifecycle-native-controller/65e8e27585ca47289cf27e87e147e83e/package-pins.json`;
add `-TraceBoundaries` only for the traced control.

## Remaining acceptance

The subsequent full chain passes **32/32**, live roles **48/48**, and Create
Warehouse **15/15**, retaining every preceding accepted check identity. The
unmodified candidate runs16:16:33--16:21:39 UTC, exits0 and closes Excel without
intervention. Settings, five packages and all three existing report files are
restored/preserved; the Application audit records zero matching Excel failures.
Private controller: `reports/runtime/production-paths-regression/chain-f9b28916dad344768d7ba3c3a2fd0dcd`;
receipt: `reports/runtime/projection-boundary-chain-verification.json`.

Static evidence in `reports/runtime/projection-boundary-trace-static` retains
253 components/6058 procedures/133166 lines,9 literal/45 unresolved calls and191
duplicate groups. All28 existing size limits are unchanged, all3 schemas pass,
and291 PowerShell files parse without errors. No runtime/form source changes
require new product geometry or screenshots for this diagnostic checkpoint.

Preserve the earlier failed chain and its assisted cleanup: its cause remains
unresolved. These later successful gates establish current-candidate acceptance
for their scope, not a root-cause fix. Current-candidate draft/reusable Production regression and the broader
Slice4be/control/visible/human/NAS requirements remain open in the
[acceptance index](plan022_slice4be_remaining_acceptance.md).

The user explicitly requests stopping the goal if desktop error5 returns.
Coordinate screen-dependent work with the user and preserve that instruction.
