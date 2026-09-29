# Slice 4be Production batch-boundary diagnostic

Last verified: 2026-09-29 UTC. Architecture v4.11 D12/D13/D18 and Plan022 remain
unchanged. This is diagnostic tooling for the native failure recorded in
[lifecycle continuation evidence](plan022_slice4be_production_lifecycle_visible_results.md),
not an invSys runtime repair or complete Production/Release1 acceptance.

The failed full gate stopped at `mProduction.RunProductionBatchScaleContractTest`
with RPC0x800706BE and ntdll.dll0xc0000028. That adapter re-enters the public
launcher, reads the form's visibility, then calls the typed scale test. Its single
progress marker did not distinguish those operations.

## Calibrated boundary evidence

`ProductionBatchBoundaryTrace.ps1` inserts31 fixed markers in unsaved Operations
memory, across that adapter, `BtnOpenProductionForm`, `ShowProductionForm` and
`frmProduction.TestBatchScaleContract`. It resolves all anchors before mutation,
compiles the four loaded workflow packages before forms, and never saves an XLAM.
It records no runtime values. Its logger rejects unknown stage names before any
write; the live transport calibration proves this with an unregistered value.

`Test-ProductionBatchTracePlacement.ps1` passes70/70: exact statement placement,
fixed unique stages, unchanged original statements, VBA case normalization,
missing/duplicate-anchor rejection and the bounded logger allowlist. This is
tool calibration, not product behavioral RED/GREEN.

`Test-ProductionBatchBoundary.ps1` derives a scoped copy of the actual standard
validator. It retains initial/repeated public launches and the original scale
assertion, stops immediately afterward, redacts report details, and retains the
existing cleanup decisions. Seven checks cover setup, launcher calls, saved
station-local workbook, reuse and exact scale results. The scope excludes later
Production workflows and restart.

| Control | Result | Cleanup and limits |
|---|---|---|
| Traced |7/7, four instrumented compiles,55 allowlisted stage entries; all three launcher entries and scale calculation return. |Automatic final termination requested. Settings/packages/validator preserved; zero Excel Application events. |
| Untraced |7/7 with the exact same check identities and no VBA instrumentation. |Automatic final termination requested. Settings/packages/validator preserved; zero Excel Application events. |

The55 traced entries match the complete expected ordered success sequence,
including all three public launcher entries, with no exception branch recorded.

Both controls observe three newly opened workbooks, exactly one station-local
operator workbook and zero new workbooks on the second launch. An earlier
diagnostic incorrectly required one new workbook overall and recorded6 PASS/two
failures, including its prerequisite-stop marker. The original validator instead
requires one station-local operator workbook; correcting the diagnostic preserves
that existing contract. This was a harness assertion error, not product RED.
Earlier JSON-array, anchor and output-path setup failures occurred before product
callbacks and are likewise not meaningful RED.

The traced and untraced controls run16:53:19--16:54:32 UTC. Their successful
boundaries do not explain the preceding crash or establish normal unassisted
shutdown. Keep the earlier failure and automatic-termination qualifications.
The subsequent standard full run at16:55:16--16:55:39 UTC stops0 PASS/one harness
failure at the same batch-scale adapter, with RPC0x800706BE and an Excel
ntdll.dll0xc0000028 Application Error. Full-only steps and restart are unreached.
The new cleanup receipt records TerminationRequested=False; the native failure
precedes cleanup. Settings/packages restore and Excel is closed. This is not
normal shutdown or full acceptance. Stop unchanged broad retries: the pair did
not reproduce or explain the failure. Further diagnostics must distinguish the
failed execution itself rather than infer a repair from successful scoped runs.

A read-only native-callback source review finds no call site for
MouseScroll.EnableMouseScroll outside its own unused declarations/comments,
consistent with the earlier Receiving-denial investigation. That lead does not
explain this failure and does not authorize deleting the module.

## Exact evidence

- Offline calibration: `reports/runtime/production-batch-trace-placement/2d2ca67c5db44d50b5e395b9bc36e891`.
- Invalid first workbook assertion: `reports/runtime/production-batch-boundary/c0b1a9daa23f44689f6041bdbaecd82d`.
- Traced control: `reports/runtime/production-batch-boundary/d18ba441a91f4bdfb619a11b3b734425`.
- Untraced control: `reports/runtime/production-batch-boundary/40e254ac78c6477fadf045865836333a`.
- Pair verification: `reports/runtime/production-batch-boundary-pair-verification.json`.
- Exact ordered stages: `reports/runtime/production-batch-boundary-stage-verification.json`.
- Full follow-up: `reports/runtime/production-paths-regression/reusablefull-3e7d9de859b442b4ba7be665d6ff8fea` and `reports/runtime/production-batch-full-followup-verification.json`.

The candidate remains `deploy/validation-production-paths`, pinned to the five
hashes in `production-lifecycle-native-controller/65e8e27585ca47289cf27e87e147e83e/package-pins.json`
under runtime reports. No Windows policy, normative contract, runtime VBA or
operational workbook changes. The user's conditional instruction remains: stop
the goal if desktop error5 returns. No such error is observed in these runs.

Refreshed `reports/runtime/production-batch-boundary-static` retains253 components,
6058 procedures,133166 lines,9 literal/45 unresolved dynamic calls,191 duplicate
groups and the exact28 previous size limits. All three schemas and295 PowerShell
parses pass. Runtime/form source is unchanged; no product rebuild or geometry
change follows from these diagnostic helpers. Prior compiled and layout evidence
retains its scope. Unrelated user files and the five package hashes are preserved.
