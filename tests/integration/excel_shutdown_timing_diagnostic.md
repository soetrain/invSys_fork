# Excel shutdown timing diagnostic

2026-10-27: 4be-A is paused at the user's request. This is separate, observation-only
harness work. No product source, XLAM, teardown policy or acceptance gate changes.

## Question and measurement

Does the slow interval occur in Excel.Quit, explicit COM release, GC.Collect,
WaitForPendingFinalizers, or after those operations? Existing Inventory traces combine
the two GC calls and cannot answer that question.

`Invoke-ExcelShutdownDiagnostic.ps1` makes a temporary instrumented copy of the
existing Config-command worker, retaining its original helper directory and calls.
It runs the previously slow Inventory worksheet Evidence workload on frozen packages.
This workload is diagnostic only; its existing product RED assertions are not new
4be-A acceptance work.

The trace records assertion ordinals, workbook close attempts (Workbook/XLAM),
Quit, COM release, GC, finalizer wait and worker exit, with UTC and monotonic ticks.
The baseline leaves package closure to Quit; `PackagesCloseViaQuit` records that
fact rather than inventing an explicit XLAM-close result. A controller thread holds
the exact Excel process handle, bound to the worker's HWND-derived PID and native
creation time. It records both native exit time and observation time independently
of a blocked worker. Exit code alone is not native process-exit proof.

Reports contain fixed labels, counts and process metadata only. No workbook names,
cell contents, credentials, macro arguments or exception text are added to the trace.
Raw runtime evidence stays ignored. Parent settings restoration and package/tooling
hash checks remain required after cleanup. There is no timeout kill or assisted
teardown in this baseline instrument.

## Prior evidence

In main `reports/runtime/slice4be-admin-inventory-entry/`, run
`17184532fae7486aa3acfd6f44d9d8ba/inventory-shutdown.jsonl` records Quit returning,
COM release taking about 0.415 seconds, then **380.229 seconds** between the combined
Before/AfterGarbageCollection markers. Other completed combined intervals include
1022.288 seconds (`f469ebd2536a4fa3b52a4f99e852fbe0`) and 1002.781 seconds
(`fc7b1f469a53487ba745fd4babbead2c`). These locate a slow harness interval but do not
distinguish collection from finalizer waiting or prove causality for every slow exit.

## Validation

`Test-ExcelShutdownTiming.ps1` passes 14/14: separate intervals, monotonic timing,
fixed fields, rejection of unapproved labels, generated-worker parsing, unchanged
Quit/release/GC call counts, original helper resolution, refusal on anchor drift,
process creation-identity mismatch refusal and independent native exit observation.
The native observer test uses a short-lived disposable PowerShell process, not Excel.

The first controller attempt (`d28e3fab1d5142f0a719cfa1269e9eb3`) failed before Excel
started because the worker requires a repository-relative deployment argument.
It is excluded; settings, packages and original tooling were preserved. The
controller now resolves that argument explicitly and retains its worker handle.

## Measured result

Controller `7a9733a34d404c1487dd6b7a936e6c6a` under main
`reports/runtime/excel-shutdown-timing/`, worker
`slice4be-admin-inventory-entry/060f370b1fb64ecd97505c881812e96d`, completed
unassisted. Timing code1439172e plus controller correction2f1d146a; the manifest
and ignored package/tooling pins identify the exact generated worker and inputs.

| Phase | Duration |
| --- | ---: |
| Last assertion to cleanup entry | about1ms |
| Cleanup workbook enumeration | about10ms; no workbooks enumerated |
| XLAM close | left to Quit; no explicit close in this baseline |
| Excel.Quit | 5.804ms |
| FinalReleaseComObject(application) | 537.571ms |
| GC.Collect | 40.019ms |
| GC.WaitForPendingFinalizers | **334,012.778ms (5min34sec)** |

Finalizer waiting began23:39:29.6669369Z and returned23:45:03.6796077Z.
The controller's native handle signaled; kernel exit time was23:45:03.6699529Z,
about10ms before the wait returned. Worker exit was observed23:45:03.8186399Z.
Thus the long blocking interval is confirmed inside finalizer waiting. This
supports the COM-lifetime hypothesis; it does not identify the exact retained
objects or prove that Excel contributes no work to their release.

Closure23:45:04.748Z confirms Excel closed, settings restored and frozen packages
and original tooling unchanged. Delayed audit23:45:24.408Z found zero Excel
1000/1001/1002 events. Desktop monitor recorded zero lock/input-failure/error5 samples
and was stopped normally. The workload's50 PASS/28 FAIL retains its existing
unimplemented-recording/Config REDs; this diagnostic claims no product GREEN.

No teardown behavior has been changed. Next, isolate ownership/release of one
high-reference helper and compare on this measured workload. Preserve unassisted
release qualification; any later development-only assisted exit needs an explicit
label and exact native process ownership, not a process-name kill.

## Reproduce

From a checkout containing the diagnostic, with Excel closed and a fresh desktop
monitor/cursor proof, supply the actual paths to the frozen candidate and known worker:

```powershell
powershell.exe -NoProfile -File tests/tooling/Test-ExcelShutdownTiming.ps1 `
  -WorkerScript ../invSys_edit_receipts/tests/tooling/Test-Slice4beConfigCommands.ps1
powershell.exe -NoProfile -File tests/tooling/Invoke-ExcelShutdownDiagnostic.ps1 `
  -RepoRoot ../invSys_fork `
  -WorkerScript ../invSys_edit_receipts/tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -DeployRoot ../invSys_fork/deploy/validation-inventory-uom-02 `
  -DesktopMonitorReport <fresh-monitor-report>
```

Do not resume or build the partial 4be-A product changes as part of this diagnostic.
Any later development-only assisted teardown must remain explicitly labeled and
cannot satisfy the unassisted final qualification gate.
