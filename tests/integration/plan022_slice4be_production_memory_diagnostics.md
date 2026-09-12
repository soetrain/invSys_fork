# Slice 4be.1 Production memory-metadata diagnostics

Continues the [exit/execute investigation](plan022_slice4be_production_exit_and_execute_results.md)
under Architecture v4.11 D12/D13/D18. No runtime, package builder or architectural
contract changed. Production acceptance remains unresolved.

## Calibrated observation

Three added assertions against the existing synthetic exception fixture produced
**24 PASS / 3 FAIL**: the observer did not report memory state, type or protection.
The bounded metadata extension then passes **27/27**, retaining every earlier
identity, redaction, exception-preservation, survival and detach check. This is
diagnostic-tool calibration, not Product behavioral RED/GREEN.

The query opens a process handle with query-information access only and uses
[VirtualQueryEx](https://learn.microsoft.com/en-us/windows/win32/api/memoryapi/nf-memoryapi-virtualqueryex).
It reads page metadata, not page contents. Outputs contain only queried status,
finite state/type/protection labels and a nullable guard flag. Both the exception
site and, for an access violation, the attempted access target are classified;
only a Boolean comparison of those addresses is retained.

Ten independent live-allocation checks pass **10/10**: reserved state, undefined
reserved protection, committed private read/write, no access, guard without touching
the page, free state, undefined free fields, unavailable invalid address, mapped
read-only memory, and a bounded output schema without addresses. Undefined fields
for free/reserved regions follow Microsoft's
[MEMORY_BASIC_INFORMATION contract](https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-memory_basic_information).
The synthetic exception also proves classification of committed executable image
memory. No memory contents, stack arguments, operational values or dumps persist.

## Full workflow observations

The original candidate's `native-stack-probe-v5` run completes **2/2**, including
clean restart. The observer attaches at the initial batch-scale boundary, reaches
its 180-second limit with no exception observed and detaches without error. The
full test continues beyond that window. It therefore supplies no fault-region
classification and does not establish crash-free behavior outside the window or
repair the earlier plain/native failures.

The restart-specific probe retains the complete standard initial workflow and
attaches to the separately owned fresh Excel process before restart package loading.
Its placement and PowerShell syntax were checked before execution. It uses the
pinned compile-before-save candidate, whose earlier plain run failed at restart.
This full run passes **2/2**, with attach/detach, no engine error and no exception
observed through restart completion. It supplies no fault-region metadata.

One bounded repeat against the original candidate (`native-stack-probe-v6`) sought
the still-missing fault capture, not a passing replacement for the full gate.
Its initial observer window also ends without an exception and detaches cleanly.
The full run completes **2/2**. None of these three full runs supplies the missing
fault-region metadata; no further identical capture was launched.

Read-only inspection of the installed oleaut32 image places the prior return
site, module-relative 0x9f77f, immediately after an indirect call at 0x9f77d. This
supports an invalid call target but does not identify the supplying VBA procedure.
Only static installed-image code was inspected; no target-process memory was read.

## Reproduction and limits

All diagnostic scripts and machine reports remain ignored under
`reports/runtime/slice4be-launcher-denial/`:

- `NativeFrameProbe.cs`, `test-native-frame-probe.ps1`,
  `native-calibration-checks.json` and `native-calibration-helper-sha256.txt`.
- `test-native-memory-labels.ps1`, `native-memory-calibration-checks.json` and
  `native-memory-calibration-helper-sha256.txt`; both calibration pins must match
  the helper before a native run.
- `native-stack-production-probe.ps1`, `native-stack-probe-v5/`,
  `native-stack-probe-v6/` and their logs.
- `create-native-restart-memory-probe.ps1`,
  `native-restart-memory-production-probe.ps1`, `native-restart-memory-probe/`
  and its log. No reduced-workflow switches are used.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File reports/runtime/slice4be-launcher-denial/test-native-frame-probe.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File reports/runtime/slice4be-launcher-denial/test-native-memory-labels.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File reports/runtime/slice4be-launcher-denial/native-restart-memory-production-probe.ps1 -RepoRoot . -DeployRoot deploy/validation-receiving-denial-compiled -OutputDirectory reports/runtime/slice4be-launcher-denial/native-restart-memory-probe -CallbackFilter Production -WorkbookState ProductionReusable
```

Preserve existing output directories and use fresh paths for future diagnostics.
Owned PID, creation identity and name checks remain required; observers preserve
normal exception handling. Instrumented passes do not replace the unchanged full
gate, the remaining role/compile/layout/static/chain gates or human acceptance.

A separate read-only desktop check on 2026-09-08 reports input desktop unavailable
and no foreground window. Native worksheet input remains unproved; this finding
does not explain the Production fault or invalidate packaged handler evidence.

Final verification, 2026-09-08: all three harness handles are terminal, Excel is
closed, and all 20 pinned original/rebuilt/saved/compiled packages are unchanged
(`memory-diagnostic-package-preservation.json`). Runtime, builder and static
baseline remain unchanged. Diagnostic scripts parse and local evidence links,
diffs and Git status were checked. No accepted deployment or NAS changes occurred.

Next: calibrate fixed, value-free markers at the Process Release path's queue,
processor, status-query and list-refresh boundaries, then combine that finer
localization with the native observer in an isolated unsaved diagnostic of the
same packaged handler. Preserve the full gate and report any instrumentation
effect separately. Do not repeat the broad captures unchanged hoping for a pass.

The [Release-boundary/recovery follow-up](plan022_slice4be_production_release_boundary_diagnostics.md)
records the calibrated finer markers and a later unmodified standard-gate GREEN
after the user's reported outage, without attributing or declaring a native fix.
