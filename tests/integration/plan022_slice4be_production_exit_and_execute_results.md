# Slice 4be.1 Production exit verification and execute-fault diagnosis

Follow-up to the [Production diagnostic matrix](plan022_slice4be_production_provenance_diagnostics.md).
Architecture v4.11 D12/D13/D18 and the full Release 1 goal remain unchanged.
No runtime or build implementation was changed. Production acceptance remains
unresolved; the new build-boundary test protects already-correct behavior.

## Exit boundary

`test-exit-boundary.ps1` extracts and executes the standard validator's actual
restart teardown block and its actual Release-ComObject helper. Three owned
disposable Excel fixtures cover empty Excel, Core only, and all four Operations
dependencies opened read-only. Each verifies the original process identity and
observes that process terminal at the point where restart would begin. All three
entered the existing stop branch and were terminal without an added pre-restart
exit wait. These fixtures do not exercise the full business workload.

The complete `ProductionReusable` sequence then ran against the pinned
compile-before-save candidate with in-memory fixed restart markers and an owned
process terminal/identity assertion. It passes **2/2**, with
`TerminalBeforeRestart=True` and `DistinctRestartIdentity=True`. The markers are
written only during final cleanup, not to disk between restart actions. No
VBProject inspection, VBA edit, reduced-workflow switch or new wait is introduced.
This establishes correct exit ordering for that run, not the cause of the earlier
plain failure. There is no evidence supporting an exit-wait repair.

## Native exception classification

Three new synthetic metadata checks first produced **15 PASS / 3 FAIL** for the
absent exception site, access kind and first-chance status. After the bounded
extension, the expanded calibration passes **24/24**, including read/write/execute,
unsupported/missing metadata, exclusion of other exception codes, site redaction,
process identity, non-target exception preservation, fixture survival and detach.

The original candidate's full native-observed run (`native-stack-probe-v4`) fails
**0 PASS / 1 failure** at batch scale. The preceding observed exception is
`0xc0000005`, `FirstOtherAccessKind=Execute`, `FirstOtherIsFirstChance=True`, with
`FirstOtherSite=UNRESOLVED_FRAME`. Its next frames are oleaut32, VBE7 and Excel;
the subsequent `0xc0000028` stack remains ntdll/repeated VBE7. The observer detached
without an engine error and the harness cleaned up Excel.

The classification reads only the documented operation flag and first-chance
metadata, retaining an enum/Boolean and sanitized module-relative site. Operation
8 identifies an execute/DEP violation in Microsoft's
[EXCEPTION_RECORD64 contract](https://learn.microsoft.com/en-us/windows/win32/api/winnt/ns-winnt-exception_record64).
No raw addresses, memory contents, stack arguments, operational values or dumps
are persisted. An unresolved site does not prove that memory was freed or identify
the supplying VBA procedure. The exact invalid execution target remains unknown.

## External package-edit boundary

Source order initially suggested that RibbonX was edited while its artifact was
loaded: the builder calls SaveAs before RibbonX and closes the source workbook
afterward. Live calibration disproved that interpretation. With IsAddin=True,
this SaveAs-to-XLAM flow writes the artifact while the in-memory source workbook
retains its original name and empty path; it is not bound to the saved XLAM.

The first probe was inconclusive: Workbooks enumeration omits add-ins and the
PowerShell JSON result array was counted incorrectly. The corrected detector uses
exact named lookup and path comparison. Named lookup for open add-ins is documented
in Microsoft's [Workbooks reference](https://learn.microsoft.com/en-us/office/vba/api/excel.application.workbooks).
The exact detector was then calibrated against a disposable artifact: saved copy
accepted, original source unbound, explicitly opened XLAM rejected, closed XLAM
accepted, and artifact bytes preserved, **5/5**.

The unchanged real five-package builder passes **7/7** external-edit checks:
five build-identity writes and Operations/Admin RibbonX writes. The new
`tests/tooling/Test-PackagedExternalEdits.ps1` runs an observed copy of the actual
builder, asserts the target is not loaded before each external writer, delegates
to the original writer, and verifies the checked-in builder was not modified.
It rejects existing candidate/evidence directories and requires Excel closed
before starting. Reports contain fixed package names, stages and Boolean state.

No RibbonX reorder or close-failure correction was implemented. There is no
Product behavioral RED here; the already-correct boundary and the erroneous
initial harness assumption must not be recast as a repaired product defect.
The generated package sets are diagnostic outputs, not accepted replacements.

## Commands and evidence

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-PackagedExternalEdits.ps1 -RepoRoot . -OutputRoot deploy/validation-ribbon-closed-red-v2 -EvidenceDirectory reports/runtime/slice4be-closed-package/red-v2
```

That directory name was chosen before the observed GREEN result and is retained.
Use fresh directories for a future run. Other exact ignored scripts and evidence:

- `reports/runtime/slice4be-launcher-denial/test-exit-boundary.ps1` and
  `exit-boundary.json`: three focused teardown cases.
- `create-exit-workload-probe.ps1`, `exit-workload-production-probe.ps1`, and
  `compiled/exit-workload/`: full sequence, memory-only stages and process evidence.
- `test-native-frame-probe.ps1`, `NativeFrameProbe.cs`, calibrated helper SHA and
  `native-stack-probe-v4/`: 24-check calibration and execute-fault trace.
- `reports/runtime/slice4be-closed-package/red/`: initial invalid probe, preserved.
  `red-v2/`: actual builder 7/7 with original builder bytes unchanged.
- `saveas-calibration/` and `boundary-guard-calibration/`: live SaveAs semantics
  and exact detector calibration, including the opened-artifact rejection.

The full exit-workload run uses `deploy/validation-receiving-denial-compiled`;
native v4 uses `deploy/validation-receiving-launcher-denial`. Both retain the full
ProductionReusable workflow and existing public launcher/form-action boundaries.
No runtime/static implementation changed, so this discovery adds no new claim
for the outstanding compile/layout/static/live-role/full-chain/human gates.
Final check, 2026-09-08: Excel is closed; all diagnostic handles are terminal and
all 20 pinned original/rebuilt/saved/compiled package hashes remain unchanged.
No accepted deployment, operational workbook or NAS changes were made. The test
and diagnostic scripts parse; evidence links, diffs and Git status were reviewed.

Next: calibrate a read-only classification of the failing execution site's memory
state/type/protection, then capture those bounded labels from an owned native
failure without reading memory contents. Do not repeat the rejected exit-wait,
save/compile preparation or RibbonX-order hypotheses as proposed repairs.
