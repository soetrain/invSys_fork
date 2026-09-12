# Slice 4be.1 Process Release observations and outage recovery

Continues the [memory investigation](plan022_slice4be_production_memory_diagnostics.md)
under Architecture v4.11 D12/D13/D18. Runtime, builder, managed schemas and
architectural contract remain unchanged. These markers are unsaved diagnostics,
not user-activity events or a new business authority.

## Calibration and recovered state

The placement test first reports **2 PASS / 11 FAIL** against untouched source;
after the diagnostic transform it passes **13/13**. It checks both sides of the
quiet-UI, pending-message, service, processor, status-query and refresh calls,
release-only filtering, and exact original-source preservation after marker
removal. A disposable Excel workbook passes **5/5** transport checks: actual macro
entry, exact ordered markers, out-of-range exclusion, unchanged template and
preservation of the caller's VBA error number/description/source/help state.

The first full diagnostic stopped on 2026-09-08 during package setup:
`Expected unique diagnostic boundary 101`, **0 PASS / 1 harness failure**. It did
not reach the operator workflow or native observer. This is not Product RED.
After the user's reported power outage, on 2026-09-12 the old session handle was
absent, Excel was closed and the terminal failure report remained available.
No work was inferred to be running from the chat window or a progress file.

Read-only inspection of the owned diagnostic package finds that Excel normalized
`mOperatorWorkbook.Name` to `mOperatorWorkbook.name`. The transform's case-sensitive
match was incorrect. The new actual-loaded-source test goes **0/3 -> 3/3** after
case-insensitive boundary matching, while preserving the exact matched source.
The existing placement and fresh transport checks remain **13/13 and 5/5**.
Native observer and memory calibrations are rerun after recovery: **27/27 + 10/10**.
All 20 package pins survive the outage unchanged.

## Packaged diagnostic

The corrected `release-boundary-probe-casefix` runs the complete standard
ProductionReusable workflow with unsaved markers in the actual packaged
`frmProduction.SubmitProcessAction` and
`modProductionReusableDesigns.SubmitReusableDesignEvent`, plus the calibrated
native observer. It passes **2/2**, including fresh restart.

The log contains exactly **204 numeric markers**: **11 complete Process Release
sequences** and two three-marker transport calibrations. The completed sequence is:

```text
101 102 103 104 105 201 202 203 204 205 206 106 107 108 109 110 111 112
```

101/102 bracket quiet-UI entry; 103/104 pending UI; 105/106 the service call;
201/202 queueing; 203/204 processing; 205/206 status lookup; 107/108 quiet-UI exit;
109/110 list refresh; 111/112 status display. No operational values are logged.
Both initial/restart marker calibrations pass. The native observer detaches
without error after its bounded initial window, with no exception captured.
This run does not localize the earlier failure or establish a native fix.

## Unmodified standard gate after recovery

The recovered environment has input-desktop access and a foreground window,
whereas the 2026-09-08 check lacked both. This is an observed environment change,
not proof of the native failure's cause or native worksheet-click acceptance.
It warrants checking the unchanged standard gate in the current environment.

The post-outage standard run uses the original pinned candidate and the checked-in
validator, without diagnostic markers, debugger attachment, project inspection or
reduced-workflow switches. It passes **2/2**, including clean restart, on
2026-09-12. This is a current standard-gate GREEN for the unchanged original
candidate, not a diagnostic substitute. The earlier native failure remains
unexplained; no runtime repair, causal attribution to the outage/reboot, or human
acceptance is claimed. Retain that history when evaluating future regressions.

## Exact retained evidence

Ignored files under `reports/runtime/slice4be-launcher-denial/`:

- `test-release-marker-placement.ps1`, `release-marker-transform.ps1`,
  `release-marker-placement/`, `release-marker-placement-casefix/`.
- `ReleaseBoundaryProbe.bas.txt`, `test-release-marker-transport.ps1`,
  `release-marker-transport-*/checks.json`; template/transform SHA pins.
- `release-boundary-loaded-procedure.txt`, `test-release-marker-live-case.ps1`.
- `create-release-boundary-probe.ps1`, `release-boundary-production-probe.ps1`,
  `release-boundary-probe/` (setup failure), `release-boundary-probe-casefix/`
  (2/2) and their logs.
- `post-outage-standard-production/` and its log: unchanged standard gate.
- `post-outage-package-preservation.json`: final 20/20 package preservation.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_plan022_packaged_launchers.ps1 -RepoRoot . -DeployRoot deploy/validation-receiving-launcher-denial -OutputDirectory reports/runtime/slice4be-launcher-denial/post-outage-standard-production -CallbackFilter Production -WorkbookState ProductionReusable
```

Preserve existing directories and use fresh outputs for future runs. Do not treat
case normalization as a product defect, diagnostic passes as standard acceptance,
or desktop availability alone as proof of native handler entry. All remaining
Release 1 contract, regression, visible and human acceptance requirements remain.

Final verification, 2026-09-12: both resumed test handles are terminal, Excel is
closed, and all 20 package pins remain unchanged. Six diagnostic scripts parse;
evidence links, diffs and statuses were reviewed. No runtime, builder, static
baseline, accepted deployment, operational workbook or NAS change occurred.
Next: use the restored desktop to calibrate actual worksheet-button input in a
minimal disposable fixture before the packaged Receiving native-caller test.
