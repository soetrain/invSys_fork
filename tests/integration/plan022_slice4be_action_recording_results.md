# Slice 4be.4 recording lifecycle

Last verified 2026-09-13. Architecture v4.11 D18 and Plan022 govern the explicit
actor/warehouse sequence, immutable observations and advisory conclusions.
Slice4be and full Release1 acceptance remain open.

## Test-first evidence

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-events-maintenance -Phase RED -CheckActionRecording
```

The frozen five-package candidate from source checkpoint `1d6a2ad` records
**34 PASS / 16 FAIL**, 50 unique check identities, exit1 as expected for RED.
All 30 existing published-reader checks retain their exact identities and pass;
comparison uses the preceding `events-maintenance-filters.log` result. No
`Harness.Exception` occurs.

The extension opens Viewer through `modInventoryViewer.OpenInventoryViewer`,
enters Events through its existing handler, locates recording buttons by their
approved captions and delivers CommandButton.Value to the actual event binding.
Missing product controls return MISSING; a missing Viewer/test seam raises a
harness error. The status probe reads `lblRecordingStatus` without refreshing or
creating a record. Instrumentation is installed only in unsaved test projects.

The actual Admin capture checkbox and Save Tracking Policy handlers configure
the disposable warehouse. Repeated Save Value actions enter
`frmAdminSettings.mBtnSaveConfig_Click`, changing synthetic BatchSize values.
Each action must produce exactly one REQUESTED and one COMPLETED record with
the same ActivityId and registered ADMIN_SETTINGS_SAVE_VALUE identity before
its sequence assertions run. Missing ordinary activity is a fixture failure,
not recording RED. No synthetic owner outcome or sequence is supplied.

| Observation | Result | Meaning |
|---|---|---|
| Prior published-reader checks | 30/30 PASS | Existing public Viewer behavior retained within this test's scope. |
| Recording buttons, disabled explanation and disabled Start | 5 FAIL | The approved surface is absent. |
| Start/counter, first/second ordinals and distinct repeated occurrence | 5 FAIL | Valid ordinary actions have blank SequenceId and Ordinal0. |
| Stop control and stopped-without-conclusion status | 2 FAIL | No lifecycle transition exists. |
| Second Start, new sequence, Cancel and cancelled status | 4 FAIL | No new-run/cancel lifecycle exists. |
| Ordinary activity before/after attempted recording and earlier byte retention | 4 PASS | Existing behavior is preserved; these passes do not prove Stop/Cancel worked. |

## Preservation and limits

No VBA, form, package, schema implementation, accepted deployment or NAS workbook
changes are made. Admin-generated disposable fixture configuration changes are
intentional. The harness removes those fixtures and restores its saved personal
settings after the run. All five frozen candidate package hashes remain equal
to their pre-run pins. Excel is closed. No Excel Application Error1000 appears
between creation of this run's log and the post-run observation; this does not
resolve the previously recorded intermittent native faults.

The architecture, Plan022 and controls name the already approved recording
surface consistently. This refinement changes no business authority, capability,
activity catalog identity, collection rule or accepted behavior. No new build,
compile, static-maintenance, role/full-chain or visible acceptance is claimed for
this test-only checkpoint. The candidate's preceding gates retain their recorded
scope in the Events filter evidence.

This first lifecycle RED does not yet protect saved Action Path schema/hash,
atomic incremental persistence, 1MiB overflow, 256-action closure, tracking
failure, mid-sequence policy/context invalidation, restart, all Operations roles,
exact multi-event submissions or evaluation. Add those focused cases before
their implementation. Then implement actual Viewer handlers and headless Core
recording, preserve these 50 checks, and complete all D18/D13 gates. How-To,
Diagnostic, Compare both, library/version/import/export and human acceptance
remain required; Stop alone must never stand in for a conclusion.

## Ignored raw evidence

- `reports/runtime/action-recording-red.log`
- `reports/runtime/slice4be-viewer-published-read/d3e82e13f0f04c5aa9a6f0881f0775a1/red.json`
- `reports/runtime/events-maintenance-package-pins.json`
- `reports/runtime/events-maintenance-filters.log`

The maintained report contains only check identities, aggregate results and
technical scope. Runtime fixture identities and row values are not copied here.
