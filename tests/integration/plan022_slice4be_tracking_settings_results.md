# Slice 4be.2 Event Tracking Settings

## Current result: packaged surface RED

Last verified 2026-09-13. Architecture v4.11 D18 already requires Admin Settings
General/Event Tracking tabs, Tracking/Event Detail/Action Paths sections and
separate policy, profile and personal preference actions. Plan 022 designates
this as 4be.2. This checkpoint adds protecting tests; it changes no runtime
implementation, package or normative contract. Remaining 4be.1 coverage is not
accepted by moving to Settings work.

The isolated candidate is `deploy/validation-shipping-prewrite-facts`, built from
the previously recorded Shipping correction. The test opens all five packaged
XLAMs read-only and creates disposable fixtures through Admin Generate Warehouse.
An unsaved observer inspects the live `frmAdminSettings` created by the existing
`TestD5Commands.OpenSettings` helper and shown modelessly. This exercises the real
form initialization/layout, followed by the preserved real selection/save handler
and Production UOM handler checks. It does **not** claim to exercise the Ribbon
`modAdmin.Open_Settings` authorization path or any nonexistent tracking save
handler. Only fixed Boolean observations enter its result JSON.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-prewrite-facts -Phase RED -CheckTrackingSettings -CaptureEvidence
```

The completed run has **26 checks: 21 PASS / 5 FAIL**, with no harness exception.
All **18/18** prior D5 identities and GREENs remain, with no duplicates. The three
new passes prove the actual form is shown, existing Config/carrier/UOM editors
are present, and opening the surface leaves Config file bytes unchanged.

| Failing check | Observed missing behavior |
|---|---|
| `TrackingSettings.GeneralAndEventTrackingTabs` | No General/Event Tracking pages |
| `TrackingSettings.TrackingDetailAndActionPathSections` | No dedicated Tracking, Event Detail and Action Paths sections |
| `TrackingSettings.SeparateSaveReloadResetActions` | No tracking-page policy/profile/preference save, Reload and Reset to Default controls |
| `TrackingSettings.CaptureDefaultsOff` | No Capture recorded controls checkbox; this does not claim an existing checkbox defaults incorrectly |
| `TrackingSettings.PersonalViewChoices` | No personal selector for Use warehouse default, How-To, Diagnostic and Compare both |

The observer accepts nested page/frame layouts and fixed product captions. The
personal selector name `cmbPreferredActionPathView` is a proposed implementation
identifier, not a newly approved architectural requirement. These assertions
prove missing surface behavior only. Control presence cannot prove persistence,
authorization, immutable versions, isolation, atomicity or recorded conclusions.

## Calibration and preserved evidence

An initial setup attempt installed the observer after constructing the form.
VBA project editing reset the helper's module state, producing run-time error 91
when it tried to show the lost form reference. The run stopped at 2 PASS / 1
harness failure and recorded an Excel native fault. It is **not** product RED.
Install all observation code before fixture setup/form creation; do not repeat
live project edits with an existing form instance. The final helper follows that
ordering. The empty recovery Excel instance was verified, quit and then stopped
after it lingered; no operational workbook was closed or changed.

Ignored evidence under `reports/runtime/slice4be-tracking-settings/`:

- `e10fff53d4a74ead986b84a44a2c312f/red.json`: completed 21/5 surface RED.
- The same directory's `tracking-settings-open.png`: inspected actual form with
  General editors and no tabs. Internal disposable-fixture capture; not human UAT
  and not an artifact to publish without review/redaction.
- `ecc8a9e2c5b64cb98c4ed2b2353f8cd9/red.json`: rejected setup attempt.
- `surface-red-comparison.json`: all 18 baseline identities/GREENs retained.
- `surface-native-windows.json`: one observed Excel fault in initial setup,
  zero in the completed run, using bounded report-creation/completion windows.

Excel is closed. The existing 65 package and 16 runtime source pins match after
the run. No accepted deployment, operational workbook or NAS runtime changed.
PowerShell parse/diff checks pass. Static maintenance is regenerated for this
checkpoint: 183 components, 5,527 procedures, 8 literal Application.Run calls,
45 unresolved calls and 195 duplicate-body candidates remain unchanged. All 28
module limits and all three report schemas pass. No fresh build, explicit
compile, full Release 1 chain, live-role GREEN or human acceptance is claimed.

## Required continuation

Implement the approved Settings surface and add focused packaged action tests
before its persistence implementation. Exercise the actual tracking policy save,
detail profile save, personal save, Reload, Reset and Close handlers; missing
seams/compile failures must never substitute for behavioral RED. Keep these
requirements explicit:

- D5 Core ownership, captured session/warehouse/station, ADMIN_MAINT at editor
  open/save and rejection after session/context/capability change.
- Whole-request validation, expected version, append-only policy/profile rows,
  required audit collection, command-on/navigation-off/capture-off defaults,
  unknown-column preservation and atomic compatibility views.
- Stale/unknown/duplicate/invalid requests, dirty/read-only/locked/missing Config,
  malformed policy errors without repair and unchanged bytes on rejection.
- Synthetic preview; separate scopes; Reset stages only and Close discards edits.
- Operations personal Settings without Admin installed, Windows/invSys-user and
  warehouse isolation, restart restoration and invalid-preference fallback.
- Both Action Path presentations and current policy-aware evidence availability;
  no collection, business authority or guide success inferred from a preference.

Retain all current GREEN regressions and complete packaged compile/layout/static,
live-role, full-chain and visible comparison gates before Slice 4be acceptance.
D8-A approval and the separate native gate failures remain unresolved.
