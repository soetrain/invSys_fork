# Slice 4be.2 Event Tracking Settings

## Current checkpoint: tabs implemented, editors incomplete

Last verified 2026-09-13. The isolated candidate
`deploy/validation-tracking-settings-tabs-final` now has the approved General and
Event Tracking pages. General retains the Config, connection, carrier and UOM
editors and their existing handlers. The Close button and status label remain
shared outside the pages. Event Tracking has Tracking, Event Detail and Action
Paths sections, explicitly labelled unavailable until their editors are built.
This is an intermediate implementation, not completion of 4be.2 or a substitute
for the approved policy/profile/preference behavior below. No normative contract
changed. Page activity instrumentation remains part of outstanding 4be.1 coverage.

The expanded old-package RED is **31 checks: 23 PASS / 8 FAIL**. The final
candidate completes the same **31 checks: 28 PASS / 3 FAIL**, retaining every
check and prior GREEN, with no duplicate identities. The focused tab/layout group
is **10/10 GREEN**. Both pages are selected and measured in their displayed
state; opening and switching leave Config bytes unchanged. All **18/18 D5**
checks still pass, including real scalar Settings/UOM handlers, stale target,
denials, dirty/read-only Config and unknown-column preservation.

The three remaining failures are exactly `SeparateSaveReloadResetActions`,
`CaptureDefaultsOff` and `PersonalViewChoices` under `TrackingSettings`.
The overall command deliberately remains in RED phase and exits 1. No incomplete
policy/profile/preference workflow is relabelled GREEN.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-tracking-settings-tabs-final -Phase RED -CheckTrackingSettings -CaptureEvidence
```

All five packages build and explicitly compile, including the Operations
cold-start dependency check. Packaged smoke passes **86/86**. Comparing the 176
compiled component hashes before/after the footer correction changes only
`invSys.Admin.xlam/frmAdminSettings`. The form is 744 by 696 points; the larger
height fixes a measured status-label overflow after the native resizable frame
is applied. The 484-line form remains below the new-module line limit, and all
28 existing module limits pass. Static metrics remain 183 components, 5,527
procedures, 8 literal and 45 unresolved Application.Run calls; duplicate-body
candidates decrease from 195 to 190 as the form's five control builders now use
the selected parent container. No source was deleted based on scanner findings.
All three regenerated report schemas pass. All 75 package pins (65 existing,
five initial tabbed and five final tabbed) match, the 16 preserved runtime source
pins match, and Excel is closed. Nine bounded validation/build/compile/smoke
windows show no observed Excel Application 1000 fault; this does not erase the
separate historical native failures.

Additional ignored evidence under `reports/runtime/slice4be-tracking-settings/`:

- `4f396dea9ddc42879b361cb24e812556/red.json`: expanded 23/8 old-package RED.
- `061b605eb6df473cb212b30ba7cbb82b/red.json`: the first tabbed candidate's
  26/5 result, including the real footer overflow on both displayed pages.
- `7876672f44634d98a1fad1e33355c209/red.json`: final 28/3 result.
- That final directory's `settings-save.png` and `tracking-settings-page.png`:
  inspected General save/status/footer and Event Tracking sections. The early
  `tracking-settings-open.png` has incompletely painted child controls and is
  not the General visual acceptance image. These are internal fixture captures,
  not human acceptance or publishable unredacted operational evidence.
- `tabs-final-build.log`, `tabs-final-compile.log`,
  `tabs-final-compiled-sources.json`, `tabs-final-package-hashes.json`,
  `tabs-final-test.log`, `tabs-packaged-smoke.log`, `tabs-packaged-report.md`,
  `tabs-final-comparison.json`, `tabs-static.log`, `tabs-native-windows.json`,
  `tabs-final-pin-verification.json`.

Do not measure an inactive MultiPage's cached initial client dimensions as its
displayed layout. The first new observer saw 141 points for the never-selected
Event Tracking page. The corrected observer checks General, selects Event
Tracking and independently checks that displayed page. It then exposed the real
footer defect: status bottom 650 versus client height 647.6 points. Increasing
the form height from 682 to 696 fixes that defect; no controls or checks were
removed. The fixture makes its isolated Excel instance visible for painting and
restores its previous visibility in `finally`; native rendering can still be
incomplete in an immediate PrintWindow capture, so use the inspected saved-form
capture above. These test calibrations do not alter runtime authority or policy.

Tracking policy/profile saves, personal preference isolation/restart, Operations
Settings without Admin, recording, method comparison and comprehensive event
coverage remain required. No new full-chain/live-role GREEN, NAS proving or human
acceptance is claimed by this intermediate UI change.

## Initial protecting surface RED

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

At that initial checkpoint, Excel was closed and the existing 65 package and
16 runtime source pins matched after the run. No accepted deployment, operational
workbook or NAS runtime changed.
PowerShell parse/diff checks pass. Static maintenance is regenerated for this
checkpoint: 183 components, 5,527 procedures, 8 literal Application.Run calls,
45 unresolved calls and 195 duplicate-body candidates remain unchanged. All 28
module limits and all three report schemas pass. No fresh build, explicit
compile, full Release 1 chain, live-role GREEN or human acceptance is claimed.

## Required continuation

Next implement the staged tracking-policy editor and its real Save Tracking
Policy/Reset handlers, establishing packaged behavioral RED for version
publication and rejection paths before implementing the Core persistence body.
Continue with the detail profile and personal preference editors. Exercise the actual tracking policy save,
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
