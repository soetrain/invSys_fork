# Slice 4be.2 Event Tracking Settings

## Current checkpoint: cancelled saves cannot report success

Last verified 2026-09-13. Existing D18/D5 requires a verified persisted policy
version before reporting save success. A real Excel `WorkbookBeforeSave` test
observer cancels only the disposable Config workbook's save. The old candidate
leaves disk bytes/version unchanged but returns success, displays a saved message
and reloads away staged edits. This is meaningful RED through the same Settings
Save action used by the operator, not an injected exception or replaced writer.
All probe code is installed before form/fixture creation; the temporary observer
is disarmed and Excel's prior EnableEvents state restored after the action.

Expanded RED is **65 checks: 61 PASS / 4 FAIL**. The two new failures are
`TrackingPolicy.CancelledSaveNeverReportsSuccess` and
`TrackingPolicy.CancelledSaveRetainsStagedEdits`; the other two remain the broad
profile/preference gaps. Candidate `deploy/validation-tracking-policy-cancel`
checks `Workbook.Saved` after the single save call and enters the existing
unverified-save cleanup on cancellation. Settings reports uncertainty and keeps
staged edits. No normative contract or Auth provisioning/read semantics change.

The corrected candidate completes **65 checks: 63 PASS / 2 FAIL**, including
**34/34 policy** and **18/18 D5** checks. The same expanded test also proves all
three per-control flags stage without writes, Reload discards those edits, the
flags persist together after an authorized save, and the actual Close handler
discards staged capture changes without saving. Every earlier GREEN/check
identity remains, with no duplicates. The overall suite intentionally remains
RED until the profile/preference workflow is implemented.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-tracking-policy-cancel -Phase RED -CheckTrackingSettings -CheckTrackingPolicy -CaptureEvidence
```

Five packages build/explicitly compile and Operations cold start passes;
packaged smoke is **86/86**. Compiled-source comparison across all 180 components
finds exactly one changed component, `modTrackingPolicyCommand`. The previous
70/70 activity and 854/854 full Receiving runs remain evidence for their prior
candidate; they were not rerun for this isolated policy-save correction. The
current focused suite covers the changed command. No new full-chain/live-role,
physical NAS or human acceptance is claimed.

The new saved-policy capture was inspected: selected Receiving Confirm has all
three flags off, capture remains off, Compare both remains the warehouse default,
and version 2 is displayed after successful save. This is fixture evidence,
not a screenshot of the cancellation notice or human UAT. Older-policy editor,
protected/missing Config and interrupted-save/caller-open rollback edge cases,
new-control activity coverage and policy-save observations remain required.
Proceed with Event Detail profile and personal preference implementation under
D18 while retaining those cases as explicit acceptance work.

Ignored evidence in `reports/runtime/slice4be-tracking-settings/`:

- `b078a0e9dcdf490184f7eddb959e7071/red.json`: expanded cancelled-save RED.
- `de5b71ff932d413a9a3b8865fbc382a9/red.json`: corrected 63/2 result and
  inspected `tracking-policy-saved.png`.
- `policy-cancel-red.log`, `policy-cancel-green.log`, `policy-cancel-comparison.json`.
- `policy-cancel-build.log`, `policy-cancel-compile.log`,
  `policy-cancel-compiled-sources.json`, `policy-cancel-component-comparison.json`.
- `policy-cancel-smoke.log`, `policy-cancel-smoke-report.md`.
- `policy-cancel-static.log`: 187 components / 5,565 procedures; literal and
  unresolved dynamic calls remain 8/45 and duplicate-body groups remain 190.
  All three regenerated report schemas pass.
- `policy-cancel-package-hashes.json`, `policy-cancel-pin-verification.json`:
  five new candidate hashes recorded; all 85 prior package pins and 16 preserved
  runtime source pins match. Operational workbooks and accepted deployment are
  untouched; Excel is closed.
- `policy-cancel-size-ratchets.json`: all 28 existing module limits hold; the
  policy command is 143 lines and its Save procedure is 109 lines.
- `policy-cancel-native-windows.json`: no observed Excel Application 1000 fault
  across five bounded RED/build/compile/corrected-test/smoke windows. Historical
  native failures remain unresolved.

## Previous checkpoint: staged tracking policy and Core persistence

Last verified 2026-09-13. The isolated five-package candidate
`deploy/validation-tracking-policy-save` adds the Admin-owned tracking editor
and headless Core whole-policy save under existing Architecture v4.11 D18/D5.
No architectural rule changes. Detail profiles, personal preferences and the
new Settings controls' activity coverage remain incomplete; 4be.2 stays open.

The protecting no-write candidate `deploy/validation-tracking-policy-red`
completes **47 checks: 44 PASS / 3 FAIL**. The real class Save action is entered,
but `TrackingPolicy.AuthorizedSavePublishesVersion` fails. The other two failures
are the existing broad Settings assertions for separate profile/preference
actions and personal view choices. This is behavioral RED, not a missing seam
or compile failure. The persistence candidate completes **56 checks: 54 PASS /
2 FAIL**, including **25/25 TrackingPolicy** and **18/18 preserved D5** checks.
Both broad failures remain; all earlier tabbed-candidate GREENs are retained.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-tracking-policy-save -Phase RED -CheckTrackingSettings -CheckTrackingPolicy -CaptureEvidence
```

The tests stage and reset through the same class action bodies used by the
operator, and save/reload versions 1 and 2 through that action. Supplemental Core
tests reject unknown fields/controls, missing/duplicate controls, invalid flags
and view, stale versions, signed-out/capability/stale-session/cross-target
contexts and read-only/dirty Config. The second save shifts managed columns by
inserting an unknown first column in both tables, then verifies all prior rows
and unknown values survive unchanged. Requests remain in memory; report JSON
contains only fixed check identities and Boolean outcomes.

Core validates the existing latest policy through the existing reader, stages
the entire request, appends metadata and control rows, then saves Config once.
Both compatibility flags and the warehouse default are in that same version.
This does not claim a new generic scalar Config compatibility API. Both tables
absent yields labelled version-0 defaults; controls absent from a valid older
catalog remain unavailable until explicit policy update. Default sequence
eligibility in the editor is permission to include registered controls, not
automatic recording: capture remains off and navigation collection defaults off.
Required audit/business collection cannot be edited here.

The inspected `tracking-policy-saved.png` shows version 2, Compare both as the
warehouse default, capture off, per-control flags and all three policy actions.
It is an internal disposable-fixture capture, not human acceptance. New-control
coverage, policy-save version/outcome observations, Close-discard proof, detailed
per-control interaction/older-policy editor tests, cancelled/interrupted-save recovery and
the full profile/preference/Operations-without-Admin scope remain outstanding.
No new full-chain, live-role, NAS or visible user-comparison acceptance is claimed.

Ignored evidence under `reports/runtime/slice4be-tracking-settings/`:

- `14f7b73e7f6c468f984bbd7aef41fe08/red.json`: meaningful policy save RED.
- `79e30c97bd484d8da1a27f21e023372d/red.json`: final 54/2 candidate and inspected
  `tracking-policy-saved.png`.
- `policy-red-build.log`, `policy-red-compile.log`, `policy-red-compiled-sources.json`;
  `policy-save-build.log`, `policy-save-compile.log`, `policy-save-compiled-sources.json`:
  both candidates build and all five projects explicitly compile, including
  Operations cold start.
- `policy-action-red-test.log`, `policy-save-final-test.log`.
- `policy-activity-regression.log`: shared activity foundation **70/70**.
- `policy-comparison.json`: all 28 prior tabbed GREENs and all 44 policy-RED
  GREENs retained, no duplicate identities; all 70 prior activity GREENs retained.
- `policy-static.log`, `policy-size-ratchets.json`: regenerated static reports
  contain 187 components and 5,565 procedures. Literal/unresolved Application.Run
  counts remain 8/45; duplicate-body groups remain 190. All 28 prior oversized
  module limits hold. The four new modules are at most 170 scanner-counted lines,
  their procedures at most 107; Settings is 490. All three report schemas pass.
  These are maintenance checks, not authority to delete scanner candidates.
- `policy-full-receiving.log`, `policy-full-receiving-green.json`,
  `policy-receiving-comparison.json`: **854/854**, preserving every prior check
  identity and GREEN without duplicates. The runner's original output is
  `reports/runtime/slice4be-receiving-activity/launcher-denial-green.json` for
  this flag combination, not the historical shorter `green.json`.
- `policy-packaged-smoke.log`, `policy-packaged-report.md`: **86/86**.
- `policy-source-import.log`, `policy-source-import-report.md`: source-backed
  harness imports all three new Core helpers and passes selected test 1/314
  (**1/1**); this does not claim all 314 source-backed cases were run.
- `policy-red-package-hashes.json`, `policy-save-package-hashes.json`,
  `policy-additional-pin-verification.json`: 10 policy-candidate hashes pinned
  and matching after smoke; all 75 earlier package pins and 16 preserved runtime
  source pins also match. Accepted deployment and operational NAS workbooks were
  not changed. Excel is closed.
- `policy-native-windows.json`: no observed Excel Application 1000 faults in
  eleven bounded setup/build/compile/test windows. Earlier historical native
  faults remain open; this evidence does not explain or erase them.

The earlier `policy-red-test.log` attempt failed in test setup because a disk
hash was requested while Excel held a read/write fixture handle. It is not
product RED. The corrected test inspects in-memory dirty state first, closes
only its disposable fixture without saving, then compares the file hash.

## Previous checkpoint: tabs implemented, editors incomplete

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

Continue with the detail profile and personal preference editors using focused
packaged behavioral RED before their persistence implementation. Exercise the actual tracking policy save,
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
