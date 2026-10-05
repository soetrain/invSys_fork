# Slice 4be.1 Receiving native input target calibration

This repairs validation input routing under Architecture v4.11 D13/D18. It changes
no Receiving implementation, activity contract, business authority or package.
The original calibration used `deploy/validation-shipping-prewrite-facts`;
the latest regression below uses `deploy/validation-user-policy-02`.
Release1 and comprehensive Slice4be acceptance remain open.

## Per-user policy candidate regression, 2026-10-05 UTC

Candidate02 retains **854/854** exact prior ordered checks, including native input,
redaction, launcher reuse/denial and preservation. Controller
`receiving-owner-regression/821e058c51134bd39ec2da9045705bdc` runs
04:15:00.165--04:25:42.174 UTC; comparison/closure and50 evidence files are retained,
including `evidence/launcher-denial-green.json`. Settings/packages are preserved,
Excel closes normally, and fresh Receiving/Returns captures are reviewed. No
test-helper correction or runtime change was needed. The final native audit has
zero Excel failures; exact receipts and limits are in
[A1 policy evidence](plan022_slice4be_user_tracking_policy_results.md).

## Receiving runner regression, 2026-10-05 UTC

Frozen candidate05 passes **854/854**, retaining the exact prior854 identities
and order with no duplicates. No runtime source or package changed in this
checkpoint. Initial execution returned796 PASS/58 FAIL:32 supported-read checks
and26 navigation-caption checks used a stale catalog3--11/6--11 allowlist, while
Core now emits catalog26. These are harness failures, not product RED.

Both helpers now retain each control's introduction floor and call the real Core
reader for every emitted record before combining assertion results. Core owns
supported-version/control validation; the tests no longer impose a stale ceiling.
Unsupported catalog999 and a control absent from catalog1 remain rejected.
Actual keyboard/mouse routing, fixed captions, redaction, exact references,
staging/System_Key/unknown columns, context/rights refusal, launcher reuse and
authority/unrelated-workbook preservation all pass.

Ignored receipts under `reports/runtime/receiving-owner-regression/`:

- `f59f51ec18344f6e84e5700c0f080bda/`: initial796/58,
  02:06:10.594--02:16:30.227 UTC; `harness-failure-summary.json` classifies all58.
- `320fd57912514137be403e585af3e7f6/`: corrected854/0,
  02:17:24.495--02:28:22.563 UTC; `comparison.json` proves retention.
- Each controller has `worker.log`, `closure.json`, `evidence-manifest.json` and
  preserved `evidence/launcher-denial-green.json` plus fresh captures.

Both runs restore settings, preserve all five package hashes and close Excel
normally. The final controller's native audit finds zero Excel1000/1002 events
from01:57:51 through02:28:45 UTC, covering the93-check runner-close regression too.
Desktop cursor probe02:28:45.680 UTC (Oct4 19:28 PDT) succeeds; no desktop error5
was observed. The earlier full-chain native failure remains unresolved.

Inspected fresh `coverage-lifecycle-open.png` and `coverage-disposition.png` show
the generated-fixture Receiving/Returns surfaces, buttons and status. This is
scoped operator evidence, not Action Path long/multiline or human acceptance.
Static `receiving-run-validation-static-06/` retains all metrics/caps and passes
three schemas and397 script parses. Candidate05's five cold compiles remain valid.
The protecting command below uses candidate05 and the full-regression switches
for this run. Next complete A1 recording policy/observations and A2 visible proof;
broader live-role/full-chain gates and B0/A/B/R1 acceptance remain open.

## Protecting evidence

The preserved Shipping checkpoint exposed Receiving navigation stopping before
its first enabled tab action. Existing traces showed native KeyDown/Change/owner
entry for the default-off action and none after the unrelated workbook activated.
The native helper checked only Excel-process ownership of the focused window.

A read-only native ancestry probe now calibrates the exact failing transition:
focus is inside the intended form after preparation, then outside it after the
unrelated fixture workbook activates. The test calls the existing native key
helper and expects refusal. Handles remain in memory; reports contain fixed
check names and booleans. Independent checks retain no form callback, no activity
and unchanged unrelated-workbook bytes.

An initial probe before policy setup did not reproduce that transition and failed
calibration; it is not behavioral RED. At the actual transition, the focused run
completes113 checks:111 PASS/2 FAIL. `Navigation.NativeTarget.UnrelatedFocusRejected`
fails because the helper accepts the wrong window; the existing keyboard-action
harness exception also remains. The other five new calibration checks pass.

## Harness correction

The native key helper requires the focused window to equal the intended form
window or descend from it, in addition to the existing process/visibility checks.
For legitimate actions, an unsaved Excel helper activates the form's top-level
window on Excel's own UI thread and focuses only the requested control. It checks
the actual thread-active window and the form's logical active control. Native
KeyDown/Change and real owner dispatch still determine the test outcome; no
selection, refresh, owner call or activity record is manufactured by refocusing.

The activation call follows the [SetActiveWindow thread requirement](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setactivewindow).
Calling MSForms SetFocus alone left native focus outside the form; requesting
foreground activation from PowerShell did not pass its native check. Both
approaches are superseded. No desktop/foreground-lock settings are changed.

A provisional guard requiring Excel's active workbook never to change during
form focus stopped at the third enabled tab. That is not an architectural
invariant: D18 excludes focus/window mechanics while requiring captured workbook
authority. The harness instead accepts an activation change only when the newly
active workbook is the exact captured live object, using the existing binding
probe. It rejects every other change. Existing staging/System_Key/unknown-column,
authority-byte, unrelated-workbook and stale/closed-context checks remain.

Navigation now makes the isolated Excel fixture visible, following the existing
native worksheet fixture, and restores its prior visibility during cleanup.
Mouse still requires its real screen point to belong to the isolated Excel
process. The hidden fixture did not establish that condition; discarded
foreground/topmost attempts are not retained. Failure diagnostics now distinguish
whether a screen point has a window and lies on a monitor, without logging
coordinates.

The final closed-workbook fixture creates and shows the form while the unrelated
workbook hosts its native window, while retaining the operator workbook as its
explicit business binding. It proves the exact captured object before closure,
the same native form after closure, and the unchanged stored workbook reference.
It then delivers the actual key and verifies no detail-owner entry, visible
reopen guidance, no activity and no unrelated-workbook redirection. This avoids
losing the form window when closing its native host and does not rebind, recreate
or initialize the form after the captured workbook closes.

## Current validation

Focused GREEN passes266/266, retaining all111 prior RED-run GREENs with no
duplicate check identities. The wrong-focus refusal is now GREEN. All13 keyboard
and13 mouse controls pass their activity and real Core-read checks, as do policy,
stale-session and closed-workbook cases. Full Receiving passes **854/854**,
retaining every check and GREEN from the preserved **845/845** baseline, with
nine added calibration checks and no duplicates. The catalog8 allowlist
corrections committed in807c7e1 now pass their actual Core reads throughout this
full run; unsupported-version rejection remains protected.

All65 package pins and16 source pins match with Excel closed. No runtime source
or package changed, so the prior candidate's explicit package compiles remain
applicable. Static evidence is regenerated:183 components,5527 procedures,
8 literal Application.Run targets,45 unresolved calls and195 duplicate-body
candidates, unchanged. All28 module limits and all three JSON schemas pass;
both changed/new harness scripts parse. No Excel Application1000 fault is observed
in the twelve calibration/focused/full run windows with three seconds' margin.
This does not erase the separately recorded live-role/full-chain/launcher faults.

The fresh Receiving form capture was inspected: tabs, result/history lists,
tally areas and action buttons are visible. It is generated-fixture operator
evidence, not human release acceptance. D8-A approval, separate native stability
failures, remaining Operations/Admin coverage and comprehensive Viewer/Action
Path implementation/comparison remain open.

Ignored evidence under `reports/runtime/slice4be-shipping-activity/` includes
`navigation-target-initial-calibration.json`, `navigation-target-boundary-red.json`,
the matching RED logs, and the separate focus/window/thread/classified attempts.
`navigation-target-focused-green.json`, `navigation-target-focused-comparison.json`
and `navigation-target-hosted-green.log` retain the completed focused proof.
`navigation-target-full-receiving-green.json`, `navigation-target-full-comparison.json`
and `navigation-target-full-receiving-exit.json` retain the full854 proof.
`navigation-target-native-windows.json`, `navigation-target-final-pin-verification.json`,
`navigation-target-visible-evidence.json` and `navigation-target-operator-evidence.png`
retain the bounded fault, integrity and visual evidence.
The exact protecting command is:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-prewrite-facts -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CheckReceivingLocalActivity -CheckReceivingLifecycleActivity -CheckReceivingNavigationActivity -ReceivingNavigationOnly
```

GREEN uses the same route with `-Phase GREEN`. Full regression adds
`-CheckReceivingSurfaceCoverage -CheckReceivingLauncherDenial -CaptureEvidence`
and omits `-ReceivingNavigationOnly`.
