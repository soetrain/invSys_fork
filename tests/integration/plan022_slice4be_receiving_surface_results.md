# Slice 4be.1 Receiving worksheet surface discovery

Architecture v4.11 D18 requires accounting for reachable controls and evidence
before excluding retired controls. This is coverage discovery, not a new runtime
contract or an activity implementation checkpoint. D13 behavioral RED/GREEN for
worksheet activity remains pending. No VBA, packages or normative rule changed.

The preserved candidate is `deploy/validation-receiving-navigation-identity`.
[Its prior navigation/identity record](plan022_slice4be_receiving_navigation_results.md)
retains 771/771 and the complete technical gates; those gates are not claimed as
rerun by this discovery-only test.

## Packaged surface probe

Last verified 2026-09-08: **105/105 PASS**, zero duplicate check identities,
normal test exit and Excel closed. All five candidate hashes match the preserved
navigation/identity manifest. PowerShell parsing and diff whitespace checks pass.
Runtime source, accepted deployment and NAS were not changed.

`Slice4beReceivingSurface.ps1` runs through the actual generated Operations
Receiving Ribbon callback using the existing unsaved callback/form seams.
Fixtures enter through Admin Generate Warehouse and Seed. It checks provisioned,
reused, sole-visible-ReceivedTally and saved/reopened operator workbooks.

- Provisioned and ordinary reused support sheets remain VeryHidden.
- When ReceivedTally is the only visible worksheet, the launcher preserves that
  visible sheet and its Confirm Writes button, including after save/reopen.
  The launcher accepts and captures this existing operator workbook.
- The expected public handler remains assigned. Excel reports it as unqualified
  `modTS_Received.ConfirmWrites` in these cases. Assignment is distinct from
  evidence of native invocation into the intended package.
- The unknown header, saved operator bytes, warehouse authority files and an
  unrelated workbook are preserved.
- Reports explicitly mark `NativeInvocationVerified=false`. This probe does not
  claim a click, posting result, activity correlation or visible human acceptance.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-receiving-navigation-identity -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CheckReceivingLocalActivity -CheckReceivingLifecycleActivity -CheckReceivingNavigationActivity -CheckReceivingSurfaceCoverage -ReceivingSurfaceOnly
```

## Rejected test assumptions and unresolved input proof

The first probe had 98 PASS / 4 FAIL because it required a package qualifier in
OnAction. The observed unqualified handler invalidates that string assumption;
it does not establish a runtime routing defect. A posted-message probe had
98 PASS / 8 FAIL without handler entry or a macro-unavailable notice. Input
delivery was not calibrated, so these failures are not meaningful D13 RED.

A separate calibration guard stopped at 92 PASS / 1 harness exception. Minimal
disposable Excel transport experiments then showed that the session could not
establish foreground ownership and Windows rejected cursor positioning. No
native-click or worksheet activity GREEN is claimed. Preserve the diagnostic
attempts; do not change runtime routing to satisfy an unproved test assumption.

Ignored local evidence is under `reports/runtime/slice4be-receiving-activity/`:
`surface-reachability.json`, `diagnostic-surface-green.json`,
`surface-visibility-discovery.log`, the earlier surface discovery logs and
`surface-native-input-diagnostic.ps1`. Reports contain counts/booleans only;
the diagnostic script contains unsaved test seams, not a runtime correction.

## Native caller proof after desktop recovery, 2026-09-12

The current discovery checkpoint is **115/115 PASS**, with all previous 105 check
identities retained and no duplicates. The candidate is the original pinned
`deploy/validation-receiving-launcher-denial`, whose unchanged standard Production
and restart gate is 2/2; see the [recovery record](plan022_slice4be_production_release_boundary_diagnostics.md).
This run does not repeat or replace that Production gate or the full 845-check
activity regression.

The minimal disposable input calibration passes **8/8**. It verifies foreground
ownership, actual handler entry, exact shape caller, workbook binding, a distinct
programmatic call, rejection of another foreground workbook window, rejection of
an inactive fixture, and native entry after switching between two workbooks.
The developer helper uses native mouse input, checks the exact workbook window,
and calculates the click from worksheet origin, window DPI and worksheet zoom.
It supports the deliberately unscrolled fixture; it is not operator runtime code.

The packaged test adds an unsaved counter and caller observation at the real
`modTS_Received.ConfirmWrites` entry. A separate fixture button calibrates input
in that same package/workbook before the original `btnConfirmWrites` is clicked.
Both OnlyStagingVisible and SavedReopen enter the intended owner exactly once,
with `Application.Caller = "btnConfirmWrites"` and no macro-unavailable notice.
Their `NativeInvocationVerified` values are true. Hidden provisioned/reused
surfaces remain unclicked and false. Saved workbook bytes, unknown header,
authority files and unrelated workbook checks pass. Both actual worksheet
captures were visually inspected; the marked input point is on Confirm Writes.
These automated fixture captures are not human operator acceptance.

Earlier recovery attempts stopped at **94 PASS / 1 harness failure**. A marked
minimal screenshot first showed unscaled coordinates above its button. Packaged
captures later showed the unrelated sentinel workbook in front; broad Excel
process ownership was insufficient. An intermediate COM parent lookup and then
explicit active-context/foreground guards also stopped before native entry.
The working setup explicitly shows and activates the disposable workbook window,
then activates its worksheet and verifies context. The window's Visible property
was already true before preparation in both passing cases; do not claim a hidden
workbook runtime defect or isolate visibility assignment as the cause. These
setup failures are not D13 product RED, and no runtime routing repair was made.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beWorksheetInput.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-receiving-launcher-denial -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CheckReceivingLocalActivity -CheckReceivingLifecycleActivity -CheckReceivingNavigationActivity -CheckReceivingSurfaceCoverage -ReceivingSurfaceOnly -CheckReceivingNativeSurface
```

The native switch requires the separate surface-only run and writes to a unique
ignored directory, preserving earlier reports. Exact local evidence:

- `reports/runtime/slice4be-receiving-activity/native-surface-230a1587ef9d46d2b307ddbce4e11810/`:
  `diagnostic-surface-green.json`, `surface-reachability.json`, and the two
  `confirm-*.png` captures;
- `native-surface-visible-window.log` in the parent directory;
- `native-calibration-*/checks.json` records the small input calibrations;
- earlier `native-surface-*.log` files retain the unsuccessful setup attempts.

All 20 original/rebuilt/saved/compiled candidate pins remain unchanged. Excel is
closed. PowerShell parsing, diff checks and reference checks protect this test
and documentation checkpoint. Runtime, normative architecture, accepted deployment,
NAS and static baselines are unchanged; no new build/compile/layout/live-role/
full-chain acceptance is claimed.

Next: refine the explicit worksheet activity identity under D18 and run meaningful
RED through this calibrated native control before changing its implementation.
The existing worksheet handler calls the posting owner directly; its form-action
activity tests do not cover this route. Captured context, source references and
owner outcomes need protecting tests. Other comprehensive Operations/Admin
coverage and all later 4be deliverables remain pending.
