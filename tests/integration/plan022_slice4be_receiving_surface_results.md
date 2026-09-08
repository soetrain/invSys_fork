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

Next: calibrate native worksheet input in an available desktop and prove the
actual Shape caller/handler before specifying and implementing its activity
identity, captured context and owner-returned outcomes. Other comprehensive
Operations/Admin coverage and all later 4be deliverables remain pending.
