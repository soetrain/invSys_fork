# Plan 022 Slice 4be.1 Shipping access interruptions

Last verified: 2026-09-12. Runtime remains code **22e14b6**, packaged in
`deploy/validation-shipping-capability-guard`. This test-only checkpoint extends
the [permission evidence](plan022_slice4be_shipping_capability_results.md);
it does not implement Shipping activity or approve a new Auth contract.

## Protecting scope

`tests/tooling/Slice4beShippingAccessInterruptions.ps1` is installed by the existing
packaged Shipping activity fixture. The unsaved form hook changes the generated
fixture's SHIP_POST at the actual `ShowPersistencePending` DoEvents return for Add,
To Shipments and Shipments Sent. Core authorization executes normally. Separate
cases move the generated Auth or Config file aside before each of the seven real
mutation handlers. The existing calibrated probes stop at mutation-owner entry;
they do not replace the form handler or the Core permission check.

The 17 cases each check interruption/file state, unavailable Core permission,
unchanged signed-in session, no owner entry, exact permission notice, unchanged
active/held staging with keys and unknown values, and retained workbook binding.
Four final checks cover restored Auth/Config bytes, canonical inventory bytes and
the unrelated workbook. These 123 checks supplement all435 prior checks/389 GREENs.
They prove dispatch/access handling, not successful business submission or rollback.

## Retained incomplete run and fixture correction

The first run, `access-interruptions-green.log` and
`d9b6adf2bff54372aedafd8ec0188813/green.json`, stops with401 PASS/48 FAIL.
All three actual-yield revocations pass. The first missing-Auth Add rejects the
mutation but recreates Auth. Cleanup intentionally refuses to overwrite that
unexpected file; a subsequent cleanup sign-in fails with fixed status3 and masks
the first cleanup exception. This is incomplete matrix evidence, not435-check
regression preservation or a fully calibrated Auth-contract RED.

The corrected helper retains the failed creation assertion, moves each unexpected
file to a unique sibling inside the verified generated-fixture root, and restores
the original file. It neither overwrites nor deletes unexpected authority files.
After an incomplete helper, cleanup does not attempt another sign-in that could
mask its error. Only disposable fixture files are moved. This is harness repair;
no Shipping/Core implementation or package changed.
Unexpected Auth copies are preserved within the helper until the outer harness
disposes the generated fixture root. Raw Auth workbooks are not retained as
diagnostic artifacts; the sanitized reports preserve the observations.

## Completed unchanged-package result

The corrected run completes **505 PASS /53 FAIL across558 checks**, exit1:

- All435 previous check identities and389 previous GREENs are retained, with
  no duplicate identities or unexpected failures.
- The123 added checks produce116 PASS and seven missing-Auth creation findings.
  All three actual-yield revocations and all seven unavailable-Config cases pass.
  Each unavailable-Auth action rejects at the permission guard, but creates Auth.
- The53 failures are the46 unchanged missing-activity assertions plus those seven
  proposed D8-A no-creation assertions. This is not an all-GREEN activity suite.
- Actual workbook-close checks still pass after the access matrix. Auth/Config
  bytes are restored; canonical inventory and unrelated workbook bytes match.

Exact command, with output captured as `access-interruptions-complete.log`:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-capability-guard -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity
```

All raw paths in this report are relative to the ignored
`reports/runtime/slice4be-shipping-activity/` root. The complete report is
`3f7c34d55b434f09a0ab18628b13c90c/green.json`; the identity/result comparison is
`access-interruptions-comparison.json`. The initial incomplete run is retained
separately. Both focused runs are terminal and Excel is closed.

All50 package/eight current Receiving/Shipping source pins match in
`access-interruptions-pin-verification.json`. Static maintenance is regenerated
(`access-interruptions-final-static.log`); all three JSON contracts and all28
preceding module limits pass. Runtime counts remain177 components,5511 procedures,
8 literal Application.Run targets,45 unresolved dynamic calls and195 duplicate
body groups. Both changed PowerShell scripts parse. Static differences are
timestamps, report hashes and additional test references; runtime source is unchanged.
These proportional checks do not repeat or replace the earlier candidate gates.

## Architecture decision discovered

Shipping correctly rejects unavailable permission, but `modAuth.LoadAuth` calls
`ResolveAuthWorkbook`, which uses `OpenOrCreateAuthWorkbookRuntime` before its
other fallbacks. The common runtime opener creates, seeds and saves missing Auth.
D5's explicit read-only rule applies to Config. Architecture v4.11 also retains a
Phase6 checked Config/Auth auto-bootstrap acceptance entry; D14 prohibits dirtying
unchanged healthy Auth but does not explicitly prohibit missing-Auth provisioning.
Therefore the test's blanket no-recreation expectation cannot be promoted into
an approved Auth rule by this evidence or by Plan022.

The normative specification, Plan022 and controls catalog now carry **pending
D8-A: Auth read/provisioning separation**. It proposes read-only exact-target Auth
load/refresh and fail-closed unavailable authority, retaining explicit authorized
warehouse/station setup. User approval is required before implementing that new
contract. Missing-Auth creation assertions remain visible as proposed-contract
failures; the existing permission rejection is separately passing.

## Remaining acceptance

Shipping's46 missing activity assertions remain open. Pending/uncertain submission,
optional tracking failure, policy behavior and precise Shipping/Boxing catalog
definitions remain necessary before comprehensive observation implementation.
Broader Core Auth tests would be required if D8-A is approved; Shipping probes do
not prove invalid-schema handling, wrong-target rejection or explicit provisioning.
The preceding candidate's build/compile/layout/live-role/full-chain gates retain
their documented scope and failures. No new broad gate or human acceptance is
claimed by these test-only additions. Full4be.1-4be.6 and Release1 remain active.
