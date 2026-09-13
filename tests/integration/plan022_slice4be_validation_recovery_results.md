# Slice 4be.1 candidate validation recovery

Verified 2026-09-12 local / 2026-09-13 UTC against the isolated Shipping catalog-8
candidate from code `469703f`. This checkpoint changes test tooling only, under
Architecture v4.11 D13/D18. Runtime and accepted deployment are unchanged.

The complete instrumented Shipping route executes **863 checks: 810 PASS / 53
FAIL**. Identity comparison preserves all 650 preceding checks and every preceding
GREEN; all 213 added catalog/reference/policy checks pass. No duplicate identities
exist. Failures remain 46 missing actual-handler activity assertions and seven
missing-Auth recreation findings for pending D8-A. D8-A is not approved here.

## Viewer setup proof

The harness now stops after rejected fixture sign-in, before its public Viewer
action block. It emits an allowlisted status and credential-match/entry booleans,
never the raw sign-in response or credential. `-RejectFixtureSignInForTest` calls
actual Core sign-in with an invalid disposable-fixture credential; stored
credentials remain unchanged.

| Run | Result | Signed in | Status | Credential matches fixture | Entry reached |
|---|---|---|---|---|---|
| Initial normal diagnostic | PASS, exit 0 | True | OK | True | Marker not yet added |
| Deliberate rejection | Expected setup FAIL, exit 1 | False | AUTH_STATUS_4 | True | False |
| Final normal calibration | PASS, exit 0 | True | OK | True | True |

Entry marks the harness block, not an independent callback observer. The normal
gate retains Viewer reuse, inventory filtering, table export, receipt/Production/
Shipping Events, hidden internal reservations, publication refresh, remembered
date preferences through close/reopen and snapshot-byte preservation. Rejection
is harness calibration, not product D13 RED. The earlier unexplained sign-in
failure remains recorded. The final numeric-status restriction uses enum 1-9;
packaged calibration exercises OK and 4, not every malformed response.

## Shipping bootstrap experiment

Moving unsaved Operations probe installation before fixture/form use did not
resolve the fault: **281 PASS / 1 harness failure** at six-value Admin bootstrap.
That experiment was reverted. The retained `-TraceBootstrapForTest` diagnostic
injects unsaved, fixed-label calls into actual Core bootstrap procedures. Only
phase names are written, without values, identities, credentials or payloads.
Exact anchor-count checks reject unexpected source shape. Initial trace setup
stopped at 1 PASS / 1 harness failure because VBE normalized identifier case;
case-insensitive matching corrected this diagnostic setup.

The calibrated trace completes all 863 checks, including the real Shipping
launcher and handlers. Five bootstraps reach BOOTSTRAP_SUCCESS. The fault did not
reproduce under instrumentation. This supplies candidate functional evidence under
that diagnostic; it neither isolates a native cause nor proves uninstrumented
native acceptance.

## Native observations and remaining gates

Application1000 windows use log creation/completion times with three seconds of
margin. Both normal Viewer windows and the calibrated Shipping window contain no
observed Excel fault. Deliberate rejection has a subsequent Excel fault at
`05:17:49.6934855Z`, module unknown, exception `c0000409`; negative setup proof
therefore does not establish clean native shutdown. The reverted early-probe run
records `ntdll.dll/c0000028` at `05:22:37.3468527Z`, offset `0000000000012d2f`.
These temporal observations do not establish causality.

All 55 retained package pins, six unchanged source pins and three current Core
pins match after testing; Excel is closed. Static generation completes with
178 components, 5,514 procedures, 123,883 lines, eight literal dynamic-call targets,
45 unresolved calls, 195 duplicate bodies and 28 existing module limits. No XLAM
rebuild is needed for these unsaved diagnostic/test-only edits.

Full live-role/Release 1 chain, uninstrumented native stability, full Receiving
candidate regression, visible comparison and human acceptance remain open.
Shipping owner facts and actual activity remain required, followed by remaining
Operations/Admin, Settings, comprehensive Viewer, recording and guide work.

## Reproduction and evidence

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_inventory_viewer.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-catalog-eight
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_inventory_viewer.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-catalog-eight -RejectFixtureSignInForTest
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-catalog-eight -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity -TraceBootstrapForTest
```

Ignored raw evidence under `reports/runtime/slice4be-shipping-activity/`:

- `viewer-signin-{diagnostic,rejected,calibrated}.log` and corresponding results.
- `shipping-early-probes-green.log`, `shipping-bootstrap-phases.log` and
  `shipping-bootstrap-phases-calibrated.log` retain each experiment.
- `cec897841bed462fa46540c1ef6ec0ec/diagnostic-bootstrap-green.json` and
  `bootstrap-phases.log` retain the full calibrated route.
- `validation-recovery-comparison.json`, `validation-recovery-native-windows.json`,
  `validation-recovery-final-pin-verification.json` and
  `validation-recovery-final-static.log` retain comparison/validation scope.

See [catalog and prior gate evidence](plan022_slice4be_shipping_catalog_results.md).
The subsequent [exact-outcome RED](plan022_slice4be_shipping_exact_outcomes_results.md)
retains a later trace failure and proves the full890-check Shipping-first route.
Only sanitized evidence is committed; raw runtime reports remain ignored.
