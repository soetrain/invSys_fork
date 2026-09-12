# Plan 022 Slice 4be.1 Shipping activity D13 entry

Last verified: 2026-09-12. Runtime source remains `01891bb`, with validation
checkpoint `6c9d2a1`. This is test-first work toward D18 comprehensive activity,
not Shipping implementation or Slice4be acceptance.

The isolated test uses `deploy/validation-receiving-worksheet-rebuild`, the
existing generated-warehouse fixture and a real SHIP_POST authorization row.
Unsaved facades call the public Shipping launcher and its actual form handlers;
they never bypass authentication, replace a business owner or fabricate activity.
Only report presentation is intercepted. An owned-process observer handles the
existing Box Save single-OK information notice; it cannot accept multi-button
confirmations and discards all dialog text. No operational or NAS workbook is used.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-receiving-worksheet-rebuild -Phase RED `
  -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity
```

The first attempt ends **68 PASS / 1 harness failure** before Add: no shippable
box is available. The launcher/captured-form check passes. Source confirms that
`ShipmentsFormLoadShippables` reads the saved-box projection; seeded inventory
alone does not establish that fixture. This is not meaningful behavioral RED.
The revised fixture creates a box using the real Box Designer New/Add/Save
handlers and makes stock through Box Maker, then explicitly refreshes Shipping.
Its first assisted run passes105 checks and fails32 missing-activity assertions.
The self-contained run repeats **105 PASS / 32 expected FAIL** with no harness
exception; no manual notice dismissal is needed.

Target actions are Add, Update Row, Send Hold, Return, Remove, a second distinct
Add, To Shipments and Shipments Sent. For each, first assert independent staging
quantity/exact System_Key, captured workbook, unknown headers and unrelated-book
preservation. Missing attempt/result, correlation, trusted owner/context and
sanitized activity are the expected RED. The existing foundation checks remain
part of the run, and Config bytes must remain unchanged by Shipping commands.

The expanded source-evidence run is **125 PASS / 40 expected FAIL**, with no
harness exception, no non-activity failures and no duplicate check identities.
It preserves all earlier checks and adds read-only canonical log inspection per
action, exact new applied-event identities, unknown values through Update/Stage,
and independent Domain proof: one SHIP application of quantity2 under the exact
box key, with BOX_BUILD quantity10 and final balance8. Activity comparisons
require those independently observed event references; staging clearance alone
does not establish application. All40 failures are five missing-activity checks
for each of eight actions (including the second distinct Add).

This normal applied sequence does not cover pending/uncertain submission,
rejection/context/policy/failure cases or every source-reference edge. Add those
protecting cases and precise catalog/owner outcomes under D18 before runtime implementation.
Source mapping is in [Shipping coverage](plan022_slice4be_shipping_coverage.md).
Boxing fixture construction does not claim comprehensive Boxing activity proof.

Ignored runtime evidence: `reports/runtime/slice4be-shipping-activity/initial-red.log`,
`boxed-fixture-red.log`, `self-contained-red.log`, `source-evidence-red.log`
and separate per-run JSON directories:

- `eedf0425afcf46bba3b73f306c575031/red.json`:68/1 fixture failure.
- `90c82cdd179a462aad9cb4567451d988/red.json`:105/32 assisted fixture proof.
- `39cc76a1aa074f7384240688ac9ff069/red.json`:105/32 self-contained RED.
- `e107f5a936da4385b36eb4ffda7ce28d/red.json`:125/40 expanded RED.

All137 prior check identities and all105 prior GREEN checks are retained in the
expanded run. Static maintenance and JSON contracts pass, with175 components,
5503 procedures,8 literal Application.Run targets,45 unresolved dynamic calls,
1097 scanner/1099 reviewed candidates and all28 module-size limits unchanged.
All30 preserved package hashes and four runtime source hashes match after the
expanded run; Excel is closed. Reports contain
fixed check identities and booleans; raw values and credentials are not committed.
