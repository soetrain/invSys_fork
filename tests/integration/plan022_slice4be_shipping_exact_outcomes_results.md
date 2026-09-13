# Slice 4be.1 exact Shipping outcome RED

The protecting test now requires exact D18 Shipping evidence through actual
packaged handlers, before owner-fact or observation implementation. The unchanged
catalog-8 candidate comes from code `469703f`; no runtime source is changed by
this test checkpoint. Normative D18 and controls already define these outcomes.

Each of eight normal actions and invalid quantity adds three independent checks:
registered ControlId/SHIPPING_WORKFLOW/surface; REQUESTED EventCode/Info/Unknown
with empty references; and exact owner outcome/EventCode/severity/effect.
Expectations are literal contract facts, not values read from the implementation's
catalog or inferred from arbitrary status text. Existing independent checks still
prove local owner effects, submitted identities, acknowledgments, exact-key Domain
state, captured workbooks, unknown columns and unrelated-workbook preservation.

| Established fixture branch | Expected outcome | Severity | Effect |
|---|---|---|---|
| Add, Add again, Remove | PENDING | Notice | Unknown |
| Delta-only Update, Hold, Return, already-reserved Stage | STAGED | Info | Changed |
| Send processing/refresh finishes | CONFIRMED | Info | Unknown |
| Invalid quantity before owner/submission | REJECTED | Warning | Unchanged |

CONFIRMED still does not prove individual Domain application. Partial/mixed-source
failures, optional tracking failures, every remaining control family and broader
release/visible acceptance remain required. Seven missing-Auth recreation findings
remain pending D8-A; they are not converted to approved architecture by this test.

## Execution evidence

The first invocation with explicit bootstrap tracing stopped at281 PASS/1 harness
failure, before the new assertions. Windows records ntdll.dll/c0000028 at
`2026-09-13T05:49:45.5347332Z`, offset `0000000000012d2f`. This is not behavioral
RED and disproves treating the earlier successful trace as a stable native fix.

A new diagnostic `-ShippingBeforeSharedFormsForTest` runs the entire Shipping
group before the shared Settings/Production form exercises, then retains every
remaining shared/catalog check. It uses no bootstrap trace. The switch requires
full Shipping mode and rejects submission-only scope; output filenames identify
the altered order. No business behavior, authorization or package changes result.

The invocation completes **890 checks: 810 PASS / 80 FAIL, exit1**. All863 preceding
identities remain, with no duplicates or lost GREENs. The only additions are the27
exact contract assertions, all behavioral RED at the eight real normal handlers
and invalid quantity because activity records are absent. The remaining failures
retain their prior partition:46 missing activity assertions and seven pending
D8-A findings. Owner-state/submission evidence executes independently.

No Excel Application1000 fault is observed between this run's log creation and
completion, with three seconds of margin. The complete reordered route is proven
for this candidate, not a general native repair or full release acceptance. All55
package pins and nine current source pins match; Excel is closed. Three static
JSON schemas, all28 previous module limits, unchanged runtime maintenance metrics,
four PowerShell parsers and local Markdown links pass. No runtime source changed.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-catalog-eight -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity -ShippingBeforeSharedFormsForTest
```

Ignored evidence under `reports/runtime/slice4be-shipping-activity/` includes
`shipping-exact-outcomes-red.log`, its native-fault JSON,
`shipping-exact-outcomes-first-red.log`, and `exact-outcomes-final-static.log`.
The exact full result is
`90c5dafbb30948a0aa36c22046b6b8a1/diagnostic-shipping-first-red.json`;
`shipping-exact-outcomes-comparison.json`,
`shipping-exact-outcomes-first-native-faults.json` and
`exact-outcomes-final-pin-verification.json` retain the comparison and boundaries.
See [candidate validation recovery](plan022_slice4be_validation_recovery_results.md)
for the prior863-check result and preserved native limitations.

Next implement typed owner facts and actual Shipping observations under D18,
without deriving clean completion from a generic True return or borrowing IDs
from earlier actions applied during catch-up. No GREEN or slice completion is
claimed by establishing this RED.
