# Slice 4be.1 Shipping owner-reference RED

Last verified: 2026-09-12 (run timestamps are UTC on September 13).
Authority: Architecture v4.11 D18 exact source references and D13 packaged actions;
Plan 022 Slice 4be.1. Runtime remains `22e14b6`; this checkpoint changes only tests
and evidence. Shipping activity implementation and full Release 1 acceptance remain open.

## Result and protecting scope

The complete packaged route finishes **650 checks: 597 PASS, 53 FAIL**. It retains
all 605 preceding full-run check identities and all 552 preceding GREENs, adds
45 passing checks, and has no duplicate identities. All 66 submission checks from
the separate focused run now pass within this full route. Of the additions, 19
are the previously focused persistence/path checks, 25 calibrate owner references,
and one verifies fixture bootstrap roots.

The remaining failures are unchanged in kind: 46 missing Shipping activity checks
and seven missing-Auth-file recreation findings. The latter concern the explicitly
pending D8-A proposal; they do not authorize changing Auth or establish an existing
D8 breach. This is meaningful pre-implementation RED for Shipping activity, with
passing source-reference calibration, not an overall GREEN or completed slice.

The test invokes the actual Add, Update Row, Send Hold, Return, Remove, To Shipments
and Shipments Sent form handlers through the packaged launcher and captured form.
An unsaved observer records exact IDs only at actual positive returns from
`QueueShippingPayloadEventServerFirst`. Each action resets the observer and verifies
valid unique identities, agreement with queue-entry counts, no uncertain returns,
and the fixture-specific submission cardinality. Runtime IDs are neither generated
nor normalized by the observer. Unknown acknowledgment behavior retains its separate
66-check submission proof; positive-only observation is not a general result API.

The historical `.Activity.ExactAppliedSourceReferences` assertion name is retained
for comparison continuity. Its expected values now mean exact owner-submitted IDs,
including pending work. Newly applied inventory-log IDs remain separate evidence.

| Actual action | Submitted by this action | Newly applied | Own pending | Earlier actions applied |
|---|---:|---:|---:|---:|
| Add | 1 | 0 | 1 | 0 |
| Update Row | 0 | 0 | 0 | 0 |
| Send Hold | 0 | 0 | 0 | 0 |
| Return | 0 | 0 | 0 | 0 |
| Remove | 1 | 0 | 1 | 0 |
| Add again | 1 | 0 | 1 | 0 |
| To Shipments | 0 | 0 | 0 | 0 |
| Shipments Sent | 1 | 4 | 0 | 3 |

This fixture uses a delta-only Update and already-reserved Stage. Their zero-source
results do not cover other branches. Send applies three earlier submissions as well
as its own source; attributing all four to Send would violate D18. Submitted means
queue acceptance, never Domain application. Exact shipment application, captured
workbook binding, immutable keys, unknown columns and unrelated-workbook checks pass.

## Fixture and native evidence

The first corrected run stops at 67 PASS / 1 harness exception before new reference
checks execute. Excel reports a seven-argument Run failure. Source call order and
arity identify the six-value `BootstrapWarehouseLocalAdmin` fixture call, before
the zero-value Shipping launcher; this is an inference, not an internal crash stack.
Windows Application event 1000 records `ntdll.dll` / `c0000028` at
`2026-09-13T04:05:12.4563024Z`. This failure is not D13 behavioral RED.

The completed run adds an unsaved read-only bootstrap-root getter and explicitly
sets/verifies fixture roots before generation. Both expected-root flags are True
before Operations instrumentation, after it, and immediately before generation.
These observations do not support the hypothesis that those edits cleared the
roots. The passing run also changes instrumentation and call sequence; it cannot
establish a root cause or a native runtime repair. No Excel Application event 1000
was observed during its recorded window; historical native failures remain open.

The common Run diagnostic now names a failed source macro and argument count only,
without emitting argument values, adding a retry, or suppressing the exception.
Disposable fixtures remain under the existing protected cleanup. No operational
workbook, accepted deployment, NAS target or runtime source is changed.

## Reproduction and evidence

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-capability-guard -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity
```

The command exits 1 because the expected activity and pending-decision failures
remain. Raw files below are ignored under
`reports/runtime/slice4be-shipping-activity/`; only this sanitized account is committed.

- Initial failure: `owner-reference-red.log`,
  `2e558e98a3b64e5aad172b005dcceda4/red.json`,
  `owner-reference-native-faults.json`.
- Complete run: `owner-reference-explicit-roots-red.log`,
  `63b82b33064940ecb656129356407e2d/red.json`, and that directory's
  `shipping-source-counts.json` and `shipping-fixture-root-observations.json`.
- Comparison: `owner-reference-comparison.json`; native observation:
  `owner-reference-explicit-roots-native-faults.json`.
- Package/source preservation: `owner-reference-final-pin-verification.json`;
  static regeneration: `owner-reference-static.log`.

Excel is closed and all 50 preserved package pins and eight current source pins
match. No new build/compile, layout, live-role, full Release 1 chain, visible or
human acceptance is claimed for this test-only change. Their earlier candidate
results and unresolved native limitations retain their original scope.

Static evidence was regenerated and all three JSON contracts pass. All 28 previous
module limits are preserved. Counts remain 177 components, 5,511 procedures, eight
literal Application.Run targets, 45 unresolved dynamic calls and 195 duplicate
bodies; scanner/reviewed candidate counts remain 1,097/1,099. Both edited PowerShell
scripts parse, 52 local Markdown targets resolve, and diff/status review preserves
the unrelated user changes in handoff 067 and untracked critique 023.

## Next implementation boundary

Register the seven Shipping controls and owner-reported outcomes/source eligibility
under D18 in Architecture v4.11, Plan 022 and the controls catalog before runtime
observation changes. The subsequent [catalog checkpoint](plan022_slice4be_shipping_catalog_results.md)
implements Core definitions/reference eligibility after normative clarification;
actual handler observation is still pending. Preserve the full 650-check route and strengthen it for exact
catalog outcomes and partial/mixed submissions. Update/Remove partial failure,
Send pending processing, policy/tracking failures, and remaining Operations/Admin
coverage remain required. D8-A approval is independent and still pending.
