# Slice4be Production Check In correctness prerequisites

Last verified:2026-10-01 UTC. Check In now requires a selected reusable Process,
preserves exact worksheet identity and custom columns, suppresses loading/nested
entry, rejects stale context and aligns worksheet values beneath their headings.
Latest unpromoted candidate is `deploy/validation-production-check-in-correctness-03`:
chain32/32, live roles48/48, Create Warehouse15/15, smoke86/86 and six-page geometry
pass. Its only change from correctness02 is a corrected two-batch test fixture.
Correctness02 owns focused103/103 and worksheet257/257 evidence; correctness01 owns
the503-check Run regression445 PASS/58 known Scale failures. Keep those candidate
distinctions explicit. These results do not establish Check In observations,
full populated-layout, native reusable, human or Release1 acceptance. Catalog24
and63/68 constructed-button wiring are unchanged.

Architecture v4.11 D18, "Check In correctness prerequisites," clarifies existing
D14 exact identity/header-extension, D15 selected-Process execution and D18
loading/nested-entry requirements. Plan022, controls1.409 and coverage1.149 are
synchronized. No new activity catalog entry, permission, authority or alternative
contract is established by these corrections.

## Expanded correctness and display evidence

Before runtime edits, the64-check RED below expands to85 checks (69 PASS/16 FAIL),
then94 (73 PASS/21 FAIL), retaining every prior ordered identity and PASS. Added
cases protect raw guard restoration, real nested invocation, missing/unresolved
keys, missing required Check headers and changed target/session/signed-out context.
Controllers `0e8ab1433e9746b5b4edd3566d60590e` and
`6c6dff5ef9b6401280457f12bd8c7809` retain results
`0324c676ad054ea89ba1396dda57bbef/red.json` and
`f424cf559de74a53bf1caf4f85623e1a/red.json` under the same baseline result root.

Correctness01 passes94/94, result
`5ed03cd7d37b4c28aec38ed0f8540b8c/green.json`, controller
`a683ea08c1ae4bdca1719b4768649c23`,15:26:34.6970334--15:28:55.3359919 UTC.
Typed guarded entry suppresses loading/nested calls and rejects stale context;
one selected reusable Process is required before note freezing. Both worksheet
identity reads require the exact selected key, and Check clearing validates its
six managed headers before clearing only those managed bodies. Unknown values,
formulas and positions survive. This does not yet protect all current-capability
and post-yield boundaries or register Check In observations.

Three principal captures are individually reviewed and hash-pinned in that
result's `capture-review.json`. They show selected-Process success and blank-Process
guidance, but expose worksheet fields beneath the wrong nine-column headings.
Architecture D18, Plan022 and the controls catalog clarify the existing projection:
place the six managed values beneath their matching headings, leaving the three
unavailable Type/Process/source context fields blank.

The103-check test on unchanged correctness01 produces101 PASS/2 FAIL, retaining
all94 ordered passes. Only `CheckInBaseline.Worksheet.ValuesUnderMatchingHeadings`
and `CheckInBaseline.LocalProjection.DisplayColumns` fail. A disposable local
InventoryManagement projection with two real same-SKU entities verifies exact
selection and unchanged inventory values/formulas independently of the Domain
fallback read. Reusable identity remains under its correct heading. Controller
`34e0aaf710f04af499f0e99354c57fbb`, result
`f55f85c44aa24b25a0357d83cff8dea8/red.json`,15:45:24.0870381--15:47:41.5776980 UTC.
All expanded focused gates retain42 shared checks, five instrumented compiles,
frozen package hashes, saved authority/operator bytes, restored settings, normal
unassisted closure and delayed zero Excel Application1000/1001/1002 failures.

Correctness02 changes only the worksheet display mapping in `frmProduction`
relative to correctness01. Focused GREEN is103/103 in exact prior order, controller
`7137d1c7ab2242e78ef8868e3254fee4`, result
`533bdf50b4564f2f953e881ddb890b91/green.json`,15:49:44.9882812--15:52:07.8872967 UTC.
Its three principal images are individually reviewed and hash-pinned in
`capture-review.json`: selected-Process success, missing-Process guidance and
worksheet fields under matching headings with unavailable context blank. Long
synthetic values remain clipped; this is not full layout or human acceptance.

Build receipts `check-in-correctness-build-01` and `check-in-correctness-build-02`
record cold startup/five compiles, preserved settings/prior packages, normal
closure and delayed zero audit. Correctness01 changes only `frmProduction` among
282 prior compiled components and adds `modProductionCheckInActions`; correctness02
changes only that form among283, retaining all282 other components.

Correctness02 shared worksheet regression passes257/257 in exact prior order,
controller `67253178192f4ebfb10759bfd176bf01`, result
`slice4be-production-run-worksheet-owner/f1468216787e435da2736c5d8b4a29cd/green.json`,
15:52:32.3706007--15:57:28.2804540 UTC. Clear/Refresh owner and interruption cases,
five compiles, canonical package pins, authority/operator/older-record preservation,
restored settings, normal unassisted cleanup and delayed zero audit all pass.

Correctness01 also preserves all503 prior Run checks in order with identical
results (445 PASS/58 known Scale failures),33 owner cases,63 stale-context checks
and seven actual closed-boundary receipt sets. Controller
`571625211a804bfc9f16326b3c6439b6`, result
`slice4be-production-run-local/c72cfd7d24274b00a17cddd922cd71bc/green.json`,
15:33:31.0226134--15:42:19.3132183 UTC. Five compiles, preservation, normal
unassisted closure and delayed zero audit pass. This is not full reusable171.

### Full-chain fixture correction

Correctness02 full-chain controller
`check-in-correctness02-regression/chain-c746bf2398834b2c977c50854a10e50c`,
15:57:55.0860611--16:03:13.9920015 UTC, returns27 PASS/5 FAIL/32 chain checks,
47 PASS/1 FAIL/48 live-role checks and15/15 Create Warehouse checks. The two-batch
Production form test fails at its first Check In; dependent batch/final-balance
assertions fail. Settings, packages and tracked reports are restored, Excel is
closed and the delayed Application audit is clear. Original reports and
`failed-fixture-verification.json` remain retained.

Source diagnosis: `PrepareRunChoiceForActionTest` explicitly supplies an empty
identity in `values(1, 4)`. `HydrateRunInventoryDisplay` receives that identity
ByVal and only fills descriptive/availability fields; it cannot select an entity.
The old test relied on Check In's prohibited SKU-based identity substitution.
This is an invalid fixture, not meaningful product RED. Correctness03 changes
only this packaged test helper: it selects one real entity returned by the Domain
for the fixture SKU/location, refuses missing/ambiguous selection, and supplies
its unchanged exact key before invoking the same Apply/Check In/Complete handlers.
No runtime identity fallback is restored and no test assertion is removed.

Correctness03 then passes32/32 full-chain,48/48 live-role and15/15 Create Warehouse
checks in exact prior order. Controller
`check-in-correctness03-regression/chain-1ace6a7bebb44b538ce6ef809e8e8b37`,
16:05:28.8726134--16:10:45.8008165 UTC, preserves settings, package hashes and all
three tracked reports, closes Excel and passes the delayed zero Application audit.
Its build03 cold startup/five compiles and source comparison retain282 other
components; only the form's fixture helper changes relative to correctness02.
This ordered chain is not the separate full reusable171/replay37 acceptance gate.

Correctness03 layout controller
`check-in-correctness03-regression/layout-1f024b01e894419fb99c0925879d0048`,
16:11:07.9450983--16:11:24.8216134 UTC, passes all18 requested size/page pairs,
six maximized-page checks and prior native transitions/representative geometry.
Minimum/default requests clamp to the same size; two distinct actual sizes are
proved. All six Run List/Settings screenshots are individually reviewed and
hash-pinned in `capture-review.json`. They show empty-form geometry, not populated
long-value acceptance. Settings/packages, normal closure and delayed zero audit pass.

Correctness03 packaged smoke passes86/86 in exact prior order, controller
`check-in-correctness03-regression/smoke-bf9d5f567bb84f2f9517c5295a248bd3`,
16:12:04.0726251--16:12:25.9265694 UTC. Both Initial/Final native shutdown receipts
under `packaged-smoke-closure/11ebb44ad8a5495bab878403e179e9a7` show unassisted exits
without termination requests. Settings/packages/tracked report preservation and
delayed zero audit pass. Focused103 and worksheet257 evidence remain attributed
to correctness02; correctness03 changes only the separate two-batch fixture.

### Excluded native failures

- First correctness01 focused attempt, controller
  `1a52265b43d94ff093b3a1f4f5201346`, result
  `e5dbe338c03c4c8d862568ff5a0e7c32/green.json`, fails during Admin bootstrap before
  Check In executes: RPC800706BE, Excel/ntdll exception `c0000028` at15:25:49.5083187 UTC.
- Ordinary run-only67 controller
  `check-in-correctness-regression/reusable-e0bd5283038447039d2cdad50482ac9d`
  fails in `mProduction.RunReusableProductionRunActionContractTest`, RPC800706BE,
  ntdll `c00000ff` at15:30:22.5091887 UTC. No67-check workflow result is available.
- One frozen allocate01 comparison,
  `check-in-before-regression/reusable-8f70d909ca7d4c2da4e1d08126cee0a1`, fails
  earlier in `mProduction.RunProductionBatchScaleContractTest`, RPC800706BE,
  ntdll `c0000028` at15:32:17.1213642 UTC. Different entry/signature does not prove
  the same cause or qualify the newer candidate.

Each retains an `excluded-native-verification.json`, restored settings, unchanged
package pins and closed Excel. Workflow/closure failures are retained, not recast
as meaningful RED, GREEN or desktop error5. Native root cause remains unresolved;
do not blindly repeat the full/run-only/replay gates.

## Focused packaged RED

Run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beProductionRunLocal.ps1 -DeployRoot deploy/validation-production-run-allocate-01 -Phase RED -CheckInBaselineOnly
```

`Slice4beProductionCheckInBaseline.ps1` installs unsaved fixture adapters into
the packaged project and invokes the unchanged `mBtnManagerCheckIn_Click` handler.
Admin Seed and the receiving owner/processor create two distinct real entities
with the same SKU/location. Process and Recipe definitions are saved/released
through their owners. The reusable cases select the actual Process or the native
blank Process entry. Worksheet staging explicitly selects the entity other than
the first currently returned by the inventory bridge, with a complete allocation,
its corresponding item/process/base-quantity lookup entries, shuffled managed
headers and unknown value/formula columns in Inventory Check. A separate active
workbook protects the original captured-book setup.

Final result: **58 PASS /6 FAIL /64 unique checks**. All42 prior shared checks retain
their exact order and PASS status. The22 new checks produce16 PASS/6 meaningful
behavioral failures:

| Failing assertion | Governing rule / source finding |
|---|---|
| `Reusable.NoProcess.CheckedStateNoteFreezeAndNoCompletion` | D15 requires a selected Process. The form's empty-selection branch invokes `CheckInReusableRun`. |
| `Reusable.Selected.Loading.CheckedStateNoteFreezeAndNoCompletion` | D18 suppresses loading entry. The Click handler unconditionally invokes Check In. |
| `Reusable.Selected.Busy.CheckedStateNoteFreezeAndNoCompletion` | D18 suppresses nested entry; the same unconditional handler bypasses the busy guard. |
| `Worksheet.ExactSelectedKey` | D14 requires the selected exact identity. `BuildRunUsedPayloadJson` and `WriteProductionCheckRowsFromRunPalette` call `ResolveRunSystemKey`, which searches by SKU/name/location and returns the first match. |
| `Worksheet.CustomValue` | D14 preserves unknown values; the Check writer calls a helper that clears the entire table body. |
| `Worksheet.CustomFormula` | The same table-wide clear erases unknown formulas. Preserving headers alone is insufficient. |

Assertion names in the result have the prefix `CheckInBaseline.`. The worksheet
reaches `RefreshManager`, writes the expected used quantity and displays the
existing checked-in status. Thus the identity/custom-value failures occur after
the real write boundary, rather than at an unavailable fixture or early refusal.
Selected-Process success, insufficient allocation refusal, source quantities/keys,
palette values/formula, table headers, saved authority and operator bytes pass.
That original64-check baseline does not cover all Check In capability, binding,
yielding, missing-column, invalid-key, partial-write or observation paths; later
expanded gates above address only their explicitly listed cases.

Ignored runtime evidence:

- Controller `production-run-local-controller/40ded7fa78234b988a62af5fa7ed2ca4`;
  result `slice4be-production-check-in-baseline/3b7c859af4c14a01bcf15dc2531f06e4/red.json`.
- Interval15:09:19.4778926--15:11:10.8405968 UTC. Five instrumented package
  compiles pass. `check-in-boundary.json` contains fixed stage/Boolean evidence,
  not keys, cell values or operational payloads.
- `verification.json` and `verify-check-in-baseline-01.ps1` verify all64 unique
  identities, the exact six failures, prior42 order/PASS, canonical package pins,
  restored settings, normal unassisted closure and zero Excel Application
  1000/1001/1002 failures through the delayed12-second audit window.

## Excluded fixture attempts

The first controller `a06fe112e84545388df913861c213323`, result
`db1c0b80d1a84572941794aab997b077`, expected a one-row Process selector. The actual
selector has a blank entry plus the Process. Its prerequisite failure is not
product RED; the corrected adapter selects the real Process at index1.

Controllers `eca172c3bbac417893681e04cd065b5a` and
`0bcd35c62e1842408d2776c76b662e6d`, results
`52603f2a34c34fb282affa389e53a1df` and `347aed043503431ca7eefe663fdcc05b`, retained
the real reusable failures but did not reach the worksheet write. Changing the
fixture's selected key without updating its item-code lookup caused an early
`BuildPayload` refusal. Their apparent worksheet preservation is not accepted
evidence. The final fixture updates its item/process/base-quantity maps and
explicitly validates those prerequisites. No runtime workaround was introduced.
All three attempts closed normally, restored settings/preserved packages and
passed delayed zero Excel audits. Their original reports remain retained.

## Static and remaining work

`reports/runtime/check-in-correctness-static-03` records290 components,6174 procedures,
135482 lines: feature growth1 component/3 procedures/48 lines relative to the
baseline. Nine literal/45 unresolved dynamic calls and190 duplicate candidates
remain unchanged; all28 oversized caps remain non-growing. Three JSON schemas
and369 tooling PowerShell parses pass.

Complete protecting regressions and current-permission/post-yield guarding before
accepting Check In observations. Define and test actual owner outcomes separately;
successful local Check In does not
submit or apply inventory events. Independent Action Paths, populated layout,
live roles, full Release1/reusable chains and human/NAS acceptance remain open.
RUN-SCALE-01 and RUN-UI-01 remain unapproved. Desktop error5 still requires the
user-authorized stop, safe cleanup, timestamp record and goal pause.
