# Slice4be Production Check In correctness prerequisites

Last verified:2026-10-01 UTC. This is pre-implementation RED, not observation,
workflow, layout or Release1 acceptance. Runtime source remains `73fd1d25` and the
unpromoted frozen candidate remains `deploy/validation-production-run-allocate-01`.
Catalog24 and63/68 constructed-button wiring are unchanged.

Architecture v4.11 D18, "Check In correctness prerequisites," clarifies existing
D14 exact identity/header-extension, D15 selected-Process execution and D18
loading/nested-entry requirements. Plan022, controls1.408 and coverage1.148 are
synchronized. No new activity catalog entry, permission, authority or alternative
contract is established by this baseline. No runtime behavior has changed.

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
This does not yet test all Check In capability, binding, yielding, missing-column,
invalid-key, partial-write or observation paths.

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

`reports/runtime/check-in-baseline-static-02` retains289 components,6171 procedures,
135434 lines,9 literal/45 unresolved dynamic calls,190 duplicate candidates and
28 non-growing oversized caps. Three JSON schemas and369 tooling PowerShell
parses pass. Runtime metrics are unchanged from `run-allocate-path-static-02`.

Correct the proven violations under the existing D14/D15/D18 rules with focused
GREEN and protecting regressions before accepting Check In observations. Define
and test its actual owner outcomes separately; successful local Check In does not
submit or apply inventory events. Independent Action Paths, populated layout,
live roles, full Release1/reusable chains and human/NAS acceptance remain open.
RUN-SCALE-01 and RUN-UI-01 remain unapproved. Desktop error5 still requires the
user-authorized stop, safe cleanup, timestamp record and goal pause.
