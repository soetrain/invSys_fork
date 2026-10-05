# Slice 4be-A: Production Print Recall

## Entry binding, 2026-10-05

Print Recall previously entered its report owner and diagnostic after a warehouse
or session change, sign-out, or loss of the captured Production sheet. The missing
sheet allowed the legacy resolver to search other open workbooks. Under existing
D18, the actual `frmProduction.mBtnManagerPrint_Click` now requires the original
live context and workbook's Production sheet before either report call. The typed
`modProductionRunBinding.RequireWorksheetContext` composes existing guards.

Packaged RED68/12 becomes GREEN80/80 with identical ordered checks and all68
previous passes retained. Each of four invalid cases proves zero owner entries,
zero report reads and the exact visible refusal. The valid case retains one owner
entry and both existing report reads with a decoy active; the observed workbook
is the captured workbook. Its missing output table produces the existing native
refusal before preview. This deliberately protects entry binding, not printing.
Captured custom values/formulas, decoy, saved operator bytes and the other generated
warehouse remain unchanged; no report is created. Five focused captures reviewed.

Run `tests/tooling/Test-Slice4beProductionRunLocal.ps1 -PrintBaselineOnly` with
`-Phase RED -DeployRoot deploy/validation-next-closed-01`, then
`-Phase GREEN -DeployRoot deploy/validation-print-binding-01`.

Ignored receipts, relative to `reports/runtime/`:

| Gate | Receipt | Result |
|---|---|---|
| RED | `production-run-local-controller/725a40197bc54fa98d48d07e4f1c8918`; `slice4be-production-print-baseline/8d7ade8fe0f846079dac647f78986ec7/red.json` | 68 PASS / 12 behavioral FAIL; 19:48:33.691–19:50:39.032 UTC |
| Build | `print-binding-build-01/verification.json` | Five cold compiles; 303 components, 301 unchanged; 19:51:47.044–19:52:25.516 UTC |
| GREEN | `production-run-local-controller/1471e10de51a48d8b9dcb6533e2723e4`; `slice4be-production-print-baseline/79ad954f71e449a888ad672e0ffa8441/green.json` | 80 PASS; 19:52:47.571–19:55:07.594 UTC |
| Layout | `print-binding01-regression/layout-c77344b88a924a30961115f7214efbfb/verification.json` | 18 page/size checks, five native-window checks, zero bounds/interactive-overlap failures |
| Smoke | `print-binding01-regression/smoke-9c23c9a207054cf3b8ec434d66e4a52e/verification.json` | 86 PASS, exact prior ordered checks; 20:00:04.117–20:00:32.435 UTC |
| Static | `print-binding-static-01/ratchet-verification.json` | 310 components, 6301 procedures, 137879 lines; 9 literal/45 unresolved dynamic calls, 190 duplicate candidates; 28 non-growing oversized caps; three schemas and 420 PowerShell parses pass |

Only compiled `frmProduction` and `modProductionRunBinding` change; form geometry
is unchanged. The added typed guard accounts for one procedure and six lines;
dynamic-call and duplicate counts do not grow. Focused runs compile all five
instrumented packages. Build/tests preserve settings and package hashes, close
Excel normally, and have zero Excel failures in their delayed Windows-event audits.
Smoke also restores tracked reports and proves both unassisted shutdown stages.
The candidate is not promoted. No desktop error5 occurred during these checks.

## Remaining acceptance

This checkpoint does not accept Print Recall as recorded or prove preview/physical
printing. Permission changes, yielding/closure paths, truthful owner outcomes and
observation integration remain open. The existing report builder clears the report
sheet and tables; required D14 unknown-column preservation needs protecting tests
and correction before acceptance. Its diagnostic rebuild and unconditional
"completed" prefix cannot establish successful preview or printing.

Existing unchanged-workflow evidence remains scoped to its recorded candidates.
The full-chain native Boxing failure is unresolved; no fresh full-chain or live-role
acceptance is claimed here. Comprehensive A1/A2 and 4be-A completion remain open.
