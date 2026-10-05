# Slice 4be-A: Production Print Recall

## Refusal preservation, 2026-10-05

The actual Print Recall handler previously created or cleared RecallCodesPrint
before discovering no recall-coded output. Both the command and its diagnostic
could erase custom columns/formulas, another table and unrelated cells while
refusing preview. Under existing D14, report preparation now follows eligibility
validation. The private renderer returns the prepared worksheet by reference;
existing refusal text and successful report generation remain the contract.

Focused RED99/7 on print-binding01 retains all80 prior ordered checks. Seven
failures prove unwanted empty-report creation, and lost report/tables/cells for
blank recall codes and a missing recall column. Empty-output refusal preserves
the independently prepared report. Tests fingerprint table names, addresses,
headers, values and formulas, plus all used cells, without publishing their content.
An earlier trial92/9 reused the damaged report between cases; its final three
failures were carryover, so isolated RED99/7 is the governing evidence.

Run `tests/tooling/Test-Slice4beProductionRunLocal.ps1 -PrintBaselineOnly` with
`-Phase RED -DeployRoot deploy/validation-print-binding-01`, then
`-Phase GREEN -DeployRoot deploy/validation-print-preserve-01`.

Ignored receipts below are relative to `reports/runtime/`:

- Trial: `production-run-local-controller/75686ae8593c4ac2a6f89c56676617ef`,
  worker `slice4be-production-print-baseline/0039161b60764c0c8c228a8dc286b30f`.
- RED: `production-run-local-controller/2c7287ce7d5c4c898fe6d5cda0043b5d`,
  worker `slice4be-production-print-baseline/050cf1ef46064a56bc23012726e9ebed/red.json`,
  20:05:49.556–20:08:04.458 UTC. Five instrumented compiles, normal closure,
  restored settings, preserved package hashes and delayed zero Excel errors pass.
- Build: `print-preserve-build-01/verification.json`, 20:08:23.044–20:09:01.495 UTC.
  Five cold compiles; only mProduction changes, with302 other components unchanged.
  Forms are identical; entry-binding candidate layout18/five native checks remains
  applicable. Prior packages/settings are preserved; delayed zero Excel errors.
- GREEN: `production-run-local-controller/81202914b767410d902244a765b9ca92`,
  worker `slice4be-production-print-baseline/092cb0aa34e1482d8222a1a53fc56f2e/green.json`,
  20:09:15.890–20:11:39.265 UTC. 106/106, exact RED/GREEN identities, all80 prior
  checks retained; five compiles, normal closure, preservation and delayed zero
  Excel errors. Four new refusal captures are readable; the misleading existing
  completed prefix remains open and is not accepted as an owner outcome.
- Smoke: `print-preserve01-regression/smoke-39e81c62a83e427d908ce55accfab4d2`,
  20:11:57.773–20:12:24.763 UTC. All86 exact ordered checks, both unassisted exits,
  settings/package/tracked-report preservation and delayed zero Excel errors.
- Static: `print-preserve-static-01/ratchet-verification.json`: 310 components,
  6301 procedures, 137878 lines (one fewer); 9 literal/45 unresolved dynamic calls,
  190 duplicate candidates, 28 non-growing oversized caps, three schemas and
  420 PowerShell parses pass.
- Full chain: `print-preserve01-regression/chain-bc24fbf5fe894a4fb90384442ad2be52`,
  20:12:43.912–20:18:39.499 UTC. Chain32, live-role48 and warehouse15 all pass
  with exact prior ordered checks. The positive recall-report diagnostic produces
  its expected row; versioned Boxing, shipment, restart and reconciliation pass.
  Normal cleanup, settings/package/tracked-report preservation and delayed zero
  Excel errors pass. Earlier Boxing native failure does not reproduce; its cause
  is not established by this report-only correction.

Candidate: `deploy/validation-print-preserve-01`, not promoted. No error5 occurred.
Successful rebuild
preservation, preview/printing and recording integration are still open.

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

At code checkpoint `7b1d8b5c`, run
`tests/tooling/Test-Slice4beProductionRunLocal.ps1 -PrintBaselineOnly` with
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
The refusal-preservation candidate above now has fresh chain32/live48/warehouse15
evidence; the earlier intermittent Boxing failure's cause remains unproved.
Comprehensive A1/A2 and 4be-A completion remain open.
