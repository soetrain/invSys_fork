# Slice 4be.3 full-chain package cleanup

Last verified 2026-09-13. This is a D13 validation-tooling correction under D12's
package dependency boundaries. It changes no XLAM, operator action, authority,
permission or publication contract. Release 1 acceptance remains open.

The reconciliation and five-package inspection phases acquired Core first but
attempted to close it before its dependent packages. An isolated real-Excel trace
recorded a first-package Close exception in each phase while all13 phase checks
passed. Those swallowed exceptions did not reproduce the earlier native fault;
they do not establish its cause.

`Test-Release1PackageCloseOrder.ps1` executes the actual two finally blocks with
disposable dependency-aware workbook substitutes. The test preserves the real
acquisition order and rejects a dependency close while consumers remain open.
It checks ordinary-workbook ordering, complete explicit closure before Quit,
no rejected dependency close, one attempt per owned book, no implicit save, and
clearing the Core override before package closure.

- Before implementation: **5 PASS / 8 FAIL**, no setup or parse failure.
- After implementation: **13/13 GREEN**.
- Existing ordered-child exit/report decisions: **4/4 GREEN**.

The two phases now close a reversed copy of their acquired-workbook list:
ordinary workbooks first, then leaf packages, then dependencies. They retain
Close(False), existing scope cleanup and the original acquisition list.
The change is limited to `tools/validate_release1_full_chain.ps1`. The real-Excel
after-trace passes **13/13**, with **zero** package-close exceptions in both
phases and Excel closed. The fresh full-chain attempt fails earlier during live
projection rebuild: top-level4/1, live32/1. Neither corrected finally block is
reached in that failed attempt. The correction therefore removes the observed
cleanup exceptions but does not resolve or explain the live native fault.

The unchanged-package full-chain retry subsequently finishes31/31 with live48/48.
A native Excel fault follows the report; its verified empty recovery child is
closed without saving and all three loaded add-in files remain unchanged. This
does not invalidate the observed dependency-order correction, but clean native
shutdown and the full Release 1 gate remain unproven. Both deleted-handle and
unchanged-handle live diagnostics pass48/48; no projection-fixture patch follows.

Ignored evidence under `reports/runtime/`:

- `events-chain-close-order-red.json`, `events-chain-close-order-green.json`.
- `events-chain-final-phases-summary.json`, `events-chain-process-trace.jsonl`:
  original real-Excel13/13 with two first-package Close exceptions.
- `events-publication-cleanup-static.log`: static regeneration for this tooling
  change completed;206 components,5700 procedures,127035 lines,9 literal/45
  unresolved calls,189 duplicate groups and28 size ratchets remain unchanged.
- `events-chain-final-phases-after-summary.json`,
  `events-chain-process-trace-after.jsonl`, `events-chain-close-order-native-green.json`:
  real-Excel13/13, no package-close exceptions and Excel closed.
- `events-publication-reviewed-cleanup-chain.md`,
  `events-publication-reviewed-cleanup-live-failed.json`: fresh4/1 and32/1 failure
  before the corrected phases. Owned recovery closes without saving and preserves
  all three loaded add-in files (`events-publication-reviewed-cleanup-recovery.json`).

The full Receiving retry passes854/854 against the unchanged reviewed XLAMs,
retaining every earlier check identity/GREEN without duplicates. Its earlier690/1
readback failure and focused266/266 diagnostic remain recorded; no native-input or
runtime patch is inferred from those results.
