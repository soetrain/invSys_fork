# Slice 4be Process worksheet picker header alignment

Last verified: 2026-09-30 UTC. Focused115 GREEN, build/compile and static pass;
regression gates remain pending. No new tracking registration or release acceptance.

## Governing contract and test

Architecture v4.11 D4/D14/D15 requires the shared Core picker to update the exact
selected numbered Process item/SKU pair under normalized managed headers.
Documentation d099b48 records the clarification before runtime changes. Existing
ownership, record types, allocation and explicit save behavior remain unchanged.

The packaged test invokes the real Production process-item picker and its actual
Core CommitSelection handler. Unsaved adapters prepare Admin-generated disposable
fixtures and read expected values from the actual displayed selection. Nine cases
cover canonical pairs1/2/5, lowercase pair2, space-padded pair2, normalized pairs1/2/5,
and normalized OUTPUT pair1. Other cells, custom values/formulas, header sequence,
the active decoy workbook, saved workbook bytes and canonical authority bytes are
protected. Fixture values remain in memory; runtime captures remain ignored.

## RED

Frozen baseline: `deploy/validation-process-worksheet-headers`.
Controller `reports/runtime/process-worksheet-picker-controller/2d220813d9fd491c8ce2fa795d740974`;
result `reports/runtime/slice4be-process-worksheet-picker/32758714dac647c09667ae70cc41abf4/red.json`.
2026-09-30 18:01:08.6416716–18:02:39.6465393 UTC: **110 PASS,5 FAIL/115**.
Five instrumented compiles and all42 prior shared GREEN checks pass. All five
noncanonical INPUT cases change the item label but leave its SKU unchanged.
Canonical and OUTPUT cases and all preservation assertions pass. Normal
unassisted shutdown, package/settings preservation and delayed Excel Application
audit pass. No desktop error5 observed.

Initial controller d7782185d7a7427b9b5fe2554b9ed731/result76f18831811a4beeb3da1923145b7dc6
reported111 PASS/4 FAIL. Its direct case-only ListColumn rename did not establish
lowercase coverage. The final fixture renames through a temporary header and
verifies the exact resulting header with binary comparison before executing the
same handler. This is fixture calibration, not a runtime workaround. The final
five failures above establish the protecting RED before Core edits.

The captured picker shows a populated actual item selection before commit; the
item/SKU mismatch is proven by assertions, not a post-commit screenshot. No claim
of human acceptance or comprehensive Slice4be completion.

## Candidate and focused GREEN

Only Core `cDynItemSearch.ProcessAlternativePairNumber` changes: trim the actual
header and compare its prefix with vbTextCompare. No module/procedure/line growth,
new bridge, identity allocation or ownership/save change.
Candidate `deploy/validation-process-worksheet-picker`; build record
`reports/runtime/process-worksheet-picker-build`,18:03:32.3325656–18:04:10.7948007 UTC.
Five compiles and Operations cold start pass. Compiled-source comparison finds
only cDynItemSearch changed among265 components;264 are exactly unchanged.
Frozen worksheet-header candidate/settings preserved, normal cleanup and delayed
Excel Application audit pass.

GREEN controller `reports/runtime/process-worksheet-picker-controller/fdc4c379abe54795a3a4cf8bffbca9b3`;
result `reports/runtime/slice4be-process-worksheet-picker/467da303e9e1453086791d9f553aa915/green.json`.
18:04:30.2768494–18:06:07.7087906 UTC: **115/115**, exact RED check sequence and
all42 prior shared GREEN checks. Five instrumented compiles, preservation, normal
unassisted cleanup and delayed Excel Application audit pass. The actual populated
picker capture was directly reviewed; it is a pre-commit surface, not visual proof
of the worksheet outcome. Automated assertions prove the latter.

Static `reports/runtime/process-worksheet-picker-static`:272 components,6116
procedures,134480 lines,9 literal/45 unresolved Application.Run,190 duplicate
body candidates and28 non-growing oversized caps. All baseline metrics unchanged;
three schemas and340 PowerShell parses pass, no exception added.

## Regression status

First full reusable controller
`reports/runtime/process-worksheet-picker-regression/reusablefull-320f6c4899ab4ba290a20d9fbb6b9de9`,
18:06:41.2076434–18:07:07.5249073 UTC, is excluded: HARNESS/HARNESS_CLEANUP RED,
RPC HRESULT0x800706BE at mProduction.RunProductionBatchScaleContractTest.
Application1000 at18:07:01.9175659 UTC records ntdll.dll/c0000028/offset12d2f;
Application1001 follows. Normal workbook closure was not completed. Settings and
packages were restored/preserved, no process termination requested, Excel absent
afterward. This matches a previously observed native failure signature, but does
not establish its cause or a repair. Desktop probes stayed healthy, zero error5.
One unchanged isolated rerun is pending; no successful regression claim yet.
