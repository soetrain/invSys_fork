# Slice 4be Process worksheet header preservation

Last verified: 2026-09-30 UTC. Operations worksheet maintenance/retrieval completes
its scoped gates: focused107, build/compile/static, reusable171, smoke86,
chain32/live48/Create15, layout and Close/query312. A separately discovered Core
picker normalization case remains to be tested before worksheet tracking.
No worksheet tracking registration or release acceptance.

## Governing contract

Architecture v4.11 D14 requires normalized managed-header access and preservation
of unknown columns. D15 explicitly applies that rule to retained Process tables,
including rejected Retrieve. Documentation02ed4da records the clarification before
runtime edits; existing metadata coordinates, identities, all-selected validation,
explicit save/import and confirmed selected-table removal remain unchanged.

The unsaved test adapters call the actual Send Process to Sheet and Retrieve
Selected Process handlers. Fixtures are Admin-generated and disposable. Tests
cover inserted value/formula columns, shifted ID, normalized headers, reordered
fields, and ordinary/shifted/normalized mixed-UOM import. Fixture values stay in
memory; persisted evidence contains named checks, counts and boolean facts.

## RED

Frozen baseline: `deploy/validation-production-close-observations`.
Final controller `reports/runtime/process-worksheet-headers-controller/54cc9f7d722547beba819ff962b67e2e`;
result `reports/runtime/slice4be-process-worksheet-headers/dc6644b62d5d49f6b8ef2a2fd7e07c05/red.json`.
2026-09-30 17:23:17.0932534–17:24:53.0785339 UTC: **97 PASS,10 FAIL/107**.
Five instrumented compiles, ordinary import, rejected-table byte preservation,
import Auth/Config byte preservation and exact inventory business-table values
pass. Packages/settings preserved, normal unassisted shutdown, zero delayed
Excel Application1000/1001/1002 events. Desktop probes remain healthy.

Expected failures: custom values/formulas and the custom column before ID are
overwritten; normalized Requirement ID/all headers and reordered Record Type
fail formula restoration; shifted/normalized imports fail and retain their tables.
These are behavior failures, not compilation or fixture failures.

Initial five-case RED: controller3b8d266737044cc3b3744e48117cbb06,
result7c7bedb352354cd08012e0c66774e3fe,79 PASS/4 FAIL/83,17:16:54.9639062–
17:18:19.2319755 UTC. It first proves the three overwrites and normalized-header
failure before runtime edits. The custom-column screenshot visibly shows generated
IDs replacing the fixture values. The normalized-header image is blank Excel
chrome and is excluded from visible evidence.

Expanded calibration362d14f7133c41a78569cc77f77c468d and diagnostic
d4e49a05c7b04eee81082e3b185e6911 each report95 PASS/11 FAIL/106. Their generic
post-import byte assertion is not a product requirement: successful explicit
Retrieve runs Core.RunBatch, which owns locks, inboxes and publication. The latter
trace proves Auth/Config bytes unchanged. The final test retains byte identity for
all generated authority files after rejected retrieval, and checks Auth/Config
bytes plus six exact inventory business tables after actual import. One diagnostic
CustomValue selection failed to reach the intended missing-name rejection; final
fixture selection explicitly activates/selects the captured table immediately
before invoking Retrieve. No runtime selection handler is replaced. Blank worksheet
captures are excluded; visible evidence remains pending the candidate capture.

## Candidate and remaining gates

Only Operations `modProductionProcessWorksheet` changes, with new typed helper
`modProcessWorksheetColumns`. All managed reads/writes, formula targets and
structured references resolve actual normalized headers. Default positions remain
only in initial layout construction. No Core/Domain or observation catalog change.

Candidate `deploy/validation-process-worksheet-headers` builds and compiles all five
packages, including Operations cold start. Build record
`reports/runtime/process-worksheet-headers-build`,17:25:08.8167484–17:25:47.6477536 UTC.
Compiled-source comparison:264 prior components,265 current; one changed and one
added Operations component,263 exact unchanged. Settings/frozen Close candidate
preserved, normal shutdown and zero delayed Excel Application failures.

## Focused GREEN and static evidence

Controller `reports/runtime/process-worksheet-headers-controller/497de0e5ea8349f0889fef2cb58792ac`;
result `reports/runtime/slice4be-process-worksheet-headers/82b1c1a9f15a426daea54fc98a676ed1/green.json`.
17:28:41.1123142–17:30:27.4853877 UTC: **107/107**, exact RED check order and all42
prior shared GREEN checks retained. Five instrumented compiles, package/settings
preservation, normal unassisted shutdown and zero delayed Excel Application
failures pass. Three directly reviewed images show retained custom values and the
actual rejection status on both custom-column and normalized-header cases.
The normalized-header worksheet image shows blank Excel chrome and is excluded;
its formula restoration is proven by assertions, not that image. No claim of
every-case visual coverage or human acceptance.

Earlier candidate controller8e1d7fce71f040b183e83b16aaf10baa reaches79 passes but
fails in screenshot setup with DISP_E_BADINDEX; it is excluded as GREEN. The test
now resolves/displays the disposable workbook window through a typed VBA adapter,
creating a window if absent. This changes no runtime handler or persisted package.
The completed rerun above replaces that incomplete gate.

Static evidence `reports/runtime/process-worksheet-headers-static`:272 components,
6116 procedures,134480 lines,9 literal/45 unresolved Application.Run,190 duplicate
body candidates,28 oversized modules. All existing oversized caps are non-growing;
the worksheet module shrinks1362 to1331 lines. Three report schemas and338
PowerShell parses pass. No dynamic-call or duplicate-body regression/exception.

The scoped regression gates below are complete. Full reusable specifically protects worksheet roundtrip/mixed-UOM/
multi-selection/item-search/restart plus the accepted execution behavior affected
by the worksheet call paths. Close/query protects captured/public launcher and
read-only authority behavior. Core, Domains, Admin, form/Ribbon and observation
catalog source hashes are unchanged; their other completed gates remain frozen
evidence, not rerun claims. No native-crash repair or comprehensive Slice4be coverage.

## Completed regressions on the frozen candidate

All roots below are under `reports/runtime/process-worksheet-headers-regression/`.

- `reusablefull-37d487963887409d9583325486c6e0d9`,17:31:09.8184249–
  17:39:53.5897909 UTC: both aggregates and171 boolean observations retain exact
  prior identity/order. Restart2677ms/final2502ms exit normally, zero reference
  release failures, no termination requested; four packages close in reverse
  order. Settings/packages preserved and zero delayed Excel Application failures.
- `smoke-1c9d102d48984a70a6282c879ebdfcf2`,17:40:20.3250785–
  17:40:42.9367915 UTC:86/86 exact prior checks, normal unassisted closure,
  settings/packages/tracked report preserved, zero delayed Excel failures.
  Shutdown receipt `reports/runtime/packaged-smoke-closure/b249f4f4a5284f8aa33bfe9989094455`.
- `layout-b15abea396814c9086b54182abe64bf7`,17:41:06.7421789–
  17:41:21.2669173 UTC: exact prior three-size/five-page geometry, no bounds or
  overlap violations, normal cleanup/preservation and zero delayed Excel failures.
  Three blank Run List screenshots are directly reviewed and hash-recorded;
  this is layout evidence, not populated workflow acceptance.
- `chain-8a920d82d928427f9c24c20bd051cf8a`,17:41:54.8462404–
  17:47:16.1659205 UTC: chain32/live-role48/Create Warehouse15 retain exact prior
  checks/order. Packages/settings/tracked reports preserved, normal cleanup and
  zero delayed Excel Application failures.
- Close/query controller `reports/runtime/production-close-controller/fa47fc65a8ff4803b297ec11e5c77c49`,
  result `reports/runtime/slice4be-production-close/121ef07e9c564b55b54ab3380003ec07/green.json`,
  17:47:39.2600580–17:51:13.5173942 UTC:312/312 exact prior ordered checks,
  including42 shared checks and80 query checks; five compiles, normal cleanup,
  preservation and zero delayed Excel failures. Four captures reviewed: the two
  pre-dismissal surfaces, public reopened form and Settings notice. Assertions
  prove dismissal and source preservation; these images alone do not.

## Next discovered prerequisite: Core picker headers

Read-only source audit finds `cDynItemSearch.ProcessAlternativePairNumber` uses
a case-sensitive, untrimmed prefix comparison. `CommitSelection` first updates
the selected item label, then uses that pair number to locate its Accepted SKU
column. A normalized numbered header can therefore leave the paired SKU unchanged.
This is a source-predicted label/SKU mismatch, not runtime-proven RED. There is no
pair1 fallback in this path; the initial conversational inference was corrected.
D4 owns the shared Core picker; D14/D15 require matching the selected managed
pair by normalized name. The next focused test must exercise the real packaged
picker commit, preserve other pairs/custom columns/captured workbook and check
the exact selected item/SKU pair. No Core picker implementation changed here.
