# Slice 4be Process worksheet header preservation

Last verified: 2026-09-30 UTC. Focused RED established; isolated correction is
under validation. No worksheet tracking registration or release acceptance.

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

Pending: exact focused GREEN and visible captures, static ratchets, packaged smoke,
full Release1 chain/live roles/Create Warehouse, layout, full reusable Production
and Close regression on this candidate. Existing Close evidence remains frozen;
no broader coverage or native-crash repair is claimed.
