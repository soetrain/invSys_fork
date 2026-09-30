# Slice4be Process worksheet observations — test-first baseline

Authority: Architecture v4.11 D18's Process worksheet discovered-control
refinement, committed in docs `ba3c206` before these tests; Plan022 and Controls
1.346 synchronized. D14/D15 retain local-save/import behavior and authority.
Catalog22 is specified, not implemented. The frozen baseline is
`deploy/validation-process-worksheet-picker`, whose scoped gates are complete in
`plan022_slice4be_process_worksheet_picker_results.md`.

## Focused RED

Controller:
`reports/runtime/process-worksheet-activity-controller/d1b17c9d89774c1fad17fdbfb36266bc`.
Result:
`reports/runtime/slice4be-process-worksheet-activity/55e84e2d17b1447cad76687175837592/red.json`.
Run: 2026-09-30,19:21:29.0507595–19:22:56.7300898 UTC.
**62 PASS / 158 expected FAIL / 220 checks**, exact expected failure sequence.
All42 shared prior GREEN checks and five instrumented package compiles pass.

The110 catalog failures cover the missing catalog22 extension, preservation of
all106 earlier control definitions within that absent version, and three missing
metadata definitions. The48 other failures arise from six absent attempt/result
pairs: two Send actions, Add, rejected Retrieve, single-table Retrieve and
two-table Retrieve. Their correlation, context, outcome, reference, redaction,
integrity and terminal assertions cannot pass without those records; this is not
evidence of corrupt records or leaked values. Actual handlers run successfully.

Unchanged packaged behavior passes: Send saves only the captured workbook despite
an active decoy; Add appends its pair and saves; invalid Retrieve leaves both
tables; valid single retrieval removes only its selected table; multi-area
retrieval removes both selected tables. Local actions preserve canonical bytes.
Successful imports preserve Auth/Config bytes and the exact six Inventory
business tables. Unrelated workbook values remain intact. Queue-return observation
captures actual IDs/states in memory: zero for local/rejected actions, one for
single retrieval, two for multi retrieval. No fabricated identity, status-text
parsing, Domain handler replacement or source-value log is used. The unsaved
adapter calls the real form handlers and observes the existing submission return;
it does not intercept an error handler's Err reads.

Settings/package pins are preserved, Excel closes without forced termination,
and the delayed Application1000/1001/1002 audit is clear. Desktop probes report
no error5. One operator capture is directly reviewed and hashed in verification:
the populated Process Designer shows three DRAFT definitions and the actual
"Retrieved 2 selected Process table(s) as DRAFT." status. It proves the visible
business result, not tracking, human acceptance or a completed Action Path.

No runtime VBA, package, catalog registration or static runtime metric changes.
The three changed/new PowerShell test files parse; all347 PowerShell files parse.
Generated fixture/report/image data remains ignored. This is meaningful initial
RED, not a compile/setup failure or an acceptance claim.

## Next required work

Before runtime implementation, extend this actual-handler baseline for partial
multi-table outcomes, failed removal/save, uncertain submission acknowledgment,
exact per-event Submitted/Unknown references, context/capability/loading/nested
guards, tracking disabled/unavailable and older-policy behavior. Preserve all220
check identities and every passing baseline assertion. Do not treat this initial
happy/rejected-path RED as comprehensive protection of those untested cases.
The current DeleteProcessWorksheetTable performs lo.Delete before wb.Save. A
post-confirmation workbook-save failure can leave a local deletion in memory;
the failure observation must retain the confirmed Designs reference and Unknown
effect, not claim restoration. D15 still prohibits removal before confirmed
Designs draft save. The normative observation section explicitly distinguishes
these cases without changing the existing save algorithm.

Then implement catalog22 and typed owner observations without growing existing
oversized modules or changing the worksheet algorithm. GREEN, packaged build,
compile, static ratchets, independent recording/publication/Event Detail and
How-To/Diagnostic/Compare evidence, layout/live-role/full-chain/reusable regressions
and visible operator evidence remain outstanding. No promotion or Slice4be
completion is claimed.
