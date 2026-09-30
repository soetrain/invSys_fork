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

## Extended partial-result and guard RED

Controller `reports/runtime/process-worksheet-activity-controller/5e5d13b95ab2498dbb4322400e20d706`,
result `reports/runtime/slice4be-process-worksheet-activity/8c2df34b83594a73aedab379eea2d2d7/red.json`,
2026-09-30,19:29:22.9170273–19:33:04.8811071 UTC:
**119 PASS / 214 expected FAIL / 333 checks**. Every original220 check retains its
exact result and relative order, including all42 shared GREEN. Five compiles,
package/settings preservation, normal closure and delayed Excel audit pass.

Three faults are injected only into the unsaved fixture package, with each fault
verified to occur exactly once. The real handlers and real Designs queue writes
still execute. UncertainSecond suppresses the second actual write's acknowledgment:
one table remains, with exact owner states Submitted/Unknown. RemoveSecond returns
failure before the second removal: one table remains, with Submitted/Submitted.
SaveSecond follows the real second lo.Delete, then takes the existing failure
return before wb.Save: no tables remain in memory, the workbook is unsaved, and
both references are Submitted. This simulates a save refusal; no real disk fault
is claimed. The tests do not fabricate identities, parse report text or use
unhandled Domain Err.Raise/VBE interruption. Later processing may catch up the
earlier queued event; each action's expected references still come only from its
own submission returns. Missing FAILED pairs account for24 new failures.

The32 other new failures prove current worksheet/save guards are absent for all
three actions under Loading, Nested, changed Session and changed Target, and for
Send/Add while SignedOut. Retrieve submits in the first four contexts; signed-out
Retrieve does not submit. Inventory business state remains exact. These are
behavioral failures under D18, not compile or setup failures. No runtime fix yet.

## Closed-workbook fixture calibration — unresolved in the full sequence

Broader capability/policy tests and form-draft preservation checks are written,
but the full sequence has not reached valid completion. Do not count those test
definitions or partial output as additional accepted coverage.

- Excluded controller `80ed43dbd7464c2c96f1cd26a229604a` under the same controller
  parent,19:36:50.9240298–19:43:57.1351238 UTC, stops at
  TestProductionDesigner.WorksheetActivityGuard after workbook closure. The owned
  Excel host shows a VBA automation dialog,0x80010007, with Continue disabled.
  VBE End/Reset is not used. After verifying the exact owned PID/start identity,
  only that disposable host is forcibly terminated at19:43:43.6633634 UTC so the
  live controller can restore settings. The subsequent RPC error and harness
  failure are not product RED. Package/settings restoration and delayed audit
  pass, but assisted termination excludes the attempt.
- Isolated `-ClosedDiagnostic` controller `0e95014cd1c547b488fbc29e99b5a1a9`,
  19:46:46.9664771–19:48:03.0987213 UTC, result
  `reports/runtime/slice4be-process-worksheet-activity/33dc04141f654c84aff029a08438f634`,
  shows the actual form before closing its captured workbook. An outer adapter
  trap prevents an unavailable form reference from becoming an unhandled VBE
  error. The form adapter is entered and returns normally, outer error0, workbook
  bytes preserved and no activity. Shared42 and five compiles pass; normal
  closure, preservation and delayed audit pass. This is diagnostic-only, not a
  full guard gate or proof of a causal repair.
- Excluded controller `a16832f3a7c3457c9dbf6a493297d435`,
  19:49:03.5524662–19:52:45.0175255 UTC, uses the visible-form setup in the longer
  sequence. The first closed-workbook guard cannot enter its form adapter. The
  defensive entry assertion stops the test as a fixture failure. It closes
  normally without another dialog or forced termination; helper/package/settings
  pins and delayed audit pass. This does not establish that the actual operator
  handler failed, or that the form remained reachable after workbook closure.

All three delayed Application1000/1001/1002 audits have zero Excel events.
Desktop probes remain healthy, with no error5. The HRESULT/fixture failure must
not be confused with the user's conditional desktop-error5 stop instruction.
No native-crash repair is claimed. All347 PowerShell files parse.

## Next required work

Before runtime implementation, calibrate the closed-workbook case against actual
operator visibility, workbook lifetime and form-entry evidence after the longer
sequence. Distinguish an already-dismissed surface from an available handler;
do not force a disconnected private form reference or count inability to enter it
as behavioral RED. Use fixed boundary facts and a targeted diagnostic, not another
unchanged full rerun. Then run the written capability-loss, optional-tracking and
older-policy cases to completion. Preserve all333 check identities and every
passing assertion; retain the original220 sequence as a subset. Full closed-binding,
form-draft, capability and policy protection remains unproven.
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
