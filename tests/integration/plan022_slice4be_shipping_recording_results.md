# Plan 022 Slice 4be.4 Shipping recording evidence

Last verified 2026-09-14: **40/40 recording checks pass** on the unchanged
candidate. The complete standard-order route reports **1009 PASS / seven existing
D8-A FAIL**, with no harness exception or duplicate check identity. All 1016
checks and 1008 passes from the preceding calibrated run are retained. All 963
comparable passes from the earlier 974-check prepared-fixture route remain;
its four ordering-specific preparation checks are not executed by this route.
The seven unavailable-Auth recreation findings remain unapproved and unresolved.

All five instrumented projects compile before fixtures; Excel closes normally.
Static regeneration preserves all runtime metrics and component line counts,
including all 28 module limits. Final verification preserves 325 package pins,
protected runtime sources and both unrelated user documents; three scripts parse
and 84 local links resolve. No Excel Application Error 1000 is recorded in either
completed calibrated run window. Earlier setup failures below remain evidence;
this is not a general native repair or full Shipping/Release 1 acceptance.
No new build, deployment, full Release 1 chain or human acceptance is claimed.

This test advances D18's comprehensive recorded-control coverage. It adds no
runtime behavior or architecture decision. The frozen runtime is source
`5e2c45a`, candidate `deploy/validation-recording-notice`; no XLAM is rebuilt.
Missing recording evidence would be behavioral RED, but a passing existing
implementation does not require a manufactured failure.

The generated Shipping fixture enables capture through the actual Admin policy
editor before signing in as its Shipping actor. Actual Viewer Start/Stop controls
surround Add, Update Row, Send Hold, Return, Remove, a second Add, To Shipments and
Shipments Sent. The existing Shipping form handlers, independent submission
observer, staging/source checks, captured-workbook checks, unknown-column checks
and unrelated-workbook preservation remain in the same run.

The added assertions require one attempt/result pair per action, a single sequence
with ordinals 1-8, eight distinct ActivityIds, sixteen exact ordered observations,
all four independently observed submitted event IDs and an eighteen-version
integrity-linked journal. Stopped remains a capture lifecycle with no invented
expected conclusion. Earlier activity bytes must remain unchanged. Test output
contains fixed check names and booleans; fixture payloads are not committed.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-recording-notice -Phase GREEN `
  -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity -CheckShippingRecording
```

This standalone Shipping recording does not prove one sequence spanning Admin,
Receiving, Production and Boxing. Those combined role bindings and comprehensive
missing control coverage remain required, as do guide authoring, How-To,
Diagnostic/Compare, import/export and human acceptance. The prior 374-check
recording/evaluation gate retains its independent scope. Pending D8-A findings
remain separate from this evidence addition.

The first standard-order run stops at **283 PASS / one harness failure** in
`ActivityShippingCreateBox`, before Start Recording. The RPC failure is not
behavioral recording RED. A surviving child of the test Excel process has a
verified empty workbook collection and three recognized add-in projects; normal
Quit closes it without forced termination or recovery-file deletion. The verified
Application Error 1000 query finds no events in this run window, so a native
crash cause is not established.

The unchanged retry uses the existing complete Shipping-first route with fixtures
prepared before Shipping probes (`-ShippingBeforeSharedFormsForTest
-PrepareShippingFixturesBeforeProbesForTest`). This ordering previously completed
the expanded Shipping suite. Its diagnostic filename identifies that ordering;
it does not erase the failed standard-order run or substitute for full acceptance.

That retry stops at **6 PASS / one harness failure** during Seed, also before
recording. Its window contains an Excel `ntdll.dll / c0000028` Application Error
1000. The WER-linked recovery instance has no open workbooks; normal Quit reaches
the recovery prompt. **Yes, I want to view these files later** is selected and
visually verified before confirmation. Excel closes and recovered files remain.

The next harness calibration installs Shipping and recording probes once before
fixture/form activity, then explicitly compiles all five instrumented candidate
projects using the existing compile verifier. The existing Shipping route reuses
those installed probes; business owners and on-disk XLAMs are unchanged. All five
compile checks and the no-loaded-forms check pass. This is not a native-crash
repair claim. The first calibrated run completes 1008 PASS / eight FAIL, including
39/40 recording passes. Its only new failed assertion compares the storage hash
with the observation body. The correction excludes only that outer metadata field;
all sixteen complete bodies must match in order and journal integrity remains
independently checked. The repeated complete route supplies the result above.

Ignored evidence under `reports/runtime/` includes the four
`shipping-recording*-green.log` files, compiled/verified summaries, native-window
verification, final package/source/link verification and the regenerated
`shipping-recording-verified-static/` reports. Exact completed result files are:

- `slice4be-shipping-activity/c7f86c3644df40edb5167d95ee820076/shipping-recording-green.json`
  (1008/8, initial body-comparison assertion).
- `slice4be-shipping-activity/28ec64b86c9446e1aaa3ec838553a667/shipping-recording-green.json`
  (1009/7, all 40 recording checks pass).

The full objective remains active. Next protect actual Boxing Make/Unbox activity,
captured-context and exact source references, followed by the combined cross-role
recording; do not substitute this standalone Shipping result for those obligations.
