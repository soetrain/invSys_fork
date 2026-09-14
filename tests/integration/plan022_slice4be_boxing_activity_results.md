# Plan 022 Slice 4be.1 Boxing Make/Unbox activity RED

Current verified result, 2026-09-14: **1123 PASS / 57 expected FAIL** across
1180 checks. All preceding 1074 checks and 1041 GREENs remain, including Shipping
recording40/40. The added context/permission matrix is **82 PASS / 24 expected
FAIL**: 16 context failures and eight absent denial-observation assertions.
The other failures retain their original scope: 26 missing Boxing observations
and seven D8-A findings. No duplicate check or harness exception occurs.

Healthy Make/Unbox calls calibrate both entry counters. Signed-out calls still
dispatch to the service and show a permission message rather than context rejection;
the service's existing authorization check stops mutation. Same-user reauthentication
and another SHIP_POST-authorized target let the old form enter both service and
mutation owner. The test stops at mutation entry, so this proves missing pre-owner
guards, not an unauthorized inventory write. Revoked permission within the same
signed-in session correctly stops mutation and shows denial, but emits neither
required observation. Both warehouses' authority bytes, staging/custom values,
captured workbook and unrelated workbook remain unchanged. Fixture Auth bytes are
restored, and explicit relaunch/reuse passes.

All five instrumented projects compile. Excel closes normally, with no Excel
Event1000 in the verified run window; earlier native failures remain unresolved.
All 325 package pins, protected runtime sources and unrelated user documents are
preserved. Fresh static output matches every runtime metric/component line count
and all 28 limits; three changed scripts parse and 84 local links resolve.
No runtime, package build/deployment, full-chain rerun or human acceptance is claimed.
The next D13 work remains failed staging/refresh, uncertain submission and optional
tracking/policy cases before structured Boxing outcome wiring.

Current ignored evidence: `reports/runtime/boxing-context-red.log`,
`boxing-context-red-verification.json`, `boxing-context-red-native-windows.json`,
`boxing-context-red-final-verification.json` and `boxing-context-red-static/`.
Exact report:
`reports/runtime/slice4be-shipping-activity/096ad8bf00a64d4aa42152c05c3fb6a3/boxing-activity-shipping-recording-red.json`.

The initial RED checkpoint below is committed and pushed as code `2593385` and
docs `deea147`. The next test entry adds `Slice4beBoxingContext.ps1` under the same
packaged command, before the existing real-workbook-close gate. Healthy Make/Unbox
calls calibrate service and mutation entry. Signed-out, same-actor reauthenticated,
other authorized target and revoked-SHIP_POST cases retain the actual form/service
authorization, with a mutation-entry probe preventing writes. It does not substitute
for the real accepted Make/Unbox actions below. Both warehouse authority files,
staging/unknown values, captured workbook and unrelated workbook are protected;
permission denial additionally requires its REQUESTED/DENIED observation pair.
The unchanged normative Boxing refinement already requires these behaviors.

Initial run, verified 2026-09-14: the complete run finishes **1041 PASS / 33 expected
FAIL**. Those failures are exactly the 26 missing Boxing-observation assertions
and seven pre-existing D8-A findings. Every preceding 1016 check and 1009 GREEN,
including all 40 Shipping recording checks, is retained with no duplicate or
harness exception. Excel closes normally; the verified run window has no Excel
Application Error 1000. This does not establish a native-crash repair.

All 325 package pins, protected runtime sources and unrelated user documents
remain unchanged. Static regeneration preserves every runtime metric/component
line count and all 28 module limits; three scripts parse and 84 links resolve.
No implementation has changed. The next D13 work is the remaining context,
permission and explicit failed-staging/refresh matrix before owner-fact wiring.

Architecture v4.11 D18's Boxing refinement names the existing Make Boxes and
Unbox controls, their BOXING_WORKFLOW owner, SHIP_POST eligibility, structured
outcomes and exact Inventory source references. Catalog 9 implementation remains
pending; the frozen candidate is `deploy/validation-recording-notice`, runtime
source `5e2c45a`. This compatible registration follows approved semantic inheritance
and does not change business authority, permissions or canonical schemas.

The packaged test preserves the preceding eight-action Shipping recording and
starts a separate Boxing recording in the same generated warehouse/actor context.
Actual Make/Unbox handlers each post one unit, followed by zero-quantity rejection
through each handler. Unsaved probes collect actual accepted submission IDs and
the owning processing/refresh Boolean before the legacy report-text fallback.
The observer does not replace an owner, fabricate activity, parse status text for
completion, or infer application from the activity result.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-recording-notice -Phase RED `
  -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity -CheckBoxingActivity
```

All five instrumented projects compile before fixtures or forms. The focused
Boxing result is **32 PASS / 26 expected FAIL** across 58 checks. Both accepted
actions have one independently observed submitted event and exactly two owning
Inventory log lines: one package and one component under exact System_Key values.
Make/Unbox restore both entity balances. Zero quantity reaches owner validation
without submission or new log entries. Captured workbook, unknown headers,
unrelated workbook and earlier Shipping journal preservation pass.

The 26 failures are six missing-observation assertions for each of four actions,
plus missing four-action/eight-observation journal content and its expected
ten-version chain. Privacy/outcome checks fail because the records are absent;
they do not demonstrate a data leak or incorrect emitted outcome. This is
meaningful behavioral RED, distinct from compile, fixture or automation failure.

Before runtime completion, extend actual-handler evidence to captured-context and
permission loss, pending/uncertain/failed submissions, tracking-policy/store
failure and old-policy exclusion of newly registered controls. Both the form
(2938 lines) and owner module (22386 lines) are at their existing size limits;
implementation needs reviewed, protected extraction while preserving those limits.
Do not interpret the existing report-text fallback as structured completion proof.

Source review identifies another required failure case: the existing
`RunShippingRuntimeQueueRefresh` can return True after a successful read-model
refresh while `stagingOk=False`, leaving a staging warning in its report. The
normal-return probe does not prove that failure branch is clean. Before runtime
implementation, protect failed staging/refresh and explicit owner failure facts;
a generic successful return must not override D18's partial/failure rules or
justify parsing the report in the observer.

No Boxing activity implementation, package rebuild, deployment, full Release 1
chain or human acceptance is claimed. Other Boxing/Production/Admin controls,
the combined cross-role recording and full guide/How-To/Diagnostic/Compare scope
remain required. D8-A remains a separate unapproved architecture decision.

Ignored evidence: `reports/runtime/boxing-activity-red.log`,
`boxing-activity-red-verification.json`, `boxing-activity-red-native-windows.json`,
`boxing-activity-red-final-verification.json` and `boxing-activity-red-static/`.
The exact complete result is
`reports/runtime/slice4be-shipping-activity/b5779eb1b3954a65bafcc98616871ae0/boxing-activity-shipping-recording-red.json`.
