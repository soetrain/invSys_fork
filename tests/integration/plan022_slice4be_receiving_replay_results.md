# Slice 4be-A / early B0: Receiving replay

Last verified: 2026-10-04. D18-REPLAY-01 is approved. This checkpoint establishes
the shared profile/run wire and a meaningful packaged entry RED, not replay GREEN.

## Evidence

`Test-Slice4beReceivingReplay.ps1 -Phase RED` against frozen
`deploy/validation-warehouse-purpose-02` returns **16 PASS / 4 FAIL**, exit1,
2026-10-04 22:52:06.159--22:54:33.826 UTC.

All five disposable packages compile. The fixture generates a Training warehouse,
starts recording through Viewer, opens Receiving through its Ribbon callback,
then invokes the ordinary Refresh, Clear, selection, Add and Confirm handlers.
Selection invokes the existing mouse-down/up provenance handlers in a disposable
test seam; it never directly sets the token or constructs an activity. Native
keyboard/mouse behavior remains covered by the existing navigation gate.

The original transaction has exactly one matching applied EventId and inventory
log row with its exact `System_Key` and quantity delta. Successful Confirm clears
staging and preserves the custom header. The journal chain and exact source
reference are intact; observations omit the dummy reference/location inputs.
Actual guide/expectation Save produces an integrity-checked six-action guide bound
to that recording and a SourceEventsApplied conclusion. These prerequisites pass
before testing execution entry.

Four product failures isolate missing Configure execution, its editor, exact guide
binding in that editor, and Run How-To. Setup inspection leaves journal/activity
bytes and staging unchanged. No harness exception contributes to this RED.

Local receipts under `reports/runtime/`:

- `receiving-replay-controller/a4fce409bb2f47ec9365f9e329984d6b/closure.json`
- `slice4be-receiving-replay/af384feb9b90493691bbe00db7a2ffe9/red.json`
- Same worker root: `receiving-replay-scope.json` records six source steps,
  ReplayExecuted=False, FreshReplayProof=False and B0Accepted=False.

Closure confirms settings restored, all five frozen XLAM hashes preserved and
Excel closed. The follow-up Application audit finds zero Excel native failures;
all 387 repository PowerShell scripts parse and scoped diff checks pass.
There are no runtime source changes in this checkpoint; existing
purpose02 static/business evidence remains applicable. This test does not replace
layout, visible, live-role, full-chain or acceptance gates.

## Next protecting test

Extend this same packaged flow through reviewed typed inputs, profile save/reopen,
explicit Start Run, fresh recording and exact new owner results before implementing
those consumers. Protect target/rights/context refusals and setup non-execution.
Do not satisfy B0 with empty editor controls or reuse the original transaction's
success as replay evidence. Full B0, remaining A scope and broader B remain open.

Earlier attempts were fixture/harness failures, not product RED: unsupported
compiled-probe flag combinations, an incorrect COMPLETED expectation (Receiving
uses CONFIRMED), an uninitialized evidence-workbook list, and programmatic selection
correctly omitted by the recorder. The final run corrects all four causes without
altering the runtime contract.
