# Slice 4be-A / early B0: Receiving replay

Last verified: 2026-10-04. D18-REPLAY-01 is approved. Profile authoring is focused
GREEN on `validation-execution-profile-03`; the combined B0 gate remains RED:
**42 PASS / 1 FAIL**, because Run How-To is not implemented. No replay is claimed.

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
The initial entry-RED checkpoint changed no runtime source; existing
purpose02 static/business evidence remains applicable. This test does not replace
layout, visible, live-role, full-chain or acceptance gates.

## Next protecting test

Extend this same packaged flow through Run setup, target-local entity selection,
explicit Start Run, fresh recording and exact new owner results before implementing
the runner. Protect target/rights/context refusals and setup non-execution.
Do not satisfy B0 with empty editor controls or reuse the original transaction's
success as replay evidence. Full B0, remaining A scope and broader B remain open.

Earlier attempts were fixture/harness failures, not product RED: unsupported
compiled-probe flag combinations, an incorrect COMPLETED expectation (Receiving
uses CONFIRMED), an uninitialized evidence-workbook list, and programmatic selection
correctly omitted by the recorder. The final run corrects all four causes without
altering the runtime contract.

## Profile authoring checkpoint

The actual Configure execution callback opens a reusable input editor only after
Core validates authoring rights, exact guide and existing profiles. Core owns typed
input validation, strict JSON and immutable profile chains. Operations owns the
form; no business handler dispatch exists yet. Entity input is a registered local
prompt, never a copied original key. Quantity uses validated invariant decimal text.

- Expanded baseline: purpose02 **RED16 PASS/13 FAIL**, controller
  `a302d19568464d298262e5b5933f17f0`, worker `28a60afc192848cc9637442a9b831acb`.
- Candidate01: **28 PASS/1 FAIL** proves input edit/save/reopen and unchanged
  inventory/activity. Extended guards expose **35 PASS/7 FAIL**: six invalid-profile
  refusals retain a blank loaded form, while Core correctly rejects their files.
  Candidate02 lifetime cleanup alone retains those failures. Candidate03 moves
  validation ahead of form creation and resolves all six.
- Candidate03: **42 PASS/1 FAIL**, 23:25:42.119--23:28:57.735 UTC. All prior passing
  identities remain passing. Profile scope covers exact guide/conclusion/step order,
  missing inputs, version preservation, close, unknown/duplicate fields, unsupported
  adapter version, original-key literal, quantity expression and wrong-guide hash;
  also minimum/default/enlarged geometry, target-change refusal and reader denial.
  The sole remaining failure is `ReceivingReplay.RunHowToControlPresent`.
- All five candidate03 cold compiles pass. Compiled comparison against purpose02
  preserves284/286 existing components; only the journal child allowlist and
  published-guide form change, with five new components. Receiving, Production,
  Shipping, Boxing, Domain and other Core/Admin component bodies are preserved.
- Regenerated static evidence:298 components/6222 procedures/136361 lines
  (+5/+40/+690). All28 oversized limits hold; dynamic calls9/45 and duplicate
  groups190 are unchanged. A shared draft-match predicate removes the new duplicate.
  All389 PowerShell scripts parse; all three evidence schemas validate.
- Test controllers preserve package hashes/settings and close Excel normally.
  Visible operator capture and observations for the new editor controls remain open.
- Published-guide regression on candidate03: **58 PASS/0 FAIL**, including editing,
  cancel, versions, selection/reader closure, permission/policy/context guards and
  unchanged prior files. Closure preserves settings/packages and closes Excel.
  Application audit from 23:05:32 through 23:36:32 UTC finds zero Excel native failures.

Additional local receipts under `reports/runtime/`:

- `receiving-replay-controller/640c20a715d849adb094f1c4a8053ac9/closure.json`
- Guard RED: `receiving-replay-controller/aa0b9c0090cd4034ad2a9e7091af1217/closure.json`
- Refusal diagnostic: controller `7681d6928fd143b6a0bb550f5134f457` (candidate01),
  `e8af2592e21e439bb3b2bb023ae3c49d` (candidate02).
- Final: `receiving-replay-controller/779c8e3752ed416598de4190d412b940/closure.json`,
  `slice4be-receiving-replay/8fce7f9f92344146b29e6ee653510f21/green.json`.
- `execution-profile-focused-verification.json`, `execution-profile-build-03/`,
  `execution-profile-static-03/ratchet-verification.json`.
- `execution-profile-guide-regression/2e894aab2551401593d449f6b20f4910/closure.json`
  and `worker.log` (23:29:18--23:35:54 UTC).

The Phase=GREEN filename does not override the exit1/remaining Run failure. Profile
proof is a scoped checkpoint; B0, 4be-A, broader B, visible acceptance and the open
native full-chain failure remain separate outstanding work. No deployment.
