# Slice 4be-A / early B0: Receiving replay

Last verified: 2026-10-04. D18-REPLAY-01 is approved. Profile authoring is focused
GREEN on `validation-execution-profile-03`; the expanded B0 gate remains RED:
**45 PASS / 17 FAIL**, because Run How-To and fresh proof are not implemented.
All 42 prior passing identities remain passing. No replay is claimed.

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

The same packaged flow now protects Run setup, target-local entity selection,
explicit Start Run, fresh recording and exact new owner results. Implement the
runner against its RED; extend target/rights/policy/stop guards before changing
their behavior. Setup must not execute, and guide use must not require authoring rights.
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

## Runner protecting RED

The same actual recording and authored profile now feed `Slice4beReceivingRun.ps1`
and `Slice4beReceivingRunProof.ps1`. Their ordinary form actions cover setup,
exact guide/profile/Training target, missing local entity, changed context,
explicit Start Run and read-only Verify. Independent reads require an immutable
run chain, fresh recording/activity/source identities, exact new entity and
quantity2.5 application, cleared staging and the preserved custom column. The
evaluator must conclude from this fresh recording and exact applied event;
ordinary Viewer Refresh is explicit and Verify cannot publish or dispatch.

Candidate03 final RED: **45 PASS/17 FAIL**, exit1, 23:48:22.856--23:51:55.601 UTC.
All 42 preceding GREEN identities remain GREEN; zero harness exceptions. The 17
failures are missing Run entry/setup/dispatch/fresh proof. Three new preservation
checks pass; passing preservation alone is not evidence of replay. Setup checks
permit non-executing run metadata, but require unchanged owner files and recordings.

Local receipts under `reports/runtime/`:

- `receiving-replay-controller/05849c8b66cd49acbf982ccf8ceab1db/closure.json`
- `slice4be-receiving-replay/2ae83ac13f9044058937fe16db7a4895/red.json`
  and `receiving-replay-scope.json`
- `receiving-run-red-verification.json`

Closure preserves all five package hashes/settings and closes Excel. Five saved
disposable probes compile; all391 repository PowerShell scripts parse. No runtime
source changed since the validated profile checkpoint, so its static evidence
remains applicable. Application audit since23:05:32 UTC finds no new Excel native
failure. Desktop probe23:52:10.198 UTC (16:52 PDT) passes cursor/input desktop/capture
with errors0/0/0. No visible acceptance or deployment is claimed.

Earlier controller225d94ac62dd4fafafb94c3afa47e8ca ended with a harness error:
an empty activity list produced a non-Boolean regex assertion. It is not the RED
receipt. Explicit array JSON fixes that; controller11ec9bee3e7742799000c40f46201cdf
then passes the harness with45/17, and the final run repeats it after removing an
unnecessary read of a null unselected list value. Further permission/policy/stop,
packaged owner dispatch, layout, visible and full-chain gates remain open.

## Packaged runner checkpoint (2026-10-04)

The corrected preimplementation test repeats **45 PASS/17 FAIL** on frozen
profile03. It checks all run revisions, including Ready, and distinguishes
REQUESTED observations from completed owner outcomes.

Core now binds the generated Training runtime, exact guide/profile, permissions,
policy and loaded builds; it appends immutable run revisions and reuses the recorder
and evaluator. Operations supplies typed inputs to ordinary Receiving handlers,
retaining the exact generated operator workbook object. The runner has Start,
Next, Stop and separate read-only Verify; no retry, rollback or repair.
Receiving list preparation and selection handling are factored without growing
their oversized source modules. Replay deliberately arms the existing selection
input boundary; ordinary programmatic list changes retain their previous rules.

- Candidate01: five compiles pass, **46/16**; expanded controls give **47/25**.
  Disposable stage markers identify package discovery as the setup refusal.
  Training authority, guide/profile, policy and permissions already pass.
- Candidate02 uses named loaded-XLAM lookup instead of workbook enumeration:
  **71 PASS/1 FAIL**, 00:18:02.877--00:22:54.285 UTC. All six Receiving actions
  complete through ordinary owners. Independent proof finds fresh identities,
  an exact applied event/new System_Key/quantity2.5, preserved custom column and
  cleared staging. Verify concludes from that exact event and leaves business,
  publication and run evidence unchanged.
- Step through waits for Next, executes one owner, rejects repeat Start, retains
  Stop after one step and refuses later dispatch. A stopped run cannot borrow the
  completed run's success. All prior profile checks remain GREEN.
- Sole failure: minimum-width Stop/Next overlap. Default/enlarged layouts pass.
  Candidate03 fixes the Stop anchor; its final evidence follows below.
- Candidate02 static:306 components/6278 procedures/137359 lines (+8/+56/+998);
  dynamic calls9/45, duplicates190 and all28 oversized caps retained. All393
  PowerShell scripts parse. Compiled comparison preserves284/291 existing
  components; seven intended existing components change and eight are added.

Local receipts: controller directories below are under
`reports/runtime/receiving-replay-controller/`; other paths are under `reports/runtime/`.

- Corrected RED: controller`03ae5425810a4b419c642ca797e25205`,
  worker`slice4be-receiving-replay/51c4b9d019fc433e8623df58f4c1e434/red.json`.
- Candidate01: controller`20e97490c8724dc49e4e6b0df11f7c9c`;
  diagnostic/controls controller`70064844ae024fb5b762441436e40214`.
- Candidate02: controller`e77afae56ef843a2b1b4df3ec2de5070/closure.json`,
  worker`slice4be-receiving-replay/9c7e4c4cd6c24a6d83da1eba646b700f/green.json`;
  `receiving-run-build-02/`, `receiving-run-static-02/`.

All completed controllers restore settings, preserve packages and close Excel.
No deployment or B0/A/B/R1 acceptance is claimed. Focused permission/policy/nested
entry and workbook-replacement guards, Receiving/native and published-guide
regressions, control observations, visible proof and the open full-chain native
failure remain outstanding.

Final candidate03: **72 PASS/0 FAIL**. Minimum/default/enlarged layouts all pass;
all71 candidate02 GREEN identities and all42 profile-checkpoint GREEN identities
are retained. Five cold compiles pass. Static evidence is306 components/6278
procedures/137360 lines (+8/+56/+999); dynamic calls9/45, duplicates190, all28
oversized caps and all three evidence schemas pass. Source comparison still
preserves284/291 existing components. All393 PowerShell scripts parse.

Final receipts:

- `receiving-replay-controller/8d659cbd855d4179b07169f4491260e9/closure.json`
- `slice4be-receiving-replay/7acfae6f3dfb41829e456c24a9d99d8f/green.json`
- `receiving-run-build-03/` (including`source-preservation.json`)
- `receiving-run-static-03/ratchet-verification.json`
- `receiving-run-focused-verification.json`, `receiving-run-native-audit.json`

Desktop probe00:28:03.762 UTC (17:28 PDT) passes cursor/input desktop/capture0/0/0.
Application audit23:54:23.853--00:27:03.641 UTC finds zero new Excel native failures.
These checks do not close the earlier full-chain native failure. Next, add focused
permission/policy/nested-entry and exact-workbook replacement tests against this
frozen candidate before further runtime changes; then run affected owner regressions.
