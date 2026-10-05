# Slice 4be-A / A1: per-user recording policy

Last verified: 2026-10-05 UTC. Approved D18-REPLAY-01 boundary7 now has a concrete
policy-v2 wire definition in Architecture v4.11; Plan022/Controls1.453 agree.
Initial **52/18 RED** becomes **91/0 GREEN** on `validation-user-policy-02`,
retaining all83 preceding checks. Settings202/202, observations535/535, replay94/94,
five cold compiles and static ratchets pass. A1/4be-A
acceptance and deployment remain open.

## Contract and protecting route

The versioned policy adds sparse `{UserId, Record}` overrides, enabled by default,
and a counted append-only user table. Preserve v1 reads/history, reject downlevel
writes that would erase v2 flags, preserve unknown columns and fail closed on
malformed policy. Admin stages/saves through its existing D5 authority; Core owns
optional collection/capture decisions. Visibility, rights, history, required audit
and ordinary business work remain independent. Removed-user flags are retained;
roster reads must not repair/write Auth or expose credentials.

`Test-Slice4beUserTrackingPolicy.ps1 -Phase RED` uses the existing compiled Settings
route with all five packaged projects compiled before forms. The new live-control
probe looks up the real named ListBox/CheckBox/buttons and changes their values;
Save/Reload/Reset use their actual handlers. Missing controls return MISSING,
never a constructed policy. Supplemental public Core recording Start/Status tests
and real General Settings saves check actor behavior and earlier record visibility.
No activity or recording outcome is fabricated.

## Evidence

- Initial59-check gate:45 PASS/14 FAIL; controller
  `user-tracking-policy-controller/26a1d3862ee9443a824978c2ca53ae21/`, worker
  `slice4be-tracking-settings/9cdf5308ede04a5da4d1eb050bbeeb0c/red.json`,
  02:34:35.687--02:35:53.891 UTC.
- Expanded70-check gate:52 PASS/18 FAIL; controller
  `user-tracking-policy-controller/9c0acc29300f48aeaf1d7671e26ccf6c/`, worker
  `slice4be-tracking-settings/4f5d396072bc4457a3ab43f2ea49eb69/red.json`,
  02:37:10.612--02:38:51.917 UTC. All18 failures are UserPolicy checks; no duplicates.
- Each controller preserves `worker.log`, `package-pins.json` and `closure.json`;
  the expanded controller also has `verification.json`. All paths are under ignored
  `reports/runtime/`; raw runtime values are not committed.

Failures isolate missing v2 projection/user controls, staging/save/reload/reset/close
behavior and protection against downlevel loss. Actor checks reproduce the intended
disabled user being able to start capture and emit optional activity. Enabled
recording, real Settings work, earlier activity visibility and policy-change partial
closure pass. Malformed-request negatives currently pass because v2 is unsupported;
they must remain passing once v2 is accepted. The first rejected harness option
(`ba79a76f159f446194f27ea5332af8f1`) never opened Excel and is not product RED.

Both RED controllers restore settings, preserve the five frozen packages and
close Excel normally. At that preimplementation checkpoint, no runtime source had
changed; runner93, guide58 and Receiving854 remained applicable to candidate05.
Regenerated `user-policy-static-red-01/`
retains307 components/6279 procedures/137389 lines, dynamic calls9/45 and190
duplicate groups; all28 size limits and three schemas pass. All399 repository
PowerShell scripts parse. The expanded gate retains all45 initial passes.
Native audit02:34:35--02:40:22 UTC finds zero Excel1000/1002 events; desktop probe
02:40:23.133 UTC (Oct4 19:40 PDT) succeeds. No desktop error5 was observed.
That RED gate alone does not prove the later implementation or its acceptance.

## Remaining acceptance

Add removed/new-user lifecycle guards and disabled-user ordinary Receiving
required-audit proof. Preserve guide58 and Receiving854 on the changed candidate.
Replay94 retains all93 prior checks and proves per-user capture-loss stopping.
Broader A1/A2/B0 and
the previously recorded full-chain native failure remain open.

## Implementation checkpoint

Core reads v1/defaults as v2 in memory, appends counted v2 user rows through D5,
rejects downlevel loss and applies the current actor's flag to collection/capture
while retaining visibility. Read-only roster projection exposes identities and
availability only; personal Settings excludes other users' flags. Admin uses the
actual user selector/checkbox and existing Save/Reload/Reset/Close handlers.
Catalog27 introduces two navigation/staging IDs and rejects them in older catalogs.

Under `reports/runtime/user-tracking-policy-controller/`:

- `f38f1860746a4de8b94606efd4ef0a2f`:70/0,02:48:36.383--02:50:36.433 UTC;
  worker `slice4be-tracking-settings/4dc88dbe375c4bb3841fbca3f4d121c9/green.json`.
- `6a1103d7dbdd407ba3beb30ccaf9b392`:82/1,02:51:01.634--02:54:30.598 UTC;
  worker `5abd9aee4aea431ab575660ee4c8b6cb`. The worksheet-deletion fixture did not
  prove its table absent; its lone missing-table assertion is not product RED.
- `95c1e413590041b98ffbd03e85d0b22f`:83/0,02:55:16.858--02:58:42.905 UTC.
  Direct table deletion now verifies absence, with separate editor/capture/byte
  facts. No runtime change was needed for this guard; the initial fixture's exact
  failure cause remains unproved. Missing table/count, bad count/flag, duplicate/
  orphan rows and unknown schema fail closed. Cancellation, unknown columns and
  v1 history/upgrade pass.

These controllers preserve settings/packages and close Excel normally. Candidate01
passes five cold compiles; its inspected `tracking-settings-page.png` shows the
new roster, Record user, audit explanation and Save/Reset/Reload without overlap.
Static01 exposes one identical FindTable/FindPolicyTable body (duplicates190->191).
Candidate02 shares the Core helper instead.

## Candidate02 checkpoint

- Controller `0079c69c45a04d32bdbe56cc3e3cefe0`: **91/0**, retaining all83 prior
  checks without duplicates; worker `c765babab3ef4521b48073114671054a/green.json`,
  03:06:44.289--03:10:27.941 UTC. Eight added access checks prove minimal Admin
  roster projection, denied reader access/writes, own-user-only personal flags,
  unavailable dirty/missing Auth and preserved Auth/Config bytes. Settings and
  all five packages are preserved; Excel closes normally.
- The preceding controller `28c5d46dee0249fd8e8d49cada3aee68` records67 passes and
  one harness failure: access helpers were loaded only inside the installer scope.
  Loading them in the caller fixes the harness; no runtime change/product RED.
- The successful worker's `user-policy-saved.png` is directly reviewed: selected
  disabled user, unchecked Record user, explanation and save/reset/reload controls
  are legible without overlap. This is scoped Settings evidence, not A2 acceptance.
- `user-policy-build-02/compile.log`: cold Operations start and all five compiles
  pass. `source-delta.json` compares302 components with candidate05's300: seven
  intended changed components, two new helpers, no removals. The293 untouched
  components retain case-insensitive code and exact string-literal hashes.
- `user-policy-static-02/ratchet-verification.json`:309 components,6296 procedures,
  137767 lines; dynamic calls9/45 and duplicate groups190 unchanged. All28 existing
  size caps and three schemas pass. Feature growth is2 components/17 procedures/
  378 lines; no oversized-module growth. All401 PowerShell scripts parse. No
  exception or deployment is claimed.
- Settings regression controller `user-policy-regression/settings-9e9e2880d151454a8c3d3bf9ea6cca1d`:
  **202/202**, all202 previous identities/GREENs retained without duplicates,
  03:10:30.700--03:15:06.266 UTC. Worker
  `slice4be-tracking-settings/03f839315bc84267b59b8a825c5ca155/green.json`; ten
  captures directly reviewed. New-process preference, Operations without Admin,
  role/context/binding, cancellation, layout and policy/profile guards pass.
  Settings/package preservation and normal Excel closure pass. The stale policy
  assertion now explicitly requires schema2/catalog27/131 controls/zero defaults;
  this harness expectation update is not product RED.
- Observation controller `observations-a1acb27dda9943ae9708e93a5bd1bba2` exits2
  before a worker report, with empty log and PowerShell1001 parameter-matching
  events at03:15:28 UTC. No Excel fault is identified. An explicit string-array
  argument retry starts normally; exact host fault cause is not proved. This is
  setup failure, not behavioral RED.
- Observation retry `user-policy-regression/observations-911672a8cc0848aea7d035ab3696c280`:
  **535/535**, 03:16:31.334--03:29:37.107 UTC, worker
  `slice4be-settings-activity/33e8f090f1d94f328bfdc6fbf1a070d6/green.json`.
  All506 referenced prior checks remain in exact relative order;29 additions are
  sixteen new-user-control checks and thirteen current-catalog older-policy
  exclusions. New handlers emit exactly one fixed-vocabulary pair each, omit
  selected values and complete a separate two-action recording. Original25-action
  recording, policy interruption, cancellation and context guards remain GREEN.
  Eight captures are reviewed; delayed unassisted closure, settings restoration
  and five package hashes pass. The current event assertion uses catalog27.
- Candidate02's Receiving template is copied unchanged from candidate05 with an
  exact hash match after Excel closed. Cursor probe03:30:03.965 UTC succeeds;
  no desktop error5 is observed.
- Receiving controller `receiving-replay-controller/9a235a37f49d4745b80176525474d881`:
  **94/94**,03:30:10.465--03:39:27.864 UTC; worker
  `slice4be-receiving-replay/4b2912b5330b40d1b0cabf4292664b88/green.json`.
  All93 prior checks retain exact relative order. New
  `ReceivingRun.Guard.UserDisabledStopsLaterDispatch` uses actual Admin selection,
  toggle and Save handlers while capture remains globally enabled. Next preserves
  one completed step and business/activity bytes and becomes Blocked. Own recording
  is restored through the same handlers. Fresh exact owner/evaluator proof and
  both close boundaries remain GREEN. Five instrumented compiles, settings/package
  preservation and normal closure pass. No runtime change was needed for this
  supplemental guard; it is GREEN coverage, not a new product RED.

Final audit through03:41:04.801 UTC finds zero Excel1000/1001/1002 faults and no
query errors; the earlier PowerShell setup failure remains recorded above. Excel
is closed. Cursor access succeeds03:41:06.633 UTC (Oct4 20:41 PDT), with error0.
The four unrelated user-document hashes remain unchanged. No promotion, handoff
or full-slice acceptance is included in this checkpoint.

## Membership lifecycle coverage, 2026-10-05 UTC

Unchanged candidate02 passes **104/104**, retaining all91 preceding checks in
exact relative order. Actual Settings handlers retain unchanged removed-user
flags as unavailable, refuse changes to unavailable users, default a newly
appearing user to enabled, preserve a numeric-looking text identity exactly,
restore a returning user's saved flag and stage/reset all overrides only on Save.
A supplemental direct command rejects a changed unavailable-user override.
Generated roster membership changes inspect only UserId cells; this fixture does
not claim coverage of user-management creation/deletion handlers. Policy commands
preserve Auth bytes, and final cleanup restores both generated fixtures exactly.

Controller `user-tracking-policy-controller/ab9a65de6a9f46b4bd3446fc55d91c66`
runs03:53:30.728--03:58:15.970 UTC. Worker
`slice4be-tracking-settings/bd013edf38dc4a369ac60f11e0fbe8e0/green.json` and the
controller's verification/ordered-comparison receipts record no duplicates,
preserved settings/packages and normal Excel closure. These are supplemental
GREEN checks under the existing contract; no runtime fix or new product RED.

The first disabled-user Receiving audit extension stops after84 passes with one
harness failure: it assumed staging was empty before the ordinary Clear action.
Controller `receiving-replay-controller/4e344209fea84a1f9c9aa78a13c6f817` runs
03:45:17.109--03:52:52.985 UTC and preserves settings/packages with Excel closed.
The initial row's origin is unproved; do not attribute it to Stop. The corrected
test invokes the existing Clear handler, records only staging counts and verifies
empty staging plus unchanged business bytes before adding the audit receipt.
This setup correction is not product RED and changes no runtime behavior.

Corrected audit extension: **101/101**, all94 prior checks in exact relative order,
no duplicates. Controller `receiving-replay-controller/7b97d29b39404453ab103ca3b9da9a43`
runs03:58:29.554--04:07:57.909 UTC; worker
`slice4be-receiving-replay/e3766508e935446c90c2375707781790/green.json`.
The count-only staging receipt finds one initial row with neither EventId nor
System_Key; ordinary Clear leaves zero rows and unchanged business bytes. Actual
Add/Confirm while this actor's recording is disabled creates one exact applied
event and matching inventory audit entry, retains the custom header, changes no
Auth/Config bytes and emits no optional activity. The global capture switch stays
enabled. Settings/packages are preserved and Excel closes normally. No runtime
change is needed. Cursor probe04:08:03.670 UTC succeeds with error0.

Affected published-guide regression on candidate02 is **58/58**, with exactly the
prior58 check identities/order. Controller
`execution-profile-guide-regression/4b7b4eb9bc904cfd91255863290e736b` runs
04:08:05.577--04:14:49.159 UTC; worker
`slice4be-viewer-published-read/eccb3a440e04402dbd42a5fb59e542a0/green.json`.
Five instrumented compiles, immutable versions, draft/cancel/conflict behavior,
layout, role/policy/context guards and source/business preservation pass.
Settings/packages are preserved and Excel closes normally. Cursor access succeeds
04:14:58.218 UTC with error0. This does not close Action Path visible acceptance.

Full Receiving on candidate02 is **854/854**, retaining the exact prior854
identities/order with no duplicates. Controller
`receiving-owner-regression/821e058c51134bd39ec2da9045705bdc` runs
04:15:00.165--04:25:42.174 UTC. Its comparison, closure and50 preserved evidence
files include `evidence/launcher-denial-green.json`. Native keyboard/mouse,
observations/redaction, staging/custom columns, required identity, launcher reuse,
denial and authority/unrelated-workbook preservation pass. Settings/packages are
preserved and Excel closes normally. Fresh Receiving/Returns captures are directly
reviewed; their visible-review receipt does not claim long-value/multiline acceptance.

Final native audit03:45:01--04:26:08.599 UTC finds zero Excel1000/1001/1002 events
and no query errors. Cursor access succeeds04:26:08.874 UTC (Oct4 21:26 PDT), error0;
Excel is closed. All404 PowerShell scripts parse, runtime source has no diff, and
the four unrelated user-document hashes remain unchanged. Candidate02's cold
compile/source/static receipts remain applicable to this test-only increment.
The lifecycle/audit tests are supplemental GREEN coverage, not a new D13 product
RED. Comprehensive observations, Action Path visible proof and broader acceptance
remain open; the older full-chain native failure is not resolved by these results.
