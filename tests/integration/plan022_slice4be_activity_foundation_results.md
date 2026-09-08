# Slice 4be.1 activity foundation candidate

Last verified: 2026-09-07. Architecture v4.11 D18 and Plan 022 remain governing.
This is an implementation checkpoint within 4be.1, not completion of its
comprehensive control catalog or Release 1 acceptance. Candidate packages are
in `deploy/validation-activity`; the accepted `deploy/current` set is unchanged.

## Behavior and ownership

The real Admin Settings Save Value and Production Retrieve UOM Catalog handlers
record separate immutable attempts/results through Core. Core retains authorized
configuration-write ownership and the existing PROD_POST UOM route. The recorder
does not execute or retry commands. A direct service call produces no user-click
evidence. Captured session/target checks reject stale form commands, including
re-authentication as the same user. Unknown columns and denied staging survive.

Activity records use generated GUIDs, verified UTC, fixed catalog messages,
logical owners, stable event codes, truthful data-effect classifications and
SHA-256 integrity. Atomic per-record publication stays in the selected warehouse
Training/Activity library. Reads reject missing/corrupt/cross-warehouse records
without creating files. Idempotent delivery preserves bytes; repeated actions
retain distinct IDs. Policy reads are non-mutating and fail closed for invalid
or partial policy; visibility applies on reads. Optional store failure leaves
the command result visible and adds **Tracking unavailable**.

## Test-first evidence

- Initial packaged activity RED: **19 PASS / 12 FAIL**; see
  [initial evidence](plan022_slice4be_activity_red_results.md).
- Expanded integrity/policy boundary RED: **62 PASS / 5 FAIL**. Valid rehashed
  records incorrectly accepted numeric schema strings, impossible UTC, invalid
  policy values and sequence ordinals without a sequence. Saved policy also
  accepted impossible UTC. These were behavioral failures with valid fixtures.
- Subsequent real-handler session RED: **68 PASS / 2 FAIL**. A Settings form and
  a Production form retained across sign-out/sign-in could still change Config.
- Candidate focused GREEN: **70 PASS / 0 FAIL**, including all 18 accepted D5
  checks, three initial action/result cases, independent hash verification,
  JSON round-trip/rejection, idempotency, context, policy and failure isolation.
- Package metadata RED was reproduced against unchanged `deploy/current` with
  an independent package-XML probe: both required properties are absent there
  and present in all five candidates. The original PowerShell COM property
  probe was replaced because it could not read the properties reliably.
- Two preliminary supplemental runs had PowerShell COM fixture-assignment
  exceptions. Corrected fixtures were rerun against unchanged candidates;
  those exceptions are excluded from behavioral RED.

## Verification status

| Gate | Candidate evidence |
|---|---|
| Cold-start package references | PASS; Operations resolves dependencies inside the candidate directory |
| Explicit VBE compile | 5/5 PASS |
| Focused packaged form/boundary suite | 70/70 GREEN |
| Packaged smoke | 81/81 PASS |
| Live-role | 48/48 PASS on a fresh serial run; first attempt lost Excel COM at Production Complete Run after 39 passes |
| Full Release 1 chain | 30/30 PASS; first attempt lost Excel COM after four passes; fresh serial retry passed without code changes |
| Viewer | PASS; displayed-list export, event labels/reservation exclusion, refresh/date preference, read-only behavior and snapshot bytes preserved |
| Production layout | PASS; three requested sizes across five pages, no control overlap/out-of-bounds, minimize/restore/maximize/restore; all screenshots inspected. The smaller request clamps to the approved 1110x800 minimum, so minimum/default captures match |
| All-role launcher callbacks | 3/3 PASS; no eligible workbook initially open, saved role workbook/form creation and reuse preserved |
| Reusable Production/restart | 2/2 aggregate checks PASS, including worksheet/lifecycle/run scenarios and a fresh Excel restart using the same saved workbook and exact released Recipe |
| Static maintenance | Regenerated; 1,077 candidates, 192 duplicate groups, 45 unresolved dynamic calls, eight literal Application.Run targets; unchanged |
| Existing oversized modules | 28 ratchets; Auth 1,802 lines versus 1,804 baseline, Production 11,700 unchanged; no growth exception |
| Visible operator evidence | Actual Settings capture inspected: configuration saved plus tracking-unavailable warning; no new Settings tab or human acceptance claimed |

Source contract regressions also pass: UOM 7/7, Batch Note/Viewer 5/5,
Production layout 8/8, launcher contracts 24/24, full-chain contract 13/13 and
variable quantity modes 6/6. The layout source check requires explicit
`-RepoRoot .` in this Windows PowerShell environment.

Core session state and role action orchestration are bounded modules. The former
private Auth status setter's four callers now perform its same direct assignment;
no authorization rule changes. Explicit source-fixture imports include the new
session dependency; PowerShell parsing and diff whitespace checks pass. Runtime
JSON, screenshots, fixtures and logs remain ignored and are not release evidence
for an operational NAS rollout.

## Remaining scope

Complete every reachable Operations/Admin control's catalog entry and eligible
handler/outcome coverage, publication and source references. Implement the
dedicated Event Tracking Settings policy/profile editor, personal preferences,
comprehensive Viewer, recorded sequences, versioned How-To authoring and
Diagnostic/Compare both. Required canonical audits remain effective. Remaining
policy compatibility, recording, publication, restart and visible acceptance
gates cannot be inferred from these first two controls.

## Reproduction

Run Excel validators serially with Excel closed:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/build-xlam.ps1 -OutputRoot deploy/validation-activity -Apply
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-PackagedVbaCompile.ps1 -DeployRoot deploy/validation-activity
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/validation-activity -CheckActivityEvidence -CheckActivityFoundation -Phase GREEN -CaptureEvidence
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_phase6_packaged_xlams.ps1 -DeployRoot deploy/validation-activity
```

The RED reports describe earlier candidate revisions. Passing current packages
with `-Phase RED` does not reproduce an earlier behavioral failure. Local ignored
reports `foundation-red.json`, `session-red.json` and `green.json` under
`reports/runtime/slice4be-activity` preserve check names/booleans only.

## Candidate package hashes

SHA-256 of the candidate used for the focused, compile and regression checks:

| Package | SHA-256 |
|---|---|
| invSys.Core.xlam | 87d58d189f5d175565fc75b9d5bccc1924fdb006b149796de33ab3faf768206f |
| invSys.Inventory.Domain.xlam | 2e1fc957ced4ef87d654a1fbfcc13819360a2f02e6cfa68f2ae7ac6fb4d06394 |
| invSys.Designs.Domain.xlam | e8f9afb2c74eeeaa2f59c8a1a26a45fd858314a2889b862e38f42334f915c80e |
| invSys.Operations.xlam | 2dbf7906c6a23f3c1f75a09db208793cbee64531f03b712c613a0d76bb744433 |
| invSys.Admin.xlam | ba61084bcf1b59022bdeba419c21c4ee4abb891cc230a89d5d2c7e316e644a11 |
