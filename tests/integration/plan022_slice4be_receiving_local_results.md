# Slice 4be.1 Receiving Refresh/Clear activity

Architecture v4.11 D18 governs these discovered shared Receiving/Returns
controls. The catalog-4 Refresh/Clear/freshness candidate checkpoint is GREEN;
comprehensive 4be coverage, user comparison and Release 1 acceptance remain open.
The existing Add/Confirm checkpoint remains GREEN; see its
[evidence and five baseline hashes](plan022_slice4be_receiving_staging_results.md).

## Test-first entry, 2026-09-07

Runtime code **54ec2cb** and its unchanged `deploy/validation-activity` packages
are the baseline. Tests enter the actual `mBtnRefresh_Click` and
`mBtnClear_Click` handlers through unsaved fixture-only access seams. All
warehouses and operator workbooks are generated disposable fixtures. Direct
services/internal refreshes are supplemental negative-attribution checks.

The first run is **322 PASS / 74 FAIL**. The previous 262 checks remain GREEN;
there is no harness exception. Sixty failures concern missing attempt/result
records, two expose false refresh success, six expose stale owner execution,
and four expose absent tracking-failure feedback. Two combined header/binding
assertions in the moved/protected-table fixture were excluded from intended
behavioral RED pending diagnosis.

Independent guards establish local Refresh/staging preservation, actual Clear
effects, authority workbook byte preservation and direct/internal negative
attribution. The partial Clear fixture moves the aggregate table to a separate
sheet and protects it: the actual owner clears the first table before Excel
rejects the second deletion. This is not a clean rollback. A Core read-model
fault seam returns False with a fixed cause; the real Refresh handler currently
discards that report and displays success. Both RED form captures were inspected.

The expanded run separates current-table header checks from cached Excel
references, verifies the moved fixture's header before the action, adds valid
catalog-3 policy checks and tests a closed captured workbook with another active.
The expanded run is **365 PASS / 74 FAIL**, with no harness exception and all
previous 262 checks GREEN. Current-table header and unrelated-workbook checks
pass in both partial-failure cases on unchanged runtime; the first combined
cached-reference checks do not establish runtime header loss. The expanded
failures are the 72 intended cases above plus two missing closed-workbook
rejection messages. Both closed actions already preserve the other workbook and
authority bytes. The five baseline hashes are unchanged; Excel closed normally.
No runtime implementation has been changed.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/validation-activity -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CheckReceivingLocalActivity -CaptureEvidence
```

Ignored evidence: `reports/runtime/slice4be-receiving-activity/local-first-red.json`
and `local-red.json` contain names/booleans only. Raw records, source values and captures remain
ignored. Architecture D18, Plan 022 and controls v1.68 record the catalog-4 target
before implementation. No new business permission or canonical writer is added.

## Initial candidate and discovered freshness gap

The initial catalog-4 implementation passes **439/439** through the actual form
handlers in `deploy/validation-receiving-local`. Five package builds, cold start
and explicit VBA compiles pass. This is focused evidence, not slice completion.
The ignored `local-first-green.json`, `local-first-green-package-hashes.json`
and `local-initial-runtime.patch` preserve this checkpoint for the next RED.

Source inspection found that `modOperatorReadModel.RefreshInventoryReadModelForWorkbook`
returns True for retained cached inventory and a loaded stale fallback. The
initial action controller labels either result REFRESHED and overwrites the
owner's report. This would misstate evidence used by Diagnostic conclusions.
The additional `Slice4beReceivingFreshness.ps1` tests remove only a generated
fixture's canonical snapshot from discovery, or expose it as a fallback, and
invoke the actual Refresh handler on both tabs. Independent guards check owner
freshness metadata, Boolean compatibility, exact identity/quantity/condition,
unknown values, staging, authority bytes and unrelated-workbook isolation.
Expected behavioral RED is lost visible stale cause and incorrect REFRESHED
activity. No runtime freshness implementation precedes that RED.

Expanded freshness RED is **463 PASS / 28 FAIL** (491 total), with no harness
exception and all previous 439 checks passing. Each of the four cases fails
the visible-cause check and six correlated outcome checks because STALE is
missing; these cascading assertions do not establish a payload leak or identity
loss. All independent owner-state, data-preservation and Boolean compatibility
guards pass. The cached Receiving form capture was inspected and visibly claims
refresh success. The five pre-fix package hashes remain unchanged.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/validation-receiving-local -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CheckReceivingLocalActivity -CaptureEvidence
```

Architecture, Plan and controls v1.69 were synchronized and pushed as docs
`bd15052` before freshness implementation. The ignored report is
`reports/runtime/slice4be-receiving-activity/freshness-red.json`.

Initial static evidence has 171 components / 5,465 procedures, 1,077 candidates,
192 duplicate groups, 45 unresolved dynamic calls and eight literal targets.
Candidate/duplicate/dynamic counts and all 28 oversized-module limits do not
regress. A supplementary historical `Test-Slice12ReviewedCleanup.ps1` run with
explicit RepoRoot returns 7 PASS / 6 FAIL on both this candidate and the unchanged
`f813f2f` comparison checkout: its older roots/count/growth assertions are not
current GREEN evidence. The default RepoRoot invocation fails in parameter
binding; neither harness/legacy result is D13 behavioral RED for this change.

## Freshness candidate GREEN

Code **3b06425** preserves the initial implementation plus the expanded failing
test before the freshness correction. The new candidate is built independently
in `deploy/validation-receiving-freshness`; accepted deployment is unchanged.
The existing Core owner returns explicit state through optional primitive
ByRef parameters, retaining old Boolean callers. Its three cached-result
branches share local metadata/cleanup handling; snapshot selection is unchanged.
The Receiving controller records STALE/Warning/Changed and retains the owner's
visible cause. Missing state cannot assert REFRESHED or erase an existing cause.

All five builds, cold start and explicit package compiles pass. Source regressions
pass: behavior locks 13, Receiving stabilization 10, persistence feedback 4,
Disposition 6, Receiving 4o 5, Receiving 4p 8 and ListBox 7. Static regeneration
retains all counts above; all 28 current oversized-module ratchets hold, and
`modOperatorReadModel` shrinks from 1,919 to 1,916 lines. The expanded packaged
suite is **491/491 GREEN**, with no harness exception. Checkpoint release gates
pass as recorded below. All four cached/fallback form captures were inspected: both tabs
retain the real cached/stale explanation with staging still displayed. Successful
Clear and tracking-unavailable Refresh captures were also inspected. All 22
current captures were reviewed: the 16 tab/action/local-outcome cases, two
closed-workbook cases and four cached/fallback cases. Messages are visible and
the form retains its existing staging/detail surfaces. These are automated
operator evidence, not human acceptance.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/validation-receiving-freshness -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CheckReceivingLocalActivity -CaptureEvidence
```

The exact candidate hashes below identify the five packages used for GREEN and
the subsequent release gates. They are not an accepted deployment manifest.

| Package | SHA-256 |
|---|---|
| Core | `82b021ea4c454ece0189f4df76456b541d588d86778c539bb0d9f97beced341f` |
| Inventory.Domain | `6c2cb6f3f59ada94b5a107e2a0daa4cdd6f594d3e1f3568ee44d695310cbecf8` |
| Designs.Domain | `0d8cab0f1f40511869ebe6a6ae344147851c05dd029678f832c87f0850fc4180` |
| Operations | `8587fd5f29953652e8592c3eff0e4afe39ce803ecff23bf2f2334206f3c8b2ed` |
| Admin | `33b209633909c32faa2b6ca4b53a91b3b6889d2be505d498aaa6945eab6428de` |

## Package-gate isolation correction

The initial packaged gate reported 81/81 but left its first Excel session at a
Designs save prompt; the later layout preflight correctly refused to start.
The validator had neither bound its runtime root nor tracked every workbook
opened by smoke tests. Its old tracked-list cleanup therefore missed that
workbook. The identified test-session save was declined, without saving or
terminating an unidentified process, and Excel exited. The pre-existing saved
Designs file's last-write time predates this run; there was no pre-run byte
baseline for that initial gate.

The harness now binds and verifies Core's temporary root before role work in
each Excel session, provisions the required Config fixture through the existing
explicit Core setup boundary, and closes all owned workbooks. It refuses a
workbook outside fixture/package roots. Process cleanup uses the actual Excel
HWND and requires an available Workbooks collection with Count=0; null COM
properties are never treated as proof of emptiness.

The first isolated run returned **84 PASS / 1 FAIL**: Admin Settings exposed the
previous external Config dependency. Both root and cleanup checks passed. After
explicit fixture provisioning, the corrected gate is **86/86**, retaining all
81 prior checks plus five fixture/root/cleanup checks. No Excel process remains,
and the saved default Designs file is byte-for-byte unchanged across both
corrected runs. This is harness isolation/fixture evidence, not D13 product RED;
no runtime contract or XLAM was changed for it. The failed isolated report is
retained under ignored `reports/runtime/slice4be-receiving-activity/`.

Live-role workflows pass **48/48**, the ordered full chain/restart passes
**30/30**, and Viewer passes its packaged refresh/filter/export/read-only checks.
Production layout passes across three sizes/five inspected pages and native
window transitions; all three layout captures were inspected. Packaged launchers
pass **3/3**. Dedicated reusable Production passes **2/2**, including a clean Excel
restart, reusable design/run actions, workbench/edit/export/import/output-picker,
quantity regulation and Chai fork/convergence coverage. No focused-only switches
were used. Its terminal report contains two passing rows, no failure row/RPC
failure, and the final layout/launcher interval has no Excel Application Error
event. Excel is closed and all five candidate and five pre-fix package hashes
remain unchanged after verification.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_phase6_packaged_xlams.ps1 -DeployRoot deploy/validation-receiving-freshness
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_phase6_live_role_workflows.ps1 -DeployRoot deploy/validation-receiving-freshness
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_release1_full_chain.ps1 -DeployRoot deploy/validation-receiving-freshness
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_inventory_viewer.ps1 -DeployRoot deploy/validation-receiving-freshness
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_slice9_production_layout.ps1 -RepoRoot . -DeployRoot deploy/validation-receiving-freshness -OutputDirectory reports/runtime/slice4be-local-layout -ResultPath reports/runtime/slice4be-local-layout/layout-results.md
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_plan022_packaged_launchers.ps1 -DeployRoot deploy/validation-receiving-freshness -OutputDirectory reports/runtime/slice4be-local-launchers -WorkbookState NoEligible
powershell -NoProfile -ExecutionPolicy Bypass -File tools/validate_plan022_packaged_launchers.ps1 -DeployRoot deploy/validation-receiving-freshness -OutputDirectory reports/runtime/slice4be-local-production -CallbackFilter Production -WorkbookState ProductionReusable
```

The code candidate is not a NAS rollout. This checkpoint does not complete
4be.1's remaining control coverage, publication, Event Tracking Settings,
comprehensive Viewer, recorded conclusions, guide management or user comparison.
