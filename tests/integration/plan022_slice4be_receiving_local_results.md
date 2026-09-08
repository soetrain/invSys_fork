# Slice 4be.1 Receiving Refresh/Clear activity

Architecture v4.11 D18 governs these discovered shared Receiving/Returns
controls. Initial catalog-4 focused GREEN is recorded below; freshness correction
and final changed-package gates are pending.
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
