# Slice 4be.1 Receiving Open/Close activity

Architecture v4.11 D18's Receiving Open/Close clarification governs this
checkpoint. The normative specification, Plan 022 and controls v1.71 were
synchronized and pushed as docs **b22b3bd** before runtime implementation.
This inherits the approved comprehensive coverage, owner-outcome and captured
context rules. The full Slice 4be and Release 1 Goal remain incomplete.

## Test-first entry, 2026-09-07

Runtime source **130df29** and `deploy/validation-receiving-freshness` are the
unchanged catalog-4 baseline. The five exact hashes and completed previous gates
are in [Refresh/Clear evidence](plan022_slice4be_receiving_local_results.md).

`Test-Slice4beConfigCommands.ps1 -CheckReceivingLifecycleActivity` adds actual
`modRibbonGenerated.RibbonOnActionOperations` dispatch using a fixture class
implementing Office.IRibbonControl, with `btnOperationsReceivingForm` as its ID.
The generated callback and its authorization guard are not replaced. Unsaved
form access seams invoke `mBtnClose_Click` and `UserForm_QueryClose` with
vbFormControlMenu; the latter performs the existing close handler followed by
Unload when allowed. This proves the actual handler contract, not native mouse
or title-bar dispatch. Direct macro calls and internal Unload are separate
negative-attribution cases.

The existing launcher alone creates/reopens/reuses its station-local workbook.
A Core bridge fault seam counts calls and can return False with a fixed cause
before owner provisioning. The existing Receiving message presenter is
intercepted only in the unsaved fixture to inspect its exact cause without
modal dialogs or raw report output. Activity bodies and source/credential values
remain private to disposable fixtures and ignored runtime evidence.

The first completed lifecycle run is **515 PASS / 37 FAIL**, with all **491
prior GREEN checks preserved** by check-name comparison and no harness exception.
Thirty-six failures describe missing correlated Open/Close observations; one
proves reuse of a form captured before sign-out/reauthentication. The unchanged
launcher only checks workbook name and visibility. The existing form's workbook
resolver already checks exact live object identity.

Independent guards pass for initial visible/captured launch, ordinary reuse,
button/window dismissal, stale-context dismissal without new-session attribution,
direct-macro/internal-unload exclusion, retained launcher failure cause and one
owner invocation, authority bytes, unknown columns and unrelated workbook values.
Excel exits on its own after cleanup. All five baseline package hashes match.

An earlier run stops before Receiving at fixture sign-in: **75 PASS / 1 harness
exception**, credential-rejected status. It is not product RED. The fixture now
writes credential fields as text and checks exact in-memory round trips without
emitting their values. The rejected run's underlying cause was not captured;
do not claim a proven authentication defect or a proven Excel-conversion cause.
The subsequent complete run establishes meaningful RED independently.

## Expanded protection

The expanded test additionally checks a valid catalog-4 policy does not enable
Open/Close, retains staged System_Key identities and unknown values, exercises
optional store failure on both dismissal routes, and reopens the same workbook
filename after closing its original object. The latter must bind the new live
object rather than reuse a form holding the old one. Visible captures cover
initial open, repeated launch and launch after reauthentication.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/validation-receiving-freshness -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CheckReceivingLocalActivity -CheckReceivingLifecycleActivity -CaptureEvidence
```

The expanded run is **531 PASS / 47 FAIL** across **578 checks**, preserving all
491 earlier GREEN checks by name with no harness exception. Forty-two failures
describe the absent observation pairs, one retains the stale-session reuse RED,
and four expose missing tracking-failure notices. Both close routes still dismiss
under a blocked activity store, and opening invokes its owner once. Catalog-4
policy compatibility, staged identities/unknown values, source authority and the
unrelated workbook are preserved. Same-name workbook reopening already binds
the correct live object on the unchanged package; this is a regression guard,
not a second confirmed binding defect.

The three full-size lifecycle captures were inspected. They show initial open,
ordinary reuse and reuse after reauthentication; the latter remains the old
form, independently confirmed by its captured-context check. These captures are
automated visible evidence, not human acceptance or native close-button evidence.

Static evidence is regenerated: **171 components / 5,465 procedures / 1,077
candidates / 192 duplicate groups / 45 unresolved dynamic calls / eight literal
targets**. All **28** existing oversized-module limits hold against HEAD. The
source scanner additionally links the new protecting test to existing handlers.

Runtime source, generated Ribbon route and deployed packages have not changed.
All five baseline hashes remain unchanged and Excel is closed. GREEN,
compile/build, packaged, layout, static, live-role, full-chain/restart and visible
acceptance gates remain required after implementation; previous checkpoint
evidence does not prove this feature.

Ignored evidence is under `reports/runtime/slice4be-receiving-activity/`:
`lifecycle-first-harness.json`, `lifecycle-first-red.json`, and the eventual
`lifecycle-red.json`. No generated authority workbook or raw activity is committed.
