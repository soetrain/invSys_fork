# Slice 4be.1 Receiving Open/Close activity

Architecture v4.11 D18's Receiving Open/Close clarification governs this
checkpoint. The normative specification, Plan 022 and controls v1.71 were
synchronized and pushed as docs **b22b3bd** before runtime implementation.
This inherits the approved comprehensive coverage, owner-outcome and captured
context rules. The full Slice 4be and Release 1 Goal remain incomplete.

**Current checkpoint, 2026-09-08:** Catalog-5 Receiving Open/Close is technically
GREEN at 596/596 with all required checkpoint gates below passing. The final
candidate is `deploy/validation-receiving-lifecycle-dismissal`. Native close
records committed UI dismissal without premature reference release or delayed
next-launch evidence. Human acceptance and comprehensive 4be coverage remain open.

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

At RED, runtime source, generated Ribbon route and deployed packages were
unchanged, with all five baseline hashes preserved and Excel closed.

Ignored evidence is under `reports/runtime/slice4be-receiving-activity/`:
`lifecycle-first-harness.json`, `lifecycle-first-red.json`, and the eventual
`lifecycle-red.json`. No generated authority workbook or raw activity is committed.

## Catalog-5 implementation and first GREEN, 2026-09-08

The first five-package candidate, `deploy/validation-receiving-lifecycle`, passes
**578/578** with all prior 491 checks retained, five explicit compiles and cold
start. Catalog 5 adds Open/Close and reads older versions unchanged. The generated
Receiving Ribbon dispatch calls `modTS_Received.ShowReceivingForm True`; its
existing capability guard remains. Other callers default to non-observed
compatibility behavior. The existing launcher remains the workbook owner and
reports OPENED/REUSED/FAILED; reuse now requires the form's captured session and
exact live workbook binding. Existing failure causes remain visible.

The initial Close-button controller recorded CLOSED after Unload returned and
routed menu close through that button while cancelling the original message.
The additional native test below exposes why that was insufficient. Tracking
notices use the existing Receiving message surface without retrying the owner;
no canonical business writer or new permission was introduced.

After GREEN, the pure header-caption formatter moves from the form to
`modReceivingFormWindow`; the existing message-display body moves to the
Receiving action controller behind its compatibility wrapper. Form size reduces
from 1,235 to 1,221 lines, and `modTS_Received` stays at 1,575. Existing source
checks pass: behavior locks 13, stabilization 10, persistence feedback 4,
disposition 6, Receiving 4o 5, Receiving 4p 8, ListBox 7 and launcher contracts 24.

`deploy/validation-receiving-lifecycle-final` is independently built with all five
compiles and cold start passing. Its expanded run additionally sends WM_CLOSE to
the Receiving window after verifying the window belongs to this isolated Excel
process. Normal and blocked-store cases must destroy the native window and
retain the same activity/dismissal guarantees. This supplements the existing
QueryClose seam; it does not replace the recorded RED or infer a physical click.
Expanded GREEN and remaining packaged/layout/static/live-role/full-chain/restart
gates are pending. This is not a completed checkpoint or Release 1 acceptance.

Ignored first-candidate evidence: `lifecycle-first-green.json`,
`lifecycle-first-package-hashes.json`, and `lifecycle-first-runtime.patch` in the
same private runtime report directory. The earlier catalog-4 baseline is retained.

## Native-window defect and correction

The first native supplement stops at a caption-lookup ownership guard: **542
PASS / 1 harness failure**, before sending WM_CLOSE. The corrected fixture reads
the exact launcher's form handle through its existing window owner and retains
process/visibility checks. This exposes actual product RED on the unchanged
`validation-receiving-lifecycle-final` candidate: **592 PASS / 4 FAIL**, retaining
all 578 first-candidate checks. Normal and blocked-store native close both leave
the window/form open while emitting CLOSED. A nested Unload return is therefore
not sufficient owner evidence when the original native close is cancelled.

D18's owner-fact clarification now explicitly requires committed dismissal and
permits termination to finish an existing explicit-native-close observation,
without starting another action. The correction allows the original native close,
captures its request in QueryClose, releases the launcher's cached form reference,
and finishes its pending observation in Terminate. The button keeps its direct
controller path. Programmatic unload has no native pending action and adds no
record. The corrected candidate is `deploy/validation-receiving-lifecycle-native`;
its GREEN and release gates remain pending.

For focused diagnosis, `-ReceivingLifecycleOnly` skips the older Receiving action
cases and writes `lifecycle-only-green.json` separately. It does not replace the
full 596-check suite or release regressions. Ignored evidence includes
`lifecycle-native-window-harness.json`, `lifecycle-native-red.json`,
`lifecycle-native-red-runtime.patch` and `lifecycle-final-package-hashes.json`.

The corrected native candidate passes all five explicit compiles and cold start,
but its first lifecycle-only run stops at **131 PASS / 1 harness exception**.
`Lifecycle.NativeWindow.Destroyed` passes, then the next Excel macro call loses
its COM connection. Windows records a VBE7 access violation during that close.
This is neither focused GREEN nor proof of the crash's root cause. The exact
report is preserved as ignored `lifecycle-native-crash.json`; the unchanged
catalog-4 comparison and subsequent diagnosis are recorded below.

Excel's automatic recovery instance has a real empty Workbooks collection before
Quit. Its subsequent Document Recovery prompt is resolved with **Yes, I want to
view these files later**, preserving recovery files. Null COM properties after
Quit are not treated as an empty-workbook proof. Excel exits normally; no process
is killed. The baseline comparison starts only after the Excel-closed guard passes.

Current static regeneration reports **171 components / 5,471 procedures / 1,077
candidates / 192 duplicate groups / 45 unresolved dynamic calls / eight literal
targets**, with all 28 module-size limits preserved. Relevant source checks and
PowerShell parser checks pass. Full candidate GREEN and release gates remain open.

The unchanged catalog-4 comparison completes at **129 PASS / 55 FAIL**, 184
checks, without a harness exception. Both actual native-window destructions and
subsequent macro calls pass. The failures concern absent lifecycle observations,
stale-session reuse and tracking notices, as expected on that baseline. Preserved
report: `lifecycle-native-catalog4-comparison.json`. This narrows the crash to the
changed candidate/environment interaction; it does not identify one statement.

The harness now offers two explicitly diagnostic, unsaved mutations under
`-ReceivingLifecycleOnly -Phase RED -LifecycleDiagnostic`: `SkipTerminationEvidence`
and `KeepLauncherReference`. They isolate termination publication from immediate
reference release without editing or saving a candidate XLAM. Separate diagnostic
reports cannot replace acceptance GREEN or the full suite. The first diagnostic
suppresses publication but still crashes at native close: **125 PASS / 7 FAIL**,
including six expected Window missing-record checks and one harness exception.
Windows records the same VBE7 access-violation signature. The second retains
the reference and completes without crashing: **176 PASS / 8 FAIL**. Native
dismissal succeeds, but its completion appears only on the next launch, and the
blocked-store close notice is late. Neither diagnostic is acceptance GREEN.

## Synchronous native UI dismissal correction

D18 defines CLOSED as UI dismissal, not a memory-lifetime event. Before the next
runtime edit, Architecture, Plan 022 and the controls clarify that native close
may synchronously hide the form, record that committed dismissal, invalidate
reuse and permit native teardown. It may not leave a hidden reusable form or
defer the observation/notice until another launch. This constrains the existing
owner-fact rule without adding a permission, business effect or authority path.

`CloseForm` now hides for the explicit native path and uses Unload for the button.
QueryClose leaves Cancel clear and marks the existing launcher invalidation flag;
it does not release the cache inside its active native event frame. Terminate
only releases the captured workbook. The same native tests still require actual
window destruction, immediate observation/notice, no next-launch activity, and
staging/authority preservation. Candidate
`deploy/validation-receiving-lifecycle-dismissal` passes its five-package build,
five explicit compiles and cold start. The full suite completes **596/596 GREEN**,
retaining every earlier 491-check and 578-check baseline identity with no duplicate
check names. Both actual native close cases destroy their windows and provide
immediate observations/notices; subsequent launch adds no delayed close activity.
Captured bindings, staged keys/unknown values and authority bytes remain protected.
Excel exits normally without a new crash. Three full-size Open/Reuse/SessionChanged
captures are inspected; these are automated evidence, not human acceptance.

The protecting command is the expanded RED command above with
`-DeployRoot deploy/validation-receiving-lifecycle-dismissal -Phase GREEN`, all
activity switches retained, and no lifecycle-only or diagnostic mode. Ignored
`lifecycle-dismissal-green.json` preserves the result and
`lifecycle-dismissal-package-hashes.json` identifies all five exact XLAMs.
Static metrics remain 171/5,471/1,077/192/45/eight; all 28 size limits hold.
Receiving form size is now 1,218 lines; the launcher remains 1,575. Relevant
source regressions pass. Broader checkpoint gates are in progress; the checkpoint
and full Goal are not complete. No accepted deployment is replaced.

Current candidate release gates: packaged XLAM validation **86/86** and live-role
regression **48/48** pass. Ignored copies are
`lifecycle-dismissal-packaged.md` and `lifecycle-dismissal-live-role.md`.
Full chain passes **30/30**, including fresh-Excel reconciliation and no new
retired static paths. Viewer regression passes. Production layout passes all
three sizes across five pages and native minimize/restore/maximize/restore;
three captures are inspected. Public NoEligible launchers pass **3/3**.
The first full reusable Production run stops during released Process edit/export
with **0 PASS / 1 HARNESS RED**, an RPC exception. Windows reports an ntdll.dll
fault (0xc0000028), distinct from the earlier Receiving VBE7 failures. This does
not establish a Production behavioral regression or a completed gate. Handoff
076 records earlier ntdll/RPC failures on both older/candidate runtimes with cause
unproven. The failed report is preserved as ignored
`lifecycle-dismissal-production-first-harness.md`. The full unchanged catalog-4
comparison completes **2/2**, including clean restart, in
`reports/runtime/slice4be-lifecycle-production-baseline`. A full candidate retry
now runs in `reports/runtime/slice4be-lifecycle-production-retry` after verifying
all five candidate hashes unchanged. No reduced Production flags or speculative
runtime correction are used. The candidate gate remains incomplete until that
run terminates; the first failure is not erased or treated as product RED.

That unchanged candidate retry also stops with **0 PASS / 1 HARNESS RED**, this
time during Run List layout, with the same ntdll fault signature. Two different
stages do not establish one failing Production operation. Both reports remain in
their separate directories. Further identical retries are replaced by a scoped
observer change and controlled harness tests; Production runtime stays unchanged.

`Test-Plan022DialogObserver.ps1` exercises the real observer with two disposable
native-dialog processes. The old observer first passes six compatibility checks;
adding an explicit OK/Cancel confirmation guard gives **6 PASS / 1 FAIL** because
it accepts that confirmation. This is meaningful harness RED, not Production RED.
The native replacement uses only owned #32770 dialogs, records fixed fixture
messages, and dismisses a sole informational OK. It does not traverse ordinary
form accessibility trees, inject Enter, accept a multiple-button confirmation,
or act in the other process. Cooperative stop and callback error capture remain.
An initial native 5/2 run reveals Windows uses IDCANCEL for the sole OK; accepting
IDOK/IDCANCEL only with the single-OK guard produces **7/7 GREEN**. Exact message,
ordinary-command, confirmation, other-process and shutdown checks all pass.

The helper is `tools/plan022-dialog-observer.ps1`, used by the existing launcher
validator. No packaged callback, Production assertion or XLAM changes. Source
launcher checks pass 24/24, restart checks 6/6, tooling contracts 62/62, and parser
checks pass. Native-observer public launchers pass **3/3**, and the unchanged
candidate's full Production workflow plus clean restart passes **2/2**. Removing
UI Automation is a diagnostic change with an independently proven confirmation
safeguard; this successful run does not establish the ntdll cause. Both prior
failed candidate runs and the successful unchanged catalog-4 comparison remain
part of the record.

## Final checkpoint gates

All commands use `deploy/validation-receiving-lifecycle-dismissal` where applicable.
No reduced Production flags are used.

| Gate | Verified result |
|---|---|
| Five-package build; explicit VBE compile; cold-start dependencies | PASS for all five projects |
| Full packaged activity/form/Ribbon suite | 596/596; every earlier 491/578 check retained by name |
| `validate_phase6_packaged_xlams.ps1` | 86/86 |
| `validate_phase6_live_role_workflows.ps1` | 48/48 |
| `validate_release1_full_chain.ps1` | 30/30, including new Excel reconciliation |
| `validate_inventory_viewer.ps1` | PASS; read-only events, filters, refresh and export retained |
| `validate_slice9_production_layout.ps1` | Three sizes/five pages and native transitions PASS; captures inspected |
| Native-observer public NoEligible launchers | 3/3 |
| Native-observer full ProductionReusable | 2/2, including workbench/edit/export/import, lifecycle/run, Chai fork/convergence and clean restart |
| Native-dialog observer | 6/1 meaningful harness RED -> 7/7 GREEN |
| Source and tooling checks | Receiving checks listed above; launcher 24/24, restart 6/6, tooling 62/62; parsers PASS |
| Static maintenance | 171 components, 5,471 procedures, 1,077 candidates, 192 duplicate groups, 45 unresolved dynamic calls, eight literal targets; all 28 size limits hold |

Final launcher evidence is under ignored
`reports/runtime/slice4be-lifecycle-native-observer-launchers` and
`reports/runtime/slice4be-lifecycle-native-observer-production`. Layout evidence
is under `reports/runtime/slice4be-lifecycle-layout`. The final Production interval
has no Excel Application Error, Excel is closed, and all five candidate and all
five accepted catalog-4 hashes remain unchanged. No accepted deployment/NAS or
operational workbook is replaced. Unrelated handoff 067 and critique 023 remain
untouched.

| Exact candidate package | SHA-256 |
|---|---|
| invSys.Admin.xlam | 0f5e6389d11e74351bbcf42eaef7bee52d4112f3be481a9083985ac223298ef0 |
| invSys.Core.xlam | f3f395cfbe8b381bed9cbd4a00351cd30c44f67374438f2f6fc0d0d0c4aca7a4 |
| invSys.Designs.Domain.xlam | 3ac99b178622b89b69a71b6a680b46837d596f2e692ffa3725dcc6040c9a732f |
| invSys.Inventory.Domain.xlam | 6ba3a95634e8d1661619641805c591c5df04636839209c13259af14aaabe049d |
| invSys.Operations.xlam | 937cd5faed9b6b3eda39d02d045cb3dd905138bf06b2209acc2c673783d61ce2 |

This completes the technical Open/Close checkpoint, not 4be.1 or Release 1 UAT.
Optional Receiving navigation/selection, worksheet reachability and launcher
denial observation remain pending, followed by comprehensive other Operations/
Admin coverage, publication, Event Tracking Settings/profiles/preferences,
comprehensive Viewer, recording/conclusions, guide management and both comparison
presentations. Physical multi-station/NAS and human acceptance remain required.
