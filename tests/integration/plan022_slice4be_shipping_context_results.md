# Plan 022 Slice 4be.1 Shipping captured-context repair

Last verified: 2026-09-12. This is a candidate under approved Architecture v4.11
D18, not complete Shipping activity, Slice4be or Release1 acceptance. Normative
clarification **089ab0d** was committed/pushed before runtime edits, with Plan022
and controls v1.92 synchronized. The clarification inherits captured-context and
workbook-preservation rules; it adds no permission, activity catalog or authority.

Continuation: [native validation and rebuilt-set evidence](plan022_slice4be_native_validation_results.md)
records the unchanged-source comparison, individual gates and unresolved native
failures. It does not replace this candidate's original evidence or claim full
Slice4be acceptance.

## Protecting test and RED

`Test-Slice4beConfigCommands.ps1 -CheckActivityEvidence -CheckActivityFoundation
-CheckShippingActivity` retains the real-owner eight-action sequence and its exact
key/source/application checks. `Slice4beShippingContext.ps1` then prepares genuine
active/held lines. A separate probe mode counts entry at existing owners and stops
there: each healthy handler must reach its owner once, while stale handlers must
stop earlier with a context notice. Probe results do not prove business effects;
the preserved normal sequence supplies that evidence. State and source-byte checks
verify the probe does not mutate business data.

- Initial matrix: **136 PASS / 50 FAIL**, including a harness failure. A picker
  threshold demanded five available units for a two-unit Add; this is not matrix RED.
- Corrected matrix: **242 PASS / 91 FAIL**. All seven healthy calibrations and all
  eight timing-status checks pass. Each of21 signed-out/reauthenticated/changed-
  target cases fails pre-owner rejection and its notice. Earlier49 failures remain.
- Explicit public relaunch: **246 PASS / 92 FAIL**. Valid reuse passes, but the
  stale form is also reused. Workbook/staging preservation and next-launch reuse pass.
- Candidate: **292 PASS / 46 FAIL**. All46 formerly failing context/relaunch
  assertions pass; no non-activity failure, lost previous GREEN check, missing check
  or duplicate exists. All46 remaining failures are missing Shipping activity.

The successful guard comparison is focused GREEN, not a passing full activity
suite. The suite retains exit1 and every unresolved activity assertion.

Ignored evidence root: `reports/runtime/slice4be-shipping-activity/`:

| Log | JSON run/report | Result |
|---|---|---|
| `context-matrix-red.log` | `9c82a4aedd264cffb5616e185e39d94f/red.json` | 136/50 harness failure |
| `context-matrix-calibrated-red.log` | `36747aa3361a4a39954ec43fe7421ca0/red.json` | 242/91 RED |
| `context-reopen-red.log` | `ce9fdc23667e403095d4857e1b85a3c2/red.json` | 246/92 RED |
| `context-guard-green.log` | `35cea1eb29e14fb59ba2a3e436bed660/green.json` | 292/46; guard GREEN |

## Candidate implementation and gates

`frmShipmentsTally` captures its trusted context once and guards all seven mutation
controls before owner entry, including after pending-status UI yields and between
multi-row Remove owner calls. Rejection stops automatic synchronization and shows
the fixed notice. `modShippingFormContext` validates the captured live workbook
and session; the explicit launcher replaces stale forms through its bounded factory,
preserving active/held staging on that same workbook. Valid reuse and ordinary Close
remain. Core authorization and business owner contracts are unchanged.

`cShippingActionTimer` receives the form's existing timing behavior, protecting the
form size limit while retaining all eight normal-action timing status checks.
The isolated candidate is `deploy/validation-shipping-context-guard`. All five
packages build; all five explicit compiles and cold-start Operations dependency
checks pass. Build/compile logs, exported source report and five SHA-256 pins are
`context-guard-build.log`, `context-guard-compile.log`,
`context-guard-compiled-source.json` and `context-guard-package-hashes.json`.

Static JSON contracts pass at177 components/5511 procedures/123758 lines. Literal
Application.Run8, unresolved dynamic calls45 and duplicate-body candidates195
are unchanged. All28 existing module limits hold: the form shrinks to3022 lines,
the launcher module to22443. The net73 runtime lines add the context/lifetime
helper and typed timing class; no size-limit exception is taken.

Candidate gates completed on the same pinned set:

- Packaged smoke **86/86** (`context-guard-smoke.log`).
- Live-role workflows **48/48** (`context-guard-live-role.log`).
- Ordered full Release1 chain **30/30**, including restart/reconciliation and
  source/static gates (`context-guard-full-chain.log`).
- Viewer launch/reuse/export/filter/read-only contract passes (`context-guard-viewer.log`).
- Combined public launchers **3/3** (`context-guard-launchers/packaged-launcher-noeligible.md`).
- Shipping layout/identity **1/1**, including fixed status anchor, grow/resize,
  headers, exact key, and no Boxing page overlap (`context-guard-shipping-layout/shipping-layout.md`).
- Full reusable Production and clean-Excel restart **2/2**, with no reduced test
  flags (`context-guard-production/production-reusable-production.md`).

Windows records Excel `c0000005` in `combase.dll` at2026-09-12T23:58:10Z, within
the full-chain window23:55:31Z-00:00:14Z. The30 passing business checks do not prove
native stability or repair the earlier unexplained native faults. The sanitized
module/code/time record is `context-guard-native-faults.json`; no causal claim is made.

The visible rerun also records292 PASS / 46 missing-activity FAIL, preserving
all338 check identities and all292 prior GREEN checks with no duplicates or other
failure. Its report is `0a955d0e608744df9358934fe91843d7/green.json`; the capture
`shipping-stale-session.png` is in that same run directory. The log is
`context-guard-visible.log` at the runtime evidence root.
The inspected capture shows the fixed session notice legibly. It also exposes
crowding at the System Key editor/Add button and adjacent table headers; the
layout code is unchanged, but baseline comparison is still required before
classifying that visible finding or claiming full form-layout/human acceptance.

All seven gate-queue stages and the visible run are terminal; Excel is closed.
All30 earlier package pins/four earlier source pins and the candidate's five
package/four Shipping source pins match. No accepted deployment or NAS runtime
is changed. Raw generated reports are retained only under ignored runtime paths.
Remaining activity
source/outcome definitions, capability-loss, during-yield/closed-workbook and automatic-sync probes,
broader Operations/Admin coverage, Settings, Viewer, recording, guides/comparison,
native stability and human/physical acceptance remain required.

## Pending-yield and automatic-sync extension

The actual Add, To Shipments and Shipments Sent handlers now have a test-only
sign-out seam immediately after their existing pending-status `DoEvents`. It
fires once, calls real Core SignOut and verifies rejection before owner entry,
the fixed notice, captured workbook, unchanged staging and no cross-context activity.
All18 assertions pass on the unchanged context candidate.

The timer test establishes actual pending rows, arms the existing timer, cancels
only the scheduled time and invokes the same public callback registered with
OnTime. It does not use the click wrapper that cancels timers after actions.
Healthy dispatch must reach the calibrated owner once and schedule another
callback. Signed-out dispatch must do neither and must show the context notice.
The owner-entry probe stops synchronization before business work; it does not
prove successful healthy synchronization or unauthorized Domain application.

Unchanged-package RED is **319 PASS / 49 FAIL**, retaining all338 prior check
identities and292 prior GREENs. Three new failures are signed-out timer owner
entry, rescheduling and missing context notice. The46 prior missing-activity
failures remain. Evidence: `interruptions-red.log` and
`3743d688580043048e2dc81d6d063482/red.json` under the same ignored evidence root.

Architecture/Plan/controls refinement **bd63ba5** precedes runtime repair and
inherits D18's existing context/internal-observation rules. The candidate adds
`RequireActionContext` before any automatic-sync or overlay work and exits
directly on rejection, preventing the cleanup path from rescheduling. The eight
timing-start callers now call the existing typed timer directly; removing that
pass-through saves four lines and the guard adds one. No limit exception, new
ControlId, permission, store or business owner is introduced.

The isolated package set is `deploy/validation-shipping-timer-context`.
All five build/compile checks and Operations cold start pass. Static maintenance
and all three JSON contracts pass at177 components/5510 procedures/123755 lines;
8 literal/45 unresolved dynamic calls and195 duplicate-body candidates are unchanged.
All28 module limits hold; the form is3019 lines. No automatic dead-code deletion
or limit exception is taken.

The first candidate run stops before Shipping checks at **67 PASS /1 harness
failure**, RPC `0x800706BE`; this is not product RED or guard GREEN. Windows records
`c0000028` in `ntdll.dll` at2026-09-13T00:29:57Z. The subsequent recovery process
has zero workbooks; after Quit fails to exit it, only that verified empty recovery
process is stopped. All five candidate package hashes remain unchanged before retry.
Evidence: `timer-context-green.log`, `98f10db6e29040b093fe15e2a65d6be9/green.json`
and sanitized `timer-context-native-faults.json`. Native stability remains open.
The unchanged-package retry records **322 PASS /46 FAIL**:
`timer-context-green-retry.log` and
`f5bdf30b4c6d4cc0b323b0e3102f903b/green.json`. All368 check identities remain,
all319 RED-run GREENs survive, and there are no duplicate, missing or non-activity
failures. All three signed-out timer failures now pass; all30 interruption/timer
checks pass. The46 missing-activity assertions still fail and the suite exits1.
This is focused timer GREEN, not complete Shipping activity acceptance.

The inspected `shipping-stale-session.png` in that same directory shows the fixed
notice legibly. Comparison with the previous guard candidate's
`0a955d0e608744df9358934fe91843d7/shipping-stale-session.png` shows the same main
key-editor/Add overlap and adjacent-header crowding. The timer change does not
introduce those visible defects; this comparison does not locate their original
introduction or grant human layout acceptance. The capture exercises the mutation
rejection surface; timer status is separately asserted through its real callback.
Candidate regression results follow; native and human acceptance remain open.

Packaged smoke passes86/86. The first live-role gate stops at **39 PASS /1 harness
failure**, step `Run Production form Complete Run`, with RPC `0x800706BE`.
Windows records another `c0000028`/`ntdll.dll` at2026-09-13T00:37:37Z. No live-role
GREEN is claimed from this run. Its raw report is preserved only as
`timer-context-live-role-failed-results.md`. The recovery process again contains
zero workbooks and is closed without saving. This prompted comparison against the
preceding pinned guard candidate;
neither a retry nor the already passing focused Shipping test clears native stability.

The preceding pinned `validation-shipping-context-guard` set reproduces the same
**39 PASS /1 harness failure** at `Run Production form Complete Run`, with
`c0000028`/`ntdll.dll` at2026-09-13T00:40:18Z. Evidence is
`timer-baseline-live-role.log` and `timer-baseline-live-role-results.md`.
This failure is therefore also observable without the timer repair; root cause
remains unknown. Identical retries are not treated as acceptance. The next native
diagnostic traces only fixed phase markers inside the existing completion owner
in an unsaved test copy; all package hashes and business assertions must survive.

Independent candidate gates pass ordered full-chain30/30, existing Viewer,
combined launchers3/3 and Shipping layout/identity1/1. Their logs use the
`timer-context-` prefix, with launcher/layout reports in matching subdirectories.
The full-chain interval also contains `c0000005`/`combase.dll` at
2026-09-13T00:44:09Z. The passing business/static/restart checks do not clear this
fault. Raw chain/Viewer reports were copied to ignored runtime evidence and their
previously clean generated integration paths restored; no raw runtime values are
committed.

The first reusable Production run returns0 PASS /2 reported RED with
`SignedIn=False` and `FAIL/FormNotOpen`. It waits at the batch-scale call on the
single-OK notice `Current invSys user is not signed in.` Read-only native-dialog
inspection identifies that notice; the existing owned-process observer dismisses
only its OK so the test can finish. This is failed fixture setup, not meaningful
Production behavioral RED or a completed restart. The run is preserved in
`timer-context-production/production-reusable-production.md`; a full unchanged-
settings retry uses a separate `timer-context-production-retry` directory.

The native diagnostic also stops39 PASS /1 harness failure at Complete Run.
Its unsaved wrapper records only fixed markers P01 throughP08. P08 precedes
`modUiQuiet.BeginQuietUi`, `ExecuteProductionSession` and `EndQuietUi`; P09,
before output restoration, is never recorded. Do not attribute the fault to the
service alone yet. Exact evidence: `timer-native-trace-run.log`,
`timer-native-trace-results.md`, `timer-native-completion-trace.txt` and
`timer-native-completion-map.json`. `prepare-native-trace.ps1` and its generated
diagnostic remain ignored developer artifacts; no Production source or XLAM is saved.
Next native action: place fixed markers around quiet-UI entry, the typed service
return and quiet-UI exit to narrow this interval before proposing any runtime fix.

The full-settings reusable Production retry establishes `SignedIn=True`, then
stops0 PASS /1 harness failure with RPC `0x800706BE` at the batch-scale contract
(fixed `progress.txt` marker). Windows records `c0000028`/`ntdll.dll` at
2026-09-13T00:54:46Z; the earlier traced Complete Run fault is00:52:51Z.
The retry's report/log are under `timer-context-production-retry`; no reusable
Production or clean-restart GREEN is claimed for this candidate. Earlier2/2 on
the preceding candidate remains historical evidence, not a substitute for these runs.

All current test handles are terminal and Excel is closed. All40 preserved/current
package pins and eight current source pins match. The timer repair is a focused
GREEN checkpoint with release gates still open, not Slice4be completion. No accepted
deployment or NAS runtime changed; Receiving's earlier845/845 source/package pins
remain preserved, and that full suite was not rerun here. Remaining Shipping
activity/outcome/capability/closed-workbook coverage, comprehensive Operations/Admin
coverage, Settings, Viewer, recording, guides/comparison and human/physical acceptance
remain required alongside resolution of the native validation failures.
