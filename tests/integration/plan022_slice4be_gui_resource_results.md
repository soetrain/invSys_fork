# Slice 4be fresh-session GUI resource diagnosis

Last verified: 2026-09-24 UTC. Architecture v4.11 D18 and the R1 operator
deployment model govern this investigation. No runtime/XLAM change or new
architecture decision is made in this investigation. The later
[paired-view layout correction](plan022_slice4be_layout_stability_results.md)
continues the latency diagnosis with a focused runtime change. The add-in-only
failure remains a failure.

The preceding restart attempt reached the configured 10,000 GDI quota and
301 native XLMAIN windows. Its rejected image, incomplete checks and assisted
closure remain in [the presentation record](plan022_slice4be_guide_presentation_results.md).

## Packaged boundary trace

`Test-Slice4beGuideResourceDiagnostic.ps1` runs the existing actual-handler
restart fixture with optional native resource tracing. `Slice4beGuideResourceTrace.ps1`
records counts, fixed developer boundary names and process identity; it never
records command arguments, credentials, workbook contents or window captions.
A separate controller samples during busy calls and retains the original local
settings snapshot in memory until Excel and its worker exit. Reports are ignored.

First diagnostic controller:
`reports/runtime/guide-resource-diagnostic/6137ed7fa1d9473197031e09c9930e98`.
Packaged report:
`reports/runtime/slice4be-viewer-published-read/ae04cab5fc53498e9bd1bab6f8daade5`.
UTC interval **00:06:37--00:12:09**. **21 PASS / one diagnostic-bound failure**;
five instrumented projects compile. This is diagnostic evidence, not product RED.

- The original process returns to six native XLMAIN windows and approximately
  600 GDI objects after actions. Its How-To capture changes counts from
  613 GDI/six windows to 612/six.
- The fresh process starts with five windows. Target selection/sign-in reaches
  eleven, Viewer opening fifteen, Events twenty-nine, and Settings thirty-one.
  Its Settings capture leaves 1,294 GDI/thirty-one windows unchanged.
- Selecting/pairing subsequently reaches sixty-six windows; opening the paired
  view reaches eighty-two and 2,956 current/3,001 peak GDI objects. The diagnostic
  stops at its eighty-window bound before the final paired capture. This bound
  is a diagnostic precaution, not a new product acceptance threshold.
- Growth continues during ordinary cleanup. Normal close messages to the owned
  read-only guide/library forms remain unprocessed; read-only COM attachment
  stalls. The verified disposable process is terminated. Unlike the preceding
  attempt, an independent original settings snapshot is retained through this
  recovery and verified restored afterward. Five package hashes are preserved.
- The passive sampler races process exit and reports a missing native handle;
  its `finally` still restores settings. The sampler now tolerates only a
  verified exited process. The failed diagnostic is retained unchanged.

These observations rule out the final capture as the first source of growth.
They do not establish an invSys form recursion defect or prove that all earlier
COM errors/crashes have the same cause.

## Excel-only controls

`Test-Slice4beExcelWindowLifecycle.ps1` creates a synthetic read workbook and
minimal XLAM, closes the creating instance, then tests twelve read-only opens/
closes in a fresh Excel process. **No invSys package is loaded.** The macro only
opens the synthetic workbook read-only, closes without saving and releases its
reference. Each control exits normally and preserves the read file's bytes.

| Control | Initial/final XLMAIN | Initial/final GDI | Report directory under `reports/runtime/excel-window-lifecycle/` |
|---|---|---|---|
| COM opens/closes | 2 / 2 | 226 / 244 | `98c4c149e9354886a3fcd75449d1624a` |
| VBA macro opens/closes | 2 / 14 | 236 / 643 | `46685cd1eac741e4b161cc2980e9787f` |
| Same macro after VBE initialization | 2 / 14 | 332 / 739 | `10fe8a389de54e759d09aa777c93656e` |
| Same macro with a separate saved workbook held open | 3 / 3 | 288 / 297 | `7dc8703318f74fd08fdd5c20a3a6b7e8` |

This reproduces the native-window retention independently of invSys. The
responsible internal Excel mechanism remains unproven. No OS quota, Office
setting, deployment, package or runtime code is changed to conceal the failure.
Receipt: `reports/runtime/guide-resource-control-verification.json`.

## Saved-workbook packaged control

The normative operator deployment model requires a saved `.xlsm`/`.xlsb`
workbook for normal acceptance and describes new-blank-workbook testing as a
diagnostic stress case. The optional `-SavedWorkbook` control retains every
existing restart assertion and adds five checks for the same saved/reopened
synthetic `.xlsm` host in both processes and its unchanged bytes. It does not
replace the preserved add-in-only stress failure or prove a role-specific
workbook binding by itself; the separate accepted full-chain gate covers that.

First saved-host calibration:
`guide-resource-diagnostic/fc33b102b46d401ca64759a330fdac42`, packaged report
`slice4be-viewer-published-read/a65d0f9738dc4f17942b7d014a3c147f` under
`reports/runtime/`. **16 PASS / one harness exception**. The close loop had
already released the retained workbook's COM wrapper; the additional release
then failed before restart. Excel exits normally, original settings restore,
and package hashes hold. The duplicate release is removed; this is not product
RED. The controller now propagates its worker's exit code after restoration.

Second calibration, controller `guide-resource-diagnostic/47ec347c954a499f97a411963fd19fb8`
and packaged report `slice4be-viewer-published-read/6f7136d121634e5b83db0c7a88f7c32c`,
also records **16 PASS / one harness exception**, normal closure, restored
settings and preserved packages. Removing the duplicate release left a null
comparison against the invalid COM wrapper. A separate `Scripting.Dictionary`
COM calibration reproduces `InvalidComObjectException` from that comparison;
using the existing Boolean mode flag passes without touching the released
wrapper. The focused restart is retried only after that calibration.

The corrected saved-host attempt, controller
`guide-resource-diagnostic/437c154421ab44e59f0adc1bea7ef61f` and packaged report
`slice4be-viewer-published-read/eb3e87dde811477a9063498fbbc6bb4e`, completes
**31 PASS / one preservation failure** at 00:23:29--00:26:41 UTC. All 27 preceding
identities are reached and five saved-host checks are added. Both Excel processes
close normally; original settings and package hashes hold. All three images are
directly reviewed: unsaved How-To, restored preference and restored Compare with
separate authored/observed panes and Not evaluated. No memory dialog is present.
Their acceptance scope is visible presentation, not the failed preservation gate.

Hash-only checkpoints in controller
`guide-resource-diagnostic/1bfaf8676a594a4796e682c62ac0b6cd`, packaged report
`slice4be-viewer-published-read/6725d87b6cb74e76b61beba3de44dc82`, retain the same
**31 PASS / one failure**, normal closure and restored settings at
00:28:33--00:32:54 UTC. Only two files are added, immediately after Save My
Preference; all 23 preceding Config/training files and both additions remain
unchanged through original shutdown, fresh target/Viewer/Settings and Compare.
The added records are the approved `VIEWER_PATH_PREFERENCE_SAVE_REQUESTED` and
`VIEWER_PATH_PREFERENCE_SAVE_COMPLETED` observations, owner
`CORE_PERSONAL_PREFERENCE`, effects Unknown/Changed, one shared activity ID,
distinct record IDs, no sequence and no source-event references. The fixed-field
receipt is `reports/runtime/guide-resource-preference-observation-facts.json`.

The old whole-interval assertion incorrectly prohibited observations required
by the already approved D18 Settings-editor contract. The correction adds a
strict pre-Save read boundary and validates exactly that correlated pair at
Save. It preserves every previous hash, admits only the two validated activity
files, and protects their hashes in all later reads. The original final check
identity remains. There is no blanket Training/Activity exception or runtime
behavior change. `Test-Slice4bePreferenceSavePins.ps1` calibrates rejection of
extra/missing files, Config/prior-file mutation, wrong directory/control/context/
sequence/outcome/effect/correlation, duplicate outcome and damaged hash; rejected
differences never advance pins. It also rejects mutation of an accepted observation
in the following interval: **32/32**, root
`reports/runtime/preference-save-pins/1851293c338d451f8025e03899a33759`.

Comprehensive coverage, transfer, pending architecture amendments and human/NAS
acceptance remain open in the [remaining checklist](plan022_slice4be_remaining_acceptance.md).

The first corrected attempt, controller
`guide-resource-diagnostic/23daee67a58f4a4c87a95ee3cab44e07`, packaged report
`slice4be-viewer-published-read/b9e80895fa7c4727a865f756d82143c4`, records
**20 PASS / one lifecycle exception** at 00:36:12--00:38:20 UTC. Both new exact
Save boundary checks pass. The single post-exit process-enumeration check stops
before restart; the controller subsequently observes Excel absent, without
assistance, and restores original settings. Package hashes hold. The precise
race is not proven because that attempt did not record both native states at
the check. The wait now observes original-process termination and disappearance
from enumeration together, still rejecting another Excel process and requiring
no Excel before restart. No extra Quit, COM reattachment or termination is added.

## Verified behavior and remaining visual/responsiveness limits

Final controller:
`reports/runtime/guide-resource-diagnostic/087a6f7fdec84b63a3c416f3761d4ee7`.
Packaged report:
`reports/runtime/slice4be-viewer-published-read/83c161e3d78346debf143bac8cf76764`.
UTC **00:39:27--00:43:41**: **34/34**, retaining all 27 preceding identities,
plus five saved-host checks and two exact Save boundaries. Five instrumented
projects compile. Both original and fresh Excel close without assistance; the
controller restores and verifies the original settings snapshot. Config, all
preceding training bytes, the two validated observations, saved probe packages
and saved host bytes remain protected. The native original-exit observation
ends with `OriginalExited=True` and zero enumerated Excel processes before
the fresh process is created.

The **00:45:46 UTC** audit records zero Application events 1000/1001/1002 and
unchanged **299 runtime/five frozen package** hashes. Independent receipt:
`reports/runtime/saved-workbook-restart-verification.json`.
The fresh process reaches at most **six XLMAIN windows at handler/capture
boundaries** and peak **563 GDI**; passive samples can include transient open
workbook windows. This verifies the saved-workbook behavior, not an add-in-only
fix or an explanation of every earlier crash/COM error.

All three current captures are directly reviewed. Presentation, saved preference,
exact provenance, authored/observed separation and Not evaluated content are
readable. A **Task Manager thumbnail overlays the lower-right form area** in
each capture, so none is accepted as an unobscured full-control image. The first
completed saved-host attempt's three clean reviewed captures remain separate
evidence; do not relabel them as captures from the final passing attempt.

The boundary around the fresh `frmActionPathView/cboActionPathView/Selected`
observation takes **84.485 seconds**, including readiness observation and macro
dispatch. Its internal cause and actual operator responsiveness remain unproven.
Stable resource counts and passing correctness assertions do not accept this
latency. Next diagnosis should count actual refresh/layout/activation entries
around that boundary in the saved-workbook fixture, and obtain unobscured
operator captures. Do not rerun the unchanged broad comparison merely because
the saved-host correctness checks now pass.

Final maintenance evidence: `reports/runtime/guide-resource-static-final`;
three schema validations pass. All preceding metrics and 28 size limits hold:
250 components, 6,045 procedures, 132,893 lines, 9 literal/45 unresolved dynamic
calls and 193 duplicate groups. All 258 PowerShell files parse. No VBA/form/
Ribbon/XLAM implementation is changed, so a new runtime build and unchanged
full-chain/curation reruns are not justified by this tooling-only correction.
Their already verified candidate results remain at their original scope.
