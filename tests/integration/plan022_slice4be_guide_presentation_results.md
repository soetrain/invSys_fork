# Plan 022 Slice 4be.5 paired Action Path presentations

**D13 RED and focused packaged GREEN verified; visible/broader acceptance open.**

## Current Settings-diagnostic candidate restart attempt

On 2026-09-23, the frozen `deploy/validation-settings-diagnostic` restart gate
records **25 PASS / two harness failures**, with two preceding checks unreached.
The five instrumented compiles, original-process closure, distinct fresh Excel,
saved Operations preference, exact explicit guide/run pairing, absence of an
inferred evaluation, and Operations-only dependency checks pass. The final
`modInventoryViewer.CloseInventoryViewerForTest` dispatch fails with 0x800A03EC
(macro unavailable/disabled); disposable fixture cleanup also fails. The training/
Config preservation and saved-probe-package checks are not reached. Preserve the
earlier candidate's 27/27 result as historical evidence, not proof of this attempt.

Two of three planned captures are accepted: the unsaved pre-restart How-To view
and restored Operations settings. The restored Compare capture is **rejected**:
an Excel insufficient-memory dialog obscures it. An additional failure capture
shows the same dialog and is diagnostic only. A prolonged pairing/capture interval
precedes failure; individual call timing and its cause are not yet isolated.

Read-only native facts from the failed Excel process show **9,991 current / 10,000
peak GDI objects**, with the machine's configured GDI quota **10,000**; USER objects
are 9,875 against a configured 10,000 quota. Native enumeration finds **301 XLMAIN**
windows, although later recovery COM inspection reports zero workbooks and zero
workbook windows. The process reached its GDI quota. Which runtime or harness path
creates/retains the windows is unresolved; do not infer it from the memory dialog
alone or increase system quotas as a substitute for diagnosis. Microsoft documents
the separate [per-process GDI quota](https://learn.microsoft.com/en-us/windows/win32/sysinfo/gdi-objects).

Under standing user authorization, the exact memory dialog is acknowledged and
Quit is requested on the identified empty test instance. It remains alive and
is then terminated after zero workbooks/windows are verified. This is assisted
recovery, **not normal shutdown**. The harness executes its original settings
restoration before terminal exit; the later recovery preserves the post-harness
settings snapshot. Independent equivalence to the original pre-test settings is
not proven. The partially cleaned generated fixture is preserved; no operational
workbook is changed. The 23:55:42 UTC audit records zero Application events, which
does not override the visible error or assisted recovery. All 299 runtime/195
test/five frozen package hashes remain unchanged.

Root: `reports/runtime/slice4be-viewer-published-read/4d9d2c76d9f14bba9afc3b815f3f2e3f`.
Receipts under `reports/runtime/`:
`settings-diagnostic-fixed-regression-restart-attempt-verification.json`,
`paired-restart-memory-resource-facts.json`, `paired-restart-gui-quotas.json`,
`paired-restart-window-counts.json`, `paired-restart-recovery-workbook-facts.json`,
and `paired-restart-assisted-closure.json`. All remain ignored local evidence.
No broad restart/comparison retry until focused resource observations localize
growth across fresh-session pairing, rendering, and foreground capture. No runtime
or architectural change has been made in response to this finding.

## Original paired-presentation contract and evidence

Architecture v4.11 D18's paired-presentation refinement names an additional
Operations **View guide and run** surface. It implements the approved How-To,
Diagnostic and Compare both views without repurposing the existing recording
library or the published reader's original-source observations. Plan 022 and
controls v1.180 carry the same contract.

`tests/tooling/Slice4beGuidePresentation.ps1` runs after the existing guide-binding
checks under `-CheckGuidePresentation`. It reuses real Admin recordings, two
immutable authored guide versions, an ordinary Operations reader, actual **Use
for selected run** and separate **Evaluate** handlers. Existing Operations
Settings **Save My Preference** supplies Compare both; cleanup restores the prior
choice. The parent harness preserves unrelated current-user settings.

The protecting checks exercise the actual new entry/method/Refresh/Close controls,
exact guide/run/result provenance, three presentation states, four layouts, view
reuse retaining its local method, local switching versus each saved preference and
the warehouse default, guide integrity, stale run/context,
an explicitly selected second version without a borrowed old conclusion, and byte
preservation. The probe only observes/delivers real control actions; no new view
implementation or form is injected. The existing generic layout probe retains its
old rules for existing forms and tests visible controls for the new single-pane/
side-by-side surface.

Meaningful RED is the missing entry/view controls and their behavior on
the frozen `deploy/validation-guide-binding` candidate. An unavailable existing
recording, guide, evaluation or Settings fixture throws as a harness failure; it
does not establish product RED. Retain all preceding 204 guide-binding identities.

The packaged RED is **207 PASS / 36 expected FAIL**, with all **204 prior GREEN
identities preserved**. Three new non-mutation checks pass; every failure belongs
to `GuidePresentation.*`, including the actual missing entry handler. No duplicate
identity or harness failure occurs. Report:
`reports/runtime/slice4be-viewer-published-read/89266b181a8744509aa4ff1e8fd1e59c/red.json`.
UTC **2026-09-21 23:48:03.5255971--2026-09-22 00:01:58.5912481**. Immediate Excel
closure is False, followed by normal closure without Quit, recovery action or
forced termination. A read-only attachment attempt during shutdown was unavailable;
it did not establish a workbook count. Zero Application events 1000/1001/1002 and
unchanged frozen-package hashes are verified. Ignored prefix:
`reports/runtime/guide-presentation-red`.

The implementation adds primitive Core `modPathPresentation`, Operations
`frmActionPathView`, and the library entry/ownership wiring in `frmActionPaths`.
It validates the exact staged intent, current journal and guide chain without
rebinding selection, displaying another guide's evaluation, saving preference or
evaluating implicitly. Existing policy readers supply visibility and the effective
method; a changed/unavailable policy fails closed. The view retains its local
method until closed, and the library's explicit selected result remains the only
candidate for saved display. The isolated `deploy/validation-guide-presentation`
candidate now builds and all five packages compile. Compiled components increase
235 to 237: only existing `frmActionPaths` changes, with the two named components
added. Event Detail and every other existing component remain identical to the
frozen guide-binding candidate. Static evidence records 244 components, 5,990
procedures and 131,826 lines (+298); duplicate groups stay 192 and literal/unresolved
dynamic calls stay 9/45. All 28 preceding module-size limits pass. Scanner findings
are not deletion approval.

Focused packaged GREEN passes **243/243**, retaining all 204 previous GREEN and
all 243 RED identities, including the 39 presentation checks. Report:
`reports/runtime/slice4be-viewer-published-read/7f306f333bb642b1b4db9111de56bea1/green.json`.
UTC **2026-09-22 00:06:28.2073233--00:27:10.7633974**. Immediate closure is False;
Excel then closes normally without Quit, recovery action or forced termination.
A read-only attachment during shutdown was unavailable, not proof of an empty
workbook collection. No recovery dialog was observed. The window has zero
Application events 1000/1001/1002 and candidate package hashes remain unchanged.
Ignored prefix: `reports/runtime/guide-presentation-green`.

The first expanded visible attempt records **272 PASS / one harness failure**,
retaining all 243 focused GREEN identities. Report:
`reports/runtime/slice4be-viewer-published-read/82eb2938896f4672b19d829575181433/green.json`.
UTC **2026-09-22 00:28:30.1907260--00:45:53.2125133**. Excel closes normally;
zero Application events 1000/1001/1002 and unchanged frozen candidate hashes are
verified. After the missing-step scenario, Excel rejects a four-argument generic
control probe; cleanup also receives `0x800AC472`. The original harness does not
identify which control/action failed. This is not product RED or visible GREEN.
All 30 captures are retained, but this record does not claim review of every image.
The retry adds only safe control/action and fixed-stage diagnostics, with no field
values or automatic replay of potentially mutating handlers.

The retry records **255 PASS / one harness failure**, with all 204 preceding
guide-binding checks and 233 of the 243 focused presentation-gate identities
reached; ten focused identities are not reached. Report:
`reports/runtime/slice4be-viewer-published-read/929b4fb5db254eb4b10b8bfcb8eca8cd/green.json`.
UTC **2026-09-22 00:47:11.7973820--01:00:51.9635954**. It stops after the saved
Diagnostic preference check, before the scenario-stage diagnostics execute.
Excel rejects the generic form-control probe and cleanup receives `0x800AC472`.
Normal closure, zero matching Application events and unchanged packages are
verified. All **25 captures were individually reviewed**: 15 prerequisite guide
screens and ten paired mode/layout/result-viewport screens. Controls fit; the
focused diagnostic bottom viewport exposes both complete matched identities.
This partial evidence does not establish a passing visible gate or a root cause.

The next harness records fixed form/control/action names and exception HRESULTs
at the shared call boundary. On failure it may copy pixels only from an already
foreground form owned by the original Excel window and with a whitelisted invSys
caption. It neither changes focus nor replays the failed call, and never logs field
values. The C# helper compiles; an invalid owner is rejected without writing an
image. A new `-CheckGuidePresentationAvailability` extension installs the existing
Admin tracking-policy probes before forms open. Its real disabled-control
recording must produce unavailable expected steps rather than missing steps; a
dirty disposable Config workbook must clear the paired content without saving it.
These are tests of existing D18 behavior, pending execution, not runtime changes.

The independent frozen-candidate Viewer/filter/Shipping-state regression passes
**94/94**, preserving every preceding identity, with immediate normal Excel
closure and zero matching Application events. Report:
`reports/runtime/slice4be-viewer-published-read/a29afdad31e44708ba4d2d70bdfad299/green.json`.
UTC **2026-09-22 01:04:17.2554733--01:06:36.9101127**. All three captures were
individually reviewed: Shipping's held-row feedback is painted, contributing-line
labels and filter controls are readable. The known adjacent Shipping headings and
pending Event Detail horizontal-scrolling defect remain; this gate does not close
either issue. Boxing/Shipping, evaluation and full-chain gates remain separate.

Boxing/Shipping now retains all **1,707 preceding GREEN identities**, with exactly
the same **seven pending D8-A missing-Auth findings** among 1,714 unique checks.
Report: `reports/runtime/slice4be-shipping-activity/cc7dbd10022a4e6eae67f3d08960b956/boxing-activity-shipping-recording-green.json`.
UTC **2026-09-22 01:06:36.9821100--01:23:39.5192895**. Immediate closure is False;
Excel closes normally without intervention. Zero matching Application events and
unchanged frozen packages are verified. All **22 captures were individually
reviewed** and their hashes checked. The denied/stale-session Shipping captures
now use the owned-form foreground helper and are painted with the expected notice.
Make/Unbox success and zero-quantity messages are correct; older-policy and
unavailable-store notices remain separate from business success. This is a
capture-harness correction, with no runtime contract change or manufactured product
RED. Existing packaged assertions and visible captures validate it proportionally.
The pending D8-A, known heading adjacency, detail scrolling, evaluation, full chain
and paired visible acceptance are not waived.

The frozen candidate's evaluation regression passes **376/376**, preserving every
preceding GREEN identity. Report:
`reports/runtime/slice4be-viewer-published-read/ee6e9f0dfb6f4431873a89bf992083af/green.json`.
UTC **2026-09-22 01:23:39.6976019--01:42:11.5430183**. Excel closes immediately
and normally, with zero Application events 1000/1001/1002. All **19 captures were
individually reviewed** and their hashes verified: pending and partial publication
retain Awaiting states, only the fully applied case shows Conclusion observed,
and the expectation editor keeps its controls reachable at minimum/default/large
sizes. The ordered editor distinguishes Pending without retries from Completed
with retries and shows the explicit conclusion selection. This regression does not
substitute for the pending paired visible gate or full Release 1 chain.

The paired candidate's full Release 1 chain passes **32/32**, ordered live roles
**48/48**, and Create Warehouse **15/15**, preserving all preceding chain/live
GREEN identities. Ignored prefix: `reports/runtime/guide-presentation-chain`.
UTC **2026-09-22 01:42:15.0432027--01:51:10.6342477**. The chain includes actual
captured-workbook Receiving and Shipping, two consecutive Production batches,
versioned Boxing, processor application, final balances, restart/reconciliation,
runtime extraction and static checks. The three tracked generated reports were
restored byte-for-byte; zero Application events 1000/1001/1002 and all 45 frozen
package hashes are verified. Cleanup was **assisted**: read-only native-owner
inspection established an integer workbook count of zero, then normal Quit was
requested. The actual recovery prompt and selected **Yes, I want to view these
files later** option were individually reviewed before confirmation. No forced
termination or recovery-file deletion occurred. This passing chain does not imply
unattended clean shutdown, paired visible acceptance, human UAT or Slice 4be completion.

The shared failure-boundary privacy/replay test passes **10/10 offline checks**:
`tests/tooling/Test-Slice4beFormFailureDiagnostics.ps1`. Entered field content is
absent, fixed control identity and inner COM code are retained, unrelated command
arguments are omitted, and a rejected call executes once. It creates/attaches to
no Excel instance. Report:
`reports/runtime/form-failure-diagnostics/8a1798273fc64be7b856ea9b7895dc82/results.json`.
This supplements diagnosis; it is not packaged product acceptance.

The expanded diagnostic run records **294 PASS / one harness failure**, retaining
all **243 focused GREEN identities**, with no reported product-check failure.
Report: `reports/runtime/slice4be-viewer-published-read/89450cea712e432a8e0c2d9d0f1489a5/green.json`.
UTC **2026-09-22 01:51:14.2621260--02:16:20.3924387**. Excel closes normally;
zero matching Application events, unchanged frozen packages and unchanged running
test sources are verified. All **39 captures were individually reviewed** and
hash-verified. Three methods, four sizes, native maximize/restore, complete bottom
match IDs, invalid guide/context, unevaluated version, missing/extra/rejected
actions, None expectations, capture-off history and current visibility loss pass.
The minimum extra-action bottom viewport exposes both complete original action
identities. These are partial results, not a passing expanded gate.

The run stops while preparing historical unavailable evidence. The fixture
incorrectly requires a Stopped closing entry after executing a command excluded
from sequence capture. Source review shows `modRecordingSession.Prepare` interrupts
that case with `TRACKING_UNAVAILABLE`, and `Interrupt` writes an Incomplete Close;
the failed run did not record which combined fixture condition differed. The test
now requires that existing Incomplete closure, removes the later Stop attempt,
verifies all three saved per-control policy flags and records only sanitized
counts/Booleans before rejecting an invalid fixture. No runtime or normative
contract changes. This is a harness correction, not product RED. Its retry and
the final unavailable-policy cases remain pending.

The corrected full visible retry passes **305/305**, preserving every one of the
preceding **294 passing identities** and all 243 focused identities. Report:
`reports/runtime/slice4be-viewer-published-read/87ecb4eaf2e64f879bce5b3aba2ff404/green.json`.
UTC **2026-09-22 02:19:15.4058741--02:46:38.0636614**. All **11 new availability
checks** pass. Sanitized fixture facts confirm two journal entries, one Incomplete
Close with TRACKING_UNAVAILABLE, no observations and unchanged activity bytes.
The evaluated gap names both expected steps as Evidence unavailable, not Missing.
Dirty borrowed Config clears both panes and disables the method while preserving
unsaved edits and on-disk Config/training bytes. Closing that unsaved fixture and
refreshing restores the same historical evaluation ID.

Excel closes normally without intervention. Zero Application events 1000/1001/1002,
unchanged frozen packages and unchanged running test sources are verified. The
retry creates 42 captures. Visible review uses **39 individually reviewed earlier
captures plus the three individually reviewed new availability captures**; every
reviewed hash is verified. This does not claim direct review of all 42 retry images.
The same frozen candidate and unchanged preceding capture-producing tests support
that combined evidence; only the availability helper changed between these runs.
The earlier COM rejections did not recur, but their root cause is not established.
All failed attempts remain recorded.

Regenerated maintenance evidence in
`reports/runtime/guide-presentation-static-availability-retry` retains 244 components,
5,990 procedures, 131,826 lines, 192 duplicate groups and 9/45 literal/unresolved
dynamic calls. All 28 preceding module-size limits and 45 frozen package hashes
hold. No source cleanup or runtime change was needed for this fixture correction.

The expanded visible gate retains the focused identities and captures all three
methods, minimum/default/larger/restored and native maximized/restored sizes,
the saved-result viewport, invalid guide/context and an unevaluated second guide.
Additional scenarios use actual missing/extra/rejected actions and explicit
Evaluate, None expectations, capture-off historical evidence and same-session
visibility loss. Fixtures must be validated before presentation assertions;
capture/method/Refresh actions must preserve training bytes. The tests do not
fabricate observation or evaluation records.
The expanded visible gate, historical unavailable comparison and unreadable-policy
recovery now pass. Fresh-process paired preference, human acceptance and the
remaining D18 guide-editing/direct-curation/transfer scope remain open. The separate
full chain passed with the assisted-cleanup limitation recorded above; these gates
do not complete Slice 4be or approve the pending Event Detail/D8-A proposals.

The first paired-preference restart attempt records **297 PASS / one harness
failure**, preserving all **294 preceding passing identities**. Report:
`reports/runtime/slice4be-viewer-published-read/184c512cafa142fb91c987cb15351267/green.json`.
UTC **2026-09-22 02:55:32.2927393--03:20:22.4379223**. The three new pre-restart
checks prove Save My Preference through Operations, initial Compare both and an
unsaved How-To switch. Its How-To image was directly reviewed and hash-recorded;
the other 39 images from this attempt are retained without claiming a new direct
review. Exact guide and different observed recording provenance remain visible.

The harness then attempts to hash writable saved probe XLAMs while Excel holds
them open, producing a sharing violation before the restart. Normal final cleanup
closes Excel without intervention. Zero matching Application events and unchanged
running test sources are verified. No fresh-session result or behavioral RED is
claimed. The correction moves probe hashes after verified normal process exit,
before read-only reopening. No runtime or architectural behavior changes.

The retry uses `-GuidePresentationRestartOnly` with the same compiled saved-copy
candidate. Its focused fixture uses actual Admin publication, Viewer recording,
Admin Save Value, Create guide and Save guide handlers, validates the immutable
source/guide hashes, then records a different observed run. It retains the original
regression evidence rather than treating a small focused gate as its replacement.
The restart helper exercises actual Operations Settings and view handlers, normal
owned-instance closure, a different native Excel process, Operations without Admin,
explicit re-pairing, absence of automatic conclusions and Config/training byte
preservation. Credentials remain in memory and the main harness restores local
settings. No new runtime fix is justified by the first attempt's harness failure.

The focused retry passes **27/27**: five instrumented package compiles, three
startup/isolation checks, four actual-handler fixture checks and all **15 restart
checks**. Report:
`reports/runtime/slice4be-viewer-published-read/21d7b8f118ed4186a010a4dbe7baba9a/green.json`.
UTC **2026-09-22 03:25:13.7245882--03:27:49.0766017**. The original owned Excel
instance has a verified integer workbook count of zero before normal Quit and
exits without intervention. A different native Excel process loads only Core,
Inventory Domain, Designs Domain and Operations. Actual Operations Settings
restores Compare both as the saved personal/effective choice. The library starts
without an inferred guide/run pair or result. After explicit exact-guide and
different observed-run selection, the view restores Compare both rather than the
unsaved How-To choice, preserves provenance, and displays Not evaluated.

Config and every training file remain byte-identical throughout the preference,
view and restart actions. Saved probe packages remain unchanged. Final Excel
closure is normal; zero matching Application events and unchanged running test
sources/candidate hashes are verified. All **three focused captures were directly
reviewed and hash-verified**: unsaved How-To, fresh-session Operations Settings,
and restored Compare both with separate instructions/observations and no conclusion.
All 45 frozen package hashes and unrelated-document pins hold. Existing runtime
maintenance metrics remain applicable because no VBA/form/Ribbon source changed.

This is regression evidence for the existing approved D18 persistence contract,
not a new behavioral RED/GREEN implementation. The separate full visible 305/305
and preceding 294 identities retained in the failed broad attempt remain evidence;
27 focused checks do not replace those gates. Fresh-process paired preference is
now verified on this isolated candidate. Human acceptance, published-guide editing,
direct curation, transfer and broader D18/Release 1 completion remain open.
