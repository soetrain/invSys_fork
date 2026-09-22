# Plan 022 Slice 4be.5 paired Action Path presentations

**D13 behavioral RED verified; runtime draft awaiting packaged GREEN.**
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

The source draft adds primitive Core `modPathPresentation`, Operations
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
are not deletion approval. The packaged 243-check run is in progress.
No GREEN, visible presentation, restart
preference, capture-off/current-policy, all five comparison outcomes or full
Release 1 gate is claimed for this draft. Those remain required before acceptance.
