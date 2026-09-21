# Plan 022 Slice 4be.5 paired Action Path presentations

**D13 test entry; runtime implementation and behavioral RED pending.**
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
reuse, local switching versus saved preference, guide integrity, stale run/context,
an explicitly selected second version without a borrowed old conclusion, and byte
preservation. The probe only observes/delivers real control actions; no new view
implementation or form is injected. The existing generic layout probe retains its
old rules for existing forms and tests visible controls for the new single-pane/
side-by-side surface.

Expected meaningful RED is the missing entry/view controls and their behavior on
the frozen `deploy/validation-guide-binding` candidate. An unavailable existing
recording, guide, evaluation or Settings fixture throws as a harness failure; it
does not establish product RED. Retain all preceding 204 guide-binding identities.
No packaged RED, new build/compile/static, visible presentation, restart preference,
capture-off/current-policy, all five comparison outcomes or full Release 1 gate is
claimed by the draft. Those remain required before acceptance.
