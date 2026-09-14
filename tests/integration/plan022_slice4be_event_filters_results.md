# Slice 4be.3 Events filters

Last verified 2026-09-13. Architecture v4.11 D18 governs the loaded-projection
filters, complete groups, captured context and publication-only Viewer reads.
Release 1 and Slice 4be acceptance remain open.

## Protecting test and RED

`tests/tooling/Test-Slice4beConfigCommands.ps1 -CheckViewerFilters` runs the existing
30 real Settings -> Admin publication -> Viewer checks before the filter extension.
The extension supplies a declared synthetic EVENTS1 wire containing 110 lines in
109 groups. It changes the actual ComboBox selections and invokes the existing
Search, Refresh, page and list-selection handlers. It does not fabricate owner
events or write synthetic inventory authority.

The complete pre-implementation result is **33 PASS / 22 FAIL**, 55 unique checks,
with all 30 reader checks passing. An earlier argument-construction error stopped
at 31 PASS / one harness failure and is excluded from behavioral RED.

The test protects Operator actions versus All published events, honest reservation
labels, AND-combined family/source/recorded-outcome criteria, unavailable metadata,
whole-projection filtering before 100-group paging, complete attempt/result detail,
page-one reset, Search, staged dates applied only by successful Refresh, Events-only
visibility, captured-session invalidation and four supported form sizes. A counter
on the actual Core projection-read entry proves selectors and Search do not reload
the projection; broader owner-read boundaries remain separately required.

## Implementation and verification

D18's filter-control refinement names the four controls and their session-only
semantics before implementation. Operations owns the controls and a typed shared
ComboBox event binding. The form disconnects the bindings on QueryClose. The pager
matches complete loaded groups without trimming their contributing detail lines.
Refreshing repopulates choices from the permitted loaded projection, preserving
an available selection and otherwise using its All choice.

The first isolated candidate passes **55/55**, Shipping/reader **56/56**, paging
**16/16**, Refresh **16/16** and Detail **34/34**, with five-package compile. Its
four identical Change wrappers increase duplicate-body groups from 189 to 190.
That maintenance regression is rejected; the shared binding replaces the wrappers.

The revised `deploy/validation-events-filters-binding` candidate builds and compiles
all five XLAMs and passes Operations cold start. Its complete no-capture focused
run passes **55/55**. Static regeneration records 209 components, 5,734 procedures,
189 duplicate-body groups, 45 unresolved calls, 9 literal Application.Run targets
and 28 oversized-module entries. Duplicate and dynamic-call counts hold. The
per-module growth review below corrects earlier size-ratchet claims. The affected
Viewer form is 861 lines; the new binding is 27 lines.

An optional foreground capture run passes 53 behavior/layout checks, then stops
because the requested Viewer is not in the foreground. No visible capture or
complete GREEN is claimed for that run. The guard is retained, and the separate
55/55 run completes the remaining Inventory-tab and sign-out checks.
The revised run retains every RED and first-candidate check identity, with no
duplicates. A later read-only Win32 observation finds no foreground window and
cannot open the input desktop. This describes the current capture environment;
it does not identify why that desktop is unavailable or establish its state at
the earlier capture instant. Visible proving must resume when it is available.

The first revised Shipping regression stops after the 30 reader checks with a COM
RPC failure in `PublicationShippingBoxForTest`; it is **30 PASS / one harness
failure**, not Shipping GREEN. The first-call failure record captures the original
Excel process and its child. The child is subsequently verified empty with only
invSys add-in projects. Normal Quit eventually closes it; the force-stop guard
rejects termination while Excel still reports content. No new Windows Application
1000 event is found at that observation, so no native fault module or cause is
assigned. The failed run remains recorded separately from the unchanged retry.
The unchanged retry passes **56/56**; paging **16/16**, Refresh **16/16** and
Detail **34/34** also pass, retaining every prior check identity without duplicates.

## Inventory layout preservation follow-up

Source review identifies a gap outside the original selector assertions: the
Events-only paging strip reduces the Inventory list by 40 points even while the
navigation controls are hidden. D18 explicitly preserves the accepted Inventory
list area. The expanded real tab/resize test checks the prior 12-point gap above
Close at minimum, default, larger and restored sizes. The expanded test records
**55 PASS / 4 FAIL** against the binding candidate: only the four new Inventory
geometry checks fail. The correction anchors the Inventory list above Close while
Events continues to reserve its paging strip. The existing 55 check identities
remain required; this preserves accepted Inventory behavior. The corrected isolated
candidate is `deploy/validation-events-inventory-layout`.
It compiles all five packages and passes Operations cold start and **59/59**,
retaining all 55 prior GREEN identities. The four Inventory geometry failures
are corrected. Maintenance regeneration retains the same component/procedure,
duplicate and dynamic-call counts; the following per-module issue remains open.

## Maintenance re-audit

Comparing each oversized-module entry with the committed baseline finds one
violation: `modWarehouseSync` is 1,759 lines versus its 1,753-line baseline. The
unchanged count of 28 oversized modules did not prove their individual growth
limits. Earlier assertions that size ratchets held are superseded by this finding;
no exception is approved and implementation acceptance remains open.

Tracked source, tests and tooling contain only the declaration of private
`AppendLocationSummariesSync`; its inspected body is an unused location-summary
projection loop. The six obsolete private Viewer helpers (`ViewerRawCell`,
`ViewerRecordedTime`, `ViewerFriendlyEventType`, `ViewerFirstNoteToken`,
`ViewerNoteToken`, `ViewerFirstNonBlank`) have no remaining runtime callers. Their
only external references are stale static-test expressions; the unused
`SNAPSHOT_EVENT_TABLE` constant likewise has no runtime reader. This explicit body
and reachability review, rather than scanner output alone, authorizes removal.

The older Slice4w static suite initially records8/4 because it expects former
Receiving and Viewer call sites. Its assertions are reconciled to the current
typed Receiving navigation/confirmation handlers and D18's published Events reader,
preserving all12 check identities and the operator requirements. It passes12/12
before and after removal. Those stale assertions are not product RED. The original
generated report is restored byte-for-byte; candidate copies remain ignored.

After the frozen Inventory-layout chain finishes, the seven private routines and
constant are removed. `modWarehouseSync` becomes1,727 lines, below its1,753 baseline;
`modInventoryViewerData` becomes184 lines. The new isolated
`deploy/validation-events-maintenance` candidate compiles all five packages and
passes Operations cold start. Public publication passes **82/82**, retaining all66
previous publication checks and adding16 grouped Viewer checks, with no duplicates.
Compiled-source comparison covers202 components: only the two reviewed Core modules
change; all other compiled components match the preceding layout candidate.

Fresh maintenance evidence records209 components,5727 procedures,1136 scanner /
1138 reviewed candidates,189 duplicate groups,45 unresolved calls and9 literal
Application.Run targets. Every one of the28 existing oversized modules is checked
against its committed line limit: none grows. All six new modules and59 new
procedures remain within1000/200 lines respectively. This individual comparison
closes the identified growth violation without an exception. Expanded filters pass
**59/59** on the maintenance candidate, retaining the preceding59 check identities.
Its full chain exits0 at **31/31**, with live roles **48/48** and Create Warehouse
**15/15**. All five package hashes remain unchanged and original reports are restored.
Windows records one `combase.dll` / `c0000005` Excel fault during this final run.
Clean native execution remains unproven despite the passing check reports.

Excel is closed. Preservation matches175 historical and15 publication package pins,
15 protected source files and one reviewed Shipping visibility-only change, plus35
reader/paging/Shipping/filter package pins. The five maintenance candidate hashes
are frozen. The source checkpoint includes publication, Viewer reads/grouping/
filters, restored Inventory space, reviewed cleanup and source-harness dependencies.
It changes no accepted deployment or NAS workbook and does not complete Slice4be.

The prior Inventory-layout candidate's full chain passes31/31, live48/48 and
Create Warehouse15/15, with all five package hashes unchanged and original reports
restored. Windows records an Excel `unknown` module / `c0000409` fault during that
run. Its cause remains unproven; the count checks do not establish clean native
execution or complete acceptance.

## Reproduction and evidence

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-events-shipping-state -Phase RED -CheckViewerFilters
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-events-filters-binding -CheckViewerFilters
```

With the four Inventory layout assertions added, the second command now reproduces
the 55/4 layout RED. Use `deploy/validation-events-inventory-layout` for the corrected
candidate's expanded 59-check run.

Ignored `reports/runtime/` evidence:

- `events-filters-red.log`: complete 33/22 behavioral RED.
- `events-filters-argument-setup-failure.log`: excluded test setup error.
- `events-filters-gates.json` and the corresponding focused/shipping/groups/refresh/detail
  logs: first candidate's 55/56/16/16/34 GREEN.
- `events-filters-binding-build.log`, `events-filters-binding-compile.log`,
  `events-filters-binding-compiled.json`: revised package build/compile.
- `events-filters-binding-static.log`: revised maintenance regeneration.
- `events-filters-binding-focused.log`: complete revised 55/55 GREEN.
- `events-filters-binding-focused-capture.log`: excluded foreground-capture failure.
- `events-filters-binding-check-comparison.json`: all 55 check identities retained.
- `events-filters-binding-regression-comparison.json`,
  `events-filters-binding-resumed-gates.json` and corresponding shipping-retry,
  groups, refresh and detail logs: 56/16/16/34 GREEN, exact prior identities retained.
- `events-filters-binding-package-pins.json`: five frozen candidate hashes before
  the expanded Inventory-layout RED run.
- `events-inventory-layout-red.log`: 55/4 RED for preserved Inventory list space.
- `events-inventory-layout-focused.log`, `events-inventory-layout-check-comparison.json`:
  59/59 and exact retained check identities.
- `events-inventory-layout-chain-summary.json`, corresponding chain/live/create
  reports and `events-inventory-layout-chain-native-faults.json`: prior candidate's
  full-chain results, preservation and unresolved native fault.
- `events-maintenance-surface-before.md`, `events-maintenance-surface-reconciled.md`,
  `events-maintenance-surface-green.md`: static assertion reconciliation8/4 to12/12,
  with12/12 retained after reviewed removal.
- `events-maintenance-build.log`, `events-maintenance-compile.log`,
  `events-maintenance-compiled.json`: post-cleanup five-package compile.
- `events-maintenance-publication.log`, `events-maintenance-publication-comparison.json`:
  82/82, all prior66 publication identities retained plus16 grouped Viewer checks.
- `events-maintenance-compiled-comparison.json`:202 components compared; only the
  two reviewed Core modules differ from the preceding layout candidate.
- `events-maintenance-static.log`, `events-maintenance-module-ratchets.json`,
  `events-maintenance-new-size-limits.json`: individual module limits and new-code
  limits verified, without a growth exception.
- `events-maintenance-filters.log`, `events-maintenance-focused-gates.json`:
  final59/59 filters/layout and82/82 publication GREEN.
- `events-maintenance-chain-summary.json`, corresponding chain/live/create reports
  and `events-maintenance-chain-native-faults.json`: final31/48/15 passing counts,
  preserved package/report bytes and the unresolved native fault.
- `events-maintenance-preservation.log`, `events-maintenance-package-pins.json`:
  protected source/package checks and five frozen final candidate packages.
- `events-filters-desktop-observation.json`: current input-desktop availability.
- `events-filters-binding-shipping.log`: interrupted Shipping regression; its
  matching `first-call-failure.json` under `slice4be-viewer-published-read/` proves
  the original/child process relationship. This is not behavioral RED.

Optional activity coverage for the four new selectors remains pending under 4be.1;
automatic population and rendering must not become user actions. Remaining source
and detail coverage, recordings, How-To/Diagnostic/Compare both, guide library,
physical NAS/multi-station proving and human acceptance remain required. Earlier
native Excel faults remain unresolved in the [Shipping evidence](plan022_slice4be_shipping_viewer_results.md).
Current source inspection finds catalog version8 with31 registered controls,
including one Production and one Admin control. These publication/Viewer results
do not establish comprehensive Operations/Admin activity coverage.
