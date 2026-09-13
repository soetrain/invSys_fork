# Slice 4be.3 selected Event Detail evidence

Last verified 2026-09-13. This step implements D18's selected-event detail,
same-read source envelope, profile rendering and captured-context rules. It
does not complete comprehensive publication, Action Paths or human acceptance.

## Contract and implementation

Architecture v4.11's selected-event detail refinement was recorded in D18,
Plan 022 and the controls catalog before runtime implementation. It constrains
the approved contract under semantic inheritance; no canonical schema change,
new business permission, legacy import or architecture exception is introduced.

Core's existing Events read now preserves exact EventID/System_Key and permitted
line fields in its `DETAIL1` envelope, with verified load UTC. The visible
Inventory event summary excludes raw Note; UOM comes from the published UOM
field/token, never a current inventory lookup. Operations owns the detail form
and controller. It groups by exact source/ID within the captured warehouse,
retains repeated keys and unlike units without aggregation, and renders the
cached profile's enabled fields/order. Profile changes apply on explicit Refresh.
Missing legacy publication/outcome fields remain Unavailable; recorded source
times remain zone unavailable. Current-state supplements gain no historical IDs.

The line selector and field list are read-only views. Detail reuses its owned
instance and closes with Viewer. Context loss invalidates retained data on
detail actions; layout also checks context, and repeated invalidation is harmless.

## Focused test evidence

`Test-Slice4beConfigCommands.ps1 -CheckViewerEventDetail` starts with Admin-generated
disposable warehouses. Typed VBA prepares three synthetic published lines for
one event, including a repeated key, another entity, unlike UOMs, a case-varied
managed header and a leading unknown user column. No canonical inventory rows
are fabricated. Tests select through the actual Viewer ListBox handler and
inspect the owned detail surface. Only fixed checks/Booleans leave the probes;
synthetic screenshots and machine reports stay ignored.

- Selection RED on `validation-viewer-refresh-final`: **2 PASS / 5 FAIL**.
  The event is selected and source bytes remain unchanged, but detail is absent.
- First `validation-viewer-detail` candidate: **7/7**, then **22 PASS / 1 FAIL**
  in expanded checks. The failed immediate restore measurement becomes GREEN
  after dispatching pending layout messages; no runtime layout fix was made.
  Settled default/larger/restore and native maximize/restore give **25/25**.
- Current-state classification expansion: **29 PASS / 2 FAIL**. Existing
  `BOX_DESIGNED` / `SHIP_HELD` codes are selected, but their classification/family
  are missing. Blank IDs already remain separate and Unavailable.
- `validation-viewer-detail-final`: **31/31**, including those corrected source
  labels, unlike-unit line values, all filtered event detail, no selection-time
  projection/profile reads, non-Admin profile access, saved visibility/order on
  explicit Refresh, sign-out line invalidation and source-byte preservation.
- Resize/context expansion: **32 PASS / 2 FAIL**. Sign-out followed by resize
  retains visible detail. The correction validates context during layout and
  makes invalidation idempotent. Its final candidate is
  `deploy/validation-viewer-detail-context`, passes **34/34** with both context
  corrections and every earlier check retained.

The final candidate compiles/cold starts all five packages, retains **16/16**
Refresh-failure and **187/187** Settings checks without retry, and passes the
populated Viewer regression and packaged smoke **86/86**. The full Release 1
chain is **31/31**, with its ordered live-role child **48/48** and successful
runner exit. `slice14_results.md` records the new 2026-09-13 10:12:07 run,
including `OrderedLiveProcessCompleted`, saved-workbook restart, exact identities,
unknown-column preservation, binding and five-package reconciliation.
Native maximize/restore and settled default/larger/
restore geometry pass. The maximized and post-restore default captures were
inspected: all three keys remain separate, fields and Close are readable, and
missing source metadata is labelled. Automated captures are not human UAT.
The immediate restore image is not used as visual acceptance evidence; the later
default capture in the same restored window supplies the inspected evidence.

Static maintenance was regenerated and all three JSON contracts validated:
**201 components / 5,664 procedures / 126,133 lines; 8 literal / 45 unresolved
Application.Run sites; 189 duplicate-body candidates**. All 28 prior oversized
modules retain their line counts. The two new components are at most 183 lines;
the 26 new procedures are at most 40 lines. Functional growth is limited to the
source envelope/profile read, detail ownership/rendering and selection mapping.
Dynamic calls, duplicate groups and oversized modules do not grow.

Compiled component review against `validation-viewer-refresh-final` identifies
four intended existing edits and two new Operations components. Seven additional
raw hash differences are identifier casing only: case-insensitive full code and
case-sensitive string-literal order/content match. The isolated inspection uses
macros disabled, read-only opens and close-without-save; inspected package bytes
remain unchanged. No unexpected compiled-code changes were found.
Final preservation verifies all **150 prior package pins, 15 detail candidate
pins and 16 protected source pins**. Excel is closed. Accepted deployment,
operational workbooks and NAS were not rebuilt or modified.

The 31-check candidate explicitly compiles/cold starts all five packages,
preserves **16/16 Refresh-failure checks**, passes the populated Viewer regression
and packaged smoke **86/86**. Its Settings retry passes **187/187**. The first
Settings attempt stopped at **54 PASS / 1 harness failure** reopening Admin
Settings; no bounded Application-1000 Excel fault was found. The residual process
was recorded during that test and had no ordinary workbooks when inspected.
It was stopped using its recorded ID/start time. No runtime cause or repair is
claimed; the unchanged candidate passed on retry.

Initial PowerShell fixture attempts failed before product RED: mixed-value COM
assignment rejected a numeric cast, then a renamed added column read back as
`Column1`. Typed VBA preparation verified the intended header and three lines.
These are setup failures, not product RED or proof of a runtime header defect.

## Evidence locations

Ignored root: `reports/runtime/`. Focused JSON/captures are in GUID directories
under `slice4be-viewer-detail/`; the 31/31 report is
`80a4bd00f54a404083a0323cb731b224/green.json`; final **34/34** is
`a61e9c8ada94489596ad840e3f87b861/green.json`, with its synthetic PNG captures.

- `viewer-detail-selection-red.log`, `viewer-detail-selection-green.log`.
- `viewer-detail-expanded-green.log`, `viewer-detail-settled-green.log`.
- `viewer-detail-state-red.log`, `viewer-detail-final-green.log`.
- `viewer-detail-resize-context-red.log`.
- `viewer-detail-final-build.log`, `viewer-detail-final-compile.log`,
  `viewer-detail-final-compiled.json`, `viewer-detail-refresh-regression.log`.
- `viewer-detail-settings-regression.log`, `viewer-detail-settings-faults.json`,
  `viewer-detail-settings-retry.log`, `viewer-detail-viewer-regression.log`,
  `viewer-detail-smoke.log`.
- Final candidate: `viewer-detail-context-build.log`,
  `viewer-detail-context-compile.log`, `viewer-detail-context-compiled.json`,
  `viewer-detail-context-green.log`, `viewer-detail-context-refresh.log`,
  `viewer-detail-context-settings.log`, `viewer-detail-context-viewer.log`,
  `viewer-detail-context-smoke.log`, `viewer-detail-context-fullchain.log`,
  `viewer-detail-context-static.log`, and `viewer-detail-static-ratchets.json`.
- `viewer-detail-component-comparison.json`,
  `viewer-detail-package-code-review.json`, `viewer-detail-package-comparison.log`,
  `viewer-detail-check-preservation.json` (all prior 31 check identities retained;
  34 final checks with no duplicates), `viewer-detail-package-pins.json`, and
  `viewer-detail-preservation.json` / `viewer-detail-preservation.log`.

Reproduction commands use `-RepoRoot . -DeployRoot deploy/validation-viewer-detail-context`:

- `tests/tooling/Test-PackagedVbaCompile.ps1` with a local `-SourceReportPath`.
- `tests/tooling/Test-Slice4beConfigCommands.ps1 -CheckViewerEventDetail -CaptureEvidence -Phase GREEN`.
- The same harness with `-CheckViewerRefreshFailure -Phase GREEN`, separately.
- The same harness with `-CheckTrackingSettings -CheckTrackingPolicy -CheckDetailProfile -CheckActionPathPreference -CheckOperationsTrackingSettings -CheckAdminSettingsClose -Phase GREEN`.
- `tools/validate_inventory_viewer.ps1 -CheckTrackingSettings`.
- `tools/validate_phase6_packaged_xlams.ps1`.
- `tools/validate_release1_full_chain.ps1 -KeepArtifacts`.

Run Excel tools serially and only on disposable generated fixtures. Static
maintenance uses `tools/create-maintenance-baseline.ps1 -OutputDirectory reports/static-baseline`;
each JSON is checked against its matching `tools/contracts` schema. This step
does not claim a fresh rerun of the complete earlier Receiving activity matrix;
its packages/source pins and the current full-chain role evidence retain their
distinct scopes.

## Remaining scope

Complete 5,000-group publication, 100-record paging, Designs/activity sources,
verified publication/coverage/outcome metadata, and migration of successful
Shipping supplement reads to owning publication boundaries remain open.
The new detail controls also require their pending activity-catalog coverage.
Show Action Path, recording/conclusions, How-To/Diagnostic/Compare both, guide
authoring/import/export, physical NAS/multi-station and fresh human comparison
remain required. Automated detail evidence does not accept those requirements.
Accepted deployment, operational workbooks and NAS remain untouched; unrelated
handoff 067 and critique 023 are preserved.
