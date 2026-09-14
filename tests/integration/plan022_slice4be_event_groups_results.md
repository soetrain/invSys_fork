# Slice 4be.3 publication and paged Events: candidate

## Current paging checkpoint

The isolated `validation-events-paging-utc` candidate passes **16/16** packaged
group/page checks, advancing the reader candidate's8/8 RED. Previous/Next and
matching/available counts now work; Search starts on page1. The boundary summary
shows Multiple for quantity and unlike units while selecting it retains all three
source lines and the repeated exact key. No previous check identity or GREEN is
intentionally removed. Core publication/read boundaries remain unchanged.

Operations `cEventPageProjection` groups the loaded rows and caches searchable
text. Its indexes are local UI references, never persisted inventory identity.
The form renders100 matching groups and retains the first source-line index only
as the selection entry into the complete existing Event Detail controller.
Records without durable source IDs remain separate; no fake state identity is
generated. The list reserves a bottom navigation strip and its controls anchor
with resize. Full visible and mixed-source acceptance remains open.

An added actual Day/All handler test finds **24 PASS / 1 FAIL**: fresh verified
UTC activity disappears in Day because the filter uses local Now. The correction
compares verified UTC groups with UTC, preserving the explicitly labelled local
approximation for unverified historical times. The expanded reader suite then
passes **25/25**. Refresh **16/16** and Detail **34/34**, including native detail
sizing and context guards, remain GREEN. All five packages compile and cold-start
Operations resolves dependencies inside the candidate.

Expanded real-form geometry checks first record **25 PASS / 4 FAIL**. The status
label extends 1.4 points below the client area at all four sizes; the page controls
fit. Moving the status label up four points preserves its height and passes minimum,
default, larger and restored-default geometry without relaxing the test tolerance.
These programmatic form-resize checks do not establish foreground or human UAT.

The `validation-events-paging-layout` candidate then records **29 PASS / 1 FAIL**:
all25 prior GREEN checks and four layout checks pass, while mixed-source equal-time
ordering fails. An in-memory synthetic `EVENTS1` envelope is supplied at the Core
serialized read boundary; actual Viewer Events/Refresh handlers order its rows.
Inventory and Designs IDs deliberately oppose source-name ordering, and Activity
must follow both business events. No canonical owner or saved artifact is replaced.
The real Settings/publication fixture, policy and context checks remain intact.

The correction reads SourceKind by the declared field name, then compares kind,
exact SourceId and Source after recorded time, matching D18 and the publisher.
The isolated `validation-events-paging-order` candidate compiles all five packages,
passes the Operations cold start and completes the expanded reader **30/30 GREEN**.
The same30 check identities advance from29/1 RED. Group/page **16/16**, stale Refresh
**16/16** and Event Detail **34/34** pass on that same candidate; each retains the
exact prior check set with no duplicates. All four child gates exit0 and Excel is
closed. Static regeneration completes at208 components/5721 procedures/127394 lines,
9 literal/45 unresolved calls,189 duplicate groups and28 oversized-module ratchets.
The ordering correction adds20 lines; no call, duplicate or size ratchet increases.
Shipping state presentation, remaining filters, maintenance, full role/chain and
visible gates remain open.

Two setup failures precede the complete29/1 RED and are excluded from it. A new
test variable was appended after a procedure, causing a VBA declaration-placement
compile error; it now enters the declaration section. A subsequent bootstrap
failure left an empty test-child Excel instance, closed without saving with all
three loaded startup add-in hashes unchanged. Neither failure justifies a runtime
patch or proves a native crash cause.

Ignored candidate evidence under `reports/runtime/`:

- `events-paging-build.log`, `events-paging-compile.log`: initial paging package.
- `events-paging-date-red.log`: complete24/1 UTC-filter RED, no harness failure.
- `events-paging-utc-build.log`, `events-paging-utc-compile.log`,
  `events-paging-utc-compiled.json`: corrected five-package build/compile.
- `events-paging-utc-groups.log`, `events-paging-utc-reader.log`,
  `events-paging-utc-refresh.log`, `events-paging-utc-detail.log` and
  `events-paging-utc-gates.json`:16,25,16,34 passing checks with child exit0.
- `events-paging-static.log`:208 components/5721 procedures/127374 lines,
 9 literal/45 unresolved calls,189 duplicate groups and28 size ratchets.
 Dynamic-call, duplicate and oversized-module ratchets do not grow; obsolete
 reader helpers and remaining maintenance candidates still require review.
- `events-paging-layout-first-failure.log`, `events-paging-layout-geometry.log`:
  complete25/4 layout RED and measured geometry.
- `events-paging-mixed-order-red.log`, `events-paging-order-red-comparison.json`:
  complete29/1 RED; all25 prior GREEN identities retained, no duplicate checks.
- `events-paging-mixed-order-harness-failure.log`,
  `events-paging-order-bootstrap-failure.log`,
  `events-paging-order-bootstrap-recovery.json`: excluded setup failures/recovery.
- `events-paging-order-build.log`, `events-paging-order-compile.log`,
  `events-paging-order-compiled.json`: corrected isolated build and compile.
- `events-paging-order-reader.log`, `events-paging-order-groups.log`,
  `events-paging-order-refresh.log`, `events-paging-order-detail.log`,
  `events-paging-order-gates.json`:30/16/16/34 GREEN, all child exits0.
- `events-paging-order-green-comparison.json`,
  `events-paging-order-regression-comparison.json`: unchanged check identities.
- `events-paging-order-static.log`: completed final maintenance regeneration.
- `events-paging-order-preservation.log`:175 historical/15 publisher package pins,
  15 exact source pins and one reviewed Shipping visibility-only change preserved.
  The five previous reader package hashes also match. No accepted deployment or
  operational/NAS workbook is changed. Runtime remains an uncommitted candidate
  pending the broader implementation gates; this checkpoint records tests/evidence.

The original test-first evidence below retains its historical baseline and scope.

Last verified 2026-09-13. Runtime baseline is code `e66ab6f`, packaged as
`deploy/validation-viewer-detail-context`; the preceding selected Event Detail
checkpoint remains GREEN. This record establishes missing D18 behavior. It does
not claim new runtime implementation, completed publication or user acceptance.

## Governing contract

Architecture v4.11 D18 requires published-only Viewer reads, 100 matching
complete source-event groups per page, every contributing detail line, separate
unlike units, deterministic timestamp/source/ID ordering and explicit coverage.
The publisher independently retains the newest 5,000 complete durable groups.
Plan 022 and controls v1.110 record the discovered Previous/Next/page-count
controls before implementation. This is semantic inheritance of those approved
rules, not a new source authority or a relaxed publication contract.

## Focused packaged RED

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-viewer-detail-context -CheckViewerEventGroups -CaptureEvidence -Phase RED
```

Result: **7 PASS / 9 FAIL**, runner exit 1 because the nine product assertions
failed. The fixture and actual Viewer handlers completed; no compile, missing
fixture or broken-harness failure is counted as behavioral RED.

Admin Generate Warehouse creates the disposable runtime. Typed VBA populates
only its published inventory-events snapshot with 5,003 lines / 5,001 source
groups. Storage order differs from descending recorded time. The group at the
100-record page boundary has three contributing lines, two exact keys with one
repeated, and EA/LB units. A leading unknown column and case-varied managed key
header protect normalized lookup. No canonical inventory rows are fabricated.

Actual Viewer Events/Refresh/Search and ListBox selection execute. Paging probes
set the discovered CommandButton's Value so the real Click handler must run;
absence of those operator controls returns false rather than introducing a
missing test procedure or compile failure. Once implemented, independent native
and layout evidence must supplement these action checks.

| Check group | Result / implication |
|---|---|
| Published fixture loaded | PASS: the real Events handler loads the volume fixture. |
| Sorted first 100 complete groups | FAIL: the existing list exposes lines in reverse storage order. |
| Page/matching count and Previous boundary | FAIL: the controls/count are absent. |
| Next and Previous actions | FAIL: no page navigation exists. |
| Refresh avoids Shipping authority-read boundary | FAIL: successful Refresh enters the existing Box Design source reader. |
| Search matches a contributing unit | PASS: LB locates the boundary group. |
| Group summary preserves unlike-unit semantics | FAIL: the visible result still selects one line's quantity/unit. |
| Selected boundary retains every exact-key line | PASS: the selected-event detail retains all three lines despite filtering. |
| Single-group search | PASS. |
| Next disabled for one match; cleared Search resets first page | FAIL: paging is absent. |
| Sign-out invalidation and source-byte preservation | PASS. |

The source-read probe counts entry to `AppendBoxDesignViewerEvents` during an
explicit Refresh. Source inspection shows that this path resolves and can open
the canonical Shipping BOM workbook; the counter does not separately prove an
actual workbook open in this run. It protects exclusion of that authority-read
boundary from Viewer, including when the source is absent. Read-only byte hashes
alone cannot establish D18's stricter published-only read contract.

The inspected synthetic capture shows 5,003 displayed event lines, repeated
boundary references, reverse storage order and no paging controls. It is RED
evidence, not layout or human acceptance. The oversized published input exercises
Viewer behavior only: it does not call or establish the publisher's 5,000-group
selection rule. A separate owning publication-boundary test remains required.

## Evidence and preservation

Ignored local artifacts:

- `reports/runtime/viewer-groups-initial-red.log`.
- `reports/runtime/slice4be-viewer-groups/dedd7010defb4b268a0f494d0a4dd229/red.json`.
- The same directory's `viewer-groups.png`.
- `reports/runtime/viewer-groups-preservation.log`: all 165 existing package
  pins and 16 protected source pins match after the RED run; Excel is closed.

Both changed PowerShell files parse and diffs pass whitespace checks. No runtime
VBA, package, accepted deployment, operational workbook or NAS change is made in
this entry. Existing compile/layout/static/live-role/full-chain GREEN retains
the scope and dates in the preceding
[Event Detail evidence](plan022_slice4be_event_detail_results.md); those gates
are not represented as rerun for an implementation that does not yet exist.

## Immediate continuation

Protect the owning publication boundary, then implement the published Events
artifact and group-level paging against these failures. Preserve all 34 detail,
16 Refresh-failure and 187 Settings checks and complete the required release
gates on the resulting candidate. Do not remove the Shipping reader without a
published replacement for accepted current-state supplements. Source coverage,
verified publication metadata, all Operation/Admin activity, recording, both
Action Path presentations, comparison/import/export and human acceptance remain
part of the full goal.
