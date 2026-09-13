# Slice 4be.3 publication and paged Events: test-first entry

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
