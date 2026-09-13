# Slice 4be.3 mixed-source publication RED

Last verified 2026-09-13. Release 1 and Slice 4be remain active. This is a
test/contract-detail checkpoint, not publisher implementation or acceptance GREEN.

## Contract and protecting test

Architecture v4.11 D18 now specifies exact ActivityId grouping, original RecordIds,
actual result records, earliest contributing chronology and named source/count
coverage. Plan 022 and controls v1.112 carry the same scope. These details implement
the approved observation, provenance and complete-group rules under semantic
inheritance; they do not authorize new authority writes or accept partial coverage
as complete. ShippingHolds is local current state, not warehouse-wide Hold history.

The publication test retains its 5,001 inventory groups / 5,003 lines fixture and
now invokes the real Admin Settings Save Value handler before the actual public
Generate Inventory Snapshot command. The handler creates one correlated activity
attempt/result pair in the generated warehouse's existing activity store. The test
requires that pair to publish as one group with both exact original RecordIds and
timestamps, the actual result/data effect, and the publisher's package/policy
provenance. It independently reads Core's immutable package metadata from the
XLAM archive. It also pins the activity files against publication-time mutation.

At the global 5,000-group bound, the controlled mixture must retain the activity
group and 4,999 inventory groups. Inventory coverage must reconcile 5,001/4,999/2
available/included/omitted groups and 5,003/5,001/2 lines; Activity must reconcile
1/1/0 groups and 2/2/0 lines. Inventory, Designs, Activity, ShippingBOM and
ShippingHolds require separate named coverage entries, availability and scope.
Counting anonymous coverage entries is insufficient. This does not yet prove
actual Designs or Shipping data publication, atomic replacement, or reader policy
enforcement; those remain independent required tests and implementation work.

## Observed results and limits

- Two combined attempts stopped before publication at the existing
  `modInventoryViewer.ViewerGroupsActionForTest` call: each **0 PASS / 1 harness
  failure**, `0x80020009`. No new publication behavioral RED is claimed from them.
  The first bounded Application-1000 query found no Excel fault; no cause is
  inferred. Each residual Excel process was verified empty and its identity
  checked before cleanup. No operational workbook was closed or changed.
- The first publication-only attempt failed before assertions with an object
  reference error. Its new direct COM document-property probe was replaced with
  the established read-only XLAM archive metadata technique. No product repair
  or behavioral RED is inferred from that fixture failure.
- The corrected publication-only run completes **6 PASS / 15 FAIL**, without a
  harness exception or duplicate check identity. The real Settings handler
  produces the pair, Admin publication succeeds, the Inventory source is read-only
  and released, and canonical/source-copy/activity bytes are preserved. All four
  source-read correction checks remain GREEN. The missing Events artifact and
  its fifteen publication assertions are meaningful RED.

The diagnostic-only flag remains prohibited from claiming acceptance GREEN. The
combined Viewer/publication gate still must run; its earlier accepted grouped
7 PASS / 9 FAIL result was not re-proven by either failed combined attempt here.
Do not substitute this diagnostic result for paging, source completeness or UI
acceptance. Prior source-read evidence is in
[publication/source-read results](plan022_slice4be_publication_results.md).

## Reproduction and preservation

Candidate: `deploy/validation-publication-source-read`, built from the runtime
correction committed as `ad08ea9`. No runtime source, package, build/deployment
script or static VBA metric changes in this checkpoint. The earlier package
compile/layout/Receiving 854/854/Settings 187/187/smoke 86/86/chain 31/31 with
live-role 48/48 evidence retains its exact unchanged-package scope; none is
represented as rerun for this test-only change.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-publication-source-read -ViewerPublicationOnly -Phase RED
```

Ignored `reports/runtime/` evidence:

- `viewer-publication-mixed-source-red.log`,
  `viewer-publication-mixed-source-retry-red.log` (combined setup failures).
- `viewer-publication-mixed-source-faults.json`,
  `viewer-publication-mixed-source-residual.json`,
  `viewer-publication-mixed-source-retry-residual.json` (bounded diagnostics/cleanup).
- `viewer-publication-mixed-source-diagnostic-red.log` (metadata probe failure),
  `viewer-publication-mixed-source-diagnostic-retry-red.log` (6 PASS / 15 FAIL).
- `slice4be-viewer-groups/93670bbebf094e9dbd0049b682f41b5b/diagnostic-publication-red.json`,
  that directory's `publication-source-preservation.json`, and
  `viewer-publication-mixed-source-result-index.json` identify the actual result.
- `viewer-publication-mixed-source-preservation.log` verifies all **170 package
  pins** and **16 protected source pins**, with Excel closed.

The changed PowerShell test parses; document links and diffs/status are checked.
Unrelated modified handoff 067 and untracked critique 023 remain unstaged.

## Immediate implementation work

Implement Core's persisted Events publication against these REDs while preserving
the source-read correction. Add protecting source-population and non-mutation
tests before adding the Designs and Shipping read boundaries: Designs' current
resolver can create/ensure/save its workbook, and Shipping's current supplements
read canonical BOM and profile-local hold state. Neither may become a Viewer
authority-read fallback. Complete atomic/schema/integrity failure coverage and
the combined grouped Viewer gate before claiming publisher or paging acceptance.
Comprehensive activity, recording, How-To/Diagnostic/Compare both, guide library,
NAS/multi-station and human acceptance remain required.
