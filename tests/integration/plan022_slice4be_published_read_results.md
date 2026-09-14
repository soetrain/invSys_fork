# Slice 4be.3 authenticated published Events read: RED

Last verified 2026-09-13. Architecture v4.11 D18 governs this test-first entry.
The existing contract requires Core-owned authenticated, policy-aware projection
reads, published-only Viewer actions, explicit unavailable/stale handling, exact
identities and captured-context invalidation. This test changes no runtime,
package, control or architectural rule. Release 1 and Slice 4be remain open.

## Protecting packaged test

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-events-publication-reviewed -Phase RED -CheckViewerPublishedRead
```

Result: **9 PASS / 14 FAIL**,23 unique checks, no harness failure. The runner
exits1 because the product assertions fail. This is the protecting behavioral RED
for the Core/Operations published read; neither missing fixtures nor compile
failures supply the RED.

Admin Generate Warehouse creates disposable fixtures. The actual Settings Save
handler records its correlated attempt/result; the actual Admin Generate Inventory
Snapshot command publishes it. The test validates the saved Events artifact with
the packaged store reader and independently checks the two distinct record IDs,
shared ActivityId and one outcome before opening Viewer. It never replaces the
runtime projection reader or fabricates a canonical business event.

Actual Viewer Open, Events, Refresh, Search and list selection execute. Unsaved
test adapters expose Boolean observations and invoke the existing form handlers;
entry counters observe the existing Shipping supplement and snapshot publisher.
Only disposable published bytes are faulted and restored. Schema and warehouse
faults retain a recomputed valid hash; the store reader rejects each fault before
the Viewer assertion. The missing-file case also leaves the generated legacy
Inventory snapshot present, protecting against an unauthorized fallback.

| Contract | Observed RED |
|---|---|
| Published activity, exact ID, selection and Search | Six assertions fail: Viewer does not load the real Settings ActivityId or its detail, publication time and provenance. |
| Published-only reads | Viewer still enters the Shipping supplement reader. The counter proves boundary entry, not a canonical workbook open in this fixture. |
| Integrity and recovery | Damaged-hash Refresh does not mark Stale or retain the expected activity; restoring the artifact does not load it. |
| Missing/incompatible publication | Schema, warehouse and missing-file cold opens do not report unavailable empty content. |
| Current visibility | Actual Admin visibility Save handlers change policy without republishing. Restoring visibility still fails to load the activity. |

The hide-activity assertion passes, but Viewer already lacks the activity; that
single result cannot establish policy enforcement. Its positive restore case is
required. Both visibility changes preserve the original publication bytes. Target
change and sign-out clear content, Viewer does not publish, and all pinned source,
activity and restored projection files remain unchanged. Only the explicitly
saved fixture policy advances its own Config pin.

## Evidence limits and preservation

Two optional foreground-capture attempts stop with2PASS/11FAIL including a harness
failure. Setting Excel visible did not resolve the foreground rejection. Those
attempts are retained as setup failures, not the complete behavioral RED or visible
acceptance. The complete23-check run omits optional capture. Native layout and
visible operator evidence remain required after implementation.

Ignored evidence under `reports/runtime/`:

- `events-viewer-published-read-red.log`, `events-viewer-published-read-checks.json`
  and `events-viewer-published-read-comparison.json`: complete9/14 with no duplicate
  identities or harness failure.
- `slice4be-viewer-published-read/be3ea3c7b80845f08c52beeba388eb15/red.json`:
  original complete result.
- `events-viewer-published-read-capture-setup-failure.log` and
  `events-viewer-published-read-visible-capture-failure.log`: incomplete capture
  attempts, excluded from the23-check result.
- `events-viewer-published-read-static.log`: regenerated maintenance/schema
  evidence;206 components,5700 procedures,127035 lines,9 literal/45 unresolved
  calls,189 duplicate groups and28 size ratchets remain unchanged.

All175 historical and15 earlier publisher package pins match. The five reviewed
package pins also match. Fifteen protected sources remain exact; the remaining
Shipping source retains only its previously reviewed visibility change. Excel
is closed. No accepted deployment, operational/NAS workbook or unrelated user
change was modified. Existing group/paging7/9 RED remains independently open.

Next implementation must supply the authenticated, current-policy-aware Core read
and Operations consumption of the complete published evidence, preserving legacy
serialized compatibility without canonical Shipping fallback. This test is one
part of that work: all-source coverage, full detail, bounded groups/paging,
recordings, How-To/Diagnostic/Compare both and physical/human acceptance remain
required. The publisher's unresolved native shutdown limitation remains recorded
in [publication evidence](plan022_slice4be_events_publication_results.md).
