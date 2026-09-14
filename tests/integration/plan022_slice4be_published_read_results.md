# Slice 4be.3 authenticated published Events read: candidate

Last verified 2026-09-13. Architecture v4.11 D18 governs this test-first entry.
The existing contract requires Core-owned authenticated, policy-aware projection
reads, published-only Viewer actions, explicit unavailable/stale handling, exact
identities and captured-context invalidation. This test changes no runtime,
package, control or architectural rule. Release 1 and Slice 4be remain open.

## Current implementation checkpoint

The subsequent paging candidate passes the expanded reader **30/30**. Actual
Day/All checks first expose the UTC/local comparison at24/1 RED; real-form layout
checks then expose status clipping at25/4 RED. A mixed-source serialized fixture
through the actual Viewer handlers establishes29/1 RED for equal-time ordering.
The corrected comparator uses the declared SourceKind, exact ID and source,
matching D18 and the publisher. Five-package compile and cold start pass.
The same candidate also passes group/page16/16, stale Refresh16/16 and Detail34/34,
retaining the exact prior regression identities; static call/duplicate/size ratchets
hold and protected package/source pins match with Excel closed.
See [paging evidence](plan022_slice4be_event_groups_results.md) for complete scope,
excluded setup failures and pending broader gates. No broad acceptance is inferred.

The isolated `validation-events-reader-compiled` candidate now passes **23/23**
focused checks, preserving all23 RED identities. All five packages compile and
Operations resolves its cold-start dependencies inside this candidate. Refresh
**16/16** and Event Detail **34/34** remain GREEN, including native detail sizing,
repeated exact keys, unlike units, profile changes and captured-context guards.
These gates do not establish full Events, paging, shipping-state presentation,
recording, Action Path or Release 1 acceptance. Runtime changes remain uncommitted
until the required broader gates and maintenance review are complete.

Core `modPublishedEventsReader` reads only the captured warehouse's validated
Events artifact and current policy. It filters activity groups using the saved
catalog/control visibility and Admin visibility gate, retaining every allowed
line. The existing Core public read delegates to it. D18's `EVENTS1` primitive
wire carries separate publication/load timestamps, coverage and named detail
values after the eighteen compatibility slots. Operations parses those values;
its Viewer reader no longer invokes Shipping supplement authority. Supported
legacy serialized payloads remain readable without that fallback.

The first isolated build failed Core compilation because a local variable
shadowed the family helper. The corrected candidate compiles all five packages;
this failure is not counted as behavioral RED. No accepted deployment is rebuilt.

Historical Detail/Groups fixtures populate a disposable published XLSB. For a
candidate containing the new reader, `Slice4bePublishedProjectionFixture.ps1`
feeds those same lines, through a read-only workbook, to the real Core publisher.
It validates the resulting artifact and verifies the source bytes remain exact.
Older package baselines retain their old fixture path. The runtime reader is
never replaced; the fixture change preserves existing assertions and supplies
the now-required published format. The current-state legacy payload cases in
the Detail suite still test supported serialized compatibility.

Candidate evidence under `reports/runtime/`:

- `events-reader-initial-build.log`, `events-reader-initial-compile.log`: initial
  compile failure, excluded from acceptance.
- `events-reader-compiled-build.log`, `events-reader-compiled-compile.log`,
  `events-reader-compiled-sources.json`: fresh package/compile evidence.
- `events-reader-focused-green.log`, `events-reader-focused-green.json`:23/23.
- `events-reader-refresh.log`, `events-reader-detail.log`:16/16 and34/34.
- `events-reader-package-pins.json`: five frozen candidate file hashes.
- `events-reader-groups-red.log`, `events-reader-groups-red.json` and
  `events-reader-group-comparison.json`:8PASS/8FAIL; all16 identities and all seven
  previous GREEN results retained, no duplicate checks. The read boundary now
  passes. Paging, page counts/navigation and unlike-unit summaries still fail.
- `events-reader-static.log`:207 components/5709 procedures/127173 lines;
 9 literal/45 unresolved calls,189 duplicate groups and28 size ratchets unchanged.
 The reader adds138 net lines against the preceding candidate. Four additional
 maintenance candidates require review; obsolete helpers are not accepted debt.

Maintenance review must remove or justify the retired XLSB-reader helpers after
reviewing their remaining source-test references. Shipping current-state labels
and complete field presentation still need comparison with the accepted package;
the23-check activity test does not protect those behaviors. Group/paging proving,
full role/chain gates and visible capture remain outstanding.

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
