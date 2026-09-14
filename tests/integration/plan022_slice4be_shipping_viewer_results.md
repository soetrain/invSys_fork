# Slice 4be.3 Shipping current-state Viewer presentation

Last verified 2026-09-13. Architecture v4.11 D18 governs published-only reads,
accepted Shipping summaries, complete detail and unavailable historical identity.
The prior paging checkpoint remains [separately recorded](plan022_slice4be_event_groups_results.md).
Release 1 and Slice4be acceptance remain open.

## Focused RED

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-events-paging-order -Phase RED -CheckViewerShippingState
```

Result: **50 PASS / 6 FAIL**,56 unique checks, no harness failure. All30 prior
reader GREEN identities remain. The Shipping extension runs after those checks;
it does not replace the positive publication/visibility or rejection assertions.

Admin Generate Warehouse and Seed create disposable inventory. Actual Box Designer
and Box Maker handlers prepare one package/alternative with two distinct component
keys, and Add/Hold prepares a held shipment. Existing calibrated fixture checks
prove saved owner rows and the durable Box Maker event. The actual Admin publication
command then writes the Events artifact; the store reader validates it and the
existing Shipping publication checks prove every permitted source line and scope.

Viewer Open/Events/Refresh, Search and list selection execute through their real
handlers. Test adapters return only in-memory display observations. Owner values
remain in memory; maintained evidence contains fixed check names and counts.

| Contract | RED result |
|---|---|
| One package/alternative summary | FAIL: the Viewer shows both component rows separately. |
| Accepted package summary values | FAIL: component item/UOM/quantity replace the package summary, and its label changes. |
| Complete selected component detail | FAIL: selecting a state row exposes only that component. |
| Component search retains package summary | FAIL: search displays the component as the summary. |
| Filtered selection retains both component lines | FAIL: only the matching component remains selected. |
| Held-shipment summary | FAIL: the state label differs from the accepted SHIP_HELD label. |
| State selection/provenance | PASS: actual list handlers execute; source event identity remains unavailable. |
| Held key and outcome | PASS: exact inventory key remains, current state is labelled and no completed shipment outcome is inferred. |
| Read and byte boundaries | PASS: Viewer does not enter the probed Shipping supplement or publisher boundaries; owner/projection/profile-local source bytes remain unchanged. |

## Contract mapping and candidate

D18 now explicitly records the compatible presentation mapping before implementation.
The existing named BomId/BomVersion fields carry the exact PackageSystemKey and
owning BomVersion as an association for current-state display. SourceId stays
unavailable. Package item/UOM/location and alternative remain the summary, with
no fabricated package quantity; named detail values remain each component's own
key, item and quantity/UOM. Related BOM lines can share a summary without grouping
unrelated blank IDs or inventing an event identity. No canonical schema, display
field allowlist, publication limit or write authority changes.

Core's reader applies that summary/detail mapping. Operations uses one loaded
group-membership projection for both page summaries and selected detail. Unknown
or older metadata retains separate blank-ID rows. The isolated
`validation-events-shipping-state` candidate compiles all five packages and passes
the Operations cold start. Focused GREEN is **56/56**, retaining the complete RED
check set. Group/page **16/16**, stale Refresh **16/16** and Event Detail **34/34**
pass on the same candidate with exact prior check identities. Static regeneration
completes at208 components/5722 procedures/127420 lines;9 literal/45 unresolved
calls,189 duplicate groups and28 oversized-module ratchets remain unchanged.

The full Release1 validator exits0 with **31/31**, ordered live roles **48/48**
and Create Warehouse source integration **15/15**. Windows nevertheless records
one Excel `combase.dll` / `c0000005` fault during the run, about24 seconds after
the live-role report and before the final chain report. This timing does not
identify the cause. Clean native execution remains unproven; no Office, credential,
permission or speculative runtime repair follows. The final check report alone
does not close that gate or establish complete Slice4be acceptance.

Excel is closed. All five candidate package hashes match their pre-chain values;
the three pre-existing generated report files were restored. Preservation matches
175 historical and15 publisher package pins,15 exact source pins and one reviewed
Shipping visibility-only change, plus all20 prior paging package pins. Runtime
implementation remains uncommitted pending broader acceptance and maintenance;
this checkpoint commits the protecting test and synchronized contract/evidence.

## Evidence and limits

Ignored artifacts under `reports/runtime/`:

- `events-shipping-viewer-red.log`, `events-shipping-viewer-red-comparison.json`:
  complete50/6 RED with all30 previous reader GREEN retained and no duplicate IDs.
- `events-shipping-state-build.log`, `events-shipping-state-compile.log`,
  `events-shipping-state-compiled.json`: isolated build and five-package compile.
- `events-shipping-state-shipping.log`, `events-shipping-state-groups.log`,
  `events-shipping-state-refresh.log`, `events-shipping-state-detail.log`,
  `events-shipping-state-gates.json`, `events-shipping-state-check-comparison.json`:
  56/16/16/34 GREEN and identical prior check sets.
- `events-shipping-state-static.log`: completed maintenance/schema regeneration.
- `events-shipping-state-chain.log`, `events-shipping-state-chain.md`,
  `events-shipping-state-chain-live.md`, `events-shipping-state-chain-create.md`,
  `events-shipping-state-chain-summary.json`: full-chain and preservation scope.
- `events-shipping-state-chain-native-faults.json`,
  `events-shipping-state-chain-fault-timing.json`: unresolved native fault.
- `events-shipping-state-preservation.log`: protected package/source checks.

The Shipping read counter proves absence of the probed owner-read entry, not every
possible workbook open. Source inspection and broader boundary gates remain relevant.
The fixture produces a held shipment, not a completed shipping sequence or Action
Path conclusion. Full role/chain, foreground operator evidence, recordings, guide
comparison, physical NAS/multi-station and human acceptance remain required. Earlier
native shutdown failures remain unresolved in the
[publication evidence](plan022_slice4be_events_publication_results.md).
