# WAN Proving Path

**What this file is:** The Codex-facing slice feed for the real two-warehouse WAN proving sequence.
This replaces `LAN WAN development cont.md`, which was scoped to WH1-only local proving and is now deleted.

**Authority:** `invSys-Design-v4.8.md` §Phase 6 and `WAN dev slices.md` (code contract frozen at Slice 1).
**Dependency:** G1 (one-account local) is already proven. Do not re-prove it here.
**Goal:** You manage two warehouses — WH1 and WH2 — and need to operate both over WAN.
WH1 processes locally, publishes to SharePoint. WH2 does the same, independently.
HQ reads only published artifacts from both. None of this requires a live connection at processing time.

---

## INVARIANTS — prepend to every Codex prompt in this series

```
INVARIANTS — do not violate:
- PathDataRoot workbooks are warehouse-authoritative. SharePoint is publish/distribution only.
- Warehouse-local processing must succeed even when SharePoint is unreachable.
- WH1 and WH2 are independent. No shared live tables. No cross-warehouse direct reads.
- Publish is replace-in-place with staged .uploading rename — never partial state.
- HQ reads only published snapshot artifacts from SharePoint. It never connects live to a warehouse.
- Operator read-model refresh modes: LOCAL (from local snapshot), SHAREPOINT (from published snapshot), CACHED (fallback, marks stale).
- CheckReceivingReadiness shows actionable messages — never raw error codes or silent failures.
- Two-warehouse HQ snapshot is advisory only — it has no write-back path and is never authoritative.
- The WAN proving path contract (this file) is the acceptance target for all slices in this series.
```

---

## Architecture recap (from WAN dev slices.md §Code-facing WAN contract)

**Artifact paths per warehouse (substitute WH2 for WH1 identically):**

```
Local authoritative:
  <PathDataRoot>\<WarehouseId>.Outbox.Events.xlsb
  <PathDataRoot>\<WarehouseId>.invSys.Snapshot.Inventory.xlsb

SharePoint published:
  <PathSharePointRoot>\Events\<WarehouseId>.Outbox.Events.xlsb
  <PathSharePointRoot>\Snapshots\<WarehouseId>.invSys.Snapshot.Inventory.xlsb
```

**Warehouse hub root contract (v4.8 alignment):**

```
RULE: <PathDataRoot> is the warehouse hub root for that warehouse.
It may be a local path during development, but the preferred LAN + WAN deployment
target is a NAS-backed path owned by that warehouse, for example:

  \\DS920\invSys\WH1\
  \\DS920\invSys\WH2\

The designated processor PC for that warehouse reads/writes the authoritative
runtime over SMB. WAN stations do not directly edit these canonical workbooks;
they relay events into the same processor lane through SharePoint publication.
```

**Key entry points already in code:**
```
modProcessor.RunBatch                                       — local process + auto publish attempt
modWarehouseSync.PublishWarehouseArtifactsToSharePoint      — publish helper
modAdminConsole.PublishWarehouseArtifacts                   — manual admin publish
modHqAggregator.RunHQAggregation                           — HQ rebuild from published artifacts
```

---

## Proving Slice A — Real-machine setup: WH1 on Machine 1

**Goal:** WH1 is fully bootstrapped on a physical machine, processes locally, and publishes to SharePoint.

**Prompt Codex:**
```
Create tests/integration/prove_wan_wh1_setup.bas.
This is a real-machine setup verification script for WH1, not a unit test.

SetupVerification_WH1() steps:
  1. Resolve PathDataRoot for WH1 and assert the warehouse runtime exists there (BootstrapWarehouseLocal must have been run already)
  2. Assert WH1.invSys.Data.Inventory.xlsb exists and is non-zero
  3. Assert WH1.Outbox.Events.xlsb exists (may be empty — presence check only)
  4. Assert WH1.invSys.Snapshot.Inventory.xlsb exists
  5. Assert PathSharePointRoot is set and the SharePoint root folder is reachable
  6. Assert <PathSharePointRoot>\Events\ folder exists
  7. Assert <PathSharePointRoot>\Snapshots\ folder exists
  8. Call modProcessor.RunBatch — assert it returns OK with no fatal errors
  9. Assert the published snapshot exists at <PathSharePointRoot>\Snapshots\WH1.invSys.Snapshot.Inventory.xlsb
  10. Assert no .uploading temp file remains at the target path

WriteProofResult(machineName As String, step As Integer, passed As Boolean, note As String):
  Appends one row to tests/integration/wan-wh1-setup-proof.md in the format:
  | <machineName> | <step> | <PASS/FAIL> | <note> | <UTC timestamp> |

Done when: all 10 steps PASS on the real machine and the results file is committed.
```

---

## Proving Slice B — Real-machine setup: WH2 on Machine 2

**Goal:** WH2 is independently bootstrapped on a second physical machine, processes, and publishes.

**Prompt Codex:**
```
Create tests/integration/prove_wan_wh2_setup.bas.
Identical structure to prove_wan_wh1_setup.bas but for WarehouseId = WH2 on a second machine.

All paths use WH2 substitution:
  <PathDataRoot for WH2>\
  WH2.invSys.Data.Inventory.xlsb
  WH2.Outbox.Events.xlsb
  WH2.invSys.Snapshot.Inventory.xlsb
  <PathSharePointRoot>\Snapshots\WH2.invSys.Snapshot.Inventory.xlsb

Steps 1–10 are identical to Slice A with WH2 substitution.
Results written to tests/integration/wan-wh2-setup-proof.md.

Additional step 11:
  Assert that after WH2 publishes, the WH1 published snapshot at
  <PathSharePointRoot>\Snapshots\WH1.invSys.Snapshot.Inventory.xlsb
  is still present and unmodified (cross-contamination check).

Done when: all 11 steps PASS on the second real machine and results are committed.
```

---

## Proving Slice C — WAN operator flow: delayed publish and stale refresh

**Goal:** Each warehouse operator can work locally when SharePoint is unreachable, then catch up on publish when it returns.

**Prompt Codex:**
```
Create tests/integration/prove_wan_operator_flow.bas.
This proves the operator WAN flow for both WH1 and WH2.

Test cases (run on each warehouse machine separately):
  C1 — LocalProcess_NoSharePoint:
    - Disconnect SharePoint (rename PathSharePointRoot in config to unreachable path)
    - Run modProcessor.RunBatch
    - Assert: returns OK (not fatal error), local outbox and snapshot updated
    - Assert: publish failure is logged to invSys.Publish.log, not raised as an error
    - Assert: operator read-model refresh mode falls back to CACHED, marks IsStale=True

  C2 — StaleOperatorRefresh_ShowsActionableMessage:
    - Open WH1.Receiving.Operator.xlsm with SharePoint unreachable
    - Call CheckReceivingReadiness
    - Assert: SnapshotStatus = "STALE", IsReady = False
    - Assert: status panel visible, message contains "Refresh Inventory" instruction
    - Assert: inbox posting is still allowed (stale snapshot must not block posting)

  C3 — CatchUpPublish_AfterConnectivityReturns:
    - Restore PathSharePointRoot to the real path
    - Call modAdminConsole.PublishWarehouseArtifacts
    - Assert: published snapshot at SharePoint target is newer than the stale one
    - Assert: no .uploading file remains
    - Assert: operator read-model refresh now returns SnapshotStatus = "OK" when using SHAREPOINT mode

  C4 — WH2_LocalProcess_IndependentOfWH1:
    - On WH2 machine: run C1–C3 identically for WH2
    - Assert WH1 published artifacts at SharePoint are unchanged throughout WH2 operations

Results written to tests/integration/wan-operator-flow-proof.md.
Done when: all four cases PASS on both warehouse machines.
```

---

## Proving Slice D — HQ aggregation from two published warehouses

**Goal:** HQ reads only published WH1 and WH2 snapshot artifacts, never a live warehouse connection.
The global snapshot is advisory — it preserves per-warehouse rows and has no write-back path.

**Prompt Codex:**
```
Create tests/integration/prove_wan_hq_aggregation.bas.
This proves the full HQ aggregation flow from real published artifacts.

Prerequisite: Slices A, B, C must all be PASS before running this slice.

Test cases:
  D1 — HqAggregation_FromBothPublishedSnapshots:
    - Assert WH1 snapshot exists at <PathSharePointRoot>\Snapshots\WH1.invSys.Snapshot.Inventory.xlsb
    - Assert WH2 snapshot exists at <PathSharePointRoot>\Snapshots\WH2.invSys.Snapshot.Inventory.xlsb
    - Call modHqAggregator.RunHQAggregation
    - Assert: returns without fatal error
    - Assert: global snapshot contains rows for both WH1 and WH2 (WarehouseId column has both values)
    - Assert: per-warehouse row counts match the individual published snapshots
    - Assert: no row has a write-back path or authority flag

  D2 — HqAggregation_StaggeredRepublish:
    - Publish WH1 again (modAdminConsole.PublishWarehouseArtifacts on WH1 machine)
    - Run RunHQAggregation again
    - Assert: global snapshot reflects the newer WH1 data
    - Assert: WH2 rows are unchanged and still present
    - Assert: run is idempotent — running again produces the same result

  D3 — HqAggregation_OneWarehouseStale:
    - Simulate WH2 snapshot is stale (age it past threshold in config)
    - Run RunHQAggregation
    - Assert: WH2 rows are included but marked with advisory stale metadata
    - Assert: WH1 rows are unaffected
    - Assert: SkippedSnapshotFiles is empty (stale-but-readable is not skipped, only unreadable is)

  D4 — HqAggregation_OneWarehouseUnreadable:
    - Rename WH2 published snapshot to simulate missing file
    - Run RunHQAggregation
    - Assert: WH1 rows still present in global snapshot
    - Assert: SkippedSnapshotFiles contains WH2 entry
    - Assert: no fatal error — HQ degrades gracefully
    - Restore WH2 snapshot, rerun, assert WH2 rows return

Results written to tests/integration/wan-hq-aggregation-proof.md.
Done when: D1–D4 all PASS and results are committed with machine names and timestamps.
```

---

## Execution order

```
Slice A (WH1 machine setup)    →  wan-wh1-setup-proof.md
Slice B (WH2 machine setup)    →  wan-wh2-setup-proof.md
Slice C (operator WAN flow)    →  wan-operator-flow-proof.md       ← depends on A and B
Slice D (HQ aggregation)       →  wan-hq-aggregation-proof.md      ← depends on A, B, C
```

Do not run Slice D until C is fully green.
Do not claim WAN proving complete until all four results files exist with PASS on all cases.

---

## What "done" means

WAN is proven when:
- `wan-wh1-setup-proof.md` — all steps PASS, real machine, real SharePoint path
- `wan-wh2-setup-proof.md` — all steps PASS, second real machine, real SharePoint path
- `wan-operator-flow-proof.md` — C1–C4 PASS on both machines
- `wan-hq-aggregation-proof.md` — D1–D4 PASS
- No result row is marked "simulated" or "assumed"

---

*Last updated: 2026-04-07. Authoritative source: `invSys-Design-v4.8.md` §Phase 6, `WAN dev slices.md` §Code-facing WAN contract.*
