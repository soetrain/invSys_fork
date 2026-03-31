Implement WAN support in these slices, in order.

| Slice | Primary purpose | Main areas touched |
|---|---|---|
| 1 | Freeze WAN contract in code-facing terms | docs, TODO map, test map |
| 2 | Config and path contract | `src/Core` |
| 3 | Warehouse publish engine | `src/Core`, `src/Admin` |
| 4 | Operator read-model freshness | `src/Core`, role workbooks |
| 5 | HQ aggregation hardening | HQ workbook modules, `src/Admin` |
| 6 | Recovery and interruption handling | `src/Core`, tests |
| 7 | Scheduler and operational wiring | scripts, admin commands, task setup docs |
| 8 | LAN + WAN proving | tests, evidence docs, smoke harnesses |

## Slice 1 status

`Slice 1 complete on 2026-03-30.`

This file now carries the code-facing WAN contract, the remaining TODO map by slice, and the validation/test map that future slices should extend.

## Code-facing WAN contract

### Authority and consistency

- `PathDataRoot` workbooks remain warehouse-authoritative.
- `PathSharePointRoot` is publish/distribution only and is never the processor authority.
- Warehouse-local processing must succeed even when `PathSharePointRoot` is unavailable.
- Cross-warehouse visibility is eventual and driven by published artifacts, not by direct shared editing.

### Artifact contract

- Local outbox path: `<PathDataRoot>\<WarehouseId>.Outbox.Events.xlsb`
- Local snapshot path: `<PathDataRoot>\<WarehouseId>.invSys.Snapshot.Inventory.xlsb`
- SharePoint events target: `<PathSharePointRoot>\Events\<WarehouseId>.Outbox.Events.xlsb`
- SharePoint snapshots target: `<PathSharePointRoot>\Snapshots\<WarehouseId>.invSys.Snapshot.Inventory.xlsb`
- Staged publish suffix: `.uploading`
- Local publish log: `<PathDataRoot>\invSys.Publish.log`

### Publish semantics

- Processor order is local outbox write, local snapshot write, then WAN publish attempt.
- WAN publish copies from local files to SharePoint targets using a staged temp file then final rename.
- Publish failures return warning/report text and publish-log entries, but do not roll back local outbox or snapshot writes.
- Safe rerun behavior is replace-in-place at the SharePoint target path.

### Main entry points

- Core auto publish: `modProcessor.RunBatch`
- Core publish helper: `modWarehouseSync.PublishWarehouseArtifactsToSharePoint`
- Manual admin publish: `modAdminConsole.PublishWarehouseArtifacts`
- HQ rebuild entry: `modHqAggregator.RunHQAggregation`

## TODO map

### Slice 1

- [x] Freeze WAN contract in code-facing terms
- [x] Record code entry points and artifact paths
- [x] Record validation ownership by phase/script

### Slice 2

- [ ] Tighten config/path contract for SharePoint root shape, required subfolders, and path normalization expectations
- [ ] Decide whether bootstrap should create `Events`, `Snapshots`, `Global`, and `Backups`

### Slice 3

- [x] Warehouse publish engine added in core/admin paths
- [x] Non-blocking WAN warning behavior implemented
- [x] Local publish log added

### Slice 4

- [x] Define/read-model freshness markers specifically for WAN-delayed snapshots
- [x] Add operator-visible stale/current status for `LOCAL`, `SHAREPOINT`, and `CACHED` refresh modes

### Slice 5

- [x] Harden HQ aggregation against mixed fresh/stale warehouse sets with explicit advisory metadata
- [x] Add temp-copy/read protections for all WAN-facing admin/HQ flows that open published artifacts

### Slice 6

- [x] Expand interruption recovery around partial publication, stale local copies, and repeated retry windows
- [x] Add recovery cases for mixed warehouse publish failure during HQ catchup

### Slice 7

- [ ] Wire scheduler/admin command surfaces for routine warehouse publish and HQ aggregation
- [ ] Add operator/admin task setup notes for WAN deployments

### Slice 8

- [ ] Add full LAN + WAN proving evidence docs
- [ ] Add smoke harnesses for delayed sync and intermittent connectivity scenarios

## Test map

### Standard validation entrypoints

- Phase 4 admin validation: `tools/run_phase4_excel_validation.ps1`
- Phase 5 warehouse/HQ validation: `tools/run_phase5_excel_validation.ps1`
- Phase 6 runtime/read-model validation: `tools/run_phase6_excel_validation.ps1`

### Current WAN-covered tests

- `TestAdminConsole.TestPublishWarehouseArtifacts_WritesAuditAndPublishesSnapshot`
  - Manual admin WAN publish writes snapshot, publishes to SharePoint path, and audits `PUBLISH_WAN`
- `TestPhase5Sync.TestWanPublish_OnlineCopy_PublishesLocalArtifactsToSharePoint`
  - Processor-driven online publish copies local outbox and snapshot to SharePoint targets
- `TestPhase5Sync.TestWanPublish_OfflineFailure_DoesNotBlockLocalProcessing`
  - SharePoint/path failure keeps local processing authoritative and records publish failure
- `TestPhase5Sync.TestWanPublish_SafeRerun_ReplacesPublishedArtifacts`
  - Repeat publish safely replaces published artifacts without leaving staged `.uploading` files
- `TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSharePoint_UpdatesReadModelAndMetadata`
  - Operator read model refreshes from SharePoint snapshots with `LastRefreshUTC`, `SnapshotId`, `SourceType=SHAREPOINT`, and `IsStale=False`
- `TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSharePoint_StaleSnapshotMarksReadModelStale`
  - Operator read model consumes a stale published snapshot variant and marks `IsStale=True` without becoming unusable
- `TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromCache_PreservesLocalStagingAndLogs`
  - Explicit cached-mode refresh marks stale metadata without mutating workbook-local staging tables or logs
- `TestPhase6CoreSurfaces.TestRefreshInventoryReadModel_MissingSharePointSnapshotMarksCachedWithoutMutatingLocalTables`
  - Missing SharePoint snapshots fall back to cached stale metadata while preserving local operator tables
- `TestPhase5Sync.TestHqAggregation_RepeatedRunsRemainStableForWH1AndWH2Fixtures`
  - HQ aggregation rebuild remains stable across repeated runs against published `WH1` and `WH2` snapshots
- `TestPhase5Sync.TestWanPublish_InterruptedReplacement_RestoresPriorArtifactAndAllowsCleanRerun`
  - Interrupted publish replacement restores the prior published artifact, cleans temporary publish files, and succeeds on deterministic rerun
- `TestPhase5Sync.TestHqAggregation_TempCopyHelper_PreservesReadableCopyWhenPublishedSourceTurnsCorrupt`
  - HQ temp-copy ingest remains readable even if the published source file becomes corrupt after the temp copy is taken

### Remaining WAN-focused proving to add

- Freshness/staleness behavior at operator read-model surfaces
- HQ aggregation behavior under stale-current mixes across more than two warehouses
- Scheduler-driven WAN publish and HQ aggregation smoke coverage
- Backup publication to SharePoint `Backups`

## Slice 5 notes

- HQ aggregation opens only temp copies of published snapshot workbooks, not the SharePoint-published files directly.
- Temp copies are staged under `%TEMP%\invSysHQ_*` and cleaned up best-effort after each aggregation run.
- Published snapshots are opened read-only with workbook automation security forced off during the open call.
- Latest-row selection is still driven by `WarehouseId + SKU + LastAppliedAtUTC`; a stale-but-readable snapshot can remain visible until the newer publish becomes readable.

## Slice 6 notes

- Publish does not retry inside the same processor pass. Failures are reported as publish warnings and publish-log rows, while local outbox/snapshot writes remain authoritative.
- Publish reruns are deterministic: stale `.uploading` work files are discarded, a prior published target is moved to `.previous` during replacement, and that prior target is restored automatically if the replacement is interrupted.
- HQ aggregation reports unreadable published snapshots through `SkippedSnapshotFiles` and skip details; it does not mutate warehouse-authoritative files and can catch up on the next run when a readable publish arrives.
- Operator read-model refresh reports missing or stale published snapshots as `SourceType=CACHED` or stale `SHAREPOINT` metadata; inbox posting remains allowed while stale.
