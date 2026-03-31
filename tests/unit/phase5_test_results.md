# Phase 5 VBA Test Results

- Date: 2026-03-31 08:03:36
- Passed: 16
- Failed: 0

| Test | Result |
|---|---|
| TestPhase5Sync.TestRunBatch_WritesOutboxAndSnapshot | PASS |
| TestPhase5Sync.TestRunBatch_SnapshotIncludesCatalogRowsWithZeroQty | PASS |
| TestPhase5Sync.TestRunBatch_SnapshotNormalizesLocationSummaryAndFormatsColumns | PASS |
| TestPhase5Sync.TestManualCopy_PublishesWarehouseArtifacts | PASS |
| TestPhase5Sync.TestWanPublish_OnlineCopy_PublishesLocalArtifactsToSharePoint | PASS |
| TestPhase5Sync.TestWanPublish_OfflineFailure_DoesNotBlockLocalProcessing | PASS |
| TestPhase5Sync.TestWanPublish_SafeRerun_ReplacesPublishedArtifacts | PASS |
| TestPhase5Sync.TestWanPublish_InterruptedReplacement_RestoresPriorArtifactAndAllowsCleanRerun | PASS |
| TestPhase5Sync.TestHqAggregation_TwoWarehousesPreservesPerWarehouseQty | PASS |
| TestPhase5Sync.TestHqAggregation_RebuildsGlobalSnapshotAfterStaggeredWarehouseUpdates | PASS |
| TestPhase5Sync.TestHqAggregation_RepeatedRunsRemainStableForWH1AndWH2Fixtures | PASS |
| TestPhase5Sync.TestHqAggregation_GlobalSnapshotStatusIsAdvisoryOnly | PASS |
| TestPhase5Sync.TestHqAggregation_TempCopyHelper_PreservesReadableCopyWhenPublishedSourceTurnsCorrupt | PASS |
| TestPhase5Sync.TestDelayedPublicationRecovery_PreservesLocalOutboxAndGlobalCatchup | PASS |
| TestPhase5Sync.TestHqAggregation_SkipsUnreadablePublishedSnapshotAndRetainsLastGoodData | PASS |
| TestPhase5Sync.TestHqAggregation_MixedWarehouseInterruption_RetainsLastGoodAndCatchesUp | PASS |
