# Phase 6 LAN + WAN Proving Results

- Date: 2026-03-31 08:05:31
- Passed: 24
- Failed: 0
- Sources: tests/unit/phase5_test_results.md; tests/unit/phase6_test_results.md; tests/unit/phase5_hq_boundary_results.md; tests/unit/phase6_packaged_wan_hq_results.md
- Scope note: proof rows below are the targeted Phase 6 LAN + WAN and central aggregation evidence items.

| Check | Result | Detail |
|---|---|---|
| Runner.Phase5Validation | PASS | PHASE5_VALIDATION_OK ; HARNESS=C:\Users\Justin\repos\invSys_fork\tests\fixtures\Phase5_Inventory.Domain_Harness.xlsm ; RESULTS=C:\Users\Justin\repos\invSys_fork\tests\unit\phase5_test_results.md ; PASSED=16 FAILED=0 TOTAL=16 |
| Runner.Phase6Validation | PASS | PHASE6_VALIDATION_OK ; HARNESS=C:\Users\Justin\repos\invSys_fork\tests\fixtures\Phase6_Inventory.Domain_Harness_20260331_080341_644.xlsm ; RESULTS=C:\Users\Justin\repos\invSys_fork\tests\unit\phase6_test_results.md ; PASSED=39 FAILED=3 TOTAL=42 |
| Runner.HqBoundaryValidation | PASS | PHASE5_HQ_BOUNDARY_OK ; RESULTS=C:\Users\Justin\repos\invSys_fork\tests\unit\phase5_hq_boundary_results.md ; SESSION_ROOT=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase5_hq_boundary_20260331_080437_989 ; PASSED=7 FAILED=0 TOTAL=7 |
| Runner.PackagedWanHqValidation | PASS | PHASE6_PACKAGED_WAN_HQ_OK ; RESULTS=C:\Users\Justin\repos\invSys_fork\tests\unit\phase6_packaged_wan_hq_results.md ; PASSED=10 FAILED=0 TOTAL=10 |
| WAN.PublishOnline | PASS | phase5 / TestPhase5Sync.TestWanPublish_OnlineCopy_PublishesLocalArtifactsToSharePoint / PASS / |
| WAN.PublishOfflineNonBlocking | PASS | phase5 / TestPhase5Sync.TestWanPublish_OfflineFailure_DoesNotBlockLocalProcessing / PASS / |
| WAN.PublishSafeRerun | PASS | phase5 / TestPhase5Sync.TestWanPublish_SafeRerun_ReplacesPublishedArtifacts / PASS / |
| WAN.PublishInterruptedRestore | PASS | phase5 / TestPhase5Sync.TestWanPublish_InterruptedReplacement_RestoresPriorArtifactAndAllowsCleanRerun / PASS / |
| WAN.DelayedPublicationCatchup | PASS | phase5 / TestPhase5Sync.TestDelayedPublicationRecovery_PreservesLocalOutboxAndGlobalCatchup / PASS / |
| WAN.UnreadablePublishedSnapshotFallback | PASS | phase5 / TestPhase5Sync.TestHqAggregation_SkipsUnreadablePublishedSnapshotAndRetainsLastGoodData / PASS / |
| WAN.MixedWarehouseInterruptionCatchup | PASS | phase5 / TestPhase5Sync.TestHqAggregation_MixedWarehouseInterruption_RetainsLastGoodAndCatchesUp / PASS / |
| LAN.SharedSnapshotTwoOperators | PASS | phase6 / TestPhase6CoreSurfaces.TestLanSharedSnapshot_TwoSavedOperatorWorkbooksRefreshWithoutCrossContamination / PASS / |
| LAN.TwoStationProcessorLockBoundary | PASS | phase6 / TestPhase6CoreSurfaces.TestLanTwoStationProcessorRun_RespectsLockAndPreservesOperatorWorkbooks / PASS / |
| Operator.SharePointStaleMetadataVisible | PASS | phase6 / TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSharePoint_StaleSnapshotMarksReadModelStale / PASS / |
| Operator.SavedWorkbookSharePointStaleVisibleNonDestructive | PASS | phase6 / TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_StaleSharePointSnapshotShowsVisibleMetadataWithoutMutatingLocalTables / PASS / |
| Operator.MissingSharePointSnapshotCachedNonDestructive | PASS | phase6 / TestPhase6CoreSurfaces.TestRefreshInventoryReadModel_MissingSharePointSnapshotMarksCachedWithoutMutatingLocalTables / PASS / |
| Operator.MissingSnapshotDoesNotBlockPosting | PASS | phase6 / TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_MissingSnapshotDoesNotBlockQueueAndRefresh / PASS / |
| HQ.StaggeredWarehouseUpdates | PASS | phase5 / TestPhase5Sync.TestHqAggregation_RebuildsGlobalSnapshotAfterStaggeredWarehouseUpdates / PASS / |
| HQ.RepeatedRunsRemainStable | PASS | phase5 / TestPhase5Sync.TestHqAggregation_RepeatedRunsRemainStableForWH1AndWH2Fixtures / PASS / |
| HQ.AdvisoryOnlyVisibility | PASS | phase5 / TestPhase5Sync.TestHqAggregation_GlobalSnapshotStatusIsAdvisoryOnly / PASS / |
| HQ.RealPublishedArtifactsBoundaryInitial | PASS | hq-boundary / Aggregate.Initial / PASS - OK/Report=Rows=2; SnapshotFiles=2; SkippedSnapshotFiles=0/QtyA=5/QtyB=8/SourceA=WH97.invSys.Snapshot.Inventory.xlsb/SourceB=WH98.invSys.Snapshot.Inventory.xlsb/Skipped=0/Warehouses=2 / |
| HQ.RealPublishedArtifactsBoundaryCatchup | PASS | hq-boundary / Aggregate.Catchup / PASS - OK/Report=Rows=2; SnapshotFiles=2; SkippedSnapshotFiles=0/QtyA=5/QtyB=11/SourceA=WH97.invSys.Snapshot.Inventory.xlsb/SourceB=WH98.invSys.Snapshot.Inventory.xlsb/Skipped=0/Warehouses=2 / |
| HQ.PackagedPublishedArtifactsInitial | PASS | packaged / Aggregate.Initial / PASS / QtyA=5; QtyB=8 / |
| HQ.PackagedPublishedArtifactsCatchup | PASS | packaged / Aggregate.Catchup / PASS / QtyA=5; QtyB=11 / |
