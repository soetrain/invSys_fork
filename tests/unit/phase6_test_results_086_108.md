# Phase 6 VBA Test Results

- Date: 2026-05-25 21:58:49
- Passed: 23
- Failed: 0
- Range: 86-108 of 124

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestResolveInventoryWorkbookBridge_PrefersCanonicalWorkbookOverOperatorSurface | PASS |
| TestPhase6CoreSurfaces.TestEnsureInventoryManagementSurface_RemovesDomainArtifacts | PASS |
| TestPhase6CoreSurfaces.TestOpenOrCreateConfigWorkbookRuntime_PrunesUnexpectedSheets | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_UpdatesReadModelAndMetadata | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSharePoint_UpdatesReadModelAndMetadata | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSharePoint_StaleSnapshotMarksReadModelStale | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromCache_PreservesLocalStagingAndLogs | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_AddsRowsWhenInvSysStartsEmpty | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_AppliesCatalogMetadataForZeroQtyRows | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_NormalizesLegacyLocationSummary | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModel_MissingSnapshotMarksStaleWithoutMutatingReceivingTally | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModel_MissingSharePointSnapshotMarksCachedWithoutMutatingLocalTables | PASS |
| TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_StaleSharePointSnapshotShowsVisibleMetadataWithoutMutatingLocalTables | PASS |
| TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_MissingSnapshotDoesNotBlockQueueAndRefresh | PASS |
| TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_FullRuntimeCloseReopenReloadsCanonicalWorkbooks | PASS |
| TestPhase6CoreSurfaces.TestSavedReceivingWorkbook_ReopenRefreshPreservesLocalTables | PASS |
| TestPhase6CoreSurfaces.TestReceivingSetupUi_ForceRefreshesRegisteredWorkbook | PASS |
| TestPhase6CoreSurfaces.TestInventoryPublisher_PublishesSnapshotForOpenInventoryWorkbook | PASS |
| TestPhase6CoreSurfaces.TestLanSharedSnapshot_TwoSavedOperatorWorkbooksRefreshWithoutCrossContamination | PASS |
| TestPhase6CoreSurfaces.TestLanTwoStationProcessorRun_RespectsLockAndPreservesOperatorWorkbooks | PASS |
| TestPhase6CoreSurfaces.TestProcessor_DiscoversClosedConfiguredStationInboxWorkbook | PASS |
| TestPhase6CoreSurfaces.TestSavedShippingWorkbook_RefreshPreservesStagingAndLogs | PASS |
| TestPhase6CoreSurfaces.TestSavedShippingWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs | PASS |
