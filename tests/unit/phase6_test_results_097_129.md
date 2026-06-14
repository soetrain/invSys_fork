# Phase 6 VBA Test Results

- Date: 2026-06-13 18:52:27
- Passed: 33
- Failed: 0
- Range: 97-129 of 129

| Test | Result |
|---|---|
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
| TestPhase6CoreSurfaces.TestShippingEventCreator_QueuesSignedInCurrentTargetEvent | PASS |
| TestPhase6CoreSurfaces.TestSavedProductionWorkbook_RefreshPreservesStagingAndLogs | PASS |
| TestPhase6CoreSurfaces.TestSavedProductionWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs | PASS |
| TestPhase6CoreSurfaces.TestProductionEventCreator_QueuesSignedInCurrentTargetEvent | PASS |
| TestPhase6CoreSurfaces.TestSavedAdminWorkbook_ReopenRefreshReissuePreservesAudit | PASS |
| TestPhase6CoreSurfaces.TestApplyReceive_RebuildsDeletedProjectionTablesInCanonicalWorkbook | PASS |
| TestPhase6RoleSurfaces.TestEnsureInventoryManagementSurface_RemovesDuplicateAliasColumns | PASS |
| TestPhase6RoleSurfaces.TestEnsureReceivingWorkbookSurface_CreatesExpectedTables | PASS |
| TestPhase6RoleSurfaces.TestEnsureReceivingWorkbookSurface_RecreatesDeletedArtifacts | PASS |
| TestPhase6RoleSurfaces.TestEnsureShippingWorkbookSurface_CreatesExpectedTables | PASS |
| TestPhase6RoleSurfaces.TestEnsureShippingWorkbookSurface_RecreatesDeletedArtifacts | PASS |
| TestPhase6RoleSurfaces.TestEnsureProductionWorkbookSurface_CreatesExpectedTables | PASS |
| TestPhase6RoleSurfaces.TestEnsureProductionWorkbookSurface_RecreatesDeletedArtifacts | PASS |
| TestPhase6RoleSurfaces.TestEnsureAdminWorkbookSurface_CreatesExpectedTables | PASS |
| TestPhase6RoleSurfaces.TestResolveAdminTargetWorkbook_PrefersActiveVisibleWorkbook | PASS |
| TestPhase6RoleSurfaces.TestResolveAdminTargetWorkbook_ExplicitWorkbookWinsOverActiveWorkbook | PASS |
| TestPhase6RoleSurfaces.TestOpenUserManagement_WithoutWorkbookArgTargetsActiveWorkbook | PASS |
| TestPhase6RoleSurfaces.TestOpenAdminConsole_WithoutRuntime_DoesNotCreateDefaultWarehouse | PASS |
