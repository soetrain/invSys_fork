# Phase 6 VBA Test Results

- Date: 2026-03-28 18:29:33
- Passed: 33
- Failed: 4

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestOpenOrCreateConfigWorkbookRuntime_CreatesCanonicalWorkbook | PASS |
| TestPhase6CoreSurfaces.TestLoadConfig_AutoBootstrapsCanonicalWorkbook | FAIL |
| TestPhase6CoreSurfaces.TestLoadConfig_BlankContextAutoBootstrapsDefaultRuntimeWorkbook | FAIL |
| TestPhase6CoreSurfaces.TestEnsureStationBootstrap_CreatesLocalConfigAndInbox | PASS |
| TestPhase6CoreSurfaces.TestLoadConfig_QuarantinesContaminatedConfigSheet | PASS |
| TestPhase6CoreSurfaces.TestLoadAuth_AutoBootstrapsCanonicalWorkbook | FAIL |
| TestPhase6CoreSurfaces.TestLoadAuth_BootstrapGrantsCurrentOperatorCapabilities | PASS |
| TestPhase6CoreSurfaces.TestResolveInventoryWorkbookBridge_PrefersCanonicalWorkbookOverOperatorSurface | PASS |
| TestPhase6CoreSurfaces.TestEnsureInventoryManagementSurface_RemovesDomainArtifacts | PASS |
| TestPhase6CoreSurfaces.TestOpenOrCreateConfigWorkbookRuntime_PrunesUnexpectedSheets | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_UpdatesReadModelAndMetadata | FAIL |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_AddsRowsWhenInvSysStartsEmpty | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_AppliesCatalogMetadataForZeroQtyRows | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_NormalizesLegacyLocationSummary | PASS |
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModel_MissingSnapshotMarksStaleWithoutMutatingReceivingTally | PASS |
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
| TestPhase6CoreSurfaces.TestSavedProductionWorkbook_RefreshPreservesStagingAndLogs | PASS |
| TestPhase6CoreSurfaces.TestSavedProductionWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs | PASS |
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
