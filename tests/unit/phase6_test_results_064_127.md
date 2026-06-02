# Phase 6 VBA Test Results

- Date: 2026-06-01 18:32:59
- Passed: 46
- Failed: 0
- Range: 64-127 of 127
- Status: PARTIAL

| Test | Result |
|---|---|
| TestWarehouseRetireLifecycle.TestRetireSourceWarehouse_SharePointUnavailableDoesNotBlockRetirement | PASS |
| TestWarehouseRetireLifecycle.TestDeleteLocalRuntime_RejectsWithoutTombstone | PASS |
| TestWarehouseRetireLifecycle.TestDeleteLocalRuntime_RejectsWithoutConfirmation | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_AllReady_ReturnsReady | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_AllReady_WhenCapabilityStationWildcard | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_SnapshotOk_WhenAuthMissingCapability | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_SnapshotStale_ReturnsStale | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_SnapshotMissing_ReturnsMissing | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_SnapshotUnreadable_ReturnsUnreadable | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_AuthOk_WhenSnapshotMissing | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_AuthNoUser_ReturnsNoUser | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_AuthMissingCapability_ReturnsMissingCapability | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_AuthInactive_ReturnsInactive | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_RuntimeOk_WhenSnapshotMissingAndNoUser | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_RuntimeMissingTables_ReturnsMissingTables | PASS |
| TestReceivingReadiness.TestCheckReceivingReadiness_RuntimePathUnresolved_ReturnsPathUnresolved | PASS |
| TestPhase6CoreSurfaces.TestOpenOrCreateConfigWorkbookRuntime_CreatesCanonicalWorkbook | PASS |
| TestPhase6CoreSurfaces.TestLoadConfig_AutoBootstrapsCanonicalWorkbook | PASS |
| TestPhase6CoreSurfaces.TestLoadConfig_BlankContextAutoBootstrapsDefaultRuntimeWorkbook | PASS |
| TestPhase6CoreSurfaces.TestEnsureStationBootstrap_CreatesLocalConfigAndInbox | PASS |
| TestPhase6CoreSurfaces.TestLoadConfig_QuarantinesContaminatedConfigSheet | PASS |
| TestPhase6CoreSurfaces.TestLoadAuth_AutoBootstrapsCanonicalWorkbook | PASS |
| TestPhase6CoreSurfaces.TestLoadAuth_BootstrapGrantsCurrentOperatorCapabilities | PASS |
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
