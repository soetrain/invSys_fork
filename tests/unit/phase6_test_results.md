# Phase 6 VBA Test Results

- Date: 2026-06-24 22:31:31
- Passed: 147
- Failed: 18
- Range: 1-185 of 185
- Status: PARTIAL

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_ReadsWarehouseIdFromConfig | PASS |
| TestPhase6CoreSurfaces.TestNasGetCurrentTarget_ReturnsDeepCopy | PASS |
| TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_RequiresStationInboxRejectsBlankStation | PASS |
| TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_AllowsRoamingBlankStationWithoutInboxRequirement | PASS |
| TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_TwoStationsHaveIndependentInboxRoots | PASS |
| TestPhase6CoreSurfaces.TestNasScanRoot_ReturnsPathStringsWithoutWarehouseInference | PASS |
| TestPhase6CoreSurfaces.TestNasScanRoot_RejectsMismatchedConfigAuthPair | PASS |
| TestPhase6CoreSurfaces.TestNasResolveRememberedTarget_UnreachableFailsClosed | PASS |
| TestPhase6CoreSurfaces.TestNasResolveRememberedTarget_ReachableRecomputesCachedHints | PASS |
| TestPhase6CoreSurfaces.TestNasFallbackPolicy_RoleRejectsFallbackAdminAccepts | PASS |
| TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_SignsInAndStatusOk | PASS |
| TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_AcceptsResetPinForUserId | PASS |
| TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_RejectsDisplayNameAsUserId | PASS |
| TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_RejectsMismatchedTargetWarehouse | PASS |
| TestPhase6CoreSurfaces.TestAuthCapabilityScope_AllowsSelectedRuntimeFolderAlias | PASS |
| TestPhase6CoreSurfaces.TestAuthFailedCredential_DoesNotReplaceSignedInUser | PASS |
| TestPhase6CoreSurfaces.TestAuthCorrectCredentialWithoutCapability_ReturnsNoCapabilities | PASS |
| TestPhase6CoreSurfaces.TestRuntimeStatusUserLabel_UnsignedShowsNotSignedIn | PASS |
| TestPhase6CoreSurfaces.TestRuntimeStatusUserLabel_TracksAuthSignIn | PASS |
| TestPhase6CoreSurfaces.TestRoleWriteCurrent_RejectsUnsignedUser | PASS |
| TestPhase6CoreSurfaces.TestRoleWriteCurrent_RejectsMissingCapability | PASS |
| TestPhase6CoreSurfaces.TestRoleWriteCurrent_RejectsFallbackTarget | PASS |
| TestPhase6CoreSurfaces.TestRoleWriteCurrent_AllowsSignedInReceivePost | FAIL |
| TestPhase6CoreSurfaces.TestAuthSignOut_ClearsUserButKeepsWarehouseTarget | PASS |
| TestPhase6CoreSurfaces.TestAuthCanPerform_SignedOutFailsClosedWithLoadedAuth | PASS |
| TestPhase6CoreSurfaces.TestAuthTtlExpiry_FailsClosedForIsSignedInAndCanPerform | PASS |
| TestAddinsPublish.TestVerifyAddinsPublished_AllPresent | PASS |
| TestAddinsPublish.TestVerifyAddinsPublished_OneMissingLogsDiagnostic | PASS |
| TestAddinsPublish.TestVerifyAddinsPublished_ZeroByteFileLogsDiagnostic | PASS |
| TestAddinsPublish.TestPublishAddins_IdempotentRepublishWritesManifest | PASS |
| TestWarehouseBootstrap.TestValidateWarehouseSpec_TrimsFieldsAndAllowsBlankSharePoint | PASS |
| TestWarehouseBootstrap.TestValidateWarehouseSpec_RejectsEmptyWarehouseId | PASS |
| TestWarehouseBootstrap.TestValidateWarehouseSpec_RejectsWarehouseIdWithSpaces | PASS |
| TestWarehouseBootstrap.TestValidateWarehouseSpec_AllowsWarehouseIdWithHyphenAndUnderscore | PASS |
| TestWarehouseBootstrap.TestValidateWarehouseSpec_RejectsWarehouseIdWithOtherSpecialCharacters | PASS |
| TestWarehouseBootstrap.TestWarehouseIdExists_LocalFolderExists | PASS |
| TestWarehouseBootstrap.TestWarehouseIdExists_SharePointArtifactExists | PASS |
| TestWarehouseBootstrap.TestWarehouseIdExists_NeitherLocalNorSharePointExists | PASS |
| TestWarehouseBootstrap.TestWarehouseIdExists_SharePointUnavailableReturnsFalseAndLogsSkip | PASS |
| TestWarehouseBootstrap.TestBootstrapWarehouseLocal_CreatesBootableLocalRuntime | FAIL |
| TestWarehouseBootstrap.TestBootstrapWarehouseLocal_FailureRollsBackPartialFolders | PASS |
| TestWarehouseBootstrap.TestPublishInitialArtifacts_PublishSuccess | FAIL |
| TestWarehouseBootstrap.TestPublishInitialArtifacts_SharePointUnavailableReturnsFalse | FAIL |
| TestWarehouseBootstrap.TestPublishInitialArtifacts_RepeatedPublishIsIdempotent | FAIL |
| test_RetireMigrateSpec.TestValidateRetireMigrateSpec_TrimsAndAcceptsArchiveOnly | PASS |
| test_RetireMigrateSpec.TestValidateRetireMigrateSpec_RejectsEmptySourceWarehouseId | PASS |
| test_RetireMigrateSpec.TestValidateRetireMigrateSpec_RejectsMissingTargetForMigrate | PASS |
| test_RetireMigrateSpec.TestValidateRetireMigrateSpec_RejectsEqualSourceAndTarget | PASS |
| test_RetireMigrateSpec.TestValidateRetireMigrateSpec_RejectsUnconfirmedWriteOperation | PASS |
| test_RetireMigrateSpec.TestValidateRetireMigrateSpec_RejectsInvalidArchiveDestPath | PASS |
| TestWarehouseRetireReAuth.TestValidateUserCredential_SucceedsWithCorrectPasswordAndRole | PASS |
| TestWarehouseRetireReAuth.TestReAuthGate_WrongPassword_ShowsInlineErrorAndDoesNotAuthenticate | PASS |
| TestWarehouseRetireReAuth.TestReAuthGate_ThreeFailures_LocksOutAndLogs | PASS |
| TestWarehouseRetireReAuth.TestReAuthGate_Cancel_LeavesUnauthenticatedWithoutLog | PASS |
| TestWarehouseRetireArchive.TestWriteArchivePackage_SuccessCreatesAtomicArchive | FAIL |
| TestWarehouseRetireArchive.TestWriteArchivePackage_PartialFailureRollsBackTempArchive | FAIL |
| TestWarehouseRetireArchive.TestWriteArchivePackage_AuthExportMasksPinHash | FAIL |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_SuccessAppendsInventoryAndTracesSource | FAIL - SuccessAppendsInventoryAndTracesSource: SetupMigrationRuntimeRetire failed for WHRETMIG1A: Processor did not apply demo inventory seed. Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHRETMIG1A-INVENTORY-20260624222953-420983 |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_RejectsMissingArchiveManifest | FAIL - RejectsMissingArchiveManifest: SetupMigrationRuntimeRetire failed for WHRETMIG2B: Processor did not apply demo inventory seed. Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHRETMIG2B-INVENTORY-20260624222955-918916 |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_RejectsMissingTargetWarehouse | FAIL - RejectsMissingTargetWarehouse: SetupMigrationRuntimeRetire failed for WHRETMIG3A: Processor did not apply demo inventory seed. Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHRETMIG3A-INVENTORY-20260624222957-374123 |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_DoesNotCopyAuthIdentities | FAIL - DoesNotCopyAuthIdentities: SetupMigrationRuntimeRetire failed for WHRETMIG4A: Processor did not apply demo inventory seed. Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHRETMIG4A-INVENTORY-20260624223000-576083 |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_PreservesTargetConfigIdentity | FAIL - PreservesTargetConfigIdentity: SetupMigrationRuntimeRetire failed for WHRETMIG5A: Processor did not apply demo inventory seed. Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHRETMIG5A-INVENTORY-20260624223002-335790 |
| TestWarehouseRetireLifecycle.TestRetireSourceWarehouse_WritesRetirementMarker | FAIL - WritesRetirementMarker: No retire lifecycle report was available. |
| TestWarehouseRetireLifecycle.TestRetireSourceWarehouse_WritesValidTombstoneJson | FAIL - WritesValidTombstoneJson: No retire lifecycle report was available. |
| TestWarehouseRetireLifecycle.TestRetireSourceWarehouse_SharePointUnavailableDoesNotBlockRetirement | FAIL - SharePointUnavailableDoesNotBlockRetirement: No retire lifecycle report was available. |
| TestWarehouseRetireLifecycle.TestDeleteLocalRuntime_RejectsWithoutTombstone | FAIL - RejectsWithoutTombstone: No retire lifecycle report was available. |
| TestWarehouseRetireLifecycle.TestDeleteLocalRuntime_RejectsWithoutConfirmation | FAIL - RejectsWithoutConfirmation: No retire lifecycle report was available. |
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
| TestReceivingReadiness.TestEnsureReceivingSurface_BlankWorkbookWithConfigLoaded_DoesNotApplyReadiness | PASS |
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
| TestPhase6CoreSurfaces.TestRefreshInventoryReadModelFromSnapshot_MatchesLocalRowWhenSkuAliasDiffers | PASS |
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
| TestPhase6CoreSurfaces.TestShippingState_TombstoneFiltersSentLineIdFromActiveCache | PASS |
| TestPhase6CoreSurfaces.TestShippingState_SentRowTombstoneFiltersLegacyActiveCache | PASS |
| TestPhase6CoreSurfaces.TestShippingWorkflowGuard_ShipmentsSentWithZeroStagedFails | PASS |
| TestPhase6CoreSurfaces.TestShippingWorkflowGuard_ToShipmentsInsufficientInventoryFails | PASS |
| TestPhase6CoreSurfaces.TestShippingWorkflowGuard_BoxesMadeInsufficientComponentFails | PASS |
| TestPhase6CoreSurfaces.TestShippingWorkflowGuard_ConfirmInventoryUseExistingWarns | PASS |
| TestPhase6CoreSurfaces.TestShippingAggregateBomMath_MultipliesComponentQtyByPackageQty | PASS |
| TestPhase6CoreSurfaces.TestBoxBuilderArchive_HidesArchivedBoxesUnlessRequested | PASS |
| TestPhase6CoreSurfaces.TestBoxBuilderForm_InitializesWithActiveArchiveFilters | PASS |
| TestPhase6CoreSurfaces.TestShippingCommitLine_MergesPostedSameRefBoxVersionCarrier | PASS |
| TestPhase6CoreSurfaces.TestShippingBoard_TwoAddsSameRefBoxVersionCarrierShowOneRow | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_DefaultsOrderToWarehouseArea | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_BlankCarrierRequiresCarrier | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_UsesDisplayedProjectedInventoryWhenVersionLedgerIsEmpty | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_UsesDisplayedProjectedInventoryWhenTotalInvIsStaleZero | PASS |
| TestPhase6CoreSurfaces.TestShippingRemove_LockedRowReleasesInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingRemove_StaleLockedRowClearsWithoutInflatingInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingHold_PreservesReservationAndLocalDeduction | PASS |
| TestPhase6CoreSurfaces.TestShippingToShipments_ReservedMultiSelectKeepsRowsAndProjection | PASS |
| TestPhase6CoreSurfaces.TestShippingUpdate_PreservesExistingReservationWithoutDoubleDeducting | PASS |
| TestPhase6CoreSurfaces.TestShippingUpdate_ReservedQtyChangeAppliesOnlyDeltaOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_NewReservedRowAppliesSingleProjectedDeduction | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_MergingExistingReservedRowAppliesOnlyDelta | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_ComposesActiveReservationWithPendingSentOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingShippables_NasInvPrefersCurrentInvSysForSingleActiveVersion | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedDisplay_SubtractsLockedAndUnreservedRows | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowDoesNotAddBackTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_UnreservedDirtyRowDeductsTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowClearsLockedReservationTotal | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_DoesNotIncreaseProjectedInventoryOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedCompletionKeepsProjectedDeductionWhenNasStale | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_FullRunNeverIncreasesProjectedInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PreservesNasBaselineAcrossSentReregister | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_EvictsStaleZeroWhenBackendPositive | PASS |
| TestPhase6CoreSurfaces.TestShippingHydrateShippables_RepairsStaleZeroAndDoesNotWriteNewZero | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_KeepsFreshSentOverlayAtBaseline | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_EvictsSentOverlayWhenNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_KeepsSentOverlayWithActiveLock | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_LocalCatchupDoesNotClearBeforeNas | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_ClearsWhenBackendRisesAboveBaseline | PASS |
| TestPhase6CoreSurfaces.TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_MergesLocalBoxBuildStagingAndClearsStaleOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_FindsBackendShippingBomViewWithoutInvSysSurface | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_SkipsBomNetworkWhenViewPopulated | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_HidesSupportSheetsAfterSurfaceRepair | PASS |
| TestPhase6CoreSurfaces.TestBoxMakerUnbox_QtyGreaterThanInventoryFailsBeforeQueue | PASS |
| TestPhase6CoreSurfaces.TestBoxMakerUnbox_UsesShippingReadModelInventoryWhenInvSysMissing | PASS |
| TestPhase6CoreSurfaces.TestShippingReservationTotals_IgnoreSameWorkbookStaleActiveReservationWithoutLocalLine | PASS |
| TestPhase6CoreSurfaces.TestShippingReservationTotals_IgnoreLocallySentActiveLedgerRows | PASS |
| TestPhase6CoreSurfaces.TestSavedProductionWorkbook_RefreshPreservesStagingAndLogs | PASS |
| TestPhase6CoreSurfaces.TestSavedProductionWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs | PASS |
