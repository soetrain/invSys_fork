# Phase 6 VBA Test Results

- Date: 2026-06-22 17:05:19
- Passed: 93
- Failed: 18
- Range: 1-170 of 170
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
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_SuccessAppendsInventoryAndTracesSource | FAIL - SuccessAppendsInventoryAndTracesSource: SetupMigrationRuntimeRetire failed for WHRETMIG1A: Processor did not apply demo inventory seed. Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHRETMIG1A-INVENTORY-20260622170410-300454 |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_RejectsMissingArchiveManifest | FAIL - RejectsMissingArchiveManifest: SetupMigrationRuntimeRetire failed for WHRETMIG2B: Processor did not apply demo inventory seed. Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHRETMIG2B-INVENTORY-20260622170412-658708 |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_RejectsMissingTargetWarehouse | FAIL - RejectsMissingTargetWarehouse: SetupMigrationRuntimeRetire failed for WHRETMIG3A: Processor did not apply demo inventory seed. Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHRETMIG3A-INVENTORY-20260622170414-696938 |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_DoesNotCopyAuthIdentities | FAIL - DoesNotCopyAuthIdentities: SetupMigrationRuntimeRetire failed for WHRETMIG4A: Processor did not apply demo inventory seed. Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHRETMIG4A-INVENTORY-20260622170417-010866 |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_PreservesTargetConfigIdentity | FAIL - PreservesTargetConfigIdentity: SetupMigrationRuntimeRetire failed for WHRETMIG5A: Processor did not apply demo inventory seed. Applied=0; SkipDup=0; Poison=0; RunId=RUN-WHRETMIG5A-INVENTORY-20260622170419-538884 |
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
