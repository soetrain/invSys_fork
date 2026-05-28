# Phase 6 VBA Test Results

- Date: 2026-05-27 12:17:59
- Passed: 62
- Failed: 0
- Range: 1-62 of 126

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_ReadsWarehouseIdFromConfig | PASS |
| TestPhase6CoreSurfaces.TestNasGetCurrentTarget_ReturnsDeepCopy | PASS |
| TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_RequiresStationInboxRejectsBlankStation | PASS |
| TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_AllowsRoamingBlankStationWithoutInboxRequirement | PASS |
| TestPhase6CoreSurfaces.TestNasSelectWarehouseTarget_TwoStationsHaveIndependentInboxRoots | PASS |
| TestPhase6CoreSurfaces.TestNasScanRoot_ReturnsPathStringsWithoutWarehouseInference | PASS |
| TestPhase6CoreSurfaces.TestNasResolveRememberedTarget_UnreachableFailsClosed | PASS |
| TestPhase6CoreSurfaces.TestNasResolveRememberedTarget_ReachableRecomputesCachedHints | PASS |
| TestPhase6CoreSurfaces.TestNasFallbackPolicy_RoleRejectsFallbackAdminAccepts | PASS |
| TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_SignsInAndStatusOk | PASS |
| TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_AcceptsResetPinForUserId | PASS |
| TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_RejectsDisplayNameAsUserId | PASS |
| TestPhase6CoreSurfaces.TestAuthValidateUserCredentialForTarget_RejectsMismatchedTargetWarehouse | PASS |
| TestPhase6CoreSurfaces.TestAuthFailedCredential_DoesNotReplaceSignedInUser | PASS |
| TestPhase6CoreSurfaces.TestAuthCorrectCredentialWithoutCapability_ReturnsNoCapabilities | PASS |
| TestPhase6CoreSurfaces.TestRuntimeStatusUserLabel_UnsignedShowsNotSignedIn | PASS |
| TestPhase6CoreSurfaces.TestRuntimeStatusUserLabel_TracksAuthSignIn | PASS |
| TestPhase6CoreSurfaces.TestRoleWriteCurrent_RejectsUnsignedUser | PASS |
| TestPhase6CoreSurfaces.TestRoleWriteCurrent_RejectsMissingCapability | PASS |
| TestPhase6CoreSurfaces.TestRoleWriteCurrent_RejectsFallbackTarget | PASS |
| TestPhase6CoreSurfaces.TestRoleWriteCurrent_AllowsSignedInReceivePost | PASS |
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
| TestWarehouseBootstrap.TestBootstrapWarehouseLocal_CreatesBootableLocalRuntime | PASS |
| TestWarehouseBootstrap.TestBootstrapWarehouseLocal_FailureRollsBackPartialFolders | PASS |
| TestWarehouseBootstrap.TestPublishInitialArtifacts_PublishSuccess | PASS |
| TestWarehouseBootstrap.TestPublishInitialArtifacts_SharePointUnavailableReturnsFalse | PASS |
| TestWarehouseBootstrap.TestPublishInitialArtifacts_RepeatedPublishIsIdempotent | PASS |
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
| TestWarehouseRetireArchive.TestWriteArchivePackage_SuccessCreatesAtomicArchive | PASS |
| TestWarehouseRetireArchive.TestWriteArchivePackage_PartialFailureRollsBackTempArchive | PASS |
| TestWarehouseRetireArchive.TestWriteArchivePackage_AuthExportMasksPinHash | PASS |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_SuccessAppendsInventoryAndTracesSource | PASS |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_RejectsMissingArchiveManifest | PASS |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_RejectsMissingTargetWarehouse | PASS |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_DoesNotCopyAuthIdentities | PASS |
| TestWarehouseRetireMigration.TestMigrateInventoryToTarget_PreservesTargetConfigIdentity | PASS |
| TestWarehouseRetireLifecycle.TestRetireSourceWarehouse_WritesRetirementMarker | PASS |
| TestWarehouseRetireLifecycle.TestRetireSourceWarehouse_WritesValidTombstoneJson | PASS |
