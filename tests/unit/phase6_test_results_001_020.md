# Phase 6 VBA Test Results

- Date: 2026-06-17 08:14:14
- Passed: 20
- Failed: 0
- Range: 1-20 of 129

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
