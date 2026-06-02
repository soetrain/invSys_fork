# Phase 6 VBA Test Results

- Date: 2026-06-02 10:56:05
- Passed: 12
- Failed: 0
- Range: 13-24 of 128

| Test | Result |
|---|---|
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
| TestPhase6CoreSurfaces.TestRoleWriteCurrent_AllowsSignedInReceivePost | PASS |
| TestPhase6CoreSurfaces.TestAuthSignOut_ClearsUserButKeepsWarehouseTarget | PASS |
