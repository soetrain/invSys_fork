# Phase 6 Ribbon Generation Contract Results

- Date: 2026-05-31 22:38:12
- Passed: 32
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Build.GetEnabledXml | PASS | RequiredCapability buttons emit getEnabled. |
| Build.GetEnabledCallback | PASS | Generated callback exists. |
| Build.GetEnabledCached | PASS | Ribbon getEnabled uses cached auth/target state. |
| Build.ReceivingCapability | PASS | Receiving buttons declare capability. |
| Build.ShippingCapability | PASS | Shipping buttons declare capability. |
| Build.ProductionCapability | PASS | Production buttons declare capability. |
| Build.RoleConnectServerButtons | PASS | Role ribbons expose Connect Server buttons. |
| Build.RoleSignOutButtons | PASS | Role ribbons expose Sign Out buttons. |
| Build.SignInLabelCallback | PASS | Current user button acts as Sign In while signed out. |
| Build.UserLabelUsesDisplayName | PASS | Ribbon user label uses display name, not account id. |
| Build.RuntimeContextNoSignIn | PASS | Runtime Context is informational and does not expose separate Sign In. |
| Build.ServerStatusLabelControl | PASS | Role ribbons emit server status label controls. |
| Build.RuntimeReferencesNormalXlams | PASS | Built operator XLAMs reference normal deployed XLAM outputs. |
| Core.RoleConnectNonModal | PASS | Role Connect Server resolves without opening the warehouse connection form. |
| Core.ConnectServerBindsWarehouseTarget | PASS | Connect Server validates the saved NAS root and binds a target from that connected root. |
| Core.ConnectServerRequiresNasRoot | PASS | Role Connect Server can reject remembered local roots. |
| Core.ConnectServerReconnectsRememberedShare | PASS | Remembered NAS roots attempt Windows SMB reattach with current/stored credentials before failing. |
| Core.ConnectServerPromptsForServerCredentials | PASS | Role Connect Server prompts for server credentials when stored SMB credentials fail. |
| Core.RoleTargetsRejectLocalPaths | PASS | Role-required NAS targets reject stale local/temp targets and report NAS probe status. |
| Core.SignOutClearsPersistedUser | PASS | Sign Out clears live auth and persisted current-user state. |
| Core.AuthStoresDisplayName | PASS | Auth cache stores and exposes signed-in display name. |
| Core.RuntimeContextShowsUserId | PASS | Runtime Context shows signed-in account id. |
| Core.RememberedTargetUsesConfigAuth | PASS | Remembered server reconnect requires config/auth, not a local inventory workbook. |
| Admin.DirectoryReadsNasRoots | PASS | Admin View Warehouses includes NAS roots remembered by Connect Server. |
| Core.SendToScansConnectedRoots | PASS | Send To scans connected NAS roots after Connect Server succeeds. |
| Core.SendToSuppressesLocalFallbackWhenConnected | PASS | Send To suppresses default/local runtime noise while a NAS root is connected. |
| Core.RibbonFullInvalidate | PASS | Auth/storage changes refresh enabled callbacks. |
| Validator.ButtonGetEnabledRead | PASS | Packaged validator reads getEnabled. |
| Validator.ButtonGetEnabledAssert | PASS | Packaged validator asserts getEnabled on required buttons. |
| Validator.CallbackGetEnabledAssert | PASS | Packaged validator asserts callback capability mapping. |
| Validator.DirectActionAssert | PASS | Packaged validator asserts direct ribbon actions. |
| Validator.StatusLabelAssert | PASS | Packaged validator asserts server status labels. |
