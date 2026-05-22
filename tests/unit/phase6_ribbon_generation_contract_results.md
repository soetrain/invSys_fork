# Phase 6 Ribbon Generation Contract Results

- Date: 2026-05-22 13:42:49
- Passed: 16
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Build.GetEnabledXml | PASS | RequiredCapability buttons emit getEnabled. |
| Build.GetEnabledCallback | PASS | Generated callback exists. |
| Build.ReceivingCapability | PASS | Receiving buttons declare capability. |
| Build.ShippingCapability | PASS | Shipping buttons declare capability. |
| Build.ProductionCapability | PASS | Production buttons declare capability. |
| Build.RoleConnectServerButtons | PASS | Role ribbons expose Connect Server buttons. |
| Build.RoleSignOutButtons | PASS | Role ribbons expose Sign Out buttons. |
| Build.SignInLabelCallback | PASS | Current user button acts as Sign In while signed out. |
| Build.ServerStatusLabelControl | PASS | Role ribbons emit server status label controls. |
| Core.RoleConnectNonModal | PASS | Role Connect Server resolves without opening the warehouse connection form. |
| Core.RibbonFullInvalidate | PASS | Auth/storage changes refresh enabled callbacks. |
| Validator.ButtonGetEnabledRead | PASS | Packaged validator reads getEnabled. |
| Validator.ButtonGetEnabledAssert | PASS | Packaged validator asserts getEnabled on required buttons. |
| Validator.CallbackGetEnabledAssert | PASS | Packaged validator asserts callback capability mapping. |
| Validator.DirectActionAssert | PASS | Packaged validator asserts direct ribbon actions. |
| Validator.StatusLabelAssert | PASS | Packaged validator asserts server status labels. |
