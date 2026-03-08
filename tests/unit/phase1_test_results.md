# Phase 1 VBA Test Results

- Date: 2026-03-08 00:11:02
- Passed: 14
- Failed: 0

| Test | Result |
|---|---|
| TestCoreConfig.TestLoad_ValidConfig | PASS |
| TestCoreConfig.TestLoad_MissingRequiredKey | PASS |
| TestCoreConfig.TestPrecedence_StationOverridesWarehouse | PASS |
| TestCoreConfig.TestGetRequired_MissingKey | PASS |
| TestCoreConfig.TestGetBool_TypeConversion | PASS |
| TestCoreConfig.TestReload_UpdatedValue | PASS |
| TestCoreAuth.TestCanPerform_Allow | PASS |
| TestCoreAuth.TestCanPerform_Deny_MissingCapability | PASS |
| TestCoreAuth.TestCanPerform_WildcardStation | PASS |
| TestCoreAuth.TestCanPerform_DisabledUser | PASS |
| TestCoreAuth.TestCanPerform_ExpiredCapability | PASS |
| TestCoreAuth.TestRequire_RaisesOnDeny | PASS |
| TestInventorySchema.TestEnsureInventorySchema_RecreatesTables | PASS |
| TestInventorySchema.TestEnsureInventorySchema_AddsMissingColumns | PASS |
