# Phase 1 VBA Test Results

- Date: 2026-03-30 18:17:51
- Passed: 8
- Failed: 6

| Test | Result |
|---|---|
| TestCoreConfig.TestLoad_ValidConfig | PASS |
| TestCoreConfig.TestLoad_MissingRequiredKey | FAIL |
| TestCoreConfig.TestPrecedence_StationOverridesWarehouse | FAIL |
| TestCoreConfig.TestGetRequired_MissingKey | PASS |
| TestCoreConfig.TestGetBool_TypeConversion | FAIL |
| TestCoreConfig.TestReload_UpdatedValue | FAIL |
| TestCoreAuth.TestCanPerform_Allow | FAIL |
| TestCoreAuth.TestCanPerform_Deny_MissingCapability | PASS |
| TestCoreAuth.TestCanPerform_WildcardStation | FAIL |
| TestCoreAuth.TestCanPerform_DisabledUser | PASS |
| TestCoreAuth.TestCanPerform_ExpiredCapability | PASS |
| TestCoreAuth.TestRequire_RaisesOnDeny | PASS |
| TestInventorySchema.TestEnsureInventorySchema_RecreatesTables | PASS |
| TestInventorySchema.TestEnsureInventorySchema_AddsMissingColumns | PASS |
