# Phase 2 VBA Test Results

- Date: 2026-03-24 21:40:31
- Passed: 21
- Failed: 11

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
| TestInventorySchema.TestEnsureInventorySchema_RemovesBlankSeedRow | PASS |
| TestInventorySchema.TestEnsureInventorySchema_CreatesProjectionTables | PASS |
| TestCoreLockManager.TestAcquireReleaseLock_Lifecycle | PASS |
| TestCoreLockManager.TestHeartbeat_ExtendsExpiry | PASS |
| TestInventoryApply.TestApplyReceive_ValidEvent | PASS |
| TestInventoryApply.TestApplyReceive_InvalidSKU | PASS |
| TestInventoryApply.TestApplyReceive_Duplicate | PASS |
| TestInventoryApply.TestApplyReceive_ProtectedSheetReturnsClearError | PASS |
| TestInventoryApply.TestApplyReceive_RebuildsProjectionTables | PASS |
| TestInventoryApply.TestResolveInventoryWorkbook_UsesConfiguredPathDataRoot | PASS |
| TestInventoryApply.TestApplyShip_MultiLineEvent | PASS |
| TestInventoryApply.TestApplyProdConsume_MultiLineEvent | PASS |
| TestInventoryApply.TestApplyProdComplete_MultiLineEvent | PASS |
| TestCoreProcessor.TestRunBatch_ProcessesInboxRow | FAIL |
| TestCoreProcessor.TestRunBatch_DuplicateMarkedSkipDup | FAIL |
| TestCoreProcessor.TestRunBatch_ProcessesShipRow | FAIL |
| TestCoreProcessor.TestRunBatch_ProcessesProdConsumeRow | FAIL |
| TestCoreProcessor.TestRunBatch_ProcessesProdCompleteRow | FAIL |
