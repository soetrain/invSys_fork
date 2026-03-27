# Phase 3 VBA Test Results

- Date: 2026-03-26 21:17:21
- Passed: 8
- Failed: 8

| Test | Result |
|---|---|
| TestCoreRoleEventWriter.TestQueueReceiveEvent_WritesInboxRow | FAIL |
| TestCoreRoleEventWriter.TestOpenInboxWorkbook_UsesStationPathInboxRoot | FAIL |
| TestCoreRoleEventWriter.TestQueueShipEvent_WritesInboxRow | FAIL |
| TestCoreRoleEventWriter.TestQueuePayloadEvent_DeniedWithoutCapability | PASS |
| TestCoreRoleEventWriter.TestBuildPayloadJson_WithObjectItems | PASS |
| TestCoreRoleUiAccess.TestCanCurrentUserPerformCapability_Allow | FAIL |
| TestCoreRoleUiAccess.TestCanCurrentUserPerformCapability_Deny | PASS |
| TestCoreRoleUiAccess.TestApplyShapeCapability_TogglesVisibility | FAIL |
| TestCoreItemSearch.TestNormalizeSearchText_CollapsesWhitespace | PASS |
| TestCoreItemSearch.TestAnyTextMatchesSearch_MatchesAcrossFields | PASS |
| TestCoreItemSearch.TestIdentifiersMatch_UsesTokenOverlap | PASS |
| TestCoreItemSearch.TestResolveSearchCaption_ReturnsRoleSpecificText | PASS |
| TestCoreItemSearch.TestShouldDefaultShippableForRole_UsesRoleDefaults | PASS |
| TestPhase3RoleFlows.TestReceivingRoleFlow_QueuesAndProcessesEvent | FAIL |
| TestPhase3RoleFlows.TestShippingRoleFlow_QueuesAndProcessesEvent | FAIL |
| TestPhase3RoleFlows.TestProductionRoleFlow_QueuesAndProcessesEvent | FAIL |
