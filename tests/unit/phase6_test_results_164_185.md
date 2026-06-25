# Phase 6 VBA Test Results

- Date: 2026-06-24 22:53:24
- Passed: 22
- Failed: 0
- Range: 164-185 of 185

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestSavedProductionWorkbook_RefreshPreservesStagingAndLogs | PASS |
| TestPhase6CoreSurfaces.TestSavedProductionWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs | PASS |
| TestPhase6CoreSurfaces.TestProductionEventCreator_QueuesSignedInCurrentTargetEvent | PASS |
| TestPhase6CoreSurfaces.TestSavedAdminWorkbook_ReopenRefreshReissuePreservesAudit | PASS |
| TestPhase6CoreSurfaces.TestAdminShipmentReconcile_AppliesSignedDeltaWithCorrectedShipEvidence | PASS |
| TestPhase6CoreSurfaces.TestAdminShipmentReconcile_RejectsOrphanAndMissingNarrative | PASS |
| TestPhase6CoreSurfaces.TestAdminShipmentReconcile_DetectsNasIncreaseAfterLatestShip | PASS |
| TestPhase6CoreSurfaces.TestAdminShipmentReconcile_RecentShipmentsSentLogShowsLast20 | PASS |
| TestPhase6CoreSurfaces.TestAdminShipmentReconcile_RecentLogIncludesShipReserveEvidence | PASS |
| TestPhase6CoreSurfaces.TestApplyReceive_RebuildsDeletedProjectionTablesInCanonicalWorkbook | PASS |
| TestPhase6RoleSurfaces.TestEnsureInventoryManagementSurface_RemovesDuplicateAliasColumns | PASS |
| TestPhase6RoleSurfaces.TestEnsureReceivingWorkbookSurface_CreatesExpectedTables | PASS |
| TestPhase6RoleSurfaces.TestEnsureReceivingWorkbookSurface_RecreatesDeletedArtifacts | PASS |
| TestPhase6RoleSurfaces.TestEnsureShippingWorkbookSurface_CreatesExpectedTables | PASS |
| TestPhase6RoleSurfaces.TestEnsureShippingWorkbookSurface_RecreatesDeletedArtifacts | PASS |
| TestPhase6RoleSurfaces.TestEnsureProductionWorkbookSurface_CreatesExpectedTables | PASS |
| TestPhase6RoleSurfaces.TestEnsureProductionWorkbookSurface_RecreatesDeletedArtifacts | PASS |
| TestPhase6RoleSurfaces.TestEnsureAdminWorkbookSurface_CreatesExpectedTables | PASS |
| TestPhase6RoleSurfaces.TestResolveAdminTargetWorkbook_PrefersActiveVisibleWorkbook | PASS |
| TestPhase6RoleSurfaces.TestResolveAdminTargetWorkbook_ExplicitWorkbookWinsOverActiveWorkbook | PASS |
| TestPhase6RoleSurfaces.TestOpenUserManagement_WithoutWorkbookArgTargetsActiveWorkbook | PASS |
| TestPhase6RoleSurfaces.TestOpenAdminConsole_WithoutRuntime_DoesNotCreateDefaultWarehouse | PASS |
