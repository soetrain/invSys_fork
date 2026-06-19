# Phase 6 VBA Test Results

- Date: 2026-06-18 11:04:50
- Passed: 13
- Failed: 0
- Range: 128-142 of 155
- Status: PARTIAL

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowClearsLockedReservationTotal | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_DoesNotIncreaseProjectedInventoryOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_FullRunNeverIncreasesProjectedInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_LocalCatchupDoesNotClearBeforeNas | PASS |
| TestPhase6CoreSurfaces.TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected | PASS |
| TestPhase6CoreSurfaces.TestShippingReservationTotals_IgnoreSameWorkbookStaleActiveReservationWithoutLocalLine | PASS |
| TestPhase6CoreSurfaces.TestShippingReservationTotals_IgnoreLocallySentActiveLedgerRows | PASS |
| TestPhase6CoreSurfaces.TestSavedProductionWorkbook_RefreshPreservesStagingAndLogs | PASS |
| TestPhase6CoreSurfaces.TestSavedProductionWorkbook_ReopenQueueProcessRefreshPreservesStagingAndLogs | PASS |
| TestPhase6CoreSurfaces.TestProductionEventCreator_QueuesSignedInCurrentTargetEvent | PASS |
| TestPhase6CoreSurfaces.TestSavedAdminWorkbook_ReopenRefreshReissuePreservesAudit | PASS |
| TestPhase6CoreSurfaces.TestAdminShipmentReconcile_AppliesSignedDeltaWithCorrectedShipEvidence | PASS |
