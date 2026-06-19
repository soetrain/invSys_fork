# Phase 6 VBA Test Results

- Date: 2026-06-18 16:14:25
- Passed: 23
- Failed: 0
- Range: 113-135 of 157

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestShippingState_TombstoneFiltersSentLineIdFromActiveCache | PASS |
| TestPhase6CoreSurfaces.TestShippingWorkflowGuard_ShipmentsSentWithZeroStagedFails | PASS |
| TestPhase6CoreSurfaces.TestShippingWorkflowGuard_ToShipmentsInsufficientInventoryFails | PASS |
| TestPhase6CoreSurfaces.TestShippingWorkflowGuard_BoxesMadeInsufficientComponentFails | PASS |
| TestPhase6CoreSurfaces.TestShippingWorkflowGuard_ConfirmInventoryUseExistingWarns | PASS |
| TestPhase6CoreSurfaces.TestShippingAggregateBomMath_MultipliesComponentQtyByPackageQty | PASS |
| TestPhase6CoreSurfaces.TestShippingCommitLine_MergesPostedSameRefBoxVersionCarrier | PASS |
| TestPhase6CoreSurfaces.TestShippingBoard_TwoAddsSameRefBoxVersionCarrierShowOneRow | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_DefaultsOrderToWarehouseArea | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_BlankCarrierRequiresCarrier | PASS |
| TestPhase6CoreSurfaces.TestShippingRemove_LockedRowReleasesInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingRemove_StaleLockedRowClearsWithoutInflatingInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingShippables_NasInvPrefersCurrentInvSysForSingleActiveVersion | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowDoesNotAddBackTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_UnreservedDirtyRowDeductsTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowClearsLockedReservationTotal | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_DoesNotIncreaseProjectedInventoryOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_FullRunNeverIncreasesProjectedInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_LocalCatchupDoesNotClearBeforeNas | PASS |
| TestPhase6CoreSurfaces.TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected | PASS |
| TestPhase6CoreSurfaces.TestShippingReservationTotals_IgnoreSameWorkbookStaleActiveReservationWithoutLocalLine | PASS |
| TestPhase6CoreSurfaces.TestShippingReservationTotals_IgnoreLocallySentActiveLedgerRows | PASS |
