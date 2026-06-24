# Phase 6 VBA Test Results

- Date: 2026-06-23 22:04:27
- Passed: 35
- Failed: 0
- Range: 121-155 of 177

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestBoxBuilderForm_InitializesWithActiveArchiveFilters | PASS |
| TestPhase6CoreSurfaces.TestShippingCommitLine_MergesPostedSameRefBoxVersionCarrier | PASS |
| TestPhase6CoreSurfaces.TestShippingBoard_TwoAddsSameRefBoxVersionCarrierShowOneRow | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_DefaultsOrderToWarehouseArea | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_BlankCarrierRequiresCarrier | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_UsesDisplayedProjectedInventoryWhenVersionLedgerIsEmpty | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_UsesDisplayedProjectedInventoryWhenTotalInvIsStaleZero | PASS |
| TestPhase6CoreSurfaces.TestShippingRemove_LockedRowReleasesInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingRemove_StaleLockedRowClearsWithoutInflatingInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingHold_PreservesReservationAndLocalDeduction | PASS |
| TestPhase6CoreSurfaces.TestShippingUpdate_PreservesExistingReservationWithoutDoubleDeducting | PASS |
| TestPhase6CoreSurfaces.TestShippingUpdate_ReservedQtyChangeAppliesOnlyDeltaOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_ComposesActiveReservationWithPendingSentOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingShippables_NasInvPrefersCurrentInvSysForSingleActiveVersion | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedDisplay_SubtractsLockedAndUnreservedRows | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowDoesNotAddBackTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_UnreservedDirtyRowDeductsTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowClearsLockedReservationTotal | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_DoesNotIncreaseProjectedInventoryOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedCompletionKeepsProjectedDeductionWhenNasStale | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_FullRunNeverIncreasesProjectedInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PreservesNasBaselineAcrossSentReregister | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_EvictsStaleZeroWhenBackendPositive | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_EvictsIdleSentOverlayAtBaseline | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_KeepsSentOverlayWithActiveLock | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_LocalCatchupDoesNotClearBeforeNas | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_ClearsWhenBackendRisesAboveBaseline | PASS |
| TestPhase6CoreSurfaces.TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_MergesLocalBoxBuildStagingAndClearsStaleOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_FindsBackendShippingBomViewWithoutInvSysSurface | PASS |
| TestPhase6CoreSurfaces.TestBoxMakerUnbox_QtyGreaterThanInventoryFailsBeforeQueue | PASS |
| TestPhase6CoreSurfaces.TestBoxMakerUnbox_UsesShippingReadModelInventoryWhenInvSysMissing | PASS |
| TestPhase6CoreSurfaces.TestShippingReservationTotals_IgnoreSameWorkbookStaleActiveReservationWithoutLocalLine | PASS |
| TestPhase6CoreSurfaces.TestShippingReservationTotals_IgnoreLocallySentActiveLedgerRows | PASS |
