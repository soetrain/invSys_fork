# Phase 6 VBA Test Results

- Date: 2026-06-24 19:05:02
- Passed: 25
- Failed: 0
- Range: 137-161 of 183

| Test | Result |
|---|---|
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
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_KeepsFreshSentOverlayAtBaseline | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_EvictsSentOverlayWhenNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_KeepsSentOverlayWithActiveLock | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_LocalCatchupDoesNotClearBeforeNas | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_ClearsWhenBackendRisesAboveBaseline | PASS |
| TestPhase6CoreSurfaces.TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_MergesLocalBoxBuildStagingAndClearsStaleOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_FindsBackendShippingBomViewWithoutInvSysSurface | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_SkipsBomNetworkWhenViewPopulated | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_HidesSupportSheetsAfterSurfaceRepair | PASS |
| TestPhase6CoreSurfaces.TestBoxMakerUnbox_QtyGreaterThanInventoryFailsBeforeQueue | PASS |
| TestPhase6CoreSurfaces.TestBoxMakerUnbox_UsesShippingReadModelInventoryWhenInvSysMissing | PASS |
| TestPhase6CoreSurfaces.TestShippingReservationTotals_IgnoreSameWorkbookStaleActiveReservationWithoutLocalLine | PASS |
| TestPhase6CoreSurfaces.TestShippingReservationTotals_IgnoreLocallySentActiveLedgerRows | PASS |
