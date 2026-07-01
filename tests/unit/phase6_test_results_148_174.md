# Phase 6 VBA Test Results

- Date: 2026-07-01 13:24:11
- Passed: 27
- Failed: 0
- Range: 148-174 of 204

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestBoxMakerShippables_MultiVersionUsesCanonicalInventoryLogFallback | PASS |
| TestPhase6CoreSurfaces.TestBoxMakerShippables_VersionNasIgnoresReserveReleaseLogRows | PASS |
| TestPhase6CoreSurfaces.TestBoxMakerHistoryExport_WritesBuildAndUnboxRowsOnly | PASS |
| TestPhase6CoreSurfaces.TestBoxMakerHistoryExport_ReportsEmptySourceDiagnostics | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedDisplay_SubtractsLockedAndUnreservedRows | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowDoesNotAddBackTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_UnreservedDirtyRowDoesNotDeductTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowClearsLockedReservationTotal | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_DoesNotIncreaseProjectedInventoryOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedCompletionKeepsProjectedDeductionWhenNasStale | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_RepairsMissingInvSysRowFromShipmentLine | PASS |
| TestPhase6CoreSurfaces.TestShippingSentPayload_UsesVisibleShipmentItemWhenInvSysCodeStale | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_FullRunNeverIncreasesProjectedInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_EmptySelectionSendsAllShipmentAreaRows | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PreservesNasBaselineAcrossSentReregister | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_EvictsStaleZeroWhenBackendPositive | PASS |
| TestPhase6CoreSurfaces.TestShippingSyncLabel_CountsOnlyQuantityDriftNotVisibleRows | PASS |
| TestPhase6CoreSurfaces.TestShippingHydrateShippables_DoesNotWriteTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_KeepsFreshSentOverlayAtBaseline | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_EvictsSentOverlayWhenNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_DoesNotEvictSentOverlayOnBlankOrZeroBackend | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_KeepsSentOverlayWithActiveLock | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_LocalCatchupDoesNotClearBeforeNas | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_ClearsWhenBackendRisesAboveBaseline | PASS |
| TestPhase6CoreSurfaces.TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected | PASS |
| TestPhase6CoreSurfaces.TestShippingRefresh_MergesLocalBoxBuildStagingAndClearsStaleOverlay | PASS |
