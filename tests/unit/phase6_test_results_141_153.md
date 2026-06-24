# Phase 6 VBA Test Results

- Date: 2026-06-24 14:17:07
- Passed: 13
- Failed: 0
- Range: 141-153 of 183

| Test | Result |
|---|---|
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
