# Phase 6 VBA Test Results

- Date: 2026-06-24 14:06:00
- Passed: 9
- Failed: 0
- Range: 145-153 of 183

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PreservesNasBaselineAcrossSentReregister | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_EvictsStaleZeroWhenBackendPositive | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_KeepsFreshSentOverlayAtBaseline | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_EvictsSentOverlayWhenNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_KeepsSentOverlayWithActiveLock | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_PersistsAcrossRestartUntilNasCatchesUp | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_LocalCatchupDoesNotClearBeforeNas | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedOverlay_ClearsWhenBackendRisesAboveBaseline | PASS |
| TestPhase6CoreSurfaces.TestShippingReserve_RunBatchRefreshUpdatesNasInvFromProjected | PASS |
