# Phase 6 VBA Test Results

- Date: 2026-06-22 21:29:35
- Passed: 14
- Failed: 0
- Range: 124-137 of 170

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestShippingAdd_DefaultsOrderToWarehouseArea | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_BlankCarrierRequiresCarrier | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_UsesDisplayedProjectedInventoryWhenVersionLedgerIsEmpty | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_UsesDisplayedProjectedInventoryWhenTotalInvIsStaleZero | PASS |
| TestPhase6CoreSurfaces.TestShippingRemove_LockedRowReleasesInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingRemove_StaleLockedRowClearsWithoutInflatingInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingShippables_NasInvPrefersCurrentInvSysForSingleActiveVersion | PASS |
| TestPhase6CoreSurfaces.TestShippingProjectedDisplay_SubtractsLockedAndUnreservedRows | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowDoesNotAddBackTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_UnreservedDirtyRowDeductsTotalInv | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedRowClearsLockedReservationTotal | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_DoesNotIncreaseProjectedInventoryOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_ReservedCompletionKeepsProjectedDeductionWhenNasStale | PASS |
| TestPhase6CoreSurfaces.TestShippingSentRows_FullRunNeverIncreasesProjectedInventory | PASS |
