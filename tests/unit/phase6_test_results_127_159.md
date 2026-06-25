# Phase 6 VBA Test Results

- Date: 2026-06-24 21:01:52
- Passed: 10
- Failed: 0
- Range: 127-159 of 185
- Status: PARTIAL

| Test | Result |
|---|---|
| TestPhase6CoreSurfaces.TestShippingAdd_UsesDisplayedProjectedInventoryWhenVersionLedgerIsEmpty | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_UsesDisplayedProjectedInventoryWhenTotalInvIsStaleZero | PASS |
| TestPhase6CoreSurfaces.TestShippingRemove_LockedRowReleasesInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingRemove_StaleLockedRowClearsWithoutInflatingInventory | PASS |
| TestPhase6CoreSurfaces.TestShippingHold_PreservesReservationAndLocalDeduction | PASS |
| TestPhase6CoreSurfaces.TestShippingToShipments_ReservedMultiSelectKeepsRowsAndProjection | PASS |
| TestPhase6CoreSurfaces.TestShippingUpdate_PreservesExistingReservationWithoutDoubleDeducting | PASS |
| TestPhase6CoreSurfaces.TestShippingUpdate_ReservedQtyChangeAppliesOnlyDeltaOverlay | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_NewReservedRowAppliesSingleProjectedDeduction | PASS |
| TestPhase6CoreSurfaces.TestShippingAdd_MergingExistingReservedRowAppliesOnlyDelta | PASS |
