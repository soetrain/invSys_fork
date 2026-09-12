# Slice 14 Full-Chain, Restart, and Reconciliation Evidence

- Date: 2026-09-12 15:40:12
- Package set: R1-5
- Ordered phases: 
GenerateFreshWarehouse -> SeedDemoInventoryThroughAdmin -> ReceiveInventory -> ProcessorApplyReceive -> RefreshAfterReceive -> ProductionTwoBatches -> ProductionConsumptionAndOutput -> BoxingVersionSelection -> ShipmentStagingAndSent -> ProcessorApplyShipment -> FinalRefresh -> RestartAndReconcile
- Passed: 30
- Failed: 0

## D13 trace

- Focused RED: 0/7 before the dedicated validator and packaged Admin seed primitive existed.
- Behavioral RED: 27/30 exposed a negative detailed entity caused by an identity-free fixture seed and an XLAM-only runtime extraction blind spot.
- Evidence RED: 7/8 exposed an unredacted generated processor RunId in the committed report.
- GREEN: focused contract 9/9 and ordered packaged full chain 30/30.

| Check | Result | Detail |
|---|---|---|
| GenerateFreshWarehouse | PASS | Packaged Admin created a fresh greenfield warehouse runtime. |
| SeedDemoInventoryThroughAdmin | PASS | OK/Demo inventory seeded./Created=21/Skipped=3/Applied=1/Processor=Applied=1; SkipDup=0; Poison=0; RunId=<redacted>; EventPersistenceSaves=3 |
| AdminEntry.InventoryCreated | PASS | The packaged entry boundary produced the canonical inventory workbook. |
| AdminEntry.SourceIntegrationRegression | PASS | Create Warehouse D14 source integration remained green. |
| ReceiveInventory | PASS | Packaged Receiving used its captured-workbook form action. |
| ProcessorApplyReceive | PASS | The processor applied the Receive event and wrote canonical evidence. |
| RefreshAfterReceive | PASS | The read-model projections rebuilt from authoritative log state. |
| ProductionTwoBatches | PASS | The packaged Production action completed two consecutive batches. |
| ProductionConsumptionAndOutput | PASS | Production consumption and output events applied through the processor. |
| BoxingVersionSelection | PASS | The packaged Box Maker service created the v1 shippable before shipment staging. |
| ShipmentStagingAndSent | PASS | The packaged Shipments Sent action posted the v1 box identity at BIN-B and cleared only its captured workbook. |
| ProcessorApplyShipment | PASS | The processor applied the shipment event and wrote its log row. |
| ExactBalancesAndLocations | PASS | The pre-Shipping projection checkpoint reconciled after Receiving and Production. |
| ProductionBatchState | PASS | Two-batch packaged evidence includes batch/Last/Total and ready-next state. |
| BoxingBomVersion | PASS | The versioned Boxing service applied v1 and returned the output System_Key. |
| OverlayPreserved | PASS | Shipping form action evidence preserved the captured operator workbook boundary. |
| FinalRefresh | PASS | A canonical snapshot was generated and all three saved operator read models refreshed after reopen. |
| HeaderPersistence | PASS | After restart, an end-user column/value survived snapshot refresh and read-model rebuild. |
| NoRowHeaders | PASS | Canonical and reopened operator runtime tables contain no managed ROW header. |
| UniqueSystemKeys | PASS | Every canonical detailed entity has one nonblank unique System_Key after the chain. |
| NoNegativeInventory | PASS | No canonical detailed entity has negative QtyOnHand. |
| ExactBalancesAndLocations.Final | PASS | Final canonical balances after two additional batches, Boxing, and shipment: SKU-BOX=1; SKU-FG=22; SKU-REC=8; SKU-SHIP=20; SKU-SUGAR=94 |
| ExactBalancesAndLocations.Location | PASS | Nonzero location balances: SKU-BOX/BIN-B=1; SKU-COMP/LINE=10; SKU-FG/BIN-A=22; SKU-REC/A1=8; SKU-SHIP/DOCK=20; SKU-SUGAR/BIN-A=94 |
| EventIdentityStatusLogAndReplay | PASS | Replaying all saved inboxes appended no log rows; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=<redacted>; EventPersistenceSaves=0 |
| LocksReleased | PASS | No active inventory locks remain after Shipments Sent and replay. |
| NoDuplicatePackagesOrCallbacks | PASS | Exactly one instance of each Release 1 package was reopened. |
| RestartReconciliation | PASS | Saved canonical and operator workbooks reconciled after a new Excel runtime opened them. |
| CanonicalWorkbooksHidden | PASS | The reconciliation runtime kept canonical workbooks out of the visible operator surface. |
| RuntimeFivePackages | PASS | Read-only extractor observed: invSys.Admin.xlam, invSys.Core.xlam, invSys.Designs.Domain.xlam, invSys.Inventory.Domain.xlam, invSys.Operations.xlam; Suspended registrations=2; Unexpected active packages=0 |
| StaticRetiredPathRatchet | PASS | New static warning paths=0; Current warnings=27; Baseline warnings=27 |
