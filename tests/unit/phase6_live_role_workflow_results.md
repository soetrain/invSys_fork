# Phase 6 Live Role Workflow Validation Results

- Date: 2026-03-22 18:34:42
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-c82eebefb4bb48b98574e10c3a2245b2
- Passed: 23
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Core.RuntimeRootOverride | PASS | C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-c82eebefb4bb48b98574e10c3a2245b2 |
| Core.AuthDiagnostic.User | PASS | ResolvedUser=Justin; SeededUsers=Justin,user1,svc_processor |
| Core.AuthDiagnostic.Config | PASS | WarehouseId=WH1; StationId=S1; PathDataRoot=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-c82eebefb4bb48b98574e10c3a2245b2 |
| Core.AuthDiagnostic.AuthLoad | PASS |  |
| Core.AuthDiagnostic.ReceiveCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ShipCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ProdCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Receiving.ConfirmWrites.QueueDiagnostic | PASS | OK |
| Receiving.ConfirmWrites.Local | PASS | RECEIVED=7; LogRows=1 |
| Receiving.ConfirmWrites.Queue | PASS | InboxRows=3; Row=3 |
| Receiving.ConfirmWrites.Process | PASS | RunBatch=2; Status=; OutboxRow=0; ErrorCode=; ErrorMessage=; Processed=2; Report=Applied=2; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260322183354-590557 |
| Receiving.ConfirmWrites.InventoryLog | PASS | InventoryLogRowsBefore=0; Row=0; OutboxRow=0 |
| Shipping.BtnShipmentsSent.QueueDiagnostic | PASS | OK |
| Shipping.BtnShipmentsSent.Local | PASS | SHIPMENTS=0 |
| Shipping.BtnShipmentsSent.Queue | PASS | InboxRows=3; Row=2 |
| Shipping.BtnShipmentsSent.Process | PASS | RunBatch=2; Status=PROCESSED; ErrorCode=; ErrorMessage=; Processed=2; Report=Applied=2; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260322183416-162927 |
| Shipping.BtnShipmentsSent.InventoryLog | PASS | InventoryLogRow=4 |
| Production.BtnSavePalette | PASS | PaletteRow=1 |
| Production.BtnToTotalInv.QueueDiagnostic | PASS | OK |
| Production.BtnToTotalInv.Local | PASS | MADE=0; TOTAL_INV=8 |
| Production.BtnToTotalInv.Queue | PASS | InboxRows=3; Row=2 |
| Production.BtnToTotalInv.Process | PASS | RunBatch=2; Status=PROCESSED; ErrorCode=; ErrorMessage=; Processed=2; Report=Applied=2; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260322183440-766321 |
| Production.BtnToTotalInv.InventoryLog | PASS | InventoryLogRow=6 |
