# Phase 6 Live Role Workflow Validation Results

- Date: 2026-06-07 18:54:27
- Deploy root: C:\Users\justu\source\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba
- Passed: 9
- Failed: 13

| Check | Result | Detail |
|---|---|---|
| Core.RuntimeRootOverride | PASS | C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba |
| Core.AuthDiagnostic.User | PASS | ResolvedUser=justin; SeededUsers=justu,Justin Jahn,user1,svc_processor |
| Core.AuthDiagnostic.Config | PASS | WarehouseId=WH1; StationId=S1; PathDataRoot=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba |
| Core.AuthDiagnostic.AuthLoad | PASS |  |
| Core.AuthDiagnostic.TargetSelect | PASS | OK/Connected - WH1 (Main Warehouse) at C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba |
| Core.AuthDiagnostic.SignIn | FAIL | FAIL/3 |
| Core.AuthDiagnostic.ReceiveCapability | FAIL | User=justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ShipCapability | FAIL | User=justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ProdCapability | FAIL | User=justin; WarehouseId=WH1; StationId=S1 |
| Core.ConfigBootstrap.CleanSurface | PASS | Load=True; Validate=; Sheets=2; WHTables=System.String[]; STTables=System.String[] |
| Core.RuntimeInventoryDiagnostic | PASS | Override=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba; PathDataRoot=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba; InventoryPath=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba\WH1.invSys.Data.Inventory.xlsb; FileExists=True; OpenFullName=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba\WH1.invSys.Data.Inventory.xlsb |
| Receiving.ConfirmWrites.Local | FAIL | ReceivedTallyRows=1; AggregateReceivedRows=1; RECEIVED=0; TOTAL_INV=10; QtyOnHand=; SourceType=; IsStale=; LogRows=0 |
| Receiving.ConfirmWrites.Queue | FAIL | InboxRows=0; Row=0 |
| Receiving.ConfirmWrites.Process | FAIL | StatusBeforeRun=; RunBatch=0; Status=; OutboxRow=0; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260607185400-549923; OpenBooks=WH1.invSys.Auth.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba\WH1.invSys.Auth.xlsb; WH1.invSys.Data.Inventory.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba\WH1.invSys.Data.Inventory.xlsb; invSys.Inbox.Receiving.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba\invSys.Inbox.Receiving.S1.xlsb; invSys.Inbox.Shipping.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba\invSys.Inbox.Shipping.S1.xlsb; invSys.Inbox.Production.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba\invSys.Inbox.Production.S1.xlsb; WH1.S1.Receiving.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba\WH1.S1.Receiving.Operator.xlsb; WH1.S1.Shipping.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba\WH1.S1.Shipping.Operator.xlsb; WH1.S1.Production.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-feeba9ac519642a7b6c92eb9f747cfba\WH1.S1.Production.Operator.xlsb; Sheet2=Sheet2; Sheet3=Sheet3 |
| Receiving.ConfirmWrites.InventoryLog | FAIL | InventoryLogRowsBefore=0; Row=0; OutboxRow=0 |
| Shipping.BtnToShipments.Preflight | PASS | AggregatePackagesRows=1; AggROW=201; AggQty=5; InvROW=201; InvCode=SKU-SHIP; InvTOTAL_INV=20; InvSHIPMENTS=0 |
| Shipping.BtnToShipments.Local | PASS | SHIPMENTS=5; AggregatePackagesRows=1 |
| Shipping.BtnShipmentsSent.Local | FAIL | SHIPMENTS=5; AggregatePackagesRows=1 |
| Shipping.BtnShipmentsSent.Queue | FAIL | InboxRows=0; Row=0 |
| Shipping.BtnShipmentsSent.Process | FAIL | StatusBeforeRun=; RunBatch=0; Status=; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260607185424-477670 |
| Shipping.BtnShipmentsSent.InventoryLog | FAIL | InventoryLogRow=0 |
| Harness.Exception | FAIL | Step=Stage Shipping hold workflow; ListObject missing. |
