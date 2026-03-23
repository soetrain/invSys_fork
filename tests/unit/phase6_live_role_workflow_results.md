# Phase 6 Live Role Workflow Validation Results

- Date: 2026-03-22 17:48:06
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-a8ea3076c2854ad3a3a12251064e2755
- Passed: 7
- Failed: 6

| Check | Result | Detail |
|---|---|---|
| Core.RuntimeRootOverride | PASS | C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-a8ea3076c2854ad3a3a12251064e2755 |
| Core.AuthDiagnostic.User | PASS | ResolvedUser=Justin; SeededUsers=Justin,user1,svc_processor |
| Core.AuthDiagnostic.Config | PASS | WarehouseId=WH1; StationId=S1; PathDataRoot=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-a8ea3076c2854ad3a3a12251064e2755 |
| Core.AuthDiagnostic.AuthLoad | PASS |  |
| Core.AuthDiagnostic.ReceiveCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ShipCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ProdCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Receiving.ConfirmWrites.QueueDiagnostic | FAIL | AggregateReceived has no rows to confirm. |
| Receiving.ConfirmWrites.Local | FAIL | RECEIVED=0; LogRows=0 |
| Receiving.ConfirmWrites.Queue | FAIL | InboxRows=0; Row=0 |
| Receiving.ConfirmWrites.Process | FAIL | RunBatch=0; Status=; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260322174640-722988 |
| Receiving.ConfirmWrites.InventoryLog | FAIL | InventoryLogRowsBefore=0; Row=0 |
| Harness.Exception | FAIL | Step=Stage Shipping workflow; Exception calling "Run" with "1" argument(s): "The remote procedure call failed. (Exception from HRESULT: 0x800706BE)" |
