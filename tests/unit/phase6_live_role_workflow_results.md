# Phase 6 Live Role Workflow Validation Results

- Date: 2026-03-28 15:37:13
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-7bd4c22d26444c00a1276d0899c3ee15
- Passed: 17
- Failed: 8

| Check | Result | Detail |
|---|---|---|
| Core.RuntimeRootOverride | PASS | C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-7bd4c22d26444c00a1276d0899c3ee15 |
| Core.AuthDiagnostic.User | PASS | ResolvedUser=Justin; SeededUsers=Justin,user1,svc_processor |
| Core.AuthDiagnostic.Config | PASS | WarehouseId=WH1; StationId=S1; PathDataRoot=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-7bd4c22d26444c00a1276d0899c3ee15 |
| Core.AuthDiagnostic.AuthLoad | PASS |  |
| Core.AuthDiagnostic.ReceiveCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ShipCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ProdCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.ConfigBootstrap.CleanSurface | PASS | Load=True; Validate=; Sheets=2; WHTables=System.String[]; STTables=System.String[] |
| Receiving.ConfirmWrites.QueueDiagnostic | PASS | OK |
| Receiving.ConfirmWrites.Local | FAIL | RECEIVED=0; TOTAL_INV=0; QtyOnHand=; SourceType=LOCAL; IsStale=False; LogRows=1 |
| Receiving.ConfirmWrites.Queue | PASS | InboxRows=2; Row=2 |
| Receiving.ConfirmWrites.Process | FAIL | StatusBeforeRun=NEW; RunBatch=0; Status=NEW; OutboxRow=0; ErrorCode=; ErrorMessage=; Processed=0; Report=Inventory workbook is read-only or locked by another Excel session. |
| Receiving.ConfirmWrites.InventoryLog | FAIL | InventoryLogRowsBefore=0; Row=0; OutboxRow=0 |
| Shipping.BtnShipmentsSent.QueueDiagnostic | PASS | OK |
| Shipping.BtnShipmentsSent.Local | PASS | SHIPMENTS=0 |
| Shipping.BtnShipmentsSent.Queue | PASS | InboxRows=2; Row=1 |
| Shipping.BtnShipmentsSent.Process | FAIL | RunBatch=0; Status=NEW; ErrorCode=; ErrorMessage=; Processed=0; Report=Inventory workbook is read-only or locked by another Excel session. |
| Shipping.BtnShipmentsSent.InventoryLog | FAIL | InventoryLogRow=0 |
| Production.BtnSavePalette | PASS | PaletteRow=1; Before=ProdWb=Book10; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=0; FirstPaletteRecipe=; After=ProdWb=Book10; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=1; FirstPaletteRecipe=R-001 |
| Production.BtnPrintRecallCodes | PASS | Diag=OK; Sheet=RecallCodesPrint; Rows=1; RecallRows=1; RecallCode=RC-001 |
| Production.BtnToTotalInv.QueueDiagnostic | PASS | OK |
| Production.BtnToTotalInv.Local | FAIL | MADE=0; TOTAL_INV=0 |
| Production.BtnToTotalInv.Queue | PASS | InboxRows=2; Row=1 |
| Production.BtnToTotalInv.Process | FAIL | RunBatch=0; Status=NEW; ErrorCode=; ErrorMessage=; Processed=0; Report=Inventory workbook is read-only or locked by another Excel session. |
| Production.BtnToTotalInv.InventoryLog | FAIL | InventoryLogRow=0 |
