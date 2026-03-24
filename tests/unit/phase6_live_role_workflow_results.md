# Phase 6 Live Role Workflow Validation Results

- Date: 2026-03-23 23:51:48
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\hotfix-inventory-canonical-20260323
- Runtime root override: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-4a7cc8c5e6424ee6a9adbe2e174886ac
- Passed: 19
- Failed: 6

| Check | Result | Detail |
|---|---|---|
| Core.RuntimeRootOverride | PASS | C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-4a7cc8c5e6424ee6a9adbe2e174886ac |
| Core.AuthDiagnostic.User | PASS | ResolvedUser=Justin; SeededUsers=Justin,user1,svc_processor |
| Core.AuthDiagnostic.Config | PASS | WarehouseId=WH1; StationId=S1; PathDataRoot=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-4a7cc8c5e6424ee6a9adbe2e174886ac |
| Core.AuthDiagnostic.AuthLoad | PASS |  |
| Core.AuthDiagnostic.ReceiveCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ShipCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ProdCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.ConfigBootstrap.CleanSurface | PASS | Load=True; Validate=; Sheets=2; WHTables=System.String[]; STTables=System.String[] |
| Receiving.ConfirmWrites.QueueDiagnostic | PASS | OK |
| Receiving.ConfirmWrites.Local | PASS | RECEIVED=7; LogRows=1 |
| Receiving.ConfirmWrites.Queue | PASS | InboxRows=3; Row=3 |
| Receiving.ConfirmWrites.Process | FAIL | StatusBeforeRun=NEW; RunBatch=0; Status=NEW; OutboxRow=0; ErrorCode=; ErrorMessage=; Processed=0; Report=Inventory workbook not found. |
| Receiving.ConfirmWrites.InventoryLog | FAIL | InventoryLogRowsBefore=0; Row=0; OutboxRow=0 |
| Shipping.BtnShipmentsSent.QueueDiagnostic | PASS | OK |
| Shipping.BtnShipmentsSent.Local | PASS | SHIPMENTS=0 |
| Shipping.BtnShipmentsSent.Queue | PASS | InboxRows=3; Row=2 |
| Shipping.BtnShipmentsSent.Process | FAIL | RunBatch=0; Status=NEW; ErrorCode=; ErrorMessage=; Processed=0; Report=Inventory workbook not found. |
| Shipping.BtnShipmentsSent.InventoryLog | FAIL | InventoryLogRow=0 |
| Production.BtnSavePalette | PASS | PaletteRow=1; Before=ProdWb=Book9; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=0; FirstPaletteRecipe=; After=ProdWb=Book9; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=1; FirstPaletteRecipe=R-001 |
| Production.BtnPrintRecallCodes | PASS | Diag=OK; Sheet=RecallCodesPrint; Rows=1; RecallRows=1; RecallCode=RC-001 |
| Production.BtnToTotalInv.QueueDiagnostic | PASS | OK |
| Production.BtnToTotalInv.Local | PASS | MADE=0; TOTAL_INV=8 |
| Production.BtnToTotalInv.Queue | PASS | InboxRows=3; Row=2 |
| Production.BtnToTotalInv.Process | FAIL | RunBatch=0; Status=NEW; ErrorCode=; ErrorMessage=; Processed=0; Report=Inventory workbook not found. |
| Production.BtnToTotalInv.InventoryLog | FAIL | InventoryLogRow=0 |
