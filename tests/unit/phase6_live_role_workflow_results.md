# Phase 6 Live Role Workflow Validation Results

- Date: 2026-03-23 22:06:36
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-fc8670f75dc74eccbb479a32c6889bcc
- Passed: 25
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Core.RuntimeRootOverride | PASS | C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-fc8670f75dc74eccbb479a32c6889bcc |
| Core.AuthDiagnostic.User | PASS | ResolvedUser=Justin; SeededUsers=Justin,user1,svc_processor |
| Core.AuthDiagnostic.Config | PASS | WarehouseId=WH1; StationId=S1; PathDataRoot=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-live-fc8670f75dc74eccbb479a32c6889bcc |
| Core.AuthDiagnostic.AuthLoad | PASS |  |
| Core.AuthDiagnostic.ReceiveCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ShipCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ProdCapability | PASS | User=Justin; WarehouseId=WH1; StationId=S1 |
| Core.ConfigBootstrap.CleanSurface | PASS | Load=True; Validate=; Sheets=2; WHTables=System.String[]; STTables=System.String[] |
| Receiving.ConfirmWrites.QueueDiagnostic | PASS | OK |
| Receiving.ConfirmWrites.Local | PASS | RECEIVED=7; LogRows=1 |
| Receiving.ConfirmWrites.Queue | PASS | InboxRows=3; Row=3 |
| Receiving.ConfirmWrites.Process | PASS | StatusBeforeRun=PROCESSED; RunBatch=0; Status=PROCESSED; OutboxRow=0; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260323220559-256477 |
| Receiving.ConfirmWrites.InventoryLog | PASS | InventoryLogRowsBefore=0; Row=3; OutboxRow=0 |
| Shipping.BtnShipmentsSent.QueueDiagnostic | PASS | OK |
| Shipping.BtnShipmentsSent.Local | PASS | SHIPMENTS=0 |
| Shipping.BtnShipmentsSent.Queue | PASS | InboxRows=3; Row=2 |
| Shipping.BtnShipmentsSent.Process | PASS | RunBatch=2; Status=PROCESSED; ErrorCode=; ErrorMessage=; Processed=2; Report=Applied=2; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260323220611-233016 |
| Shipping.BtnShipmentsSent.InventoryLog | PASS | InventoryLogRow=4 |
| Production.BtnSavePalette | PASS | PaletteRow=1; Before=ProdWb=Book9; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=0; FirstPaletteRecipe=; After=ProdWb=Book9; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=1; FirstPaletteRecipe=R-001 |
| Production.BtnPrintRecallCodes | PASS | Diag=OK; Sheet=RecallCodesPrint; Rows=1; RecallRows=1; RecallCode=RC-001 |
| Production.BtnToTotalInv.QueueDiagnostic | PASS | OK |
| Production.BtnToTotalInv.Local | PASS | MADE=0; TOTAL_INV=8 |
| Production.BtnToTotalInv.Queue | PASS | InboxRows=3; Row=2 |
| Production.BtnToTotalInv.Process | PASS | RunBatch=2; Status=PROCESSED; ErrorCode=; ErrorMessage=; Processed=2; Report=Applied=2; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260323220635-632691 |
| Production.BtnToTotalInv.InventoryLog | PASS | InventoryLogRow=6 |
