# Phase 6 Live Role Workflow Validation Results

- Date: 2026-06-02 12:25:22
- Deploy root: C:\Users\justu\source\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613
- Passed: 21
- Failed: 16

| Check | Result | Detail |
|---|---|---|
| Core.RuntimeRootOverride | PASS | C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613 |
| Core.AuthDiagnostic.User | PASS | ResolvedUser=justu; SeededUsers=justu,Justin Jahn,user1,svc_processor |
| Core.AuthDiagnostic.Config | PASS | WarehouseId=WH1; StationId=S1; PathDataRoot=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613 |
| Core.AuthDiagnostic.AuthLoad | PASS |  |
| Core.AuthDiagnostic.TargetSelect | PASS | OK/Connected - WH1 (Main Warehouse) at C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613 |
| Core.AuthDiagnostic.SignIn | PASS | OK/User=justu/DisplayName=justu |
| Core.AuthDiagnostic.ReceiveCapability | PASS | User=justu; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ShipCapability | PASS | User=justu; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ProdCapability | PASS | User=justu; WarehouseId=WH1; StationId=S1 |
| Core.ConfigBootstrap.CleanSurface | PASS | Load=True; Validate=; Sheets=2; WHTables=System.String[]; STTables=System.String[] |
| Core.RuntimeInventoryDiagnostic | PASS | Override=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613; PathDataRoot=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613; InventoryPath=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613\WH1.invSys.Data.Inventory.xlsb; FileExists=True; OpenFullName=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613\WH1.invSys.Data.Inventory.xlsb |
| Receiving.ConfirmWrites.Local | FAIL | ReceivedTallyRows=1; AggregateReceivedRows=1; RECEIVED=0; TOTAL_INV=10; QtyOnHand=; SourceType=; IsStale=; LogRows=0 |
| Receiving.ConfirmWrites.Queue | FAIL | InboxRows=0; Row=0 |
| Receiving.ConfirmWrites.Process | FAIL | StatusBeforeRun=; RunBatch=0; Status=; OutboxRow=0; ErrorCode=; ErrorMessage=; Processed=0; Report=Inventory workbook not found.; OpenBooks=WH1.invSys.Auth.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613\WH1.invSys.Auth.xlsb; WH1.invSys.Data.Inventory.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613\WH1.invSys.Data.Inventory.xlsb; invSys.Inbox.Receiving.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613\invSys.Inbox.Receiving.S1.xlsb; invSys.Inbox.Shipping.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613\invSys.Inbox.Shipping.S1.xlsb; invSys.Inbox.Production.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613\invSys.Inbox.Production.S1.xlsb; WH1.S1.Receiving.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613\WH1.S1.Receiving.Operator.xlsb; WH1.S1.Shipping.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613\WH1.S1.Shipping.Operator.xlsb; WH1.S1.Production.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-83e56db5ce5040b980b871e03f740613\WH1.S1.Production.Operator.xlsb; Sheet2=Sheet2; Sheet3=Sheet3; Sheet4=Sheet4 |
| Receiving.ConfirmWrites.InventoryLog | FAIL | InventoryLogRowsBefore=0; Row=0; OutboxRow=0 |
| Shipping.BtnToShipments.Preflight | PASS | AggregatePackagesRows=1; AggROW=201; AggQty=5; InvROW=201; InvCode=SKU-SHIP; InvTOTAL_INV=20; InvSHIPMENTS=0 |
| Shipping.BtnToShipments.Local | PASS | SHIPMENTS=5; AggregatePackagesRows=0 |
| Shipping.BtnShipmentsSent.Local | FAIL | SHIPMENTS=5; AggregatePackagesRows=0 |
| Shipping.BtnShipmentsSent.Queue | FAIL | InboxRows=0; Row=0 |
| Shipping.BtnShipmentsSent.Process | FAIL | StatusBeforeRun=; RunBatch=0; Status=; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260602122501-590353 |
| Shipping.BtnShipmentsSent.InventoryLog | FAIL | InventoryLogRow=0 |
| Shipping.Hold.ToggleNotShipped | PASS | InitialHidden=False; AfterFirst=True; AfterSecond=False |
| Shipping.Hold.Send | PASS | Result=OK/Moved=4/SourceQty=6/TargetQty=4; ShipQty=6; HoldQty=4; HoldROW=250 |
| Shipping.Hold.Return | PASS | Result=OK/Moved=4/SourceQty=0/TargetQty=10; ShipQty=10; HoldQty=0 |
| Shipping.BtnBoxesMade.Local | PASS | ComponentUSED=0; ComponentTOTAL_INV=7; PackageMADE=2; AggregatePackagesRows=0 |
| Shipping.BtnToTotalInv.Local | PASS | PackageMADE=0; PackageTOTAL_INV=2 |
| Production.BtnSavePalette | PASS | PaletteRow=1; Before=ProdWb=WH1.S1.Production.Operator.xlsb; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=0; FirstPaletteRecipe=; After=ProdWb=WH1.S1.Production.Operator.xlsb; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=1; FirstPaletteRecipe=R-001 |
| Production.BtnPrintRecallCodes | PASS | Diag=OK; Sheet=RecallCodesPrint; Rows=1; RecallRows=1; RecallCode=RC-001 |
| Production.BtnToMade.Preflight | PASS | ProcessTables=RecipeChooser_generated:Rows=0,Process=,IO=,Ingredient=,Amount=; proc_1_rchooser:Rows=1,Process=Mix,IO=MADE,Ingredient=Finished Good,Amount=8; ProcessCheckboxes=0; OutputROW=401; RealOutput=8; InvRow2Code=SKU-FG |
| Production.BtnToMade.Local | FAIL | MADE=0; TOTAL_INV=0; RealOutput=8 |
| Production.BtnToMade.Queue | FAIL | InboxRows=0; Row=0 |
| Production.BtnToMade.Process | FAIL | Status=; ErrorCode=; ErrorMessage= |
| Production.BtnToMade.InventoryLog | FAIL | InventoryLogRow=0 |
| Production.BtnToTotalInv.Local | FAIL | MADE=0; TOTAL_INV=0; ProductionOutputRows=1 |
| Production.BtnToTotalInv.Queue | FAIL | InboxRows=0; Row=0 |
| Production.BtnToTotalInv.Process | FAIL | StatusBeforeRun=; RunBatch=0; Status=; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260602122520-271379 |
| Production.BtnToTotalInv.InventoryLog | FAIL | InventoryLogRow=0 |
