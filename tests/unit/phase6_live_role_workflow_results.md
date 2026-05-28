# Phase 6 Live Role Workflow Validation Results

- Date: 2026-05-27 17:17:57
- Deploy root: C:\Users\justu\source\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d
- Passed: 37
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Core.RuntimeRootOverride | PASS | C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d |
| Core.AuthDiagnostic.User | PASS | ResolvedUser=justu; SeededUsers=justu,Justin Jahn,user1,svc_processor |
| Core.AuthDiagnostic.Config | PASS | WarehouseId=WH1; StationId=S1; PathDataRoot=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d |
| Core.AuthDiagnostic.AuthLoad | PASS |  |
| Core.AuthDiagnostic.TargetSelect | PASS | OK/Connected - WH1 (Main Warehouse) at C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d |
| Core.AuthDiagnostic.SignIn | PASS | OK/User=justu/DisplayName=justu |
| Core.AuthDiagnostic.ReceiveCapability | PASS | User=justu; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ShipCapability | PASS | User=justu; WarehouseId=WH1; StationId=S1 |
| Core.AuthDiagnostic.ProdCapability | PASS | User=justu; WarehouseId=WH1; StationId=S1 |
| Core.ConfigBootstrap.CleanSurface | PASS | Load=True; Validate=; Sheets=2; WHTables=System.String[]; STTables=System.String[] |
| Core.RuntimeInventoryDiagnostic | PASS | Override=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d; PathDataRoot=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d; InventoryPath=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d\WH1.invSys.Data.Inventory.xlsb; FileExists=True; OpenFullName=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d\WH1.invSys.Data.Inventory.xlsb |
| Receiving.ConfirmWrites.Local | PASS | ReceivedTallyRows=0; AggregateReceivedRows=0; RECEIVED=0; TOTAL_INV=10; QtyOnHand=; SourceType=CACHED; IsStale=True; LogRows=1 |
| Receiving.ConfirmWrites.Queue | PASS | InboxRows=1; Row=1 |
| Receiving.ConfirmWrites.Process | PASS | StatusBeforeRun=NEW; RunBatch=1; Status=PROCESSED; OutboxRow=0; ErrorCode=; ErrorMessage=; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260527171636-086392; OpenBooks=WH1.invSys.Auth.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d\WH1.invSys.Auth.xlsb; WH1.invSys.Data.Inventory.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d\WH1.invSys.Data.Inventory.xlsb; invSys.Inbox.Receiving.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d\invSys.Inbox.Receiving.S1.xlsb; invSys.Inbox.Shipping.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d\invSys.Inbox.Shipping.S1.xlsb; invSys.Inbox.Production.S1.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d\invSys.Inbox.Production.S1.xlsb; WH1.S1.Receiving.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d\WH1.S1.Receiving.Operator.xlsb; WH1.S1.Shipping.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d\WH1.S1.Shipping.Operator.xlsb; WH1.S1.Production.Operator.xlsb=C:\Users\justu\AppData\Local\Temp\invsys-phase6-live-50734a4b1138415d8b17a19a5ba6b61d\WH1.S1.Production.Operator.xlsb; Sheet2=Sheet2; Sheet3=Sheet3; Sheet4=Sheet4; Sheet5=Sheet5 |
| Receiving.ConfirmWrites.InventoryLog | PASS | InventoryLogRowsBefore=0; Row=1; OutboxRow=0 |
| Shipping.BtnToShipments.Preflight | PASS | AggregatePackagesRows=1; AggROW=201; AggQty=5; InvROW=201; InvCode=SKU-SHIP; InvTOTAL_INV=20; InvSHIPMENTS=0 |
| Shipping.BtnToShipments.Local | PASS | SHIPMENTS=5; AggregatePackagesRows=0 |
| Shipping.BtnShipmentsSent.Local | PASS | SHIPMENTS=0; AggregatePackagesRows=0 |
| Shipping.BtnShipmentsSent.Queue | PASS | InboxRows=1; Row=1 |
| Shipping.BtnShipmentsSent.Process | PASS | StatusBeforeRun=PROCESSED; RunBatch=0; Status=PROCESSED; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260527171706-133050 |
| Shipping.BtnShipmentsSent.InventoryLog | PASS | InventoryLogRow=2 |
| Shipping.Hold.ToggleNotShipped | PASS | InitialHidden=False; AfterFirst=True; AfterSecond=False |
| Shipping.Hold.Send | PASS | Result=OK/Moved=4/SourceQty=6/TargetQty=4; ShipQty=6; HoldQty=4; HoldROW=250 |
| Shipping.Hold.Return | PASS | Result=OK/Moved=4/SourceQty=0/TargetQty=10; ShipQty=10; HoldQty=0 |
| Shipping.BtnBoxesMade.Local | PASS | ComponentUSED=0; ComponentTOTAL_INV=7; PackageMADE=2; AggregatePackagesRows=0 |
| Shipping.BtnToTotalInv.Local | PASS | PackageMADE=0; PackageTOTAL_INV=2 |
| Production.BtnSavePalette | PASS | PaletteRow=1; Before=ProdWb=WH1.S1.Production.Operator.xlsb; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=0; FirstPaletteRecipe=; After=ProdWb=WH1.S1.Production.Operator.xlsb; RecipesSheet=Recipes; PaletteSheet=IngredientsPalette; RecipeId=R-001; IngredientId=ING-001; ChooseRecipeRows=1; ChooseIngredientRows=1; ChooseItemRows=1; FirstItem=Sugar Bin; PaletteRows=1; FirstPaletteRecipe=R-001 |
| Production.BtnPrintRecallCodes | PASS | Diag=OK; Sheet=RecallCodesPrint; Rows=1; RecallRows=1; RecallCode=RC-001 |
| Production.BtnToMade.Preflight | PASS | ProcessTables=RecipeChooser_generated:Rows=0,Process=,IO=,Ingredient=,Amount=; proc_1_rchooser:Rows=1,Process=Mix,IO=MADE,Ingredient=Finished Good,Amount=8; ProcessCheckboxes=0; OutputROW=401; RealOutput=8; InvRow2Code=SKU-FG |
| Production.BtnToMade.Local | PASS | MADE=8; TOTAL_INV=0; RealOutput=8 |
| Production.BtnToMade.Queue | PASS | InboxRows=1; Row=1 |
| Production.BtnToMade.Process | PASS | Status=PROCESSED; ErrorCode=; ErrorMessage= |
| Production.BtnToMade.InventoryLog | PASS | InventoryLogRow=3 |
| Production.BtnToTotalInv.Local | PASS | MADE=0; TOTAL_INV=8; ProductionOutputRows=1 |
| Production.BtnToTotalInv.Queue | PASS | InboxRows=2; Row=2 |
| Production.BtnToTotalInv.Process | PASS | StatusBeforeRun=PROCESSED; RunBatch=0; Status=PROCESSED; ErrorCode=; ErrorMessage=; Processed=0; Report=Applied=0; SkipDup=0; Poison=0; RunId=RUN-WH1-INVENTORY-20260527171754-438755 |
| Production.BtnToTotalInv.InventoryLog | PASS | InventoryLogRow=4 |
