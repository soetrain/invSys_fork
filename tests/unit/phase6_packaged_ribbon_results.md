# Phase 6 Packaged Ribbon Validation Results

- Date: 2026-04-01 16:48:37
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-ribbon-a357e9651d374d12b6bd0ae67c2e842e
- Passed: 76
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| invSys.Core.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Core.xlam |
| invSys.Inventory.Domain.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Inventory.Domain.xlam |
| invSys.Designs.Domain.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Designs.Domain.xlam |
| invSys.Receiving.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Receiving.xlam |
| invSys.Shipping.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Shipping.xlam |
| invSys.Production.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Production.xlam |
| invSys.Admin.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\current\invSys.Admin.xlam |
| Core.RuntimeRootOverride | PASS | C:\Users\Justin\AppData\Local\Temp\invsys-phase6-ribbon-a357e9651d374d12b6bd0ae67c2e842e |
| Receiving.RibbonXml | PASS | customUI/customUI.xml present. |
| Receiving.CallbackModule | PASS | modRibbonGenerated |
| Receiving.RibbonButton.btnReceivingSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionReceiving; Screentip= |
| Receiving.MacroExists.btnReceivingSetup | PASS | modTS_Received.EnsureGeneratedButtons |
| Receiving.CallbackMap.btnReceivingSetup | PASS | btnReceivingSetup -> modTS_Received.EnsureGeneratedButtons |
| Receiving.SafeExec.btnReceivingSetup | PASS | modTS_Received.EnsureGeneratedButtons |
| Receiving.RibbonButton.btnReceivingConfirm | PASS | Label=Confirm Writes; OnAction=RibbonOnActionReceiving; Screentip= |
| Receiving.MacroExists.btnReceivingConfirm | PASS | modTS_Received.ConfirmWrites |
| Receiving.CallbackMap.btnReceivingConfirm | PASS | btnReceivingConfirm -> modTS_Received.ConfirmWrites |
| Receiving.RibbonButton.btnReceivingUndo | PASS | Label=Undo; OnAction=RibbonOnActionReceiving; Screentip= |
| Receiving.MacroExists.btnReceivingUndo | PASS | modTS_Received.MacroUndo |
| Receiving.CallbackMap.btnReceivingUndo | PASS | btnReceivingUndo -> modTS_Received.MacroUndo |
| Receiving.RibbonButton.btnReceivingRedo | PASS | Label=Redo; OnAction=RibbonOnActionReceiving; Screentip= |
| Receiving.MacroExists.btnReceivingRedo | PASS | modTS_Received.MacroRedo |
| Receiving.CallbackMap.btnReceivingRedo | PASS | btnReceivingRedo -> modTS_Received.MacroRedo |
| Shipping.RibbonXml | PASS | customUI/customUI.xml present. |
| Shipping.CallbackModule | PASS | modRibbonGenerated |
| Shipping.RibbonButton.btnShippingSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionShipping; Screentip= |
| Shipping.MacroExists.btnShippingSetup | PASS | modTS_Shipments.InitializeShipmentsUI |
| Shipping.CallbackMap.btnShippingSetup | PASS | btnShippingSetup -> modTS_Shipments.InitializeShipmentsUI |
| Shipping.SafeExec.btnShippingSetup | PASS | modTS_Shipments.InitializeShipmentsUI |
| Shipping.RibbonButton.btnShippingConfirm | PASS | Label=Confirm Inventory; OnAction=RibbonOnActionShipping; Screentip= |
| Shipping.MacroExists.btnShippingConfirm | PASS | modTS_Shipments.BtnConfirmInventory |
| Shipping.CallbackMap.btnShippingConfirm | PASS | btnShippingConfirm -> modTS_Shipments.BtnConfirmInventory |
| Shipping.RibbonButton.btnShippingStage | PASS | Label=To Shipments; OnAction=RibbonOnActionShipping; Screentip= |
| Shipping.MacroExists.btnShippingStage | PASS | modTS_Shipments.BtnToShipments |
| Shipping.CallbackMap.btnShippingStage | PASS | btnShippingStage -> modTS_Shipments.BtnToShipments |
| Shipping.RibbonButton.btnShippingSend | PASS | Label=Shipments Sent; OnAction=RibbonOnActionShipping; Screentip= |
| Shipping.MacroExists.btnShippingSend | PASS | modTS_Shipments.BtnShipmentsSent |
| Shipping.CallbackMap.btnShippingSend | PASS | btnShippingSend -> modTS_Shipments.BtnShipmentsSent |
| Production.RibbonXml | PASS | customUI/customUI.xml present. |
| Production.CallbackModule | PASS | modRibbonGenerated |
| Production.RibbonButton.btnProductionSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionSetup | PASS | mProduction.InitializeProductionUI |
| Production.CallbackMap.btnProductionSetup | PASS | btnProductionSetup -> mProduction.InitializeProductionUI |
| Production.SafeExec.btnProductionSetup | PASS | mProduction.InitializeProductionUI |
| Production.RibbonButton.btnProductionLoad | PASS | Label=Load Recipe; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionLoad | PASS | mProduction.BtnLoadRecipe |
| Production.CallbackMap.btnProductionLoad | PASS | btnProductionLoad -> mProduction.BtnLoadRecipe |
| Production.RibbonButton.btnProductionUsed | PASS | Label=To Used; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionUsed | PASS | mProduction.BtnToUsed |
| Production.CallbackMap.btnProductionUsed | PASS | btnProductionUsed -> mProduction.BtnToUsed |
| Production.RibbonButton.btnProductionMade | PASS | Label=To Made; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionMade | PASS | mProduction.BtnToMade |
| Production.CallbackMap.btnProductionMade | PASS | btnProductionMade -> mProduction.BtnToMade |
| Production.RibbonButton.btnProductionTotal | PASS | Label=To Total Inv; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionTotal | PASS | mProduction.BtnToTotalInv |
| Production.CallbackMap.btnProductionTotal | PASS | btnProductionTotal -> mProduction.BtnToTotalInv |
| Production.RibbonButton.btnProductionPrintCodes | PASS | Label=Print Recall Codes; OnAction=RibbonOnActionProduction; Screentip= |
| Production.MacroExists.btnProductionPrintCodes | PASS | mProduction.BtnPrintRecallCodes |
| Production.CallbackMap.btnProductionPrintCodes | PASS | btnProductionPrintCodes -> mProduction.BtnPrintRecallCodes |
| Admin.RibbonXml | PASS | customUI/customUI.xml present. |
| Admin.CallbackModule | PASS | modRibbonGenerated |
| Admin.RibbonButton.btnAdminOpen | PASS | Label=Admin Console; OnAction=RibbonOnActionAdmin; Screentip= |
| Admin.MacroExists.btnAdminOpen | PASS | modAdmin.Admin_Click |
| Admin.CallbackMap.btnAdminOpen | PASS | btnAdminOpen -> modAdmin.Admin_Click |
| Admin.SafeExec.btnAdminOpen | PASS | modAdmin.Admin_Click |
| Admin.RibbonButton.btnAdminUsers | PASS | Label=Users and Roles; OnAction=RibbonOnActionAdmin; Screentip= |
| Admin.MacroExists.btnAdminUsers | PASS | modAdmin.Open_CreateDeleteUser |
| Admin.CallbackMap.btnAdminUsers | PASS | btnAdminUsers -> modAdmin.Open_CreateDeleteUser |
| Admin.SafeExec.btnAdminUsers | PASS | modAdmin.Open_CreateDeleteUser |
| Admin.RibbonButton.btnAdminCreateWarehouse | PASS | Label=Create New Warehouse; OnAction=RibbonOnActionAdmin; Screentip= |
| Admin.MacroExists.btnAdminCreateWarehouse | PASS | modAdmin.Open_CreateWarehouse |
| Admin.CallbackMap.btnAdminCreateWarehouse | PASS | btnAdminCreateWarehouse -> modAdmin.Open_CreateWarehouse |
| Admin.RibbonButton.btnAdminRetireMigrateWarehouse | PASS | Label=Retire / Migrate Warehouse; OnAction=RibbonOnActionAdmin; Screentip=Archive, migrate, retire, or delete a warehouse runtime |
| Admin.RibbonButtonScreentip.btnAdminRetireMigrateWarehouse | PASS | Archive, migrate, retire, or delete a warehouse runtime |
| Admin.MacroExists.btnAdminRetireMigrateWarehouse | PASS | modAdmin.Admin_RetireMigrateWarehouse_Click |
| Admin.CallbackMap.btnAdminRetireMigrateWarehouse | PASS | btnAdminRetireMigrateWarehouse -> modAdmin.Admin_RetireMigrateWarehouse_Click |
