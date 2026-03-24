# Phase 6 Packaged Ribbon Validation Results

- Date: 2026-03-23 20:23:15
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\hotfix-recall-print-20260323
- Runtime root override: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-ribbon-954a5d5ce35a4986939d67d7d7e0d29b
- Passed: 69
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| invSys.Core.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\hotfix-recall-print-20260323\invSys.Core.xlam |
| invSys.Inventory.Domain.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\hotfix-recall-print-20260323\invSys.Inventory.Domain.xlam |
| invSys.Designs.Domain.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\hotfix-recall-print-20260323\invSys.Designs.Domain.xlam |
| invSys.Receiving.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\hotfix-recall-print-20260323\invSys.Receiving.xlam |
| invSys.Shipping.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\hotfix-recall-print-20260323\invSys.Shipping.xlam |
| invSys.Production.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\hotfix-recall-print-20260323\invSys.Production.xlam |
| invSys.Admin.xlam.Open | PASS | Opened from C:\Users\Justin\repos\invSys_fork\deploy\hotfix-recall-print-20260323\invSys.Admin.xlam |
| Core.RuntimeRootOverride | PASS | C:\Users\Justin\AppData\Local\Temp\invsys-phase6-ribbon-954a5d5ce35a4986939d67d7d7e0d29b |
| Receiving.RibbonXml | PASS | customUI/customUI.xml present. |
| Receiving.CallbackModule | PASS | modRibbonGenerated |
| Receiving.RibbonButton.btnReceivingSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionReceiving |
| Receiving.MacroExists.btnReceivingSetup | PASS | modTS_Received.EnsureGeneratedButtons |
| Receiving.CallbackMap.btnReceivingSetup | PASS | btnReceivingSetup -> modTS_Received.EnsureGeneratedButtons |
| Receiving.SafeExec.btnReceivingSetup | PASS | modTS_Received.EnsureGeneratedButtons |
| Receiving.RibbonButton.btnReceivingConfirm | PASS | Label=Confirm Writes; OnAction=RibbonOnActionReceiving |
| Receiving.MacroExists.btnReceivingConfirm | PASS | modTS_Received.ConfirmWrites |
| Receiving.CallbackMap.btnReceivingConfirm | PASS | btnReceivingConfirm -> modTS_Received.ConfirmWrites |
| Receiving.RibbonButton.btnReceivingUndo | PASS | Label=Undo; OnAction=RibbonOnActionReceiving |
| Receiving.MacroExists.btnReceivingUndo | PASS | modTS_Received.MacroUndo |
| Receiving.CallbackMap.btnReceivingUndo | PASS | btnReceivingUndo -> modTS_Received.MacroUndo |
| Receiving.RibbonButton.btnReceivingRedo | PASS | Label=Redo; OnAction=RibbonOnActionReceiving |
| Receiving.MacroExists.btnReceivingRedo | PASS | modTS_Received.MacroRedo |
| Receiving.CallbackMap.btnReceivingRedo | PASS | btnReceivingRedo -> modTS_Received.MacroRedo |
| Shipping.RibbonXml | PASS | customUI/customUI.xml present. |
| Shipping.CallbackModule | PASS | modRibbonGenerated |
| Shipping.RibbonButton.btnShippingSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionShipping |
| Shipping.MacroExists.btnShippingSetup | PASS | modTS_Shipments.InitializeShipmentsUI |
| Shipping.CallbackMap.btnShippingSetup | PASS | btnShippingSetup -> modTS_Shipments.InitializeShipmentsUI |
| Shipping.SafeExec.btnShippingSetup | PASS | modTS_Shipments.InitializeShipmentsUI |
| Shipping.RibbonButton.btnShippingConfirm | PASS | Label=Confirm Inventory; OnAction=RibbonOnActionShipping |
| Shipping.MacroExists.btnShippingConfirm | PASS | modTS_Shipments.BtnConfirmInventory |
| Shipping.CallbackMap.btnShippingConfirm | PASS | btnShippingConfirm -> modTS_Shipments.BtnConfirmInventory |
| Shipping.RibbonButton.btnShippingStage | PASS | Label=To Shipments; OnAction=RibbonOnActionShipping |
| Shipping.MacroExists.btnShippingStage | PASS | modTS_Shipments.BtnToShipments |
| Shipping.CallbackMap.btnShippingStage | PASS | btnShippingStage -> modTS_Shipments.BtnToShipments |
| Shipping.RibbonButton.btnShippingSend | PASS | Label=Shipments Sent; OnAction=RibbonOnActionShipping |
| Shipping.MacroExists.btnShippingSend | PASS | modTS_Shipments.BtnShipmentsSent |
| Shipping.CallbackMap.btnShippingSend | PASS | btnShippingSend -> modTS_Shipments.BtnShipmentsSent |
| Production.RibbonXml | PASS | customUI/customUI.xml present. |
| Production.CallbackModule | PASS | modRibbonGenerated |
| Production.RibbonButton.btnProductionSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionProduction |
| Production.MacroExists.btnProductionSetup | PASS | mProduction.InitializeProductionUI |
| Production.CallbackMap.btnProductionSetup | PASS | btnProductionSetup -> mProduction.InitializeProductionUI |
| Production.SafeExec.btnProductionSetup | PASS | mProduction.InitializeProductionUI |
| Production.RibbonButton.btnProductionLoad | PASS | Label=Load Recipe; OnAction=RibbonOnActionProduction |
| Production.MacroExists.btnProductionLoad | PASS | mProduction.BtnLoadRecipe |
| Production.CallbackMap.btnProductionLoad | PASS | btnProductionLoad -> mProduction.BtnLoadRecipe |
| Production.RibbonButton.btnProductionUsed | PASS | Label=To Used; OnAction=RibbonOnActionProduction |
| Production.MacroExists.btnProductionUsed | PASS | mProduction.BtnToUsed |
| Production.CallbackMap.btnProductionUsed | PASS | btnProductionUsed -> mProduction.BtnToUsed |
| Production.RibbonButton.btnProductionMade | PASS | Label=To Made; OnAction=RibbonOnActionProduction |
| Production.MacroExists.btnProductionMade | PASS | mProduction.BtnToMade |
| Production.CallbackMap.btnProductionMade | PASS | btnProductionMade -> mProduction.BtnToMade |
| Production.RibbonButton.btnProductionTotal | PASS | Label=To Total Inv; OnAction=RibbonOnActionProduction |
| Production.MacroExists.btnProductionTotal | PASS | mProduction.BtnToTotalInv |
| Production.CallbackMap.btnProductionTotal | PASS | btnProductionTotal -> mProduction.BtnToTotalInv |
| Production.RibbonButton.btnProductionPrintCodes | PASS | Label=Print Recall Codes; OnAction=RibbonOnActionProduction |
| Production.MacroExists.btnProductionPrintCodes | PASS | mProduction.BtnPrintRecallCodes |
| Production.CallbackMap.btnProductionPrintCodes | PASS | btnProductionPrintCodes -> mProduction.BtnPrintRecallCodes |
| Admin.RibbonXml | PASS | customUI/customUI.xml present. |
| Admin.CallbackModule | PASS | modRibbonGenerated |
| Admin.RibbonButton.btnAdminOpen | PASS | Label=Admin Console; OnAction=RibbonOnActionAdmin |
| Admin.MacroExists.btnAdminOpen | PASS | modAdmin.Admin_Click |
| Admin.CallbackMap.btnAdminOpen | PASS | btnAdminOpen -> modAdmin.Admin_Click |
| Admin.SafeExec.btnAdminOpen | PASS | modAdmin.Admin_Click |
| Admin.RibbonButton.btnAdminUsers | PASS | Label=Users and Roles; OnAction=RibbonOnActionAdmin |
| Admin.MacroExists.btnAdminUsers | PASS | modAdmin.Open_CreateDeleteUser |
| Admin.CallbackMap.btnAdminUsers | PASS | btnAdminUsers -> modAdmin.Open_CreateDeleteUser |
| Admin.SafeExec.btnAdminUsers | PASS | modAdmin.Open_CreateDeleteUser |
