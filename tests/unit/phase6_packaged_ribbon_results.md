# Phase 6 Packaged Ribbon Validation Results

- Date: 2026-06-01 18:27:57
- Deploy root: C:\Users\justu\source\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\justu\AppData\Local\Temp\invsys-phase6-ribbon-1b472503c98f47c695b43dbaa0e348bd
- Passed: 136
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| invSys.Core.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Core.xlam |
| invSys.Inventory.Domain.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Inventory.Domain.xlam |
| invSys.Designs.Domain.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Designs.Domain.xlam |
| invSys.Receiving.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Receiving.xlam |
| invSys.Shipping.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Shipping.xlam |
| invSys.Production.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Production.xlam |
| invSys.Admin.xlam.Open | PASS | Opened from C:\Users\justu\source\repos\invSys_fork\deploy\current\invSys.Admin.xlam |
| Core.RuntimeRootOverride | PASS | C:\Users\justu\AppData\Local\Temp\invsys-phase6-ribbon-1b472503c98f47c695b43dbaa0e348bd |
| Receiving.RibbonXml | PASS | customUI/customUI.xml present. |
| Receiving.CallbackModule | PASS | modRibbonGenerated |
| Receiving.StatusLabel.lblReceivingServerStatus | PASS | GetLabel=RibbonServerStatusGetLabel |
| Receiving.RibbonButton.btnReceivingConnectServer | PASS | Label=Connect Server; OnAction=RibbonOnActionReceiving; GetEnabled=; Screentip=Connect to warehouse storage |
| Receiving.CallbackMap.btnReceivingConnectServer | PASS | btnReceivingConnectServer -> modRoleEventWriter.ConnectWarehouseStorageForCapability "RECEIVE_POST" |
| Receiving.RibbonButton.btnReceivingCurrentUser | PASS | Label=; OnAction=RibbonOnActionReceiving; GetEnabled=; Screentip=Sign in as an invSys user |
| Receiving.RibbonButtonScreentip.btnReceivingCurrentUser | PASS | Sign in as an invSys user |
| Receiving.CallbackMap.btnReceivingCurrentUser | PASS | btnReceivingCurrentUser -> modRoleEventWriter.PromptSetCurrentUserForCapability "RECEIVE_POST" |
| Receiving.RibbonButton.btnReceivingSignOut | PASS | Label=Sign Out; OnAction=RibbonOnActionReceiving; GetEnabled=; Screentip=Sign out of invSys without disconnecting storage |
| Receiving.CallbackMap.btnReceivingSignOut | PASS | btnReceivingSignOut -> modRoleEventWriter.SignOutCurrentUser |
| Receiving.RibbonButton.btnReceivingSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionReceiving; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Receiving.RibbonButtonGetEnabled.btnReceivingSetup | PASS | RibbonRequiredCapabilityGetEnabled |
| Receiving.MacroExists.btnReceivingSetup | PASS | modTS_Received.EnsureGeneratedButtons |
| Receiving.CallbackMap.btnReceivingSetup | PASS | btnReceivingSetup -> modTS_Received.EnsureGeneratedButtons |
| Receiving.CallbackGetEnabled.btnReceivingSetup | PASS | btnReceivingSetup -> RECEIVE_POST |
| Receiving.RibbonButton.btnReceivingConfirm | PASS | Label=Confirm Writes; OnAction=RibbonOnActionReceiving; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Receiving.RibbonButtonGetEnabled.btnReceivingConfirm | PASS | RibbonRequiredCapabilityGetEnabled |
| Receiving.MacroExists.btnReceivingConfirm | PASS | modTS_Received.ConfirmWrites |
| Receiving.CallbackMap.btnReceivingConfirm | PASS | btnReceivingConfirm -> modTS_Received.ConfirmWrites |
| Receiving.CallbackGetEnabled.btnReceivingConfirm | PASS | btnReceivingConfirm -> RECEIVE_POST |
| Receiving.RibbonButton.btnReceivingUndo | PASS | Label=Undo; OnAction=RibbonOnActionReceiving; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Receiving.RibbonButtonGetEnabled.btnReceivingUndo | PASS | RibbonRequiredCapabilityGetEnabled |
| Receiving.MacroExists.btnReceivingUndo | PASS | modTS_Received.MacroUndo |
| Receiving.CallbackMap.btnReceivingUndo | PASS | btnReceivingUndo -> modTS_Received.MacroUndo |
| Receiving.CallbackGetEnabled.btnReceivingUndo | PASS | btnReceivingUndo -> RECEIVE_POST |
| Receiving.RibbonButton.btnReceivingRedo | PASS | Label=Redo; OnAction=RibbonOnActionReceiving; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Receiving.RibbonButtonGetEnabled.btnReceivingRedo | PASS | RibbonRequiredCapabilityGetEnabled |
| Receiving.MacroExists.btnReceivingRedo | PASS | modTS_Received.MacroRedo |
| Receiving.CallbackMap.btnReceivingRedo | PASS | btnReceivingRedo -> modTS_Received.MacroRedo |
| Receiving.CallbackGetEnabled.btnReceivingRedo | PASS | btnReceivingRedo -> RECEIVE_POST |
| Shipping.RibbonXml | PASS | customUI/customUI.xml present. |
| Shipping.CallbackModule | PASS | modRibbonGenerated |
| Shipping.StatusLabel.lblShippingServerStatus | PASS | GetLabel=RibbonServerStatusGetLabel |
| Shipping.RibbonButton.btnShippingConnectServer | PASS | Label=Connect Server; OnAction=RibbonOnActionShipping; GetEnabled=; Screentip=Connect to warehouse storage |
| Shipping.CallbackMap.btnShippingConnectServer | PASS | btnShippingConnectServer -> modRoleEventWriter.ConnectWarehouseStorageForCapability "SHIP_POST" |
| Shipping.RibbonButton.btnShippingCurrentUser | PASS | Label=; OnAction=RibbonOnActionShipping; GetEnabled=; Screentip=Sign in as an invSys user |
| Shipping.RibbonButtonScreentip.btnShippingCurrentUser | PASS | Sign in as an invSys user |
| Shipping.CallbackMap.btnShippingCurrentUser | PASS | btnShippingCurrentUser -> modRoleEventWriter.PromptSetCurrentUserForCapability "SHIP_POST" |
| Shipping.RibbonButton.btnShippingSignOut | PASS | Label=Sign Out; OnAction=RibbonOnActionShipping; GetEnabled=; Screentip=Sign out of invSys without disconnecting storage |
| Shipping.CallbackMap.btnShippingSignOut | PASS | btnShippingSignOut -> modRoleEventWriter.SignOutCurrentUser |
| Shipping.RibbonButton.btnShippingSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingSetup | PASS | RibbonRequiredCapabilityGetEnabled |
| Shipping.MacroExists.btnShippingSetup | PASS | modTS_Shipments.InitializeShipmentsUI |
| Shipping.CallbackMap.btnShippingSetup | PASS | btnShippingSetup -> modTS_Shipments.InitializeShipmentsUI |
| Shipping.CallbackGetEnabled.btnShippingSetup | PASS | btnShippingSetup -> SHIP_POST |
| Shipping.RibbonButton.btnShippingConfirm | PASS | Label=Confirm Inventory; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingConfirm | PASS | RibbonRequiredCapabilityGetEnabled |
| Shipping.MacroExists.btnShippingConfirm | PASS | modTS_Shipments.BtnConfirmInventory |
| Shipping.CallbackMap.btnShippingConfirm | PASS | btnShippingConfirm -> modTS_Shipments.BtnConfirmInventory |
| Shipping.CallbackGetEnabled.btnShippingConfirm | PASS | btnShippingConfirm -> SHIP_POST |
| Shipping.RibbonButton.btnShippingStage | PASS | Label=To Shipments; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingStage | PASS | RibbonRequiredCapabilityGetEnabled |
| Shipping.MacroExists.btnShippingStage | PASS | modTS_Shipments.BtnToShipments |
| Shipping.CallbackMap.btnShippingStage | PASS | btnShippingStage -> modTS_Shipments.BtnToShipments |
| Shipping.CallbackGetEnabled.btnShippingStage | PASS | btnShippingStage -> SHIP_POST |
| Shipping.RibbonButton.btnShippingSend | PASS | Label=Shipments Sent; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingSend | PASS | RibbonRequiredCapabilityGetEnabled |
| Shipping.MacroExists.btnShippingSend | PASS | modTS_Shipments.BtnShipmentsSent |
| Shipping.CallbackMap.btnShippingSend | PASS | btnShippingSend -> modTS_Shipments.BtnShipmentsSent |
| Shipping.CallbackGetEnabled.btnShippingSend | PASS | btnShippingSend -> SHIP_POST |
| Production.RibbonXml | PASS | customUI/customUI.xml present. |
| Production.CallbackModule | PASS | modRibbonGenerated |
| Production.StatusLabel.lblProductionServerStatus | PASS | GetLabel=RibbonServerStatusGetLabel |
| Production.RibbonButton.btnProductionConnectServer | PASS | Label=Connect Server; OnAction=RibbonOnActionProduction; GetEnabled=; Screentip=Connect to warehouse storage |
| Production.CallbackMap.btnProductionConnectServer | PASS | btnProductionConnectServer -> modRoleEventWriter.ConnectWarehouseStorageForCapability "PROD_POST" |
| Production.RibbonButton.btnProductionCurrentUser | PASS | Label=; OnAction=RibbonOnActionProduction; GetEnabled=; Screentip=Sign in as an invSys user |
| Production.RibbonButtonScreentip.btnProductionCurrentUser | PASS | Sign in as an invSys user |
| Production.CallbackMap.btnProductionCurrentUser | PASS | btnProductionCurrentUser -> modRoleEventWriter.PromptSetCurrentUserForCapability "PROD_POST" |
| Production.RibbonButton.btnProductionSignOut | PASS | Label=Sign Out; OnAction=RibbonOnActionProduction; GetEnabled=; Screentip=Sign out of invSys without disconnecting storage |
| Production.CallbackMap.btnProductionSignOut | PASS | btnProductionSignOut -> modRoleEventWriter.SignOutCurrentUser |
| Production.RibbonButton.btnProductionSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionSetup | PASS | RibbonRequiredCapabilityGetEnabled |
| Production.MacroExists.btnProductionSetup | PASS | mProduction.InitializeProductionUI |
| Production.CallbackMap.btnProductionSetup | PASS | btnProductionSetup -> mProduction.InitializeProductionUI |
| Production.CallbackGetEnabled.btnProductionSetup | PASS | btnProductionSetup -> PROD_POST |
| Production.RibbonButton.btnProductionLoad | PASS | Label=Load Recipe; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionLoad | PASS | RibbonRequiredCapabilityGetEnabled |
| Production.MacroExists.btnProductionLoad | PASS | mProduction.BtnLoadRecipe |
| Production.CallbackMap.btnProductionLoad | PASS | btnProductionLoad -> mProduction.BtnLoadRecipe |
| Production.CallbackGetEnabled.btnProductionLoad | PASS | btnProductionLoad -> PROD_POST |
| Production.RibbonButton.btnProductionUsed | PASS | Label=To Used; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionUsed | PASS | RibbonRequiredCapabilityGetEnabled |
| Production.MacroExists.btnProductionUsed | PASS | mProduction.BtnToUsed |
| Production.CallbackMap.btnProductionUsed | PASS | btnProductionUsed -> mProduction.BtnToUsed |
| Production.CallbackGetEnabled.btnProductionUsed | PASS | btnProductionUsed -> PROD_POST |
| Production.RibbonButton.btnProductionMade | PASS | Label=To Made; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionMade | PASS | RibbonRequiredCapabilityGetEnabled |
| Production.MacroExists.btnProductionMade | PASS | mProduction.BtnToMade |
| Production.CallbackMap.btnProductionMade | PASS | btnProductionMade -> mProduction.BtnToMade |
| Production.CallbackGetEnabled.btnProductionMade | PASS | btnProductionMade -> PROD_POST |
| Production.RibbonButton.btnProductionTotal | PASS | Label=To Total Inv; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionTotal | PASS | RibbonRequiredCapabilityGetEnabled |
| Production.MacroExists.btnProductionTotal | PASS | mProduction.BtnToTotalInv |
| Production.CallbackMap.btnProductionTotal | PASS | btnProductionTotal -> mProduction.BtnToTotalInv |
| Production.CallbackGetEnabled.btnProductionTotal | PASS | btnProductionTotal -> PROD_POST |
| Production.RibbonButton.btnProductionPrintCodes | PASS | Label=Print Recall Codes; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabled; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionPrintCodes | PASS | RibbonRequiredCapabilityGetEnabled |
| Production.MacroExists.btnProductionPrintCodes | PASS | mProduction.BtnPrintRecallCodes |
| Production.CallbackMap.btnProductionPrintCodes | PASS | btnProductionPrintCodes -> mProduction.BtnPrintRecallCodes |
| Production.CallbackGetEnabled.btnProductionPrintCodes | PASS | btnProductionPrintCodes -> PROD_POST |
| Admin.RibbonXml | PASS | customUI/customUI.xml present. |
| Admin.CallbackModule | PASS | modRibbonGenerated |
| Admin.RibbonButton.btnAdminOpen | PASS | Label=Admin Console; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip= |
| Admin.MacroExists.btnAdminOpen | PASS | modAdmin.Admin_Click |
| Admin.CallbackMap.btnAdminOpen | PASS | btnAdminOpen -> modAdmin.Admin_Click |
| Admin.RibbonButton.btnAdminConnectServer | PASS | Label=Connect Server; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip=Connect to warehouse storage |
| Admin.CallbackMap.btnAdminConnectServer | PASS | btnAdminConnectServer -> modRoleEventWriter.ConnectWarehouseStorageForCapability "ADMIN_MAINT" |
| Admin.RibbonButton.btnAdminCurrentUser | PASS | Label=; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip=Sign in as an invSys user |
| Admin.RibbonButtonScreentip.btnAdminCurrentUser | PASS | Sign in as an invSys user |
| Admin.CallbackMap.btnAdminCurrentUser | PASS | btnAdminCurrentUser -> modRoleEventWriter.PromptSetCurrentUserForCapability "ADMIN_MAINT" |
| Admin.RibbonButton.btnAdminSignOut | PASS | Label=Sign Out; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip=Sign out of invSys without disconnecting storage |
| Admin.CallbackMap.btnAdminSignOut | PASS | btnAdminSignOut -> modRoleEventWriter.SignOutCurrentUser |
| Admin.RibbonButton.btnAdminUsers | PASS | Label=Users and Roles; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip= |
| Admin.MacroExists.btnAdminUsers | PASS | modAdmin.Open_CreateDeleteUser |
| Admin.CallbackMap.btnAdminUsers | PASS | btnAdminUsers -> modAdmin.Open_CreateDeleteUser |
| Admin.RibbonButton.btnAdminCreateWarehouse | PASS | Label=Create New Warehouse; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip= |
| Admin.MacroExists.btnAdminCreateWarehouse | PASS | modAdmin.Open_CreateWarehouse |
| Admin.CallbackMap.btnAdminCreateWarehouse | PASS | btnAdminCreateWarehouse -> modAdmin.Open_CreateWarehouse |
| Admin.RibbonButton.btnAdminSetupTesterStation | PASS | Label=Setup Tester Station; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip= |
| Admin.MacroExists.btnAdminSetupTesterStation | PASS | modAdmin.Admin_SetupTesterStation_Click |
| Admin.CallbackMap.btnAdminSetupTesterStation | PASS | btnAdminSetupTesterStation -> modAdmin.Admin_SetupTesterStation_Click |
| Admin.RibbonButton.btnAdminVerifyAddinsPublished | PASS | Label=Verify Add-ins Published; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip= |
| Admin.MacroExists.btnAdminVerifyAddinsPublished | PASS | modAdmin.Verify_AddinsPublished |
| Admin.CallbackMap.btnAdminVerifyAddinsPublished | PASS | btnAdminVerifyAddinsPublished -> modAdmin.Verify_AddinsPublished |
| Admin.RibbonButton.btnAdminRetireMigrateWarehouse | PASS | Label=Retire / Migrate Warehouse; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip=Archive, migrate, retire, or delete a warehouse runtime |
| Admin.RibbonButtonScreentip.btnAdminRetireMigrateWarehouse | PASS | Archive, migrate, retire, or delete a warehouse runtime |
| Admin.MacroExists.btnAdminRetireMigrateWarehouse | PASS | modAdmin.Admin_RetireMigrateWarehouse_Click |
| Admin.CallbackMap.btnAdminRetireMigrateWarehouse | PASS | btnAdminRetireMigrateWarehouse -> modAdmin.Admin_RetireMigrateWarehouse_Click |
