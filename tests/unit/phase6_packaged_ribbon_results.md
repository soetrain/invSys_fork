# Phase 6 Packaged Ribbon Validation Results

- Date: 2026-06-12 12:46:22
- Deploy root: C:\Users\justu\source\repos\invSys_fork\deploy\current
- Runtime root override: C:\Users\justu\AppData\Local\Temp\invsys-phase6-ribbon-2a06e90af02748c298492ab97e36a78e
- Passed: 227
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
| Core.RuntimeRootOverride | PASS | C:\Users\justu\AppData\Local\Temp\invsys-phase6-ribbon-2a06e90af02748c298492ab97e36a78e |
| Receiving.RibbonXml | PASS | customUI/customUI.xml present. |
| Receiving.CallbackModule | PASS | modRibbonGenerated |
| Receiving.StatusLabel.lblReceivingServerStatus | PASS | GetLabel=RibbonServerStatusGetLabel |
| Receiving.StatusLabel.lblReceivingAccessStatus | PASS | GetLabel=RibbonAccessStatusGetLabel |
| Receiving.RibbonButton.btnReceivingConnectServer | PASS | Label=Connect Server; OnAction=RibbonOnActionReceiving; GetEnabled=; Screentip=Connect to warehouse storage |
| Receiving.CallbackMap.btnReceivingConnectServer | PASS | btnReceivingConnectServer -> modRoleEventWriter.ConnectWarehouseStorageForCapability "RECEIVE_POST" |
| Receiving.RibbonButton.btnReceivingCurrentUser | PASS | Label=; OnAction=RibbonOnActionReceiving; GetEnabled=; Screentip=Sign in as an invSys user |
| Receiving.RibbonButtonScreentip.btnReceivingCurrentUser | PASS | Sign in as an invSys user |
| Receiving.CallbackMap.btnReceivingCurrentUser | PASS | btnReceivingCurrentUser -> modRoleEventWriter.PromptSetCurrentUserForCapability "RECEIVE_POST" |
| Receiving.RibbonButton.btnReceivingSignOut | PASS | Label=Sign Out; OnAction=RibbonOnActionReceiving; GetEnabled=; Screentip=Sign out of invSys without disconnecting storage |
| Receiving.CallbackMap.btnReceivingSignOut | PASS | btnReceivingSignOut -> modRoleEventWriter.SignOutCurrentUser |
| Receiving.RibbonButton.btnReceivingSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionReceiving; GetEnabled=RibbonRequiredCapabilityGetEnabledReceiving; Screentip= |
| Receiving.RibbonButtonGetEnabled.btnReceivingSetup | PASS | RibbonRequiredCapabilityGetEnabledReceiving |
| Receiving.MacroExists.btnReceivingSetup | PASS | modTS_Received.EnsureGeneratedButtons |
| Receiving.CallbackMap.btnReceivingSetup | PASS | btnReceivingSetup -> modTS_Received.EnsureGeneratedButtons |
| Receiving.CallbackGetEnabled.btnReceivingSetup | PASS | btnReceivingSetup -> RECEIVE_POST |
| Receiving.DisabledOffline.btnReceivingSetup | PASS | btnReceivingSetup enabled=False |
| Receiving.RibbonButton.btnReceivingConfirm | PASS | Label=Confirm Writes; OnAction=RibbonOnActionReceiving; GetEnabled=RibbonRequiredCapabilityGetEnabledReceiving; Screentip= |
| Receiving.RibbonButtonGetEnabled.btnReceivingConfirm | PASS | RibbonRequiredCapabilityGetEnabledReceiving |
| Receiving.MacroExists.btnReceivingConfirm | PASS | modTS_Received.ConfirmWrites |
| Receiving.CallbackMap.btnReceivingConfirm | PASS | btnReceivingConfirm -> modTS_Received.ConfirmWrites |
| Receiving.CallbackGetEnabled.btnReceivingConfirm | PASS | btnReceivingConfirm -> RECEIVE_POST |
| Receiving.DisabledOffline.btnReceivingConfirm | PASS | btnReceivingConfirm enabled=False |
| Receiving.RibbonButton.btnReceivingUndo | PASS | Label=Undo; OnAction=RibbonOnActionReceiving; GetEnabled=RibbonRequiredCapabilityGetEnabledReceiving; Screentip= |
| Receiving.RibbonButtonGetEnabled.btnReceivingUndo | PASS | RibbonRequiredCapabilityGetEnabledReceiving |
| Receiving.MacroExists.btnReceivingUndo | PASS | modTS_Received.MacroUndo |
| Receiving.CallbackMap.btnReceivingUndo | PASS | btnReceivingUndo -> modTS_Received.MacroUndo |
| Receiving.CallbackGetEnabled.btnReceivingUndo | PASS | btnReceivingUndo -> RECEIVE_POST |
| Receiving.DisabledOffline.btnReceivingUndo | PASS | btnReceivingUndo enabled=False |
| Receiving.RibbonButton.btnReceivingRedo | PASS | Label=Redo; OnAction=RibbonOnActionReceiving; GetEnabled=RibbonRequiredCapabilityGetEnabledReceiving; Screentip= |
| Receiving.RibbonButtonGetEnabled.btnReceivingRedo | PASS | RibbonRequiredCapabilityGetEnabledReceiving |
| Receiving.MacroExists.btnReceivingRedo | PASS | modTS_Received.MacroRedo |
| Receiving.CallbackMap.btnReceivingRedo | PASS | btnReceivingRedo -> modTS_Received.MacroRedo |
| Receiving.CallbackGetEnabled.btnReceivingRedo | PASS | btnReceivingRedo -> RECEIVE_POST |
| Receiving.DisabledOffline.btnReceivingRedo | PASS | btnReceivingRedo enabled=False |
| Shipping.RibbonXml | PASS | customUI/customUI.xml present. |
| Shipping.CallbackModule | PASS | modRibbonGenerated |
| Shipping.StatusLabel.lblShippingServerStatus | PASS | GetLabel=RibbonServerStatusGetLabel |
| Shipping.StatusLabel.lblShippingAccessStatus | PASS | GetLabel=RibbonAccessStatusGetLabel |
| Shipping.RibbonButton.btnShippingConnectServer | PASS | Label=Connect Server; OnAction=RibbonOnActionShipping; GetEnabled=; Screentip=Connect to warehouse storage |
| Shipping.CallbackMap.btnShippingConnectServer | PASS | btnShippingConnectServer -> modRoleEventWriter.ConnectWarehouseStorageForCapability "SHIP_POST" |
| Shipping.RibbonButton.btnShippingCurrentUser | PASS | Label=; OnAction=RibbonOnActionShipping; GetEnabled=; Screentip=Sign in as an invSys user |
| Shipping.RibbonButtonScreentip.btnShippingCurrentUser | PASS | Sign in as an invSys user |
| Shipping.CallbackMap.btnShippingCurrentUser | PASS | btnShippingCurrentUser -> modRoleEventWriter.PromptSetCurrentUserForCapability "SHIP_POST" |
| Shipping.RibbonButton.btnShippingSignOut | PASS | Label=Sign Out; OnAction=RibbonOnActionShipping; GetEnabled=; Screentip=Sign out of invSys without disconnecting storage |
| Shipping.CallbackMap.btnShippingSignOut | PASS | btnShippingSignOut -> modRoleEventWriter.SignOutCurrentUser |
| Shipping.RibbonButton.btnShippingSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabledShipping; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingSetup | PASS | RibbonRequiredCapabilityGetEnabledShipping |
| Shipping.MacroExists.btnShippingSetup | PASS | modTS_Shipments.InitializeShipmentsUI |
| Shipping.CallbackMap.btnShippingSetup | PASS | btnShippingSetup -> modTS_Shipments.InitializeShipmentsUI |
| Shipping.CallbackGetEnabled.btnShippingSetup | PASS | btnShippingSetup -> SHIP_POST |
| Shipping.DisabledOffline.btnShippingSetup | PASS | btnShippingSetup enabled=False |
| Shipping.RibbonButton.btnShippingBoxMode | PASS | Label=; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabledShipping; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingBoxMode | PASS | RibbonRequiredCapabilityGetEnabledShipping |
| Shipping.MacroExists.btnShippingBoxMode | PASS | modTS_Shipments.BtnSwitchToBoxMaker |
| Shipping.CallbackMap.btnShippingBoxMode | PASS | btnShippingBoxMode -> modTS_Shipments.BtnSwitchToBoxMaker |
| Shipping.CallbackGetEnabled.btnShippingBoxMode | PASS | btnShippingBoxMode -> SHIP_POST |
| Shipping.DisabledOffline.btnShippingBoxMode | PASS | btnShippingBoxMode enabled=False |
| Shipping.RibbonButton.btnShippingSaveBox | PASS | Label=Save Box; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabledShipping; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingSaveBox | PASS | RibbonRequiredCapabilityGetEnabledShipping |
| Shipping.MacroExists.btnShippingSaveBox | PASS | modTS_Shipments.BtnSaveBox |
| Shipping.CallbackMap.btnShippingSaveBox | PASS | btnShippingSaveBox -> modTS_Shipments.BtnSaveBox |
| Shipping.CallbackGetEnabled.btnShippingSaveBox | PASS | btnShippingSaveBox -> SHIP_POST |
| Shipping.DisabledOffline.btnShippingSaveBox | PASS | btnShippingSaveBox enabled=False |
| Shipping.RibbonButton.btnShippingDeleteBoxVersion | PASS | Label=Delete Version; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabledShipping; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingDeleteBoxVersion | PASS | RibbonRequiredCapabilityGetEnabledShipping |
| Shipping.MacroExists.btnShippingDeleteBoxVersion | PASS | modTS_Shipments.BtnDeleteBoxVersion |
| Shipping.CallbackMap.btnShippingDeleteBoxVersion | PASS | btnShippingDeleteBoxVersion -> modTS_Shipments.BtnDeleteBoxVersion |
| Shipping.CallbackGetEnabled.btnShippingDeleteBoxVersion | PASS | btnShippingDeleteBoxVersion -> ADMIN_MAINT |
| Shipping.DisabledOffline.btnShippingDeleteBoxVersion | PASS | btnShippingDeleteBoxVersion enabled=False |
| Shipping.RibbonButton.btnShippingDeleteBox | PASS | Label=Delete Box; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabledShipping; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingDeleteBox | PASS | RibbonRequiredCapabilityGetEnabledShipping |
| Shipping.MacroExists.btnShippingDeleteBox | PASS | modTS_Shipments.BtnDeleteBox |
| Shipping.CallbackMap.btnShippingDeleteBox | PASS | btnShippingDeleteBox -> modTS_Shipments.BtnDeleteBox |
| Shipping.CallbackGetEnabled.btnShippingDeleteBox | PASS | btnShippingDeleteBox -> ADMIN_MAINT |
| Shipping.DisabledOffline.btnShippingDeleteBox | PASS | btnShippingDeleteBox enabled=False |
| Shipping.RibbonButton.btnShippingConfirm | PASS | Label=Confirm Inventory; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabledShipping; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingConfirm | PASS | RibbonRequiredCapabilityGetEnabledShipping |
| Shipping.MacroExists.btnShippingConfirm | PASS | modTS_Shipments.BtnConfirmInventory |
| Shipping.CallbackMap.btnShippingConfirm | PASS | btnShippingConfirm -> modTS_Shipments.BtnConfirmInventory |
| Shipping.CallbackGetEnabled.btnShippingConfirm | PASS | btnShippingConfirm -> SHIP_POST |
| Shipping.DisabledOffline.btnShippingConfirm | PASS | btnShippingConfirm enabled=False |
| Shipping.RibbonButton.btnShippingBoxCreated | PASS | Label=Box Created; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabledShipping; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingBoxCreated | PASS | RibbonRequiredCapabilityGetEnabledShipping |
| Shipping.MacroExists.btnShippingBoxCreated | PASS | modTS_Shipments.BtnBoxCreated |
| Shipping.CallbackMap.btnShippingBoxCreated | PASS | btnShippingBoxCreated -> modTS_Shipments.BtnBoxCreated |
| Shipping.CallbackGetEnabled.btnShippingBoxCreated | PASS | btnShippingBoxCreated -> SHIP_POST |
| Shipping.DisabledOffline.btnShippingBoxCreated | PASS | btnShippingBoxCreated enabled=False |
| Shipping.RibbonButton.btnShippingBoxUnboxed | PASS | Label=Box Unboxed; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabledShipping; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingBoxUnboxed | PASS | RibbonRequiredCapabilityGetEnabledShipping |
| Shipping.MacroExists.btnShippingBoxUnboxed | PASS | modTS_Shipments.BtnBoxUnboxed |
| Shipping.CallbackMap.btnShippingBoxUnboxed | PASS | btnShippingBoxUnboxed -> modTS_Shipments.BtnBoxUnboxed |
| Shipping.CallbackGetEnabled.btnShippingBoxUnboxed | PASS | btnShippingBoxUnboxed -> SHIP_POST |
| Shipping.DisabledOffline.btnShippingBoxUnboxed | PASS | btnShippingBoxUnboxed enabled=False |
| Shipping.RibbonButton.btnShippingStage | PASS | Label=To Shipments; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabledShipping; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingStage | PASS | RibbonRequiredCapabilityGetEnabledShipping |
| Shipping.MacroExists.btnShippingStage | PASS | modTS_Shipments.BtnToShipments |
| Shipping.CallbackMap.btnShippingStage | PASS | btnShippingStage -> modTS_Shipments.BtnToShipments |
| Shipping.CallbackGetEnabled.btnShippingStage | PASS | btnShippingStage -> SHIP_POST |
| Shipping.DisabledOffline.btnShippingStage | PASS | btnShippingStage enabled=False |
| Shipping.RibbonButton.btnShippingSend | PASS | Label=Shipments Sent; OnAction=RibbonOnActionShipping; GetEnabled=RibbonRequiredCapabilityGetEnabledShipping; Screentip= |
| Shipping.RibbonButtonGetEnabled.btnShippingSend | PASS | RibbonRequiredCapabilityGetEnabledShipping |
| Shipping.MacroExists.btnShippingSend | PASS | modTS_Shipments.BtnShipmentsSent |
| Shipping.CallbackMap.btnShippingSend | PASS | btnShippingSend -> modTS_Shipments.BtnShipmentsSent |
| Shipping.CallbackGetEnabled.btnShippingSend | PASS | btnShippingSend -> SHIP_POST |
| Shipping.DisabledOffline.btnShippingSend | PASS | btnShippingSend enabled=False |
| Production.RibbonXml | PASS | customUI/customUI.xml present. |
| Production.CallbackModule | PASS | modRibbonGenerated |
| Production.StatusLabel.lblProductionServerStatus | PASS | GetLabel=RibbonServerStatusGetLabel |
| Production.StatusLabel.lblProductionAccessStatus | PASS | GetLabel=RibbonAccessStatusGetLabel |
| Production.RibbonButton.btnProductionConnectServer | PASS | Label=Connect Server; OnAction=RibbonOnActionProduction; GetEnabled=; Screentip=Connect to warehouse storage |
| Production.CallbackMap.btnProductionConnectServer | PASS | btnProductionConnectServer -> modRoleEventWriter.ConnectWarehouseStorageForCapability "PROD_POST" |
| Production.RibbonButton.btnProductionCurrentUser | PASS | Label=; OnAction=RibbonOnActionProduction; GetEnabled=; Screentip=Sign in as an invSys user |
| Production.RibbonButtonScreentip.btnProductionCurrentUser | PASS | Sign in as an invSys user |
| Production.CallbackMap.btnProductionCurrentUser | PASS | btnProductionCurrentUser -> modRoleEventWriter.PromptSetCurrentUserForCapability "PROD_POST" |
| Production.RibbonButton.btnProductionSignOut | PASS | Label=Sign Out; OnAction=RibbonOnActionProduction; GetEnabled=; Screentip=Sign out of invSys without disconnecting storage |
| Production.CallbackMap.btnProductionSignOut | PASS | btnProductionSignOut -> modRoleEventWriter.SignOutCurrentUser |
| Production.RibbonButton.btnProductionSetup | PASS | Label=Setup UI; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabledProduction; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionSetup | PASS | RibbonRequiredCapabilityGetEnabledProduction |
| Production.MacroExists.btnProductionSetup | PASS | mProduction.InitializeProductionUI |
| Production.CallbackMap.btnProductionSetup | PASS | btnProductionSetup -> mProduction.InitializeProductionUI |
| Production.CallbackGetEnabled.btnProductionSetup | PASS | btnProductionSetup -> PROD_POST |
| Production.DisabledOffline.btnProductionSetup | PASS | btnProductionSetup enabled=False |
| Production.RibbonButton.btnProductionLoad | PASS | Label=Load Recipe; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabledProduction; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionLoad | PASS | RibbonRequiredCapabilityGetEnabledProduction |
| Production.MacroExists.btnProductionLoad | PASS | mProduction.BtnLoadRecipe |
| Production.CallbackMap.btnProductionLoad | PASS | btnProductionLoad -> mProduction.BtnLoadRecipe |
| Production.CallbackGetEnabled.btnProductionLoad | PASS | btnProductionLoad -> PROD_POST |
| Production.DisabledOffline.btnProductionLoad | PASS | btnProductionLoad enabled=False |
| Production.RibbonButton.btnProductionUsed | PASS | Label=To Used; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabledProduction; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionUsed | PASS | RibbonRequiredCapabilityGetEnabledProduction |
| Production.MacroExists.btnProductionUsed | PASS | mProduction.BtnToUsed |
| Production.CallbackMap.btnProductionUsed | PASS | btnProductionUsed -> mProduction.BtnToUsed |
| Production.CallbackGetEnabled.btnProductionUsed | PASS | btnProductionUsed -> PROD_POST |
| Production.DisabledOffline.btnProductionUsed | PASS | btnProductionUsed enabled=False |
| Production.RibbonButton.btnProductionMade | PASS | Label=To Made; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabledProduction; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionMade | PASS | RibbonRequiredCapabilityGetEnabledProduction |
| Production.MacroExists.btnProductionMade | PASS | mProduction.BtnToMade |
| Production.CallbackMap.btnProductionMade | PASS | btnProductionMade -> mProduction.BtnToMade |
| Production.CallbackGetEnabled.btnProductionMade | PASS | btnProductionMade -> PROD_POST |
| Production.DisabledOffline.btnProductionMade | PASS | btnProductionMade enabled=False |
| Production.RibbonButton.btnProductionTotal | PASS | Label=To Total Inv; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabledProduction; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionTotal | PASS | RibbonRequiredCapabilityGetEnabledProduction |
| Production.MacroExists.btnProductionTotal | PASS | mProduction.BtnToTotalInv |
| Production.CallbackMap.btnProductionTotal | PASS | btnProductionTotal -> mProduction.BtnToTotalInv |
| Production.CallbackGetEnabled.btnProductionTotal | PASS | btnProductionTotal -> PROD_POST |
| Production.DisabledOffline.btnProductionTotal | PASS | btnProductionTotal enabled=False |
| Production.RibbonButton.btnProductionPrintCodes | PASS | Label=Print Recall Codes; OnAction=RibbonOnActionProduction; GetEnabled=RibbonRequiredCapabilityGetEnabledProduction; Screentip= |
| Production.RibbonButtonGetEnabled.btnProductionPrintCodes | PASS | RibbonRequiredCapabilityGetEnabledProduction |
| Production.MacroExists.btnProductionPrintCodes | PASS | mProduction.BtnPrintRecallCodes |
| Production.CallbackMap.btnProductionPrintCodes | PASS | btnProductionPrintCodes -> mProduction.BtnPrintRecallCodes |
| Production.CallbackGetEnabled.btnProductionPrintCodes | PASS | btnProductionPrintCodes -> PROD_POST |
| Production.DisabledOffline.btnProductionPrintCodes | PASS | btnProductionPrintCodes enabled=False |
| Admin.RibbonXml | PASS | customUI/customUI.xml present. |
| Admin.CallbackModule | PASS | modRibbonGenerated |
| Admin.StatusLabel.lblAdminServerStatus | PASS | GetLabel=RibbonServerStatusGetLabel |
| Admin.StatusLabel.lblAdminAccessStatus | PASS | GetLabel=RibbonAccessStatusGetLabel |
| Admin.RibbonButton.btnAdminOpen | PASS | Label=Admin Console; OnAction=RibbonOnActionAdmin; GetEnabled=RibbonRequiredCapabilityGetEnabledAdmin; Screentip= |
| Admin.RibbonButtonGetEnabled.btnAdminOpen | PASS | RibbonRequiredCapabilityGetEnabledAdmin |
| Admin.MacroExists.btnAdminOpen | PASS | modAdmin.Admin_Click |
| Admin.CallbackMap.btnAdminOpen | PASS | btnAdminOpen -> modAdmin.Admin_Click |
| Admin.CallbackGetEnabled.btnAdminOpen | PASS | btnAdminOpen -> ADMIN_MAINT |
| Admin.DisabledOffline.btnAdminOpen | PASS | btnAdminOpen enabled=False |
| Admin.RibbonButton.btnAdminConnectServer | PASS | Label=Connect Server; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip=Connect to warehouse storage |
| Admin.CallbackMap.btnAdminConnectServer | PASS | btnAdminConnectServer -> modRoleEventWriter.ConnectWarehouseStorageForCapability "ADMIN_MAINT" |
| Admin.RibbonButton.btnAdminCurrentUser | PASS | Label=; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip=Sign in as an invSys user |
| Admin.RibbonButtonScreentip.btnAdminCurrentUser | PASS | Sign in as an invSys user |
| Admin.CallbackMap.btnAdminCurrentUser | PASS | btnAdminCurrentUser -> modRoleEventWriter.PromptSetCurrentUserForCapability "ADMIN_MAINT" |
| Admin.RibbonButton.btnAdminSignOut | PASS | Label=Sign Out; OnAction=RibbonOnActionAdmin; GetEnabled=; Screentip=Sign out of invSys without disconnecting storage |
| Admin.CallbackMap.btnAdminSignOut | PASS | btnAdminSignOut -> modRoleEventWriter.SignOutCurrentUser |
| Admin.RibbonButton.btnAdminUsers | PASS | Label=Users and Roles; OnAction=RibbonOnActionAdmin; GetEnabled=RibbonRequiredCapabilityGetEnabledAdmin; Screentip= |
| Admin.RibbonButtonGetEnabled.btnAdminUsers | PASS | RibbonRequiredCapabilityGetEnabledAdmin |
| Admin.MacroExists.btnAdminUsers | PASS | modAdmin.Open_CreateDeleteUser |
| Admin.CallbackMap.btnAdminUsers | PASS | btnAdminUsers -> modAdmin.Open_CreateDeleteUser |
| Admin.CallbackGetEnabled.btnAdminUsers | PASS | btnAdminUsers -> ADMIN_MAINT |
| Admin.DisabledOffline.btnAdminUsers | PASS | btnAdminUsers enabled=False |
| Admin.RibbonButton.btnAdminWarehouses | PASS | Label=View Warehouses; OnAction=RibbonOnActionAdmin; GetEnabled=RibbonRequiredCapabilityGetEnabledAdmin; Screentip= |
| Admin.RibbonButtonGetEnabled.btnAdminWarehouses | PASS | RibbonRequiredCapabilityGetEnabledAdmin |
| Admin.MacroExists.btnAdminWarehouses | PASS | modAdmin.Open_WarehouseDirectory |
| Admin.CallbackMap.btnAdminWarehouses | PASS | btnAdminWarehouses -> modAdmin.Open_WarehouseDirectory |
| Admin.CallbackGetEnabled.btnAdminWarehouses | PASS | btnAdminWarehouses -> ADMIN_MAINT |
| Admin.DisabledOffline.btnAdminWarehouses | PASS | btnAdminWarehouses enabled=False |
| Admin.RibbonButton.btnAdminWarehouseRoot | PASS | Label=Add Warehouse Root; OnAction=RibbonOnActionAdmin; GetEnabled=RibbonRequiredCapabilityGetEnabledAdmin; Screentip= |
| Admin.RibbonButtonGetEnabled.btnAdminWarehouseRoot | PASS | RibbonRequiredCapabilityGetEnabledAdmin |
| Admin.MacroExists.btnAdminWarehouseRoot | PASS | modAdmin.Add_WarehouseDirectoryRoot |
| Admin.CallbackMap.btnAdminWarehouseRoot | PASS | btnAdminWarehouseRoot -> modAdmin.Add_WarehouseDirectoryRoot |
| Admin.CallbackGetEnabled.btnAdminWarehouseRoot | PASS | btnAdminWarehouseRoot -> ADMIN_MAINT |
| Admin.DisabledOffline.btnAdminWarehouseRoot | PASS | btnAdminWarehouseRoot enabled=False |
| Admin.RibbonButton.btnAdminCreateWarehouse | PASS | Label=Create New Warehouse; OnAction=RibbonOnActionAdmin; GetEnabled=RibbonRequiredCapabilityGetEnabledAdmin; Screentip= |
| Admin.RibbonButtonGetEnabled.btnAdminCreateWarehouse | PASS | RibbonRequiredCapabilityGetEnabledAdmin |
| Admin.MacroExists.btnAdminCreateWarehouse | PASS | modAdmin.Open_CreateWarehouse |
| Admin.CallbackMap.btnAdminCreateWarehouse | PASS | btnAdminCreateWarehouse -> modAdmin.Open_CreateWarehouse |
| Admin.CallbackGetEnabled.btnAdminCreateWarehouse | PASS | btnAdminCreateWarehouse -> ADMIN_MAINT |
| Admin.DisabledOffline.btnAdminCreateWarehouse | PASS | btnAdminCreateWarehouse enabled=False |
| Admin.RibbonButton.btnAdminSetupTesterStation | PASS | Label=Setup Tester Station; OnAction=RibbonOnActionAdmin; GetEnabled=RibbonRequiredCapabilityGetEnabledAdmin; Screentip= |
| Admin.RibbonButtonGetEnabled.btnAdminSetupTesterStation | PASS | RibbonRequiredCapabilityGetEnabledAdmin |
| Admin.MacroExists.btnAdminSetupTesterStation | PASS | modAdmin.Admin_SetupTesterStation_Click |
| Admin.CallbackMap.btnAdminSetupTesterStation | PASS | btnAdminSetupTesterStation -> modAdmin.Admin_SetupTesterStation_Click |
| Admin.CallbackGetEnabled.btnAdminSetupTesterStation | PASS | btnAdminSetupTesterStation -> ADMIN_MAINT |
| Admin.DisabledOffline.btnAdminSetupTesterStation | PASS | btnAdminSetupTesterStation enabled=False |
| Admin.RibbonButton.btnAdminSeedInventory | PASS | Label=Seed Demo Inventory; OnAction=RibbonOnActionAdmin; GetEnabled=RibbonRequiredCapabilityGetEnabledAdmin; Screentip= |
| Admin.RibbonButtonGetEnabled.btnAdminSeedInventory | PASS | RibbonRequiredCapabilityGetEnabledAdmin |
| Admin.MacroExists.btnAdminSeedInventory | PASS | modAdmin.Seed_DemoInventory |
| Admin.CallbackMap.btnAdminSeedInventory | PASS | btnAdminSeedInventory -> modAdmin.Seed_DemoInventory |
| Admin.CallbackGetEnabled.btnAdminSeedInventory | PASS | btnAdminSeedInventory -> ADMIN_MAINT |
| Admin.DisabledOffline.btnAdminSeedInventory | PASS | btnAdminSeedInventory enabled=False |
| Admin.RibbonButton.btnAdminVerifyAddinsPublished | PASS | Label=Verify Add-ins Published; OnAction=RibbonOnActionAdmin; GetEnabled=RibbonRequiredCapabilityGetEnabledAdmin; Screentip= |
| Admin.RibbonButtonGetEnabled.btnAdminVerifyAddinsPublished | PASS | RibbonRequiredCapabilityGetEnabledAdmin |
| Admin.MacroExists.btnAdminVerifyAddinsPublished | PASS | modAdmin.Verify_AddinsPublished |
| Admin.CallbackMap.btnAdminVerifyAddinsPublished | PASS | btnAdminVerifyAddinsPublished -> modAdmin.Verify_AddinsPublished |
| Admin.CallbackGetEnabled.btnAdminVerifyAddinsPublished | PASS | btnAdminVerifyAddinsPublished -> ADMIN_MAINT |
| Admin.DisabledOffline.btnAdminVerifyAddinsPublished | PASS | btnAdminVerifyAddinsPublished enabled=False |
| Admin.RibbonButton.btnAdminRetireMigrateWarehouse | PASS | Label=Retire / Migrate Warehouse; OnAction=RibbonOnActionAdmin; GetEnabled=RibbonRequiredCapabilityGetEnabledAdmin; Screentip=Archive, migrate, retire, or delete a warehouse runtime |
| Admin.RibbonButtonScreentip.btnAdminRetireMigrateWarehouse | PASS | Archive, migrate, retire, or delete a warehouse runtime |
| Admin.RibbonButtonGetEnabled.btnAdminRetireMigrateWarehouse | PASS | RibbonRequiredCapabilityGetEnabledAdmin |
| Admin.MacroExists.btnAdminRetireMigrateWarehouse | PASS | modAdmin.Admin_RetireMigrateWarehouse_Click |
| Admin.CallbackMap.btnAdminRetireMigrateWarehouse | PASS | btnAdminRetireMigrateWarehouse -> modAdmin.Admin_RetireMigrateWarehouse_Click |
| Admin.CallbackGetEnabled.btnAdminRetireMigrateWarehouse | PASS | btnAdminRetireMigrateWarehouse -> ADMIN_MAINT |
| Admin.DisabledOffline.btnAdminRetireMigrateWarehouse | PASS | btnAdminRetireMigrateWarehouse enabled=False |
