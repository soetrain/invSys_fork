# Create Warehouse Integration Results

- Date: 2026-05-15 08:56:21
- Overall: PASS
- Harness: C:\Users\justu\source\repos\invSys_fork\tests\fixtures\CreateWarehouse_Integration_Harness_20260515_085611_887.xlsm
- Warehouse: WHBOOT-E2E_01
- Station: ADM1
- Local root: C:\Users\justu\AppData\Local\Temp\invSys_createwarehouse_local_20260515_085615_3.217575E+07
- SharePoint root: C:\Users\justu\AppData\Local\Temp\invSys_createwarehouse_share_20260515_085615_3.217575E+07
- Summary: Create warehouse lifecycle completed, SharePoint artifacts were published, and duplicate rejection was proven.
- Passed checks: 9
- Failed checks: 0

| Check | Result | Detail |
|---|---|---|
| WarehouseSpec.Valid | PASS | OK |
| CollisionCheck.InitialClear | PASS | WarehouseIdExists=False |
| Bootstrap.Local | PASS | OK ; Hub=C:\Users\justu\AppData\Local\Temp\invSys_createwarehouse_local_20260515_085615_3.217575E+07 ; Inbox=C:\Users\justu\AppData\Local\Temp\invSys_createwarehouse_local_20260515_085615_3.217575E+07\inbox\invSys.Inbox.Receiving.ADM1.xlsb ; Seed=SEEDED ; Operator=C:\Users\justu\Documents\invSys\OperatorWorkbooks\WHBOOT-E2E_01\ADM1\WHBOOT-E2E_01.Receiving.Operator.xlsm |
| LocalStructure.Exists | PASS | All required runtime folders and seeded artifacts were created under C:\Users\justu\AppData\Local\Temp\invSys_createwarehouse_local_20260515_085615_3.217575E+07 |
| ConfigSeeded.Correctly | PASS | Config workbook seeded WarehouseId, WarehouseName, StationId, PathDataRoot, PathSharePointRoot, and RECEIVE defaults. |
| SharePointPublish.Initial | PASS | OK ; Config=COPIED:C:\Users\justu\AppData\Local\Temp\invSys_createwarehouse_share_20260515_085615_3.217575E+07\WHBOOT-E2E_01\WHBOOT-E2E_01.invSys.Config.xlsb ; Discovery=COPIED:C:\Users\justu\AppData\Local\Temp\invSys_createwarehouse_share_20260515_085615_3.217575E+07\WHBOOT-E2E_01.config.json |
| SharePointArtifacts.Exists | PASS | Discovery artifact and published config workbook exist under C:\Users\justu\AppData\Local\Temp\invSys_createwarehouse_share_20260515_085615_3.217575E+07 |
| CollisionCheck.DuplicateVisible | PASS | WarehouseIdExists=True |
| DuplicateRun.Rejected | PASS | WarehouseId already exists in the configured warehouse catalog: WHBOOT-E2E_01 |
