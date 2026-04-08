# Create Warehouse Integration Results

- Date: 2026-04-08 10:18:34
- Overall: PASS
- Harness: C:\Users\Justin\repos\invSys_fork\tests\fixtures\CreateWarehouse_Integration_Harness_20260408_101830_916.xlsm
- Warehouse: WHBOOT-E2E_01
- Station: ADM1
- Local root: C:\Users\Justin\AppData\Local\Temp\invSys_createwarehouse_local_20260408_101832_3.711255E+07
- SharePoint root: C:\Users\Justin\AppData\Local\Temp\invSys_createwarehouse_share_20260408_101832_3.711255E+07
- Summary: Create warehouse lifecycle completed, SharePoint artifacts were published, and duplicate rejection was proven.
- Passed checks: 9
- Failed checks: 0

| Check | Result | Detail |
|---|---|---|
| WarehouseSpec.Valid | PASS | OK |
| CollisionCheck.InitialClear | PASS | WarehouseIdExists=False |
| Bootstrap.Local | PASS | OK |
| LocalStructure.Exists | PASS | All required runtime folders and seeded artifacts were created under C:\Users\Justin\AppData\Local\Temp\invSys_createwarehouse_local_20260408_101832_3.711255E+07 |
| ConfigSeeded.Correctly | PASS | Config workbook seeded WarehouseId, WarehouseName, StationId, PathDataRoot, PathSharePointRoot, and ADMIN defaults. |
| SharePointPublish.Initial | PASS | OK ; Config=COPIED:C:\Users\Justin\AppData\Local\Temp\invSys_createwarehouse_share_20260408_101832_3.711255E+07\WHBOOT-E2E_01\WHBOOT-E2E_01.invSys.Config.xlsb ; Discovery=COPIED:C:\Users\Justin\AppData\Local\Temp\invSys_createwarehouse_share_20260408_101832_3.711255E+07\WHBOOT-E2E_01.config.json |
| SharePointArtifacts.Exists | PASS | Discovery artifact and published config workbook exist under C:\Users\Justin\AppData\Local\Temp\invSys_createwarehouse_share_20260408_101832_3.711255E+07 |
| CollisionCheck.DuplicateVisible | PASS | WarehouseIdExists=True |
| DuplicateRun.Rejected | PASS | WarehouseId already exists. |
