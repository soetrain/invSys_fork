# Phase 6 LAN Boundary Validation Results

- Date: 2026-03-25 19:01:08
- Passed: 10
- Failed: 0

| Check | Result |
|---|---|
| Setup.SharedRoot | PASS - OK|Root=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260325_190058_916\runtime|Published=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260325_190058_916\published |
| Attach.SessionA | PASS - OK|Warehouse=WH89|Station=S1 |
| Attach.SessionB | PASS - OK|Warehouse=WH89|Station=S2 |
| Lock.SessionAHold | PASS - OK|Path=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260325_190058_916\runtime\WH89.invSys.Data.Inventory.xlsb|ReadOnly=False |
| Lock.SessionBDeniedByFileBoundary | PASS - OK|EventID=E44A15D5-E85B-4ED7-A43A-3D1B706AABA5l|Processed=0|Report=Inventory workbook is read-only or locked by another Excel session.|Status=NEW|ErrorCode=|ErrorMessage= |
| Lock.SessionARelease | PASS - OK|Closed |
| Lock.SessionARetryAfterRelease | PASS - OK|EventID=CD9425D9-53C5-4B05-900B-10F9ED99B754l|Processed=1|Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH89-INVENTORY-20260325190107-800961|Status=PROCESSED|ErrorCode=|ErrorMessage= |
| Publish.SessionAToSharedSnapshot | PASS - OK|PublishedPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260325_190058_916\published\WH89.invSys.Snapshot.Inventory.xlsb |
| Operator.BuildStationB | PASS - OK|OperatorPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260325_190058_916\stationB_operator.xlsb |
| Refresh.SessionBReadsPublishedSnapshot | PASS - OK|TotalInv=7|QtyAvailable=7|SnapshotId=WH89.invSys.Snapshot.Inventory.xlsb/20260325190108|SourceType=SHAREPOINT|IsStale=False|Path=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260325_190058_916\stationB_operator.xlsb |
