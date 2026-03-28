# Phase 6 LAN Boundary Validation Results

- Date: 2026-03-27 22:36:34
- Passed: 10
- Failed: 0

| Check | Result |
|---|---|
| Setup.SharedRoot | PASS - OK|Root=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_223624_388\runtime|Published=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_223624_388\published |
| Attach.SessionA | PASS - OK|Warehouse=WH89|Station=S1 |
| Attach.SessionB | PASS - OK|Warehouse=WH89|Station=S2 |
| Lock.SessionAHold | PASS - OK|Path=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_223624_388\runtime\WH89.invSys.Data.Inventory.xlsb|ReadOnly=False |
| Lock.SessionBDeniedByFileBoundary | PASS - OK|EventID=4C5B702D-110B-4896-80C9-6F929B134477l|Processed=0|Report=Inventory workbook is read-only or locked by another Excel session.|Status=NEW|ErrorCode=|ErrorMessage= |
| Lock.SessionARelease | PASS - OK|Closed |
| Lock.SessionARetryAfterRelease | PASS - OK|EventID=BB939AA2-7B9E-48CF-B5A3-7E0DAF3B2793l|Processed=1|Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH89-INVENTORY-20260327223632-982983|Status=PROCESSED|ErrorCode=|ErrorMessage= |
| Publish.SessionAToSharedSnapshot | PASS - OK|PublishedPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_223624_388\published\WH89.invSys.Snapshot.Inventory.xlsb |
| Operator.BuildStationB | PASS - OK|OperatorPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_223624_388\stationB_operator.xlsb |
| Refresh.SessionBReadsPublishedSnapshot | PASS - OK|TotalInv=7|QtyAvailable=7|SnapshotId=WH89.invSys.Snapshot.Inventory.xlsb/20260327223634|SourceType=SHAREPOINT|IsStale=False|Path=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_223624_388\stationB_operator.xlsb |
