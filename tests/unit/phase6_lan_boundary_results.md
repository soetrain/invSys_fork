# Phase 6 LAN Boundary Validation Results

- Date: 2026-03-27 23:25:46
- Passed: 10
- Failed: 0

| Check | Result |
|---|---|
| Setup.SharedRoot | PASS - OK|Root=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_232536_791\runtime|Published=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_232536_791\published |
| Attach.SessionA | PASS - OK|Warehouse=WH89|Station=S1 |
| Attach.SessionB | PASS - OK|Warehouse=WH89|Station=S2 |
| Lock.SessionAHold | PASS - OK|Path=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_232536_791\runtime\WH89.invSys.Data.Inventory.xlsb|ReadOnly=False |
| Lock.SessionBDeniedByFileBoundary | PASS - OK|EventID=0B995F9D-61D3-4AA2-8B0C-ECE5BA012C1Fl|Processed=0|Report=Inventory workbook is read-only or locked by another Excel session.|Status=NEW|ErrorCode=|ErrorMessage= |
| Lock.SessionARelease | PASS - OK|Closed |
| Lock.SessionARetryAfterRelease | PASS - OK|EventID=78244837-AF2B-47F5-9E5F-6B7C39F86098l|Processed=1|Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH89-INVENTORY-20260327232544-483120|Status=PROCESSED|ErrorCode=|ErrorMessage= |
| Publish.SessionAToSharedSnapshot | PASS - OK|PublishedPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_232536_791\published\WH89.invSys.Snapshot.Inventory.xlsb |
| Operator.BuildStationB | PASS - OK|OperatorPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_232536_791\stationB_operator.xlsb |
| Refresh.SessionBReadsPublishedSnapshot | PASS - OK|TotalInv=7|QtyAvailable=7|SnapshotId=WH89.invSys.Snapshot.Inventory.xlsb/20260327232545|SourceType=SHAREPOINT|IsStale=False|Path=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase6_lan_boundary_20260327_232536_791\stationB_operator.xlsb |
