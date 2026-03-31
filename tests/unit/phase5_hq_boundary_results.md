# Phase 5 HQ Boundary Validation Results

- Date: 2026-03-31 08:04:55
- Passed: 7
- Failed: 0

| Check | Result |
|---|---|
| Setup.WH97 | PASS - OK|Warehouse=WH97|Station=S1 |
| Setup.WH98 | PASS - OK|Warehouse=WH98|Station=S2 |
| Publish.WH97.Initial | PASS - OK|EventID=EVT-WH97-20260331080448-567928|Processed=1|Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH97-INVENTORY-20260331080449-143394|PublishedPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase5_hq_boundary_20260331_080437_989\share\Snapshots\WH97.invSys.Snapshot.Inventory.xlsb |
| Publish.WH98.Initial | PASS - OK|EventID=EVT-WH98-20260331080450-444164|Processed=1|Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260331080451-537407|PublishedPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase5_hq_boundary_20260331_080437_989\share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Initial | PASS - OK|Report=Rows=2; SnapshotFiles=2; SkippedSnapshotFiles=0|QtyA=5|QtyB=8|SourceA=WH97.invSys.Snapshot.Inventory.xlsb|SourceB=WH98.invSys.Snapshot.Inventory.xlsb|Skipped=0|Warehouses=2 |
| Publish.WH98.Catchup | PASS - OK|EventID=EVT-WH98-20260331080453-704747|Processed=1|Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260331080453-416134|PublishedPath=C:\Users\Justin\repos\invSys_fork\tests\fixtures\phase5_hq_boundary_20260331_080437_989\share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Catchup | PASS - OK|Report=Rows=2; SnapshotFiles=2; SkippedSnapshotFiles=0|QtyA=5|QtyB=11|SourceA=WH97.invSys.Snapshot.Inventory.xlsb|SourceB=WH98.invSys.Snapshot.Inventory.xlsb|Skipped=0|Warehouses=2 |
