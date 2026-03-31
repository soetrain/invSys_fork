# Phase 6 Packaged WAN HQ Validation Results

- Date: 2026-03-31 08:05:26
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Session root: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-d5eec49b18e54a9399cb4012cb3daf87
- Passed: 10
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Setup.RuntimeRoots | PASS | SessionRoot=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-d5eec49b18e54a9399cb4012cb3daf87 |
| Packaged.OpenA | PASS | Core+Inventory.Domain |
| Packaged.OpenB | PASS | Core+Inventory.Domain |
| Packaged.OpenHQ | PASS | Core+Inventory.Domain |
| Packaged.RuntimeOverrides | PASS | WH97=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-d5eec49b18e54a9399cb4012cb3daf87\WH97; WH98=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-d5eec49b18e54a9399cb4012cb3daf87\WH98 |
| Publish.WH97.Initial | PASS | EventID=EVT-WH97-20260331080514341; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH97-INVENTORY-20260331080515-322079; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-d5eec49b18e54a9399cb4012cb3daf87\Share\Snapshots\WH97.invSys.Snapshot.Inventory.xlsb |
| Publish.WH98.Initial | PASS | EventID=EVT-WH98-20260331080514805; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260331080518-910885; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-d5eec49b18e54a9399cb4012cb3daf87\Share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Initial | PASS | QtyA=5; QtyB=8 |
| Publish.WH98.Catchup | PASS | EventID=EVT-WH98-20260331080523767; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260331080524-138389; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-d5eec49b18e54a9399cb4012cb3daf87\Share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Catchup | PASS | QtyA=5; QtyB=11 |
