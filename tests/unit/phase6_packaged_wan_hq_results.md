# Phase 6 Packaged WAN HQ Validation Results

- Date: 2026-03-30 19:43:51
- Deploy root: C:\Users\Justin\repos\invSys_fork\deploy\current
- Session root: C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-e85f205d64134e5fa51e697b3a7de6ec
- Passed: 10
- Failed: 0

| Check | Result | Detail |
|---|---|---|
| Setup.RuntimeRoots | PASS | SessionRoot=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-e85f205d64134e5fa51e697b3a7de6ec |
| Packaged.OpenA | PASS | Core+Inventory.Domain |
| Packaged.OpenB | PASS | Core+Inventory.Domain |
| Packaged.OpenHQ | PASS | Core+Inventory.Domain |
| Packaged.RuntimeOverrides | PASS | WH97=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-e85f205d64134e5fa51e697b3a7de6ec\WH97; WH98=C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-e85f205d64134e5fa51e697b3a7de6ec\WH98 |
| Publish.WH97.Initial | PASS | EventID=EVT-WH97-20260330194339967; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH97-INVENTORY-20260330194341-297512; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-e85f205d64134e5fa51e697b3a7de6ec\Share\Snapshots\WH97.invSys.Snapshot.Inventory.xlsb |
| Publish.WH98.Initial | PASS | EventID=EVT-WH98-20260330194340428; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260330194343-417553; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-e85f205d64134e5fa51e697b3a7de6ec\Share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Initial | PASS | QtyA=5; QtyB=8 |
| Publish.WH98.Catchup | PASS | EventID=EVT-WH98-20260330194348759; Processed=1; Report=Applied=1; SkipDup=0; Poison=0; RunId=RUN-WH98-INVENTORY-20260330194349-087410; C:\Users\Justin\AppData\Local\Temp\invsys-phase6-wanhq-e85f205d64134e5fa51e697b3a7de6ec\Share\Snapshots\WH98.invSys.Snapshot.Inventory.xlsb |
| Aggregate.Catchup | PASS | QtyA=5; QtyB=11 |
