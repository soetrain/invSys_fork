# WAN WH2 Setup Proof

- Warehouse: `WH2`
- Scope note: real-machine setup, publish proof, and WH1 cross-contamination check for the WAN proving path Slice B.

| Machine | Step | Result | Note | UTC timestamp |
|---|---|---|---|---|
| JWJZENBOOK | 1 | FAIL | Missing runtime root: C:\invSys\WH2 | 2026-04-08 16:03:18   |
| JWJZENBOOK | 2 | FAIL | Missing inventory workbook: C:\invSys\WH2\WH2.invSys.Data.Inventory.xlsb | 2026-04-08 16:03:18   |
| JWJZENBOOK | 3 | FAIL | Missing outbox workbook: C:\invSys\WH2\WH2.Outbox.Events.xlsb | 2026-04-08 16:03:19   |
| JWJZENBOOK | 4 | FAIL | Missing local snapshot workbook: C:\invSys\WH2\WH2.invSys.Snapshot.Inventory.xlsb | 2026-04-08 16:03:19   |
| JWJZENBOOK | 5 | FAIL | Config workbook missing: C:\invSys\WH2\WH2.invSys.Config.xlsb | 2026-04-08 16:03:20   |
| JWJZENBOOK | 6 | FAIL | SharePoint root was not resolved from config. | 2026-04-08 16:03:21   |
| JWJZENBOOK | 7 | FAIL | SharePoint root was not resolved from config. | 2026-04-08 16:03:21   |
| JWJZENBOOK | 8 | FAIL | StationId could not be resolved from config; RunBatch was not attempted. | 2026-04-08 16:03:22   |
| JWJZENBOOK | 9 | FAIL | Missing published snapshot: Snapshots\WH2.invSys.Snapshot.Inventory.xlsb | 2026-04-08 16:03:23   |
| JWJZENBOOK | 10 | PASS | No publish temp file remains at Snapshots\WH2.invSys.Snapshot.Inventory.xlsb.uploading. | 2026-04-08 16:03:23   |
| JWJZENBOOK | 11 | FAIL | Peer WH1 published snapshot missing after WH2 publish: Snapshots\WH1.invSys.Snapshot.Inventory.xlsb | 2026-04-08 16:03:24   |
