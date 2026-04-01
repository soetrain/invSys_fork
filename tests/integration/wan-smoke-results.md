# WAN Smoke Results

- Date: 2026-03-31 16:45:00
- Warehouse: `WH1`
- Station: `S1`
- SharePoint synced path: `C:\Users\Justin\Black Scottie Chai\Justinwj - invSys`
- Runtime root: `C:\invSys\WH1`
- Scope note: user-side WAN smoke proof covering sync setup, config wiring, reachable publish, and failure-path handling.

| Step | Result | Detail |
|---|---|---|
| 1. Sync SharePoint Library | PASS | SharePoint library synced locally at `C:\Users\Justin\Black Scottie Chai\Justinwj - invSys`; folder confirmed on disk. |
| 2. Wire Path Into Config | PASS | `PathSharePointRoot` saved in `C:\invSys\WH1\WH1.invSys.Config.xlsb`; `modConfig.LoadConfig("WH1","S1") = True`; loaded value matched synced path. |
| 3. Smoke Test Publish | PASS | Admin scheduler-safe publish returned `OK`; published files created at `Events\WH1.Outbox.Events.xlsb` and `Snapshots\WH1.invSys.Snapshot.Inventory.xlsb`; `invSys.Publish.log` recorded `Result=OK`. |
| 4a. Pause OneDrive Sync And Publish | PASS | With OneDrive paused, publish still returned `OK`; local snapshot write succeeded and local synced-folder copy succeeded; no unhandled error thrown. This proves pause-sync does not make the local SharePoint sync folder unavailable. |
| 4b. True Publish Failure Probe | PASS | Publish rerun against unavailable override root `Z:\invSys-offline-probe` returned structured `FAIL`; local snapshot remained available at `C:\invSys\WH1\WH1.invSys.Snapshot.Inventory.xlsb`; `invSys.Publish.log` recorded `Result=FAIL`; no unhandled error thrown. |
| 5. Record Evidence | PASS | Results captured in this file. |

## Evidence

| Check | Result | Detail |
|---|---|---|
| Config.LoadConfig | PASS | `WarehouseId=WH1`; `StationId=S1`; `PathSharePointRoot=C:\Users\Justin\Black Scottie Chai\Justinwj - invSys` |
| Publish.Online.Result | PASS | `OK|WarehouseId=WH1|SnapshotPath=C:\invSys\WH1\WH1.invSys.Snapshot.Inventory.xlsb|Publish=Root=C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\|Outbox=COPIED:C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\Events\WH1.Outbox.Events.xlsb|Snapshot=COPIED:C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\Snapshots\WH1.invSys.Snapshot.Inventory.xlsb` |
| Publish.Online.Files | PASS | `C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\Events\WH1.Outbox.Events.xlsb`; `C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\Snapshots\WH1.invSys.Snapshot.Inventory.xlsb` |
| Publish.Log.Success | PASS | `2026-03-31 15:50:29 | WarehouseId=WH1 | RunId= | Result=OK | Root=C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\|Outbox=COPIED:C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\Events\WH1.Outbox.Events.xlsb|Snapshot=COPIED:C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\Snapshots\WH1.invSys.Snapshot.Inventory.xlsb` |
| Publish.PausedSync.Result | PASS | `OK|WarehouseId=WH1|SnapshotPath=C:\invSys\WH1\WH1.invSys.Snapshot.Inventory.xlsb|Publish=Root=C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\|Outbox=REPLACED:C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\Events\WH1.Outbox.Events.xlsb|Snapshot=REPLACED:C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\Snapshots\WH1.invSys.Snapshot.Inventory.xlsb` |
| Publish.Log.PausedSync | PASS | `2026-03-31 16:31:31 | WarehouseId=WH1 | RunId= | Result=OK | Root=C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\|Outbox=REPLACED:C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\Events\WH1.Outbox.Events.xlsb|Snapshot=REPLACED:C:\Users\Justin\Black Scottie Chai\Justinwj - invSys\Snapshots\WH1.invSys.Snapshot.Inventory.xlsb` |
| Publish.TrueFailure.Result | PASS | `FAIL|WarehouseId=WH1|SnapshotPath=C:\invSys\WH1\WH1.invSys.Snapshot.Inventory.xlsb|Publish=Root=Z:\invSys-offline-probe\|Outbox=FAILED:|Snapshot=FAILED:` |
| Publish.Log.Failure | PASS | `2026-03-31 16:40:17 | WarehouseId=WH1 | RunId= | Result=FAIL | Root=Z:\invSys-offline-probe\|Outbox=FAILED:|Snapshot=FAILED:` |
| LocalAuthorityPreserved | PASS | Local canonical snapshot remained at `C:\invSys\WH1\WH1.invSys.Snapshot.Inventory.xlsb`; warehouse-local write path remained usable regardless of publish outcome. |

## Notes

- OneDrive `Pause syncing` is not equivalent to SharePoint path unavailability for this architecture. The local synced folder remained writable, so publish-to-folder still succeeded while cloud upload was deferred.
- The controlled unavailable-root probe was needed to prove the failure-handling contract: local write first, publish failure logged, no unhandled error, and no change to warehouse authority.
