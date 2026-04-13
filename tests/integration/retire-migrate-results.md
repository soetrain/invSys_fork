# Retire / Migrate Integration Results

- Date: 2026-04-12 22:20:35
- Overall: PASS
- Harness: C:\Users\Justin\repos\invSys_fork\tests\fixtures\RetireMigrate_Integration_Harness_20260412_221958_712.xlsm
- Summary: Retire/migrate lifecycle cases passed for archive-only, migrate, retire, delete, reuse rejection, and safety guards.
- Passed checks: 8
- Failed checks: 0

| Check | Result | Detail |
|---|---|---|
| ArchiveOnly.SourceUntouched | PASS | Archive package was complete and manifest-valid; source runtime remained present with WarehouseStatus=ACTIVE. |
| ArchiveMigrate.TargetSeededNoIdentityBleed | PASS | Target inventory appended source state to QtyOnHand=7, MigrationSourceId was logged, auth was not copied, and target config identity stayed intact. |
| ArchiveRetire.TombstoneAndStatus | PASS | Retirement stamped WarehouseStatus=RETIRED, wrote the local tombstone, and published the tombstone to the SharePoint folder. |
| ArchiveRetireDelete.RuntimeRemoved | PASS | Delete mode only ran after the tombstone existed, and the local runtime folder tree was removed. |
| RetiredReuse.Rejected | PASS | After retirement, the same WarehouseId was still network-visible via published artifacts and duplicate bootstrap was rejected. |
| DeleteNoManifest.Rejected | PASS | DeleteLocalRuntime rejected a hand-dropped tombstone when no archive manifest existed, and the runtime folder remained untouched. |
| DeleteNoConfirmation.Rejected | PASS | DeleteLocalRuntime rejected the unconfirmed destructive request and left the runtime folder intact. |
| RetireSharePointUnavailable.WarningOnly | PASS | Retirement completed with a local tombstone, SharePoint failure stayed advisory via PublishWarning, and diagnostics captured the warning. |
