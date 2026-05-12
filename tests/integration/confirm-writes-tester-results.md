# Confirm Writes Tester Integration Results

- Date: 2026-05-12 12:07:51
- Machine: X1-PRO-AI
- Overall: PASS
- Harness: C:\Users\Justin\repos\invSys_fork\tests\fixtures\ConfirmWritesTester_Integration_Harness_20260512_120727_111.xlsm
- Summary: Confirm Writes tester proving flow passed all eight deterministic cases.
- Runtime user: Justin
- Tester user: TESTER01
- Passed checks: 8
- Failed checks: 0

| Check | Result | Detail |
|---|---|---|
| ReadinessCheck_OK | PASS | CheckReceivingReadiness returned ready for the configured WH1 receiving workbook. |
| ReadinessCheck_MissingCapability | PASS | Readiness returned MISSING_CAPABILITY after RECEIVE_POST was removed from the effective runtime user. |
| RefreshInventory_Loads | PASS | Refresh populated the read model with TEST-SKU-001 at QtyOnHand = 100. |
| ConfirmWrites_SingleRow | PASS | Confirm Writes cleared staging, appended ReceivedLog, and wrote inbox event EF1BE2A1-9ED0-4E75-878B-5036E36C3163s with inbox status PROCESSED. |
| ProcessorApplies | PASS | Inventory QtyOnHand reached 110 and the inbox row is PROCESSED. RunBatch processed 0 rows. |
| SnapshotRefreshAfterPost | PASS | Snapshot refresh showed TEST-SKU-001 at QtyOnHand = 110 with current LOCAL metadata. |
| IdempotentSetup | PASS | Rerun reused the runtime, preserved the workbook file, and did not duplicate seed or capability rows. |
| SharePointUnavailable | PASS | Offline SharePoint refresh completed without hard failure and rendered a stale read-model banner: INVENTORY SNAPSHOT STALE  ;  Source=CACHED  ;  Refreshed=2026-05-12 12:07:45  ;  SnapshotId=WH1.invSys.Snapshot.Inventory.xlsb ; 20260512120743  ;  Snapshot workbook not found for source SHAREPOINT; operator read model kept cached. |
