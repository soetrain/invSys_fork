# Tester Setup Integration Results

- Date: 2026-04-20 15:59:54
- Overall: PASS
- Harness: C:\Users\Justin\repos\invSys_fork\tests\fixtures\TesterSetup_Integration_Harness_20260420_155932_270.xlsm
- Summary: Tester station setup passed fresh-machine, rerun-safe, offline-SharePoint, and existing-auth cases.
- Passed checks: 4
- Failed checks: 0

| Check | Result | Detail |
|---|---|---|
| FreshMachine.CreatesRuntimeAndWorkbook | PASS | Fresh setup created the runtime tree, auth/config state, TEST-SKU-001 seed, and a valid receiving workbook. |
| IdempotentRerun.DoesNotDuplicateSeed | PASS | Second setup reused the runtime and left TEST-SKU-001 at QtyOnHand = 100. |
| SharePointUnavailable.LocalSetupStillSucceeds | PASS | Local setup succeeded and recorded the unavailable SharePoint root without blocking runtime creation. |
| ExistingAuth.HashPreservedCapabilitiesUpdated | PASS | Existing tester auth kept the original hash and restored RECEIVE_POST, RECEIVE_VIEW, and READMODEL_REFRESH. |
