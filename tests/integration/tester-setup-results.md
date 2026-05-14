# Tester Setup Integration Results

- Date: 2026-05-14 15:13:31
- Overall: PASS
- Harness: C:\Users\justu\source\repos\invSys_fork\tests\fixtures\TesterSetup_Integration_Harness_20260514_151245_714.xlsm
- Summary: Tester station setup passed fresh-machine, existing-hub, rerun-safe, offline-SharePoint, and existing-auth cases.
- Passed checks: 5
- Failed checks: 0

| Check | Result | Detail |
|---|---|---|
| FreshMachine.CreatesRuntimeAndWorkbook | PASS | Fresh setup created the runtime tree, auth/config state, TEST-SKU-001 seed, and a valid receiving workbook. |
| ExistingHub.CreatesNamespacedTesterRuntime | PASS | Existing hub folder accepted a namespaced tester runtime without requiring matching warehouse artifacts first. |
| IdempotentRerun.DoesNotDuplicateSeed | PASS | Second setup reused the runtime and left TEST-SKU-001 at QtyOnHand = 100. |
| SharePointUnavailable.LocalSetupStillSucceeds | PASS | Local setup succeeded and recorded the unavailable SharePoint root without blocking runtime creation. |
| ExistingAuth.HashPreservedCapabilitiesUpdated | PASS | Existing tester auth kept the original hash and restored RECEIVE_POST, RECEIVE_VIEW, and READMODEL_REFRESH. |
