# Expert Handoff - 2026-04-20 NAS Prep

Read these first:
- [WAN proving path.md](c:/Users/Justin/repos/invSys_fork/0%20plan%20docs/xlam_invSys/WAN%20proving%20path.md)
- [invSys-Design-v4.8.md](c:/Users/Justin/repos/invSys_fork/0%20plan%20docs/xlam_invSys/invSys-Design-v4.8.md)

## Current State

- Repo-side work is pushed through commit `fc5fb9f` on `origin/main`.
- The Synology DS920+ was not reachable from the current location, so no real NAS-backed warehouse hub testing was done yet.
- Resume NAS work only when the SMB path is actually reachable from the client machine.

## Completed Repo-side Prep

- Added shared deployment path policy in [modDeploymentPaths.bas](c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modDeploymentPaths.bas).
- Replaced scattered `C:\invSys\...` fallbacks with shared helpers across bootstrap/runtime/admin/tester callers.
- Added UNC-safe folder/file/path helpers for warehouse hub roots.
- Updated [frmCreateWarehouse.frm](c:/Users/Justin/repos/invSys_fork/src/Admin/Forms/frmCreateWarehouse.frm) so `PathLocal` is treated as a warehouse hub path, with Synology/SMB guidance and a folder browse helper.
- Updated [WAN proving path.md](c:/Users/Justin/repos/invSys_fork/0%20plan%20docs/xlam_invSys/WAN%20proving%20path.md) so `<PathDataRoot>` is the warehouse hub root contract, not a hard-coded `C:\invSys\WHx`.
- Updated WAN setup proof modules to resolve runtime roots dynamically instead of assuming `C:\invSys\WH1` / `C:\invSys\WH2`.
- Fixed PowerShell harness import lists so ad hoc VBA test workbooks include `modDeploymentPaths.bas`.

## Validated Without NAS

- [phase6_test_results.md](c:/Users/Justin/repos/invSys_fork/tests/unit/phase6_test_results.md): `PASSED=99 FAILED=0 TOTAL=99`
- [create-warehouse-results.md](c:/Users/Justin/repos/invSys_fork/tests/integration/create-warehouse-results.md): `OVERALL=PASS PASSED=9 FAILED=0 TOTAL=9`
- [retire-migrate-results.md](c:/Users/Justin/repos/invSys_fork/tests/integration/retire-migrate-results.md): `OVERALL=PASS PASSED=8 FAILED=0 TOTAL=8`
- [tester-setup-results.md](c:/Users/Justin/repos/invSys_fork/tests/integration/tester-setup-results.md): `OVERALL=PASS PASSED=4 FAILED=0 TOTAL=4`

## Not Yet Proven

- No real DS920+ SMB warehouse hub path has been tested.
- No NAS-backed `Create New Warehouse` run has been executed.
- No real `WH1` / `WH2` warehouse hub on Synology has been established.
- No WAN proving slice that depends on NAS reachability has been executed.

## Next Step When NAS Is Reachable

1. Verify the actual SMB path from the client machine, for example `\\DS920\<share>\WH1`.
2. Run `Create New Warehouse` using the DS920+ warehouse hub path as `PathLocal`.
3. Confirm bootstrap can create:
   - `<hub>\WH1.invSys.Config.xlsb`
   - `<hub>\WH1.invSys.Auth.xlsb`
   - `<hub>\WH1.invSys.Data.Inventory.xlsb`
   - `<hub>\WH1.Outbox.Events.xlsb`
   - `<hub>\WH1.invSys.Snapshot.Inventory.xlsb`
   - `<hub>\inbox\`, `<hub>\outbox\`, `<hub>\snapshots\`, `<hub>\config\`
4. If Excel/VBE stops, capture the exact dialog text and highlighted symbol/line before changing code.
5. After successful NAS-backed bootstrap, resume WAN proving Slice A and Slice B against the real warehouse hub roots.

## Likely First Risk

- Remaining VBA code paths outside the already-patched bootstrap/runtime surface may still use local-only filesystem assumptions when pointed at a UNC/SMB root.
- Patch only against real DS920+ failures once the path is reachable. Do not guess ahead of the first concrete NAS error.

## Last Useful Commits

- `fc5fb9f` Fix harness imports for deployment path module
- `bce2ea0` pushed the next repo-side step toward the DS920+ model,
