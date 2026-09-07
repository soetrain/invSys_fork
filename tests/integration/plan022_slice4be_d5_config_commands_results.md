# Plan 022 Slice 4be: D5 configuration command separation

Last verified: 2026-09-06. Governing approval: Architecture v4.11 D5,
Plan 022's D5 prerequisite, and controls catalog v1.59. This evidence does not
approve the separate D18 Event Detail/recorded Action Path proposal.

The ordinary Config API reads existing authority without creating, repairing,
or saving it. Admin Settings saves through headless `modConfigCommands` after
Core checks sign-in, capability, and captured warehouse/station. The old scalar
entry forwards to that command. Production's validated UOM command retains
`PROD_POST` access to its catalog fields while denying unrelated config writes.

## D13 trace

- Before implementation, the packaged focused test on baseline `b108313`
  recorded 5 PASS / 5 FAIL: unauthorized direct writes, stale Settings target,
  read-side schema repair, and arbitrary Production writes were meaningful RED.
- Expanded comparison against the unchanged baseline packages recorded
  8 PASS / 9 FAIL. A later focused required-schema case recorded 17 PASS / 1 FAIL
  before adding the full read-only validation gate to the writer. Final
  implementation recorded **18 PASS / 0 FAIL**.
- Harness startup errors were corrected before recording behavioral RED.
  Compile errors were separate blockers, not behavioral RED.
- Tests inject unsaved instrumentation into packaged forms to invoke the real
  Settings selection/save and Production UOM Send/Retrieve handlers. Direct
  service checks supplement those actions. Fixtures use packaged Admin warehouse
  generation, disposable credentials, and isolated authority/operator roots.

| Protecting check | Baseline | Final |
|---|---|---|
| Signed-out scalar save denied without writing | FAIL | PASS |
| Missing capability denied without writing | FAIL | PASS |
| Actual Settings Save Value handler persists and reloads | PASS | PASS |
| Stale captured Settings target cannot save either warehouse | FAIL | PASS |
| Invalid type cannot write | PASS | PASS |
| Identity keys cannot be renamed | PASS | PASS |
| Reading a missing optional header does not repair it | FAIL | PASS |
| Unknown user column survives reading | PASS | PASS |
| Production UOM publication persists a real version increment | PASS | PASS |
| Production cannot write arbitrary settings | FAIL | PASS |
| Actual Production Retrieve handler publishes and removes staging | PASS | PASS |
| Denied Retrieve preserves staging and authority bytes | FAIL | PASS |
| Read-only workbook rejects saving | PASS | PASS |
| Dirty workbook and unsaved user edit remain untouched | FAIL | PASS |
| Closed-workbook read preserves file bytes | PASS | PASS |
| A different required header missing prevents all writes | Later RED | PASS |
| Missing required identity header is not repaired by reading | FAIL | PASS |
| Missing identity header prevents command persistence | FAIL | PASS |

## Package and regression evidence

- Candidate directory: `deploy/validation-config-commands`. The final five XLAMs
  were rebuilt directly in `deploy/current`; its `addins-manifest.json` records
  the distributable hashes. Comparison of all 151 packaged VBA component code
  hashes found only the final four-line full-schema validation guard in
  `modConfigCommands` differs from the long-regression candidate; all other
  150 components match, including Operations and its Production forms.
- A cold-start packaging check exposed a copied Operations package resolving
  Core from the candidate directory. That was packaging RED, distinct from D5
  behavior. Rebuilding in the final directory produces cold-start GREEN.
  The compile validator now opens Operations before preloading Core, so a stale
  reference cannot be hidden by an already-loaded correct Core project.
- Explicit VBE Compile commands: **5/5 PASS**, on candidates and final packages.
  InvSys project references resolve inside their tested package directory.
  The focused **18/18 GREEN** and visible Settings capture were repeated on the
  final build. Screenshot capture uses window rendering, with a screen fallback.
- Packaged smoke: **81/81 PASS**. Ordered Release 1 chain: **30/30 PASS**,
  including fresh generation/seed, Receiving, Production, Boxing, Shipping,
  exact identity, restart/reconciliation, replay, locks and five-package runtime.
  Both were repeated on the final build after the required-schema guard.
  One smoke attempt lost Excel's COM connection at InventoryZeroCreate and
  produced eight dependent failures; a fresh serial run passed 81/81 without
  code changes. Its cause is unconfirmed; keep Excel validators serial.
- Live-role regression: **48/48 PASS** after correcting one superseded fixture
  expectation. The old fixture expected `LoadConfig` to create missing Config;
  it now asserts non-creating failure, then explicitly provisions the surface.
  No business-role assertion was removed.
- Production layout: **PASS**, three sizes across five tested pages, no overlaps
  or out-of-bounds controls, plus minimize/restore/maximize/restore.
- Viewer on the final build: **PASS** for displayed-list export, readable
  events, internal-reservation exclusion, refresh, remembered date filters,
  read-only behavior and byte-for-byte snapshot preservation.
- Source contract checks: UOM 7/7, Batch Note/Viewer 5/5, Production layout 8/8,
  R1 controls 6/6, launcher contracts 24/24, full-chain contract 13/13,
  variable quantity modes 6/6.
- Reusable Production on the source-equivalent promoted packages: **2/2 aggregate checks PASS**,
  including its worksheet/lifecycle/run scenarios and a clean Excel restart
  that reuses the saved operator workbook and exact persisted released Recipe.
  All-role launcher callbacks: **3/3 PASS** for Receiving, Production and Shipping
  with no eligible workbook initially open; they create/reuse the saved local
  role workbook and modeless form.

Full project compilation exposed pre-existing reference defects that smoke
execution did not reach: three Shipping calls now use the existing
`DescribeLocalStagedInboxRows`; Production's inventory lookup uses its declared
`loInv` parameter; Production/Shipping event creators reference the existing
Core bridge constants for unchanged `PROD_COMPLETE`/`SHIP` events. These are
compile repairs, not identity, schema, or workflow amendments. The Create
Warehouse source harness imports the newly required command module.

## Maintenance and visible evidence

Regenerated `reports/static-baseline`: 1,077 maintenance candidates, 192 duplicate
groups, 45 unresolved dynamic calls and 8 literal Application.Run targets, all
unchanged from the pre-change scan. There are still 28 oversized module ratchets.
`modConfig` shrank from 1,683 to 1,615 lines; the new command module is 219 lines,
below the 1,000-line module / 200-line procedure limits. No scanner candidate
was automatically deleted and no metric exception was added.

The local, ignored `reports/runtime/config-commands/settings-save.png` was
visually inspected: the actual Settings form shows BatchSize 601 and
"Configuration saved." Production layout screenshots were also generated.
These are operator evidence from disposable fixtures; fresh human acceptance
of this correction and comprehensive Event Viewer acceptance are not claimed.
Raw runtime reports/screenshots remain ignored; this record contains no
credentials, operational rows, machine paths or generated event identities.

## Reproduction

Run with Excel closed, using `powershell -NoProfile -ExecutionPolicy Bypass -File`:

- `tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/current -Phase GREEN -CaptureEvidence`
- `tests/tooling/Test-PackagedVbaCompile.ps1 -DeployRoot deploy/current`
- `tools/validate_phase6_packaged_xlams.ps1 -DeployRoot deploy/current`
- `tools/validate_phase6_live_role_workflows.ps1 -DeployRoot deploy/current`
- `tools/validate_release1_full_chain.ps1 -DeployRoot deploy/current`
- `tools/validate_slice9_production_layout.ps1 -RepoRoot . -DeployRoot deploy/current -OutputDirectory reports/runtime/config-commands/production-layout -ResultPath reports/runtime/config-commands/production-layout.md`
- `tools/validate_inventory_viewer.ps1 -DeployRoot deploy/current`
- `tools/validate_plan022_packaged_launchers.ps1 -DeployRoot deploy/current -WorkbookState ProductionReusable -CallbackFilter Production -OutputDirectory reports/runtime/config-commands/reusable-production`
- `tools/validate_plan022_packaged_launchers.ps1 -DeployRoot deploy/current -WorkbookState NoEligible -OutputDirectory reports/runtime/config-commands/launchers`

The historical baseline comparison requires the pre-change packages from
`b108313`; running RED against current packages is not baseline evidence.
