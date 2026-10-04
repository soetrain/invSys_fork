# Slice 4be-A / early B0: generated warehouse purpose

Last verified: 2026-10-04. D18-REPLAY-01 is approved; this is a prerequisite,
not Receiving replay proof or 4be-A acceptance.

Admin Create Warehouse now offers Operational (default) or Training. Its actual
Create handler passes the choice through the primitive Admin/Core bridge. Core
validates before provisioning and writes `WarehousePurpose` in its creation-only
Config command. Existing runtimes cannot be relabelled. Header-based writes retain
other columns; the extracted stamp keeps Core headless and shrinks bootstrap.
Six source-import harnesses include the extracted dependency.

## Evidence

- Frozen `validation-production-next-activity-01`: meaningful RED15 PASS/8 FAIL.
  Both actual creations succeed; failures isolate missing choice/persistence.
- `validation-warehouse-purpose-01`: GREEN23/23 with identical ordered checks,
  default/Training reopening, own artifacts, cancel and byte-preserved refusal.
- Layout RED31/4 on purpose01 exposes summary/input overlap at both sizes.
  `validation-warehouse-purpose-02`: GREEN35/35, retaining all23 earlier checks.
  Minimum620x610 and enlarged800x700 geometry keeps summary/footer separated and
  visible controls inside the form. This is geometry evidence, not screenshot acceptance.
- Both candidates pass five packaged cold compiles and preserve frozen packages
  and local settings. Purpose01 changes three compiled components and adds one;
  purpose02 changes only `frmCreateWarehouse` layout (285/286 unchanged).
- Purpose01 smoke86/86 and source creation15/15 pass; source creation includes exact
  System_Key round trips and custom-column preservation. Their business/package
  sources are unchanged by purpose02's layout-only delta. Smoke exits naturally;
  generated tracked reports and settings are restored. Native audit through22:31UTC
  finds zero Excel application failures during these tests/builds.
- Regenerated static evidence:293 components/6182 procedures/135671 lines;
  feature growth is one component, four procedures and66 net lines. All28 oversized
  caps hold;9 literal/45 unresolved Application.Run and190 duplicate candidates
  remain unchanged. All384 PowerShell files parse and all three evidence schemas validate.

Local receipts under `reports/runtime/`:

- `warehouse-purpose/203d6565b54b4a3cb2e2191155100011/red.json`
- `warehouse-purpose/8e06c5ef2af145dea0f49702d9970c64/green.json`
- `warehouse-purpose/81dd49034e51482c947e1b8638af3693/red.json`
- `warehouse-purpose/91cca96abd394818b2530ff9ceb6597b/green.json`
- `warehouse-purpose-focused-verification.json`, `warehouse-purpose-layout-verification.json`
- `warehouse-purpose-build-01/`, `warehouse-purpose-build-02/`
- `warehouse-purpose-static-02/ratchet-verification.json`
- `warehouse-purpose-regression/smoke-ab54f5790dd8408a847a242ac094d0f2/`
- `warehouse-purpose-regression/warehouse-090480f4d30d4e64a8f0c7186db47109/`

Reproduce with `tests/tooling/Test-Slice4beWarehousePurpose.ps1 -DeployRoot
deploy/validation-warehouse-purpose-02 -Phase GREEN -CheckLayout`.

## Open gates

Visible capture is pending: the foreground helper failed while Task Manager held
focus; UIA root/descendant focus did not resolve it and was removed. Desktop probes
report error0; no Win32 error5 is established by those capture failures. Captures
`60827a...`, `9d9f4d...`, `55d288...` are incomplete harness attempts, never product
RED/GREEN. The initial reflection argument error `ac2757...` was likewise corrected
before meaningful RED. Do not repeat those focus workarounds.

Shared execution schemas, Receiving record/author/replay/fresh proof, remaining A
coverage and visible acceptance are open. The earlier native Boxing full-chain
failure remains unresolved; these checks do not replace that gate. No deployment
or B0/A/B/R1 acceptance is claimed.
