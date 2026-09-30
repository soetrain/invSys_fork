# Slice 4be Production output-regulation observations

Last verified:2026-09-30 UTC. Incomplete; no operational promotion or human
acceptance. Architecture v4.11 D15/D18 specifies catalog20 under approved semantic
inheritance. Documentation commit3d7d603 precedes tests and runtime changes;
Plan022, controls1.305 and Production coverage1.48 carry the same contract.

The bounded group is Apply Regulation and Clear Override in Production Settings.
Both observe existing local Process-default/Recipe-override staging. Only STAGED
can establish CommandCompleted; no Save/Release or Domain application is implied.
Preserve existing conversion, validation, clear, refresh and status semantics.

## Packaged test-first RED

`Test-Slice4beProductionRegulation.ps1 -DeployRoot deploy/validation-production-design-reads-final -Phase RED`
invokes the actual `mBtnOutputRegulationApply_Click` and
`mBtnOutputRegulationClear_Click` through unsaved typed adapters. Its real released
Process fixture is created before saved-authority pins are captured. Both local
scopes, exact Recipe binding, unrelated outputs/nodes and unknown workbook columns
are protected. Test adapters return fixed failure codes rather than raw errors.

RED:168 PASS/528 FAIL across696 unique checks,
10:19:20.1996933--10:21:49.2485182 UTC. Controller:
`reports/runtime/production-regulation-controller/6bb456b66ff849ecaacc7c133ad354f2`.
Result: `reports/runtime/slice4be-production-regulation/5bbfd42a206d44b3834e101ed60d990b/red.json`.
Five instrumented compiles pass; there are no harness failures. Settings and
package pins restore, runtime source matchesa2aba6a, Excel closes unassisted and
the delayed Application1000/1001/1002 audit finds zero Excel failures.

Every selected local effect, other-output/node/scope preservation, exact released
binding, refresh/editor reset and numeric/disabled-bound semantic check passes.
Existing blank/nonnumeric conversion exceptions are observed without mutation;
their new handled/fixed failure observations remain RED. Partial staging is not
rolled back. Saved-authority, workbook-byte and prior-activity pins pass.

Failures cover the absent catalog20/observations, current permission and captured
context guards, loading/reentry suppression, fixed exception handling and visible
optional-tracking failure. The catalog-preservation assertions are RED because
catalog20 is absent; they require all103 prior definitions to remain exact when
that catalog exists. Missing compilation, a broken fixture or an unavailable
workbook is not counted as this behavioral RED.

## Remaining work

Implement the specified observations and guards while preserving the tested local
algorithms; retain all696 identities for focused GREEN. Build a separate frozen
candidate, then complete paired paths, visible review and required regressions.
Catalog19 remains unchanged and unpromoted. Its ordinary full reusable native
failure remains unresolved; see
[designer read evidence](plan022_slice4be_production_design_read_results.md).
