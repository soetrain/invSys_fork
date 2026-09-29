# Slice 4be UOM staging preservation

Last verified:2026-09-29 UTC. **Behavioral RED; no runtime correction yet.**
This is a newly discovered blocker while accounting for Production's untracked
`mBtnUomCatalogSend_Click` control. Architecture v4.11 header-extension rules
require normalized managed-header lookup and preservation of unknown local columns.
The existing UOM worksheet is explicitly a staging surface, not Config authority.

The focused test invokes the real packaged form handler twice on a disposable
captured workbook, with a different workbook active. Between calls it inserts
two custom table columns, a formula and an unrelated worksheet note. The frozen
`validation-production-instructions-typed` candidate records **49 PASS/5 FAIL**:

- `UomStaging.Repeat.UnknownColumnsRetained`
- `UomStaging.Repeat.UnknownValuesRetained`
- `UomStaging.Repeat.UnknownFormulaRetained`
- `UomStaging.Repeat.UnknownColumnOrderRetained`
- `UomStaging.Repeat.UnrelatedWorksheetCellRetained`

All five fail after the handler reports success. Initial creation, managed row
count, captured-workbook binding, Config bytes and absence of implicit workbook
save pass. Source `modProductionUomCatalog.SendUomCatalogToWorksheet` unlists the
existing table and calls `ws.Cells.Clear`, explaining the observed loss.
Five instrumented projects compile; Excel exits normally without assistance,
settings/all five frozen packages are preserved, and zero Excel Application
events are recorded. These failures are independent of the cold Production crash.

The first attempt records39 PASS/1 harness failure because a saved test workbook
was hashed while still open. Closing, hashing and reopening corrects that fixture;
it is not a product RED. The54-check result above is the valid behavioral baseline.

## Decision before implementation

The existing caption says **Edit UOM Catalog on Sheet**, while the implementation
replaces the staging data on every call. Reopening an existing draft would preserve
custom fields and unsaved edits, but would stop implicitly reloading saved catalog
values. That workflow choice is proposed in Architecture v4.11 and synchronized
with Plan022/controls; user approval is pending. The preservation requirement is
already binding and is not waived while that choice is reviewed.

Proposed behavior: create from the saved catalog only for a new empty workbench;
otherwise reopen the existing staging table without rewriting its cells. After
successful Retrieve unlists the table, retain its cells and reopen its identifiable
managed region on the next Edit. Resolve the required header subset by normalized
name; unknown columns/formulas/order and unrelated worksheet cells remain local
and unchanged. Refuse ambiguous ownership/header shapes before mutation. Do not
claim a reused draft reflects later saved-catalog changes. Existing retrieval
authorization and Core Config authority remain unchanged.

Before implementation, expand the same packaged tests for reordered headers,
retrieval with extra columns, reopened unlisted workbenches and ambiguous shapes.
Do not claim a new tracking ID, coverage completion or Release1 acceptance here.

## Evidence

- Test controller: `tests/tooling/Test-Slice4beProductionUomStaging.ps1`.
- Actual-handler adapters/assertions: `tests/tooling/Slice4beProductionUomStaging.ps1`.
- Valid RED controller: `reports/runtime/production-uom-staging-controller/0001fc7fb1d1444080e0b395d4135951` (19:37:03--19:38:14 UTC).
- Valid RED result: `reports/runtime/slice4be-production-uom-staging/1423544f53da42d68aa4afcca4e2f1a1/red.json`.
- Preservation/event receipt: `reports/runtime/production-uom-staging-red-verification.json`.
- Rejected fixture attempt: `reports/runtime/production-uom-staging-controller/e55e7d7648c641eabd4de2667090fe2a`.
