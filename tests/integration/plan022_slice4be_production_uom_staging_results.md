# Slice 4be UOM staging preservation

Last verified:2026-09-29 UTC. **Focused GREEN; Release1 acceptance incomplete.**

**Follow-up extent defect:** The later catalog15 candidate
`validation-production-uom-activity` passes84 checks but fails reopening a draft
with an internal blank row. Cells remain unchanged while `CurrentRegion` truncates
table membership above the gap. This exposes a missing preservation assertion,
not a regression in the preceding78 checks. Expanded tests reach **96 PASS/14
behavioral FAIL**, including save/reopen and conflicting, off-sheet or broken
extent-marker refusal. One intervening fixture attempt hashes an open workbook
and fails with a file-sharing error; that harness failure is excluded from RED.
The checksum now runs after closing the disposable workbook.

Architecture/Plan/controls commit `8c3c87d` clarifies the already-approved complete
draft preservation: retain an owned hidden worksheet-local range name at successful
Retrieve, validate ownership and headers before reuse/publication, and never infer
a shorter extent from a blank-row gap. Metadata is staging only; it neither saves
the operator workbook nor becomes catalog authority. Runtime correction is in
isolated `validation-production-uom-extent`; five builds/compiles and cold loading
pass, with exactly one of249 compiled components changed relative to the activity
candidate. The separate visible gate advances96/14 to **110/110 GREEN**, preserving
all110 exact identities and every preceding78 GREEN. Five instrumented compiles,
normal unassisted closure, package/settings preservation and zero Excel Application
events in its window are verified. The draft and complete reuse-status images are
reviewed; images remain ignored. This does not close the separate combined-suite
closed-workbook automation failure or broader release gates.

Current static evidence:256 components/6073 procedures/133570 lines;9 literal and45
unresolved dynamic calls,191 duplicate groups, no growth in28 existing module caps.
All three schemas validate. Final regression acceptance remains pending.

Extent RED controller/result: `production-uom-staging-controller/8729c533bd5e4d4ea579b509aa3468aa`;
`slice4be-production-uom-staging/dc8af7409bf440fc847320b967655ac0/red.json`.
Build/source comparison: `production-uom-extent-build`. All paths are under ignored
`reports/runtime/`. [UOM observation evidence](plan022_slice4be_production_uom_activity_results.md)
records the separate catalog15 work and its path-run qualification.
Visible GREEN controller/result: `production-uom-staging-controller/878a1039734f4bf1a6c77db392943191`;
`slice4be-production-uom-staging/52213a197ea1430596d893549d058578/green.json`
(20:51:43--20:53:54 UTC). Exact check comparison:
`production-uom-extent-focused-verification.json`; static: `production-uom-extent-static`.

The following78-check checkpoint describes the preceding preservation candidate:
The approved correction in `validation-production-uom-staging` advances the
expanded actual-handler gate from **60 PASS/18 FAIL to78/78**, preserving all54
original identities and every earlier GREEN. A separate visible run also passes
78/78. Both have five instrumented compiles, normal unassisted closure, zero Excel
Application errors, and preserved settings/five package hashes.

`modProductionUomCatalog` now reuses a valid table without writing its cells,
reopens a uniquely identifiable unlisted region, rejects missing/duplicate managed
headers and nonempty unidentifiable staging, and projects normalized managed
columns into the existing seven-field Core publication envelope. Extra columns
stay local. The displayed reuse status states that edits were retained and the
saved catalog was not reloaded. Core authorization and publication are unchanged.
No new tracking control is registered by this correction.

Five package builds, five compile checks and cold Operations dependency loading
pass. Among248 compiled components, only Operations `modProductionUomCatalog`
differs from the preceding candidate (exact code/case-insensitive/literal hashes
compared). Static evidence:255 components/6068 procedures/133432 lines;9 literal
and45 unresolved dynamic calls,191 duplicate groups, all28 existing size caps
unchanged. Three schemas and304 PowerShell files under tools/tests/tooling parse.
The Production form and other components are unchanged by all three compiled-source
hash comparisons; this is not a claim of identical XLAM containers.

Two UOM images were reviewed: the retained worksheet draft and complete Production
Settings form with the new status. The additional Settings fixture image was
reviewed separately. Packaged layout geometry/native actions pass; minimum and
expanded images are complete, and the default image is byte-identical to minimum.
These are automated visual observations, not user acceptance.
Packaged smoke passes86/86 with settings/packages and tracked report restored;
normal shutdown remains unproven because that validator can force termination
without recording whether it did so.

**Full-chain gate failed:**5 PASS/1 harness failure; live-role subprocess40 PASS/1
harness failure; Create Warehouse15/15. The ordered run passed projection recovery
and the two-batch Production form action, then crashed at **Run Shipping
BtnShipmentsSent**. Windows recorded Excel/ntdll exceptionc0000028, offset
0000000000012d2f at20:02:03 UTC; transport reported0x800706BE. The same native
signature appeared earlier at other boundaries, but a common cause is unproven.
The `/restore` recovery process required assisted termination before the controller
restored settings, candidate hashes and tracked reports. This candidate is not
promoted; no unchanged retry or native workaround is claimed. Desktop monitoring
had no error5 in this run.

## Original discovery

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
values. The user approved that workflow choice on2026-09-29; Architecture v4.11,
Plan022 and controls record the approval. The preservation requirement is
already binding and is not waived while that choice is reviewed.

Approved behavior: create from the saved catalog only for a new empty workbench;
otherwise reopen the existing staging table without rewriting its cells. After
successful Retrieve unlists the table, retain its cells and reopen its identifiable
managed region on the next Edit. Resolve the required header subset by normalized
name; unknown columns/formulas/order and unrelated worksheet cells remain local
and unchanged. Refuse ambiguous ownership/header shapes before mutation. Do not
claim a reused draft reflects later saved-catalog changes. Existing retrieval
authorization and Core Config authority remain unchanged.

The expanded tests protect reordered and case/space-normalized headers, retrieval
with extra columns, retained formulas and draft edits, reopened unlisted workbenches,
missing/duplicate managed headers and an unidentifiable populated worksheet.
Two intervening fixture attempts failed on PowerShell COM Boolean/array assignment;
those were corrected before runtime implementation and are not behavioral RED.
No new tracking ID, coverage completion or Release1 acceptance is claimed here.

## Evidence

- Approved Architecture/Plan/controls checkpoint: docs `f4ce104`.
- Expanded RED controller/result: `production-uom-staging-controller/76b19103c9794816b43f204b5688569c`; `slice4be-production-uom-staging/6ed8ef7a69694c668401a63118adccad/red.json` (19:55:04--19:56:26 UTC).
- GREEN controller/result: `production-uom-staging-controller/f9ac8a1bae0a4dcbae4504059f12612c`; `slice4be-production-uom-staging/f08f526581ba4cfca2c4bbf959a2d1ff/green.json` (19:58:03--19:59:30 UTC).
- Visible controller/result: `production-uom-staging-controller/990e279510a547898ec1069c889089a5`; `slice4be-production-uom-staging/51a52a5cb36447658c97e38921ce667d` (20:03:45--20:05:17 UTC).
- Build/source comparison: `production-uom-staging-build`; focused verification: `production-uom-staging-focused-verification.json`; maintenance: `production-uom-staging-static`.
- Failed chain: `production-uom-staging-regression/chain-44946306b98c4b7aa6c90369ce9d4721` (19:59:49--20:03:23 UTC).
- Layout: `production-uom-staging-regression/layout-7bbb935b93b142d0a6dacad4f630f335` (20:05:30--20:05:44 UTC).
- Smoke: `production-uom-staging-regression/smoke-02ceb6b573a44a65b0a9adeae8cf6dff` (20:06:17--20:06:36 UTC).
- All runtime paths above are beneath ignored `reports/runtime/`; images and runtime values are not committed.
- Test controller: `tests/tooling/Test-Slice4beProductionUomStaging.ps1`.
- Actual-handler adapters/assertions: `tests/tooling/Slice4beProductionUomStaging.ps1`.
- Valid RED controller: `reports/runtime/production-uom-staging-controller/0001fc7fb1d1444080e0b395d4135951` (19:37:03--19:38:14 UTC).
- Valid RED result: `reports/runtime/slice4be-production-uom-staging/1423544f53da42d68aa4afcca4e2f1a1/red.json`.
- Preservation/event receipt: `reports/runtime/production-uom-staging-red-verification.json`.
- Rejected fixture attempt: `reports/runtime/production-uom-staging-controller/e55e7d7648c641eabd4de2667090fe2a`.
