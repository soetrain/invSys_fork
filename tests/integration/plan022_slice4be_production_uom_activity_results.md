# Slice 4be UOM Edit observations

Last verified: 2026-09-29 UTC. **Focused observation GREEN; broader acceptance open.**

Architecture v4.11 D18 catalog15 registers the existing **Edit UOM Catalog on
Sheet** handler as `PRODUCTION_UOM_EDIT`, owned by `PRODUCTION_UOM_STAGING`.
The typed Operations adapter rechecks current captured context and existing
PROD_POST permission, then records REQUESTED and the owning staging outcome.
OPENED/REUSED describe local workbench completion with saved authority unchanged;
they never assert catalog publication or Domain application. Fixed observations
contain no catalog values, formulas, notes or workbook paths. Catalogs1-14 remain.

## Focused D13 evidence

The actual packaged Send handler records **98 PASS / 134 behavioral FAIL** before
implementation and **232/232 GREEN** on `validation-production-uom-activity`.
Check identities match exactly, including all preceding78 staging checks. The gate
covers owner metadata, historical catalog definitions, OPENED/REUSED/reopened,
REJECTED/DENIED/FAILED, tracking disabled/unavailable, current-context refusal,
closed workbooks, loading/re-entry suppression, redaction and exact completion
classification. All five instrumented projects compile; closure is unassisted and
settings and frozen packages are preserved.

An earlier101/131 RED used an overly broad refusal-message oracle; it was tightened
before implementation. The98/134 result is the protecting baseline.

Five candidate builds/compiles and cold Operations dependency loading pass.
Among249 compiled components, six intentionally change relative to the preceding
UOM preservation candidate: Core activity catalog/completion map/new UOM codes,
Operations UOM action/staging owner and Production form. Compare exact source,
case-insensitive source and literal hashes, not XLAM container bytes.

## Action Path evidence and qualification

Independent guide-source and observed recordings pass **84/84** through actual
Admin publication, Viewer selection, guide authoring, reader pairing, evaluation
and How-To/Diagnostic/Compare both views. Exact original activity/journal identities
are preserved. OPENED then REUSED concludes CommandCompleted; empty source-event
references leave SourceEventsApplied incomplete. Saved authority, workbook bytes,
custom columns and immutable evidence remain unchanged.

This first path run required UI assistance confirming a verified disposable-sheet
delete dialog (`NUIDialog`). It is not unattended acceptance. The shared fixture
now suppresses alerts only around its validated disposable reset and restores the
prior value; an unassisted rerun is required. Six principal images were reviewed:
Production Settings, Event Detail, How-To, Diagnostic, Compare both and conclusion.
The first Event Detail image selects REQUESTED while showing the paired REUSED line;
the corrected fixture will select the terminal line for capture. Images/runtime
values remain ignored. These observations are not user acceptance.

Static evidence:256 components/6071 procedures/133522 lines;9 literal and45
unresolved dynamic calls,191 duplicate groups. None of28 existing oversized-module
caps grows; Production form shrinks11739 to11736 lines. Three schemas validate and
305 PowerShell files under tools/tests/tooling parse without errors.

## Remaining gates

The first combined visible run on `validation-production-uom-extent` stops at
**255 PASS/1 harness failure**, before completing the closed-workbook guard and
optional-tracking tail. Of the110 preservation-suite identities,107 have passed;
three common trailing checks have not run. This is not a completed GREEN gate.
A VBA dialog reports80010007 at
`TestProductionDesigner.SendUom`, the test adapter call into the form. Diagnostic
Debug/Reset intervention localizes the boundary but changes the run. Windows then
records Excel/oleaut32 nativec0000005 at20:50:17 UTC, offset0000000000033b41; recovery
termination is needed. The timing does not prove a common cause with the earlier
c0000028 failure or identify a product/harness root cause. This is not desktop
Win32 error5. Settings and packages are restored. Do not hide the failure by
reporting only earlier passing checks or treat assisted closure as acceptance.

Failed visible controller/result:
`production-uom-staging-controller/91119955725d4554911de2eea19aae16`;
`slice4be-production-uom-staging/d29842ba571541f788c60f04fe74d756/green.json`.
The controller contains sanitized dialog-location, reset and cleanup receipts.
Separate visible staging validation now completes110/110 with exact identities,
normal closure and preservation; a controlled comparison with capture disabled
is in progress. The larger visible guard failure remains open.

The additional blank-row preservation test exposes a separate defect: Retrieve
retains cells but Edit recreates a shortened table above the blank row. Initial
expanded result84/1 preserves all earlier78 GREEN checks. The full-extent correction
under the approved reuse decision reaches separate visible110/110 on
`validation-production-uom-extent`, following expanded RED96/14.
See [staging evidence](plan022_slice4be_production_uom_staging_results.md).

Applicable instruction, draft/lifecycle, Settings and native regressions must run
on the final candidate. The preceding candidate's full-chain/native c0000028
blocker remains unresolved. No package promotion or Release1 acceptance is claimed.

## Traceability

All runtime paths below are beneath ignored `reports/runtime/`.

- Contract before implementation: documentation commit `13291eb`.
- Final RED controller/result: `production-uom-staging-controller/7875e27bf99d4e5d999402087ba81eca`; `slice4be-production-uom-staging/e27a45e839444857b1e644c7d36a0026/red.json`.
- GREEN controller/result: `production-uom-staging-controller/95f6458f3b984b89b2a56a38d4e3f83f`; `slice4be-production-uom-staging/93bffe93d5c7432c8e109942aea1de3b/green.json`.
- Exact check verification: `production-uom-activity-focused-verification.json`.
- Build/source: `production-uom-activity-build`; maintenance: `production-uom-activity-static`.
- Assisted paths controller/result: `production-uom-staging-controller/92807e8dd2ab4cf0b095c64d11b73411`; `slice4be-production-uom-paths/43467d5b5ca94624abc53e54a57956f2/green.json`.
- Blank-row RED controller/result: `production-uom-staging-controller/bf9c0604d0a543a28d8977e4085d1a42`; `slice4be-production-uom-staging/fe3a8086b9394ff49f11a1e040f570a4/red.json`.
- Tests: `Test-Slice4beProductionUomStaging.ps1`, `Slice4beProductionUomActivity.ps1`, shared `Slice4beProductionInstructionPaths.ps1` under `tests/tooling`.
