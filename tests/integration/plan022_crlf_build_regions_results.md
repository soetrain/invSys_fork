# Plan 022: CRLF test-only import regions

Last verified: 2026-09-24 UTC. This developer-tooling correction preserves the
existing runtime exclusion of explicitly marked test-only code. It changes no
operator control, business authority or Architecture v4.11 contract.

The layout-candidate build comparison exposed retained `mProduction` test helpers
when source came from a CRLF archive. The builder's marker-count expressions
matched LF boundaries but not CRLF boundaries; a zero count returned the source
before stripping. The stripping expression itself already supported CRLF.

`Test-VbaImportTestRegions.ps1` extracts and invokes the actual
`New-NormalizedImportFile` and `Remove-VbaTestOnlyRegions` functions. It tests LF,
CRLF, mixed endings, indentation, multiple regions, a terminal marker at EOF,
malformed/nested regions, preservation of marker text inside a literal, and
equivalence of real Production imports under LF/CRLF.

- RED: **4 PASS / 8 expected FAIL**, root
  `reports/runtime/vba-import-test-regions/15abeba769ee4babbb9f47555f68e466`.
- GREEN: **12/12**, retaining every RED identity, root
  `reports/runtime/vba-import-test-regions/a67aec00190f4010b7721c51c2bdf75b`.
- Correction: permit the optional CR in the two marker-count expressions and
  the remaining-marker rejection check. Runtime text outside marked regions is
  preserved. No broad source rewrite or test-only code import is accepted.

An isolated CRLF-source control builds all five XLAMs at
`deploy/validation-guide-layout-crlf-control` and passes explicit compile and
Operations cold-start dependency checks. All **243 compiled components** match
the validated `validation-guide-layout-normalized` candidate, including separate
string-literal hashes: **zero changed components**. Reports are
`reports/runtime/crlf-build-compiled.json` and `crlf-build-component-comparison.json`.
The control is a build proof, not a replacement acceptance candidate. The validated
layout candidate and frozen predecessor retain their package bytes and earlier
40/40 and 32/48/15 evidence; no runtime change requires repeating those tests.

Maintenance evidence: `reports/runtime/crlf-build-static`. Metrics remain 250
components, 6,045 procedures, 132,897 lines, 9 literal/45 unresolved dynamic calls,
193 duplicate groups and all 28 size limits. Generated fixtures/reports remain
ignored; unrelated user changes are preserved.
All three maintenance JSON schemas and all 263 PowerShell files under tools/tests
validate. Both candidate package sets and the original runtime pins are preserved
(only the already accepted four-line Action Path layout correction differs).

## Separately discovered Production boundary failure

The existing `Test-Slice8ProductionRetirement.ps1` reports **13 PASS / one FAIL**:
`Production.Bridges.PrimitiveJsonOnly`. The same result is reproduced against
pre-builder-fix commit **af6162c**, and the frozen compiled Production module
contains the same call. This is a pre-existing D12 failure, not a CRLF or layout
regression. The tracked historical 14/14 report is restored byte-for-byte rather
than overwritten with an unsanitized runtime result.

`mProduction.CompleteProductionRunAfterCheckInForOutput` passes `wsProd.Parent`
directly to Core `modUiQuiet.BeginQuietUi`. D12 requires the existing primitive
`modOperationsPrimitiveBridge.BeginQuietUiForWorkbook` boundary instead. A passing
behavioral chain does not waive this rule. Before correction, protect the actual
packaged completion/form-action path with a focused boundary check; preserve
captured-workbook binding and the existing quiet-state restoration behavior.
No Production implementation change is included in this tooling correction.
