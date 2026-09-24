# Plan 022: Production completion primitive boundary

Last verified: 2026-09-24 UTC. This discovered Release 1 blocker belongs to
Slice 4be preservation work. It restores the existing D12 primitive cross-XLAM
boundary; it introduces no new business or operator contract.

`mProduction.CompleteProductionRunAfterCheckInForOutput` passed a Workbook to
Core `modUiQuiet.BeginQuietUi`. The pre-existing violation and predecessor
13/14 static result are recorded in [the build evidence](plan022_crlf_build_regions_results.md).
The correction passes the captured `wsProd.Parent.Name` through the existing
`modOperationsPrimitiveBridge.BeginQuietUiForWorkbook`. The owning Core bridge
resolves the open workbook. Completion, cleanup and Domain processing retain
their existing implementations.

## Focused packaged RED/GREEN

`Test-ProductionQuietBoundary.ps1` uses disposable copies of the five packages
and the real two-batch `frmProduction` action adapter, including
`mBtnManagerApplyOutput_Click`, with another workbook activated. The probe arms
immediately before the completion quiet bracket and observes the actual Core
primitive entry; earlier form/workbook initialization cannot satisfy it. No
handler or service is replaced. Four instrumented dependencies compile before
the form runs. Probe arguments stay in memory; exported facts are counts.

- RED: **50 PASS / two expected FAIL**, root
  `reports/runtime/production-quiet-boundary/0702a54c9876423eb237d55c8a274b95`.
  Interval 01:57:19--01:59:35 UTC. Two completions and two UI restorations pass;
  primitive-entry and captured-name-through-bridge checks each observe zero.
- GREEN: **52/52**, root
  `reports/runtime/production-quiet-boundary/6a5322ae32934d32bad65968ad80101e`.
  Interval 02:01:39--02:04:06 UTC. Both batches enter the bridge exactly once,
  use the captured workbook name and restore ScreenUpdating, EnableEvents,
  DisplayAlerts, DisplayStatusBar, Calculation and quiet-state activity.
- Every RED identity is retained; all 48 base live-role checks pass in both runs.
  The ordered full chain replaces two base Shipping checks with two Boxing checks;
  its exact 48-identity set is verified separately, not conflated with this set.
- Both runs close Excel normally without assistance and preserve candidate
  package bytes, local settings and the tracked report. The 02:06:22 UTC audit
  finds zero Application failure events. Receipt:
  `reports/runtime/production-quiet-focused-verification.json`.

Reproduce with `Test-ProductionQuietBoundary.ps1 -DeployRoot
deploy/validation-production-quiet-boundary -Phase GREEN`. Instrumentation is
confined to disposable package memory; no test module enters the runtime source.

## Candidate and preservation gates

Candidate: `deploy/validation-production-quiet-boundary`. All five packages build
and explicitly compile; Operations cold-start dependencies resolve inside this
candidate. Comparison with the preceding validated layout candidate covers all
**243 compiled components** with separate string-literal hashes: **only
`invSys.Operations.xlam|mProduction` changes**. Reports:
`reports/runtime/production-quiet-compiled.json` and
`production-quiet-component-comparison.json`. The change is one call site.

The original Production retirement audit now passes **14/14**. Static evidence
at `reports/runtime/production-quiet-static` retains 250 components, 6,045
procedures, 132,897 lines, 9 literal/45 unresolved dynamic calls, 193 duplicate
groups and all 28 size limits. All 265 PowerShell files under tools/tests parse.
All three maintenance JSON reports validate against their schemas.

The ordered full Release 1 chain passes **32/32**, ordered live roles **48/48**
and Create Warehouse **15/15**, retaining every preceding chain identity.
Interval: 02:04:19--02:10:26 UTC. Excel closes normally without assistance;
local settings, all three tracked reports, five package hashes and 263 tooling
pins are preserved. The 02:10:41 UTC audit finds zero Application failure events.
Receipt: `reports/runtime/production-quiet-chain-verification.json`, with the
three copied `production-quiet-chain-*.md` reports. The 263 pins cover top-level
tooling/scripts; the 265-file parse includes nested tests.

## Unresolved reusable-Production layout gate

The separate existing `ProductionReusable -ProductionRunOnly` gate is **0 PASS /
one aggregate FAIL** on both the corrected candidate and its preceding
`validation-guide-layout-normalized` candidate. Both report
`EightPaletteRows=False`; all **67 Boolean observations match**. The ordinary
two-batch reusable and Chai fork/convergence action envelopes report OK, including
exact-key consumption, distinct output keys, completed processes and retained
run identity. Their success does not waive the failing palette assertion.

- Corrected candidate: `reports/runtime/production-quiet-reusable`,
  02:10:42--02:14:59 UTC.
- Preceding candidate control: `reports/runtime/production-quiet-reusable-baseline`,
  02:16:58--02:21:20 UTC.
- Comparison: `reports/runtime/production-quiet-reusable-comparison.json`.
  Both runs restore settings and close Excel. The existing launcher harness can
  terminate its owned Excel process after Quit; these are not strict normal
  shutdown proofs. The separate full-chain normal closure remains valid.

This is a pre-existing acceptance failure, not a boundary-correction regression.
The accepted `EightPaletteRows=True` record in Plan 022/controls remains binding;
do not weaken the test. Before another broad run, measure the packaged palette at
default/minimum, expanded and restored sizes and protect its declared height and
non-overlap. A source-only hypothesis is `frmProduction.AddList` assigning Height
before disabling IntegralHeight, allowing MSForms to round the requested height.
That cause is **not confirmed** and no palette implementation change is included.
The subsequent [focused palette correction](plan022_production_palette_results.md)
confirms setup-order rounding, restores the declared height and records the
broader reusable GREEN on its own candidate; the preceding failed attempts remain.

All 299 original runtime pins remain accounted for: only the prior Action Path
layout guard and this mProduction correction differ. All three five-package sets
retain their bytes; unrelated user changes remain outside this commit.

No visible operator acceptance is claimed. The unchanged Action Path form retains
its preceding four-size evidence; comprehensive Viewer/Action Path coverage,
guide transfer, pending contract decisions, human comparison and applicable NAS
acceptance remain open in the [Slice 4be checklist](plan022_slice4be_remaining_acceptance.md).
