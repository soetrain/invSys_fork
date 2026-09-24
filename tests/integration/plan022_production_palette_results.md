# Plan 022: preserve Production ListBox heights

Last verified: 2026-09-24 UTC. Slice 4be preservation work; the accepted
Production eight-row palette requirement remains unchanged. This corrects
MSForms setup order, not the requirement or test threshold.

The preceding [Production boundary evidence](plan022_production_quiet_boundary_results.md)
records the same EightPaletteRows=False aggregate failure on both candidates,
with all 67 Boolean observations matching. Completion/Chai success did not waive
that failure.

## Focused packaged RED/GREEN

`Test-ProductionPaletteLayout.ps1` copies five packages into a disposable root,
compiles four instrumented dependencies, and runs the real public Production
launcher twice to retain its captured workbook/reuse gate. Its existing typed
form-layout adapter additionally measures minimum, default, expanded and restored
sizes. The form's geometry checker inspects every page for bounds and interactive
overlap. A temporary control calls the actual private AddList factory to measure
its requested height and a reassignment after IntegralHeight is already disabled.
No business handler is replaced and no probe enters runtime source.

- RED: **7 PASS/four expected FAIL**, root
  `reports/runtime/production-palette/7ba53b96cbc74e6db20d70c10e9821ba`,
  02:29:18--02:29:40 UTC. A requested 96 points returns **88.45** at
  minimum/default/restored sizes and from the actual factory. Expanded height
  is 111.05; all geometry checks pass. Reassigning the same height after
  IntegralHeight is disabled produces exactly **96**, isolating setup order.
- Correction: move `.IntegralHeight = False` before `.Height = heightVal` in
  `frmProduction.AddList`. Keep requested dimensions, anchors, headers and
  handlers unchanged. This honors the existing factory argument for all lists.
- GREEN: **11/11**, root
  `reports/runtime/production-palette/b36af5b4d3b94db0b734261ce93b4268`,
  02:32:11--02:32:36 UTC. Minimum/default/restored and factory heights are
  **96.000**; expanded height is **118.600**. Every RED identity remains present;
  all-page bounds/overlap and original launcher checks pass.

Both attempts restore settings and preserve candidate package bytes. Excel is
closed afterward; the existing launcher harness may terminate its owned process
after Quit, so these attempts do not establish normal unassisted shutdown.

Candidate `deploy/validation-production-palette` builds and explicitly compiles
all five packages and passes Operations cold-start dependency resolution.
Among all **243 compiled components**, only **frmProduction** differs from the
preceding boundary candidate, with string literals checked separately. See
`reports/runtime/production-palette-component-comparison.json`.

Source checks pass **7/7** for Run List layout and **8/8** for Production layout.
The latter's default path expression failed before testing; the preserved
`production-palette-layout-source.log` is a harness failure, not product RED.
Supplying RepoRoot explicitly passes; see `production-palette-layout-explicit-root.log`.

Static evidence at `reports/runtime/production-palette-static` has unchanged
250 components, 6045 procedures, 132897 lines, 9 literal/45 unresolved dynamic
calls, 193 duplicate groups and all 28 size limits. All three report schemas
validate. After the fixture correction below, all 268 PowerShell files parse.

## Separately protected fixture correction

The first broader reusable attempt stops at fixture sign-in before any Production
callback: `reports/runtime/production-palette-reusable`, 02:33:04--02:33:23 UTC.
Its setup booleans pass except SignedIn. The original failure did not record a
numeric Auth status and its private values were not retained; its precise cause
is not established and it is not product RED.

Inspection found that the shared `New-AuthWorkbook` helper wrote opaque hash
strings into General-format cells. `Test-AuthFixtureText.ps1` independently
reproduces conversion of generated leading-zero and exponent-like values through
the actual helper, both after saving and reopening: **4 PASS/eight expected FAIL**
at `reports/runtime/auth-fixture-text/e9da26c518cf430f83859e83cee0f1f2`.
Formatting the named fixture column as text before inserting rows yields
**12/12** at `auth-fixture-text/482bd86c1f9f43e18c4bb8c9b280df49`. The ordinary
hexadecimal control remains passing. Values are never reported; all three
private fixture workbooks are removed after observation, and no values or hashes
enter source or evidence. This changes test setup only, not runtime Auth or the
pending D8-A contract. The launcher now reports only an allowlisted numeric/code
failure classification, never the free-text sign-in envelope.

The broader reusable retry with corrected fixture passes **one aggregate gate**
at `reports/runtime/production-palette-reusable-text-fixture`,
02:39:22--02:43:40 UTC. All 67 Boolean observation identities are retained; the
only changed value is **EightPaletteRows=False -> True**. Two reusable batches
and Chai fork/convergence complete with the previously passing exact-key,
co-product, output, instruction, status and run-identity observations unchanged.
Receipt: `reports/runtime/production-palette-reusable-verification.json`.
Settings are restored and Excel is closed; the existing launcher cleanup limits
normal-shutdown claims as described above.

The final candidate's full chain passes **32/32**, ordered live roles **48/48**
and Create Warehouse **15/15**, with every preceding chain identity retained.
Interval: 02:44:45--02:51:13 UTC. This separate gate closes Excel normally without
assistance, restores settings and three tracked reports, and preserves five
package hashes and 266 top-level tooling pins. All 268 PowerShell parses include
nested tests. The 02:52:02 UTC audit finds zero Application failure events.
Receipt: `reports/runtime/production-palette-chain-verification.json`; the three
`production-palette-chain-*.md` copies retain the gate details. No full Slice 4be
acceptance is inferred from this bounded correction.

## Visible acceptance limit

The 02:35:20 UTC read-only desktop preflight still fails GetCursorPos with Win32
error5 (`reports/runtime/production-palette-desktop-facts.json`). No capture is
attempted. The restored height and automated EightPaletteRows predicate are not
substitutes for native visible evidence. Broader Slice 4be coverage and acceptance
remain open in the [remaining checklist](plan022_slice4be_remaining_acceptance.md).
