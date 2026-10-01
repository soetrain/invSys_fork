# Slice4be Production Run local observations

Architecture v4.11 D18 specifies nine catalog24 controls under approved semantic
inheritance: Scale, Clear, Load, Loader/Manager Refresh, List/Tree Apply and Tree
Expand/Collapse. Owner is PRODUCTION_RUN_LOCAL. Local completion is distinct from
inventory application; these controls carry no source-event references. D15's
experimental Tree status and D14's exact-key/unknown-column rules remain binding.

This is test-first work, not completed implementation or acceptance. Runtime
remains catalog23/118 IDs and55/68 constructed Production buttons. The separate
full reusable native crash remains unresolved; see
`plan022_slice4be_production_assignment_results.md` for its bounded diagnostics
and the thirty verified shared gates on the frozen Assignment candidate.

## Protecting test and candidate

`tests/tooling/Test-Slice4beProductionRunPresentation.ps1` runs the isolated
ConfigCommands harness against frozen, unpromoted
`deploy/validation-production-assignment-01`. Its unsaved adapters in
`Slice4beProductionRunPresentation.ps1` call the original packaged
`mBtnRunTreeExpandAll_Click` / `mBtnRunTreeCollapseAll_Click` handlers, not replacement
handlers. They stage synthetic local palette rows without inventory writes.
The opt-in gate is independent of the existing ConfigCommands selections.

The initial baseline covers only this presentation pair: expanded/collapsed/
process-collapsed/empty trees, repeated actions, unchanged palette contents,
fixed metadata, exact paired observations, redaction/integrity/terminal semantics,
loading/busy suppression, stale target/session/sign-out/closed workbook, saved
authority, unknown local values/formula and older-record preservation. All nine
controls, the complete optional-policy/failure/yield matrix and paired Action Path
publication/reader evidence remain required before implementation acceptance.
No runtime source, package build, deployment or promotion has changed here.

## Verified presentation-pair RED

Controller `reports/runtime/production-run-presentation-controller/397ec444067d43b68e61488b307aaea5`;
result `reports/runtime/slice4be-production-run-presentation/92792238dab347198548a78248cc4a3c/red.json`.
2026-10-01 06:07:45.6116044--06:09:33.3276574 UTC: **98 PASS /134 FAIL /232 unique checks**.
The134 failures are exactly two absent catalog definitions,112 absent observation/
terminal facts, four missing loading/busy guards and16 missing captured-context
mutation/refusal guards. Every existing tree shape, palette value and status check
passes, including preserved collapsed Process groups and empty/repeated actions.
All42 prior shared checks remain GREEN in exact order; the other56 passes establish
the preserved owner/preservation behavior. No assertion is removed or relaxed.

`verification.json` records five instrumented package compiles, canonical frozen
package pins, preserved packages/settings, normal unassisted closure, no remaining
Excel and zero delayed Application1000/1001/1002 failures. This is meaningful RED
for the two actual handlers only. It neither protects all nine controls nor
completes the optional-policy/failure/yield/recording matrix. No screenshot or
human-acceptance claim is supplied by this baseline.

Static evidence `reports/runtime/run-presentation-static-01`:280 components,
6150 procedures,135006 lines,9 literal/45 unresolved dynamic calls,190 duplicate
bodies,28 non-growing module caps. Three schemas validate and356 PowerShell
scripts parse. Runtime metrics/source are unchanged. The two unrelated document
hashes remain pinned and preserved. Desktop monitoring through this baseline
reported no new Win32 error5; it does not guarantee future unlocked access.

## Excluded fixture attempts

- Controller `reports/runtime/production-run-presentation-controller/3025ae9e856b437591c7e54d6c9f96cc`,
  result `reports/runtime/slice4be-production-run-presentation/a894e4e652ad42068a8ecedb0f978fe1/red.json`,
  2026-10-01 06:00:28.5397603--06:01:51.9510085 UTC:39 PASS/1 harness FAIL.
  An incorrectly parenthesized PowerShell argument array passed one argument to
  the two-argument catalog Definition probe. This is not behavioral RED. Corrected
  test call; runtime unchanged. Five compiles, normal closure, settings/package
  preservation and zero delayed Excel Application1000/1001/1002 failures.
- Controller `reports/runtime/production-run-presentation-controller/b65cd6db964141ce92b580ce709f20fd`,
  result `reports/runtime/slice4be-production-run-presentation/5f31a81995d34b2985a0d0cd1a99df1c/red.json`,
  2026-10-01 06:02:12.2937350--06:05:04.5627583 UTC:39 PASS/2 FAIL.
  One missing-metadata assertion preceded fixture Stage error91. The owned Visual
  Basic dialog was inspected through allowlisted error/button facts and dismissed
  using End; cleanup was assisted, so this is excluded from protecting RED.
  Settings/packages were preserved, Excel closed and delayed Application audit
  found zero failures. Error91 is not desktop Win32 error5. The next adapter returns
  only a fixed failing setup stage and numeric error, rather than leaving a dialog.
- Controller `reports/runtime/production-run-presentation-controller/c14544f577b04218a052da0d819bee85`,
  2026-10-01 06:05:53.3664289--06:07:14.8565239 UTC:39 PASS/2 FAIL.
  The bounded fixture diagnostic identifies TreeState/error91. In the adapter,
  `EnsureRunTreeState:` was parsed as a VBA label rather than a no-argument call,
  so the dictionary was never initialized. Separating the call onto its own line
  corrects only the fixture. This diagnostic is excluded from RED; five compiles,
  unassisted closure, preservation and delayed zero-failure audit pass. Do not
  repeat the colon form for no-argument procedure calls in VBA adapters.

## Remaining gates

Preserve every established presentation baseline identity through GREEN. Add the
other seven owner behaviors and complete
context/policy/failure coverage; then require packaged XLAM, compile, layout,
static maintenance, live roles, full Release1 chain, reusable and independent
Action Path evidence. Agent inspection is not human acceptance.
