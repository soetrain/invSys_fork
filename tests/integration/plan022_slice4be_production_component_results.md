# Slice 4be Process component observations — packaged RED established

Last verified:2026-09-29 UTC. **Incomplete; runtime implementation has not started.**
Architecture v4.11 D18, Plan022 and controls specify the ten requirement/output
Add/Update/Remove/Up/Down controls under approved semantic inheritance, committed
in documentation4dbbf76 before implementation. Catalog16 is specified, not built.
The unchanged candidate remains `deploy/validation-production-uom-extent`.

The new gate `Test-Slice4beProductionComponents.ps1` invokes the actual packaged
handlers through disposable adapters. Intended coverage includes existing edits,
Update identity/fallback/append behavior, ACTUAL and whole-UOM validation, output
yield defaults, no-selection Remove reset, regulation removal, movement/ordinals,
captured context, authorization, loading/nesting, disabled/unavailable tracking,
partial-write failure uncertainty, redaction and saved authority/custom columns.
The complete protecting RED below now exercises this coverage; GREEN remains pending.

## Protecting RED

After the user's explicit resume, the unchanged extent candidate completes
**180 PASS / 615 FAIL**,795 unique checks, five instrumented compiles and no harness
failure. All preceding179 passing identities remain; the corrected ACTUAL fixture
adds its missing preservation pass without removing an assertion. Controller
`production-component-controller/e7c1a894fba24da9819d087037f89014`, result
`slice4be-production-components/25881ef627e344ac9d12a0542c3d319f/red.json`,
22:52:01--22:54:28 UTC2026-09-29. Excel closes unassisted; local settings and all
five frozen package hashes are preserved. Workbook bytes/custom columns and saved
authority checks pass. No component runtime implementation precedes this RED.

Failures establish absent catalog16/action pairs/terminal facts, missing current-
context/permission/loading/nested guards and missing partial-failure cleanup.
Actual Up/Down also reproduces requirement80020005 and output80070057 after the
populated row fields move but before selection/ordinal completion. The shared
helper iterates declared columns, including unpopulated requirement slots and
unsupported output slots. Its seven/ten owned-field movement needs repair under
the existing D18 preservation requirement; this is not a new saved-state contract.

Calibration before this RED: VBE normalizes `.Text` to `.text`, so the unique
partial-write anchor now compares case-insensitively. Two complete diagnostic
runs produce179/616 with identical795 check names/results: controller95e307fc3f2f43839debd8067c4bf63b
and0e0675a0bd0448f99c2b7b6405a3088d; results246d7845465f420eb8c865d734d3af8e
anda1b37fb0eff342fc8d2a47a56fe7b070. The additional ACTUAL failure is fixture-only:
a no-argument procedure followed by a colon is interpreted as a VBA label.
Explicit `Call` statements restore setup. Boolean-only calibration now proves
both editors and written records have ACTUAL mode and empty quantity fields.
The adapter fails explicitly if that prerequisite is absent. No entered values
are written to calibration reports. Both diagnostic runs close unassisted and
preserve settings/packages; neither supersedes the corrected RED above.

## Rejected fixture attempts

1. Controller `production-component-controller/5aafcb64716b44a89a65450d686babd7`,
   result `slice4be-production-components/12869d6384c445898849e46f637b0594/red.json`:
   **48 PASS/90 FAIL**, including a harness failure;22:28:38--22:33:42 UTC.
   Five instrumented compiles pass. The first actual requirement Add preserves its
   expected rows/editor/regulations but lacks observations. The next test snapshot
   stops with VBA80070057, invalid List-property argument. The snapshot read all
   declared columns, including output slots beyond its ten populated fields.
   This is an incomplete fixture attempt, not the protecting behavioral RED.
   The verified disposable Excel process was terminated; settings and all five
   package hashes were restored. No VBE Debug/Reset intervention was used.
2. The snapshot is narrowed to the actual seven requirement/ten output fields,
   with explicit probe-error returns. A partial-write fault case is also added.
   Controller `production-component-controller/3e73876c876e47dca9db90295317ea78`,
   result `slice4be-production-components/59cfd85b775d4d2f97de81b1fc9abd08/red.json`:
   **1 PASS/1 harness failure**,22:34:31--22:34:40 UTC. The partial-write insertion
   anchor is unavailable before compilation/handler execution. This does not
   validate the corrected snapshot or constitute product RED. Closure is unassisted;
   settings and package hashes are preserved. Fix/calibrate the probe anchor first.

All runtime paths above are beneath ignored `reports/runtime/`. No raw fixture
values, screenshots or package binaries are committed. The gate is opt-in and
does not change the default ConfigCommands route. PowerShell parsing and diff
checks are static tooling checks, not substitutes for packaged RED/GREEN.

## Desktop condition and next action

Desktop cursor error5 first recurs at **22:34:46.392 UTC /15:34:46.392 Pacific**,
after the second controller already closed Excel. Last successful desktop sample:
22:34:43.368 UTC. InputDesktopError=0 and CaptureError=6. The goal was paused under
the user's explicit stop condition; no lock cause or idle timeout is established.
The finite observer was stopped. Onset receipt:
`reports/runtime/desktop-lockout-onset-20260929-153446.json`.

The user subsequently resumes work after reporting that the RDP-client display
timeout is set to Never. The new observer starts22:45:27 UTC; desktop checks pass
through this RED. This does not prove a lock-policy fix. Stop again on actual
desktop error5, preserving its first timestamp and last successful sample.

Next implement the approved catalog16 metadata/evaluator and typed Operations
owner/handler guards, restore busy/loading after failure, and preserve seven/ten
component fields during movement without growing the oversized form. Then build,
compile, obtain GREEN with all795 identities, and prove published paired paths,
relevant regressions, layout/static and current Release1 chain/reuse. The separate
native full-chain/reuse failures remain unresolved. No release acceptance or
operational deployment is claimed.
