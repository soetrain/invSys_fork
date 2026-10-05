# Slice 4be-A: Production Print Recall

## Saved test-package confirmation, 2026-10-05

On unchanged `validation-print-outcome-02`, the full Print handler gate passes
250/250 using `-PrintBaselineOnly -SavedProbeCopiesForTest`. All248 previous
checks retain their exact order; two additional checks require probes installed
before forms and saved only in disposable packages. Five instrumented compiles,
normal closure, original package hashes, settings restoration and delayed zero
Excel Application errors pass. Four reviewed captures show preview return,
failure, refusal after success and recovery. Native preview remains a test seam.

This completes the scoped owner-feedback confirmation alongside outcome02's
existing cold compile and chain32/live48/warehouse15 evidence. Fresh smoke86
retains every ordered check and both unassisted exits. Outcome01 layout18/five
native checks remain applicable: all304 component code/string-literal hashes
match outcome02 and form geometry is unchanged. No runtime or production build
tool changed. Earlier native failures remain unresolved; this is a passing test
configuration, not proof that saving compiled state cures Excel instability.

Ignored receipts, relative to `reports/runtime/`:

| Gate | Receipt | Result |
|---|---|---|
| Full focused | `production-run-local-controller/6421e06c0ba6422f91a04be1ef8341d6/verification.json`; worker `slice4be-production-print-baseline/58f4ff803eb14921845d78748c763e9e` | 250 PASS, 22:37:41-22:40:48 UTC |
| Smoke | `print-outcome02-regression/smoke-590d63d66be4484483b59bdfacc07a01/verification.json` | 86 PASS, 22:41:09-22:41:29 UTC |
| Static | `print-seed-boundary-static-01/ratchet-verification.json`, `final-parse.json` | Unchanged311 components/6306 procedures/137979 lines; 9 literal/45 unresolved calls,190 duplicate candidates;28 caps,3 schemas,424 PowerShell parses |

Narrow diagnostics retain their failures; none are behavioral RED or substitute
for the full gate. Diagnostic controller IDs below are under
`production-run-local-controller/`, with `failed-run-audit.json` or
`verification.json` recording closure/preservation and delayed events:

- SeedOnly, after shared Config/UOM exercises but without Print actions:
  `5be01534b0b44da38aa4c986fda418d6`,40 PASS/2 harness FAIL,
  22:24:33-22:27:37 UTC. Seed crashed with zero Production forms loaded;
  assisted recovery preserved23 fixture hashes. Print actions are not necessary
  for this failure.
- FreshSeed, before shared Config/UOM exercises:
  `75c1512da09949d58236a7943e68cadc`,6 PASS/1 harness FAIL,
  22:27:53-22:28:21 UTC. Bootstrap failed before Seed.
- Unmodified full-chain Admin entry control:
  `print-seed-admin-entry/092838f63c9d44da8a112a8edfcbb754/verification.json`,
  3 PASS,22:30:41-22:30:54 UTC, no delayed Excel errors. It generates a fresh
  warehouse and seeds through the packaged Admin owner.
- BareSeed, before shared forms and without Production probes:
  `a2e1468a1622466eb67d47ece222e7c8`,2 PASS/1 harness FAIL,
  22:32:19-22:32:53 UTC. Seed failed after both fixtures were generated.
  This path also omits the five instrumented compiles, so it is not a controlled
  compiled-state comparison. Production probes are not necessary for failure.
- FreshSeed with saved disposable probes:
  `b21a12a4424d4cb3801d02ad51bc5bfa`,14 PASS,
  22:34:16-22:35:06 UTC, normal closure and delayed zero Excel errors.
  Saving occurs before fixtures in the same Excel session, without a restart.

Each failed diagnostic records Excel Application1000/1001; the successful ones
restore settings and preserve original packages. Unused OpenClose/Prelude options
were removed before commit; the successful full None path is unchanged. D13's
prior231/17 behavioral RED remains the runtime evidence; these harness changes
introduce no architectural contract. No error5 occurred. Next protect remaining
Print Recall permission/interruption guards through the actual packaged handler;
native preview, report identity/provenance and observation acceptance remain open.

## Truthful owner feedback, 2026-10-05

The form previously inferred success by rebuilding the report after its owner
returned. It displayed "completed" even after refusal or preview failure. Under
D18 truthful feedback, the handler now displays the original owner's detail and
prepares the report once. Normal preview return says "Print preview closed."
Refusals and errors retain their existing native dialogs and replace prior status.
No physical-printing or observation acceptance follows from a preview return.

The public compatibility command retains no-argument use and adds optional primitive
outcome/detail outputs. Typed Operations calls use REJECTED, FAILED or
PREVIEW_RETURNED. The report helper owns feedback; the report-builder function is
public for that same-project typed call, not a new cross-XLAM bridge. The diagnostic
API remains available to explicit callers. Form geometry is unchanged.

Tests deliberately replace five incidental two-build assertions with a single
owner-build assertion. Four inventory cases now explicitly invoke the diagnostic
API and still verify both reads' captured binding. This maps all217 preceding
checks, with31 added checks; it is not an unchanged-check-identity claim.
New cases cover preview return, preview failure after success, refusal after
preview and recovery. The native-preview call is still a disposable test seam;
actual native printing and interruption acceptance remain open.

Entry: tests/tooling/Test-Slice4beProductionRunLocal.ps1 -PrintBaselineOnly,
Phase RED / deploy/validation-print-inventory-01 then Phase GREEN /
deploy/validation-print-outcome-01. Ignored receipts, relative to reports/runtime/:

- RED: production-run-local-controller/3d46e35f5a4b4575b0de3ea14e2dffbb;
  slice4be-production-print-baseline/80fff3b4ed2d457585b6ae64d598f0fd/red.json,
  21:28:55-21:32:03 UTC.231 PASS/17 behavioral FAIL:13 duplicate-build checks and
  four false-status checks. Five instrumented compiles, normal
  closure/restoration, package preservation and delayed zero Excel errors pass.
- Build: print-outcome-build-01/verification.json,21:32:46-21:33:25 UTC.
  Five cold compiles,304 components/301 unchanged; changes are mProduction,
  modProductionRecallReport and frmProduction. Settings and frozen packages are
  preserved, with delayed zero Excel errors.
- GREEN: production-run-local-controller/e2d8c74700154afdbd71cedfe89f2526;
  slice4be-production-print-baseline/5e29c0efc9ac407489a60429fc1c8887/green.json,
  21:33:38-21:36:56 UTC.248/248 with exact RED identities and prior217 mapped
  coverage. Five compiles, preservation/restoration, normal closure and delayed
  zero Excel errors pass. All four outcome captures are readable and show the
  original owner result, including failure/refusal after earlier success.
- Layout: print-outcome01-regression/layout-61e31fc8da1649199cfe23d3e35f77a6/
  verification.json,21:37:14-21:37:32 UTC.18 size/page and five native-window checks
  pass, zero bounds/interactive-overlap failures; preservation and delayed zero
  Excel errors pass.
- Static: print-outcome-static-01/ratchet-verification.json:311 components,
  6306 procedures,137979 lines; two helper procedures/nine net lines added.
  mProduction shrinks15 lines and frmProduction shrinks2. All28 oversized caps
  pass;9 literal/45 unresolved dynamic calls and190 duplicate candidates are
  unchanged. Three schemas and423 PowerShell parses pass.
- Smoke: print-outcome01-regression/smoke-b2126e7bb16b4b98b8b1bc80ce571244/
  verification.json,21:37:45-21:38:13 UTC.86 exact prior checks, both unassisted
  exits, settings/package/tracked-report restoration and delayed zero Excel errors.
- Full-chain attempt: print-outcome01-regression/chain-5c7dcadc00a747569364d7f14184b034,
  21:38:52-21:54:21 UTC. Chain5 PASS/1 FAIL; live-role35 PASS/1 harness exception;
  warehouse15 PASS. The exception occurred at ProductionFormTwoBatchActionReportForTest
  (HRESULT 0x80020009); its cause is unresolved. Ten recovered test workbooks
  required closure without saving. All twelve original fixture hashes remained
  unchanged before Quit; controller settings/packages/tracked reports were restored.
  failed-run-audit.json and assisted-recovery-closure.json retain the failure,
  assisted cleanup and delayed zero Excel Application errors. No error5 occurred.
  This run is not accepted.
- Unchanged-candidate retry: print-outcome01-regression/chain-510622ae02954797a20b4fe5576cab00,
  21:54:37-21:59:13 UTC. Production two-batch form action passed; Boxing then
  failed at modBoxingService.RunRelease1BoxingActionForTest (HRESULT 0x800706BE).
  Chain5 PASS/1 FAIL; live-role36 PASS/1 harness exception; warehouse15 PASS.
  Application1000 records ntdll.dll/0xc0000028 at21:57:04 UTC, followed by1001.
  Seventeen recovered test books (ten current, seven earlier recovery copies)
  closed without saving; all24 original fixture hashes were preserved. Settings,
  packages and tracked reports were restored. Exact receipts are failed-run-audit.json
  and assisted-recovery-closure.json. This is a second failed run, not acceptance
  or an established cause.
- Preceding-package control: print-outcome-prior-control/chain-63982828f1994d9085507d1b83d3d630/
  verification.json,21:59:36-22:04:54 UTC, validation-print-inventory-01 unchanged.
  Exact chain32/live48/warehouse15 pass with unassisted closure, restored settings/
  packages/tracked reports and delayed zero Excel errors. A read-only checkpoint
  saw zero recovery books; it is not continuous isolation proof. The comparison
  does not establish the failure cause or accept the newer package.
- Third outcome01 attempt: print-outcome01-regression/chain-fd643eff3dee4af1917a583429c0b276,
  22:05:43-22:09:07 UTC. Chain5 PASS/1 FAIL, live-role32 PASS/1 exception,
  warehouse15 PASS. Projection rebuilding failed at modProcessor.RunBatchReportForAutomation
  (HRESULT 0x800706BE), before the preceding failure points. Application1000 records
  ntdll.dll/0xc0000028 at22:07:19 UTC, followed by1001. Ten test books closed without
  saving; twelve original fixture hashes were preserved. Settings/packages/reports
  restored; failed-run-audit.json and assisted-recovery-closure.json retain evidence.
  No error5. Repeated current-candidate failure prevents full-chain acceptance.

Packaging hypothesis under investigation: build saves source-imported XLAMs;
the ordinary cold compile test opens read-only and discards compiled state on close.
A separate validation-print-outcome-02 build persisted compilation and reopened
all packages for the ordinary cold check. All304 component code/string-literal
hashes match outcome01; frozen packages/settings are preserved. Build interval
22:09:40-22:10:33 UTC; print-outcome-build-02/verification.json and
source-comparison.json retain the proof. The local compile-save diagnostic opens
candidate packages writable, executes the existing compile check, saves each,
then runs unmodified Test-PackagedVbaCompile.ps1 from a fresh Excel instance.
No runtime source or production build-tool change was made. Rebuild and saved
compilation are both differences, so this does not isolate or prove the cause.

Outcome02 chain passes32/live-role48/warehouse15 in exact prior order:
print-outcome02-regression/chain-a762d42d779a40219c0f8f50cef57c3b/verification.json,
22:11:04-22:16:27 UTC. Normal unassisted closure, settings/packages/tracked-report
restoration and delayed zero Excel Application errors pass. Original outcome01
remains unaccepted; its three failures are retained above.

The focused rerun on outcome02 did not complete: controller
production-run-local-controller/48d8c66c58c540368e788cd47f47bbd1, worker
slice4be-production-print-baseline/c8783c6e3d9e4f7b9b8f718bf8420ef5/green.json,
22:16:52-22:20:53 UTC,178 PASS/2 harness FAIL. At the Admin Seed fixture boundary,
first-call-failure.json records HRESULT0x800706BE; Application1000 records
ntdll.dll/0xc0000028 at22:19:24 UTC, followed by1001. Three recovered test books
closed without saving, all23 original fixture workbook hashes were preserved,
and settings/packages were restored. failed-run-audit.json and
assisted-recovery-closure.json retain the evidence. This is not behavioral RED
or a focused GREEN on outcome02. Desktop probes remained successful.

At that checkpoint, evidence remained partial: focused248/layout/smoke was on outcome01;
full-chain evidence was on source-identical outcome02. No candidate then had all gates
passing together. Persisting compilation did not eliminate native instability;
do not adopt it as a proven fix or repeat broad gates without a narrower diagnostic.
The next investigation isolated the post-form Admin Seed boundary with the existing packaged fixture,
keeping owner calls, saved-byte preservation and normal shutdown observable.
Source review found MouseScroll's native hook but no EnableMouseScroll caller in
src; there is no evidence to blame or delete it. No speculative native-code fix.

## Captured inventory-location lookup, 2026-10-05

The legacy inventory-sheet resolver could read another open workbook when the
captured workbook lacked InventoryManagement, even if its supported Inventory
Management alias existed. Actual Print Recall RED211/6 proves foreign reads and
foreign locations at both preview and diagnostic boundaries in those two cases.
Captured-table and missing-table cases pass. All183 prior checks remain ordered.
The report builder now calls GetInvSysTableFromWorkbook(wsProd.Parent), applying
existing D18 binding. Missing local lookup retains the existing blank location.
This changes one typed call; it does not change report schema or permissions.

Run tests/tooling/Test-Slice4beProductionRunLocal.ps1 -PrintBaselineOnly with
Phase RED / deploy/validation-print-rebuild-02, then Phase GREEN /
deploy/validation-print-inventory-01. The fixture uses an Admin Seed-created exact
System_Key in workbook-local projections. It compares source rows, both projections,
warehouse-file hashes and saved operator bytes; no new canonical inventory is
invented. The native-preview seam remains explicit and does not prove printing.

Ignored receipts below are relative to reports/runtime/:

- RED: production-run-local-controller/b22c6ae1a2e54817a17d2b33e38e58f3;
  slice4be-production-print-baseline/f52dc1ecb22b46b8b6f346ebc7681810/red.json.
  21:08:04-21:11:17 UTC:211 PASS/6 FAIL, five instrumented compiles, restoration,
  package preservation, normal closure and delayed zero Excel errors.
- Build: print-inventory-build-01/verification.json,21:11:51-21:12:30 UTC.
  Five cold compiles,304 components/303 unchanged; only mProduction changes.
  Prior package hashes/settings are preserved; delayed zero Excel errors.
- GREEN: production-run-local-controller/d8aac0c32e5646caa1dce55744b7f3cb;
  slice4be-production-print-baseline/1f34752c523d49b88a678477657b3f3a/green.json.
  21:13:07-21:16:16 UTC:217/217, exact RED check order and all183 prior checks,
  five instrumented compiles, normal closure/restoration and delayed zero Excel
  errors. Reviewed alias/missing-sheet form captures retain the existing status;
  they are not native preview or printing evidence.
- Static: print-inventory-static-01/ratchet-verification.json: unchanged311
  components/6304 procedures/137970 lines,9 literal/45 unresolved dynamic calls,
  190 duplicate candidates;28 non-growing oversized caps,3 schemas and422
  PowerShell parses pass. No runtime code growth. Forms are identical, so the
  prior binding01 layout18/five native checks remain applicable.
- Smoke: print-inventory01-regression/smoke-220f73a30f624ecfb266d1602eb860ee/
  verification.json,21:16:50-21:17:17 UTC:86 exact prior checks, both unassisted
  shutdown stages, restoration and delayed zero Excel errors.
- Full chain: print-inventory01-regression/chain-4ecad966e0494d3492dc82d1aa2cc7bb/
  verification.json,21:17:40-21:23:37 UTC:chain32/live-role48/warehouse15 pass in
  exact prior order. Normal cleanup, settings/package/tracked-report preservation
  and delayed zero Excel errors pass. Candidate remains unpromoted; no error5.

Print Recall report identity, exact lookup/header edge cases, native preview,
permission/yield/closure guards, truthful outcomes and observations remain open.
This binding correction does not accept Print Recall or comprehensive A1/A2.

Two earlier setup trials ended before the new assertions: controllers
af2507f62b72462b9b17423b590b2afc / da04a4b42a2f4bc08fc23ab82dd4066c,
workers2d2ff997ddd641768cacf56e9b8198a2 /6c8a9ce003b94deca17f059218e0fb56.
Each had178 passes and one fixture file-sharing exception, not behavioral RED.
The first attempted correction did not change the shared fingerprint helper's
direct hash operation. The corrected scoped reader excludes Excel lock files,
opens other files with read sharing, and closes only authority workbooks it opened.
Both trial-audit.json records confirm restored settings/packages, closure and
delayed zero Excel errors. No implementation changed before the governing RED.

## Successful rebuild preservation, 2026-10-05

Under existing D14, the report now reuses its named table and writes managed
fields by normalized header. Custom columns, values, formulas and positions,
other tables and unrelated cells survive reuse, growth and shrink. Missing or
ambiguous managed headers and occupied managed growth cells refuse before writes.
Existing custom growth cells remain allowed. This does not define report identity
or change inventory authority. The new typed Operations helper has three procedures;
the owning mProduction module loses eleven lines.

The actual packaged Print Recall handler protects this change. Only native
PrintOut Preview is replaced in unsaved test instrumentation with an observed
boundary; no native preview or physical printing acceptance is claimed.
Corrected RED138/40 on print-preserve01 becomes GREEN178/178 on print-rebuild02,
retaining all106 prior checks. Intermediate rebuild01 RED175/3 isolates overly
restrictive custom-cell growth. Five additional first-preview metadata checks pass
on unchanged rebuild02:183/183. That positive run was invoked with Phase RED and
keeps its red.json filename; it is not a behavioral RED.

Entry point: tests/tooling/Test-Slice4beProductionRunLocal.ps1 -PrintBaselineOnly,
using Phase RED with deploy/validation-print-preserve-01 and Phase GREEN with
deploy/validation-print-rebuild-02. The current test also includes the five later
metadata checks; the recorded original RED/GREEN pair contains178 checks.

Ignored receipts, relative to reports/runtime/:

| Gate | Controller / worker or verification receipt | Result |
|---|---|---|
| Corrected RED | production-run-local-controller/88789176c8884afaac18f24f5604199e; slice4be-production-print-baseline/01a6091a860f4af8b6a76d9d794c6ed1/red.json | 138 PASS / 40 FAIL; 20:35:45-20:38:23 UTC |
| Expansion RED | production-run-local-controller/d707b72cd084492d8f30a3fc383926f0; slice4be-production-print-baseline/66cf3e75e3ee4c458d8bbc70cf8ccb98/red.json | 175 PASS / 3 FAIL; 20:38:42-20:41:24 UTC |
| Build | print-rebuild-build-02/verification.json | Five cold compiles;304 components,303 unchanged from rebuild01; only helper changes |
| GREEN | production-run-local-controller/532c91edd064417abc6987d5426e104e; slice4be-production-print-baseline/35264c02e0484f108312fbd602bc91da/green.json | 178 PASS; exact corrected RED check order;20:42:45-20:45:33 UTC |
| Metadata proof | production-run-local-controller/3979f0c4407e401b9b33673478883321; slice4be-production-print-baseline/be655a3f8e7b4174af2f6d6f1bda61f9/red.json | 183 PASS; all178 retained plus five readable timestamp checks;20:46:09-20:48:54 UTC |
| Smoke | print-rebuild02-regression/smoke-ef9c488c2733416a976a05fc572903be/verification.json | 86 PASS, exact prior order, both unassisted exits;20:49:43-20:50:08 UTC |
| Full chain | print-rebuild02-regression/chain-61b5c7586c6d45a3899d023cc6e2473b/verification.json | Chain32/live-role48/warehouse15 PASS, exact prior order;20:53:18-20:59:14 UTC |
| Static | print-rebuild-static-02/ratchet-verification.json | 311 components/6304 procedures/137970 lines;9 literal/45 unresolved dynamic calls/190 duplicate candidates;28 non-growing caps,3 schemas,421 PowerShell parses |

Focused runs/build/smoke/chain verify five package hashes, settings restoration,
normal Excel closure and delayed zero Excel errors. Smoke and chain also restore
tracked reports. Full chain retains the positive report diagnostic, Boxing,
Shipping, restart and reconciliation checks.
Net runtime growth from preserve01 is one small module, three procedures and92
lines; oversized modules, dynamic calls and duplicate counts do not regress.
Forms are unchanged; binding01 layout18/five native checks remains applicable.
Reviewed GREEN captures show preserved custom cells during growth and below the
shrunk table, retained unrelated content, and the occupied-growth refusal. The
existing misleading completed prefix remains an explicit open defect.

Superseded trials: controller ef09f887ee5f42889a0ffd20e3c17944 / worker
54c39f7daee64799809ef919974720b4 produced133/35. Controller
c005ffdff8fc42c5b94f4d21e3b5f57d / worker adbd4a41bf014d20a538c1e2b5656127
produced165/3 on rebuild01, but its collision fixture wrote below an existing table
and Excel expanded it. Those three failures are fixture errors, not product RED.
Corrected fixtures populate collision/custom cells before table creation and assert
the initial row count; the two governing REDs above supersede these trials.

Candidate deploy/validation-print-rebuild-02 is not promoted. No error5 occurred.
Report System_Key/provenance, inventory lookup binding, native preview,
permission/yield/closure guards, truthful outcomes and observations remain open.

## Refusal preservation, 2026-10-05

The actual Print Recall handler previously created or cleared RecallCodesPrint
before discovering no recall-coded output. Both the command and its diagnostic
could erase custom columns/formulas, another table and unrelated cells while
refusing preview. Under existing D14, report preparation now follows eligibility
validation. The private renderer returns the prepared worksheet by reference;
existing refusal text and successful report generation remain the contract.

Focused RED99/7 on print-binding01 retains all80 prior ordered checks. Seven
failures prove unwanted empty-report creation, and lost report/tables/cells for
blank recall codes and a missing recall column. Empty-output refusal preserves
the independently prepared report. Tests fingerprint table names, addresses,
headers, values and formulas, plus all used cells, without publishing their content.
An earlier trial92/9 reused the damaged report between cases; its final three
failures were carryover, so isolated RED99/7 is the governing evidence.

Run `tests/tooling/Test-Slice4beProductionRunLocal.ps1 -PrintBaselineOnly` with
`-Phase RED -DeployRoot deploy/validation-print-binding-01`, then
`-Phase GREEN -DeployRoot deploy/validation-print-preserve-01`.

Ignored receipts below are relative to `reports/runtime/`:

- Trial: `production-run-local-controller/75686ae8593c4ac2a6f89c56676617ef`,
  worker `slice4be-production-print-baseline/0039161b60764c0c8c228a8dc286b30f`.
- RED: `production-run-local-controller/2c7287ce7d5c4c898fe6d5cda0043b5d`,
  worker `slice4be-production-print-baseline/050cf1ef46064a56bc23012726e9ebed/red.json`,
  20:05:49.556–20:08:04.458 UTC. Five instrumented compiles, normal closure,
  restored settings, preserved package hashes and delayed zero Excel errors pass.
- Build: `print-preserve-build-01/verification.json`, 20:08:23.044–20:09:01.495 UTC.
  Five cold compiles; only mProduction changes, with302 other components unchanged.
  Forms are identical; entry-binding candidate layout18/five native checks remains
  applicable. Prior packages/settings are preserved; delayed zero Excel errors.
- GREEN: `production-run-local-controller/81202914b767410d902244a765b9ca92`,
  worker `slice4be-production-print-baseline/092cb0aa34e1482d8222a1a53fc56f2e/green.json`,
  20:09:15.890–20:11:39.265 UTC. 106/106, exact RED/GREEN identities, all80 prior
  checks retained; five compiles, normal closure, preservation and delayed zero
  Excel errors. Four new refusal captures are readable; the misleading existing
  completed prefix remains open and is not accepted as an owner outcome.
- Smoke: `print-preserve01-regression/smoke-39e81c62a83e427d908ce55accfab4d2`,
  20:11:57.773–20:12:24.763 UTC. All86 exact ordered checks, both unassisted exits,
  settings/package/tracked-report preservation and delayed zero Excel errors.
- Static: `print-preserve-static-01/ratchet-verification.json`: 310 components,
  6301 procedures, 137878 lines (one fewer); 9 literal/45 unresolved dynamic calls,
  190 duplicate candidates, 28 non-growing oversized caps, three schemas and
  420 PowerShell parses pass.
- Full chain: `print-preserve01-regression/chain-bc24fbf5fe894a4fb90384442ad2be52`,
  20:12:43.912–20:18:39.499 UTC. Chain32, live-role48 and warehouse15 all pass
  with exact prior ordered checks. The positive recall-report diagnostic produces
  its expected row; versioned Boxing, shipment, restart and reconciliation pass.
  Normal cleanup, settings/package/tracked-report preservation and delayed zero
  Excel errors pass. Earlier Boxing native failure does not reproduce; its cause
  is not established by this report-only correction.

Candidate: `deploy/validation-print-preserve-01`, not promoted. No error5 occurred.
Successful rebuild
preservation, preview/printing and recording integration are still open.

## Entry binding, 2026-10-05

Print Recall previously entered its report owner and diagnostic after a warehouse
or session change, sign-out, or loss of the captured Production sheet. The missing
sheet allowed the legacy resolver to search other open workbooks. Under existing
D18, the actual `frmProduction.mBtnManagerPrint_Click` now requires the original
live context and workbook's Production sheet before either report call. The typed
`modProductionRunBinding.RequireWorksheetContext` composes existing guards.

Packaged RED68/12 becomes GREEN80/80 with identical ordered checks and all68
previous passes retained. Each of four invalid cases proves zero owner entries,
zero report reads and the exact visible refusal. The valid case retains one owner
entry and both existing report reads with a decoy active; the observed workbook
is the captured workbook. Its missing output table produces the existing native
refusal before preview. This deliberately protects entry binding, not printing.
Captured custom values/formulas, decoy, saved operator bytes and the other generated
warehouse remain unchanged; no report is created. Five focused captures reviewed.

At code checkpoint `7b1d8b5c`, run
`tests/tooling/Test-Slice4beProductionRunLocal.ps1 -PrintBaselineOnly` with
`-Phase RED -DeployRoot deploy/validation-next-closed-01`, then
`-Phase GREEN -DeployRoot deploy/validation-print-binding-01`.

Ignored receipts, relative to `reports/runtime/`:

| Gate | Receipt | Result |
|---|---|---|
| RED | `production-run-local-controller/725a40197bc54fa98d48d07e4f1c8918`; `slice4be-production-print-baseline/8d7ade8fe0f846079dac647f78986ec7/red.json` | 68 PASS / 12 behavioral FAIL; 19:48:33.691–19:50:39.032 UTC |
| Build | `print-binding-build-01/verification.json` | Five cold compiles; 303 components, 301 unchanged; 19:51:47.044–19:52:25.516 UTC |
| GREEN | `production-run-local-controller/1471e10de51a48d8b9dcb6533e2723e4`; `slice4be-production-print-baseline/79ad954f71e449a888ad672e0ffa8441/green.json` | 80 PASS; 19:52:47.571–19:55:07.594 UTC |
| Layout | `print-binding01-regression/layout-c77344b88a924a30961115f7214efbfb/verification.json` | 18 page/size checks, five native-window checks, zero bounds/interactive-overlap failures |
| Smoke | `print-binding01-regression/smoke-9c23c9a207054cf3b8ec434d66e4a52e/verification.json` | 86 PASS, exact prior ordered checks; 20:00:04.117–20:00:32.435 UTC |
| Static | `print-binding-static-01/ratchet-verification.json` | 310 components, 6301 procedures, 137879 lines; 9 literal/45 unresolved dynamic calls, 190 duplicate candidates; 28 non-growing oversized caps; three schemas and 420 PowerShell parses pass |

Only compiled `frmProduction` and `modProductionRunBinding` change; form geometry
is unchanged. The added typed guard accounts for one procedure and six lines;
dynamic-call and duplicate counts do not grow. Focused runs compile all five
instrumented packages. Build/tests preserve settings and package hashes, close
Excel normally, and have zero Excel failures in their delayed Windows-event audits.
Smoke also restores tracked reports and proves both unassisted shutdown stages.
The candidate is not promoted. No desktop error5 occurred during these checks.

## Remaining acceptance

This checkpoint does not accept Print Recall as recorded or prove native preview/
physical printing. Permission changes, yielding/closure paths and observation
integration remain open. Rebuild preservation and truthful owner feedback are
protected above; report identity/provenance remains open. The displayed preview
return cannot establish that a page was printed.

Existing unchanged-workflow evidence remains scoped to its recorded candidates.
The refusal-preservation candidate above now has fresh chain32/live48/warehouse15
evidence; the earlier intermittent Boxing failure's cause remains unproved.
Comprehensive A1/A2 and 4be-A completion remain open.
