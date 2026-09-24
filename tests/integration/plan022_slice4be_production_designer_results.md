# Slice 4be Production designer draft observations

Last verified: 2026-09-24 UTC. **Slice and six-control acceptance incomplete.**
Architecture v4.11 D18's discovered-control refinement governs catalog12's
Process/Recipe New, Clear and Validate controls. It preserves saved authority,
captured workbook/session/target binding and optional tracking. STAGED/VALIDATED
describe local draft results, never a saved or Domain-applied definition.

## Focused test-first evidence

`Test-Slice4beProductionDesigner.ps1` uses the existing Admin-generated fixtures
and five packages. Disposable adapters invoke each exact private click handler;
they never replace the handler or owner validator. Synthetic input remains in
fixture memory. Persisted results contain check identities/booleans only.

Initial adapter calibration required fixing an incorrect status-control name,
reserved VBA entry names, temporary lock-file hashing and hashing a saved fixture
while open. These attempts failed before the new action checks, not product RED.
An intermediate run reached meaningful observation/target/session failures but
its closed-workbook action raised an unhandled disconnected-object error; the
adapter now catches and reports only the numeric handler error. Its purportedly
valid output fixture also needed the existing three-character component ID rule.
Two compile-failure instances required assisted empty-process shutdown. The
intermediate interrupted run later found a separate Excel process with 15 already
saved workbooks outside the fixture paths; all were closed under standing user
authorization without saving or discarding unsaved edits. Its cause is unproven.
None of these calibration attempts is claimed as complete RED or clean shutdown.

The corrected first RED records **55 PASS/123 expected FAIL**, retaining the
actual six local results, valid Process validation and authority/workbook
preservation. The expanded RED includes denial and disabled-collection checks:
**63 PASS/165 expected FAIL**, 03:30:17--03:32:18 UTC. Failures are missing
observations/catalog12, permitted actions without capability, and missing stale
target/session/closed-workbook refusal. No harness exception; five instrumented
compiles pass and Excel closes normally. RED report:
`reports/runtime/slice4be-production-designer/79aee3cd63414233b4c92b8a171f11b3`;
controller `production-designer-controller/9f640aeebdea463b8edbbc40afb2857d`.

Runtime implementation then adds Core's fixed `modProductionDraftCodes`, extends
the catalog and routes six private handlers through one typed form-local action
path. Context and capability checks precede draft changes and do not depend on
collection being enabled. The live captured Workbook object must still belong
to Application.Workbooks; a replacement active workbook cannot substitute.
Existing draft helpers remain unobserved and validator errors are rethrown after
sanitized FAILED observation rather than swallowed. No new Application.Run call.

GREEN is **228/228**, 03:35:40--03:38:07 UTC, retaining every expanded RED identity.
The suite verifies linked distinct records, owner/actor/warehouse context,
sanitized payloads, hashes, local outcomes, historical catalog definitions,
disabled tracking, denial and all three stale-binding cases. Saved authority,
unknown workbook content in memory and saved bytes are preserved. Excel closes
normally, settings and candidate package bytes are restored/preserved. Zero
Application 1000/1001/1002 events at 03:38:51 UTC. GREEN report:
`slice4be-production-designer/4da00b74c27d4e2896e3625f60350b88`; controller
`production-designer-controller/3c7f92a8d25240a38e448efd4e038d4d`. Receipt:
`reports/runtime/production-designer-focused-verification.json`.

Review then found three fixed captions did not match the existing buttons. A
packaged adapter now reads all six actual Caption properties and compares the
catalog's caption/surface/class/role/capability metadata. Before correction this
records **231 PASS/three expected FAIL**, 03:53:03--03:55:16 UTC, exactly Process
New, Recipe New and Recipe Validate metadata; all 228 previous checks pass.
Architecture, Plan022 and controls were corrected to name the existing buttons;
no button wording or action behavior changes. The new vocabulary uses New Process,
New Recipe and Validate Recipe, and classifies outcomes by stable control ID.

Candidate `deploy/validation-production-designer-captions` passes **234/234**,
03:56:54--03:59:23 UTC, retaining every RED and preceding GREEN identity. Both
attempts preserve settings/packages and close without assistance; the GREEN
controller waited for eventual process exit. No Application 1000/1001/1002 events
at 04:00:00 UTC. RED report `slice4be-production-designer/cfa2db40ba404dc895a265b969209d16`,
controller `production-designer-controller/fc614684edd645aab577f34671f87d63`;
GREEN report `slice4be-production-designer/925f314b39f5450ca0ffef4cf1ad826e`,
controller `production-designer-controller/8ab132065a734d359f96701bb8c423e5`.
Receipt: `reports/runtime/production-designer-captions-focused-verification.json`.

## Build and regression scope

Candidate `deploy/validation-production-designer` builds all five packages and
passes explicit compile and Operations cold-start dependency validation.
Comparison of 244 compiled components against the palette candidate's 243 finds
only changed modActivityCatalog/frmProduction and added modProductionDraftCodes;
no component is removed. User Event Detail edits already in frozen packages are
preserved without committing them. Accepted deployment is unchanged.
The corrected-caption candidate also builds/compiles all five packages, passes
cold-start validation and changes only modProductionDraftCodes among those 244
compiled components. Both candidate directories and their receipts are retained.

Reusable Production passes one aggregate, 03:38:41--03:41:52 UTC. All **67 Boolean
observations** retain exact values and order from the palette candidate, including
two batches and Chai fork/convergence. Settings restore and Excel closes. The
existing launcher helper may terminate its owned process after Quit; strict
normal-closure proof is not claimed for this gate. Receipt:
`reports/runtime/production-designer-reusable-verification.json`.

The corrected-caption candidate also passes reusable Production, 04:04:06--04:08:05
UTC, retaining the exact values and order of all 67 Boolean observations. Settings
restore and Excel closes. The same helper cleanup limitation applies; this is
not strict unassisted-closure or native visible evidence. Receipt:
`reports/runtime/production-designer-captions-reusable-verification.json`.

Final maintenance reports record 251 components/6050 procedures/133021 lines: +1/+5/+124
for the six new observations. Plan022 records the explicit +45-line oversized
form exception; all other 27 oversized components retain their line counts.
Dynamic calls remain **9 literal/45 unresolved**, duplicate bodies **193**.
The new module has 63 lines; new procedures stay below 200. Three schemas validate,
270 PowerShell files under `tools` and `tests/tooling` parse recursively, and the existing
Production layout source checks pass **7/7 and 8/8**. Reports:
`reports/runtime/production-designer-captions-static`; the earlier static reports
are preserved separately. An early ratchet read ran before generation completed
and could not verify it; the terminal report comparison subsequently verifies all
28 ratchets unchanged from the initial designer candidate.

Initial candidate Settings regression first passed 191/191 but omitted 11
instrumented compile/install checks through an older wrapper. The corrected
wrapper passes the complete **202/202**, ending 03:52:03 UTC, with settings
restored and immediate Excel closure. Ten captures were produced; only three
were directly reviewed before the caption correction, so full image acceptance
is not claimed for that intermediate run. Current-candidate results follow
separately; the prior attempts remain preserved.

Corrected-caption candidate Settings passes **202/202**, 03:59:33--04:03:33 UTC,
retaining the exact prior 202 identities. All **ten images** are directly reviewed:
Admin General, Tracking, saved policy, detail profile, personal preference/save,
Operations Viewer entry and Settings normal/maximized/restored. Controls are
readable and unobscured; these are Settings-surface checks only. All five package
and 211 test-source pins match, settings restore, Excel closes immediately without
assistance and the Application audit is clear at 04:04:05 UTC. Report:
`slice4be-tracking-settings/1f255898c6ff48eab52d24ec5fd146a6`; receipt:
`reports/runtime/production-designer-captions-compiled-settings-verification.json`.

The test shell is still not an elevated administrator at 04:03:06 UTC
(`production-designer-captions-privilege-check.json`). Desktop capture succeeds
in this session without a permission change. This does not establish the cause
or permanent resolution of the earlier intermittent Windows error 5.

The corrected-caption full chain passes **32/32**, ordered live roles **48/48**,
and Create Warehouse **15/15**, 04:08:15--04:13:40 UTC. Every preceding palette
check identity is retained. Five package and 269 tooling hashes match, three
tracked reports and local settings are restored, and Excel closes without
assistance. No Application 1000/1001/1002 events at 04:13:54 UTC. Receipt:
`reports/runtime/production-designer-captions-chain-verification.json`; terminal
and outer settings-restoration receipts retain that same prefix. Earlier crashes
remain unexplained; this gate does not establish full Slice 4be acceptance.

## Remaining acceptance

The subsequent [Production path integration](plan022_slice4be_production_paths_results.md)
adds a real released-Process fixture and proves positive Recipe Validate at
244/244, retaining all 234 preceding checks without runtime changes.
Its later diagnostic candidate passes 390/390 through original recording,
publication, Viewer selection and expectation/Evaluate handlers after eight
expected RED failures. Two native designer captures are reviewed; remaining
Detail/diagnostic captures and owner/full-chain regressions await desktop input.
Guide integration and broader acceptance remain open. This is a partial checkpoint,
not complete acceptance of the six controls or Slice 4be.
The earlier 780 Settings diagnostic checks retain their previous
candidate scope; they have not been rerun here.
The broader [Slice 4be checklist](plan022_slice4be_remaining_acceptance.md) remains
binding, including the other Production actions and Operations/Admin coverage.
