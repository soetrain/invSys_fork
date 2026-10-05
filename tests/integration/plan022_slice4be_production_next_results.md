# Slice 4be-A: Production Next Batch binding and recording

## Recording-to-guide evidence, 2026-10-05

The unchanged `deploy/validation-complete-activity-04` candidate records the actual
Next Batch handler twice, once as guide provenance and once as an independent
observed run. Reusable preparation first completes a real batch outside recording;
worksheet preparation uses generated staging. Each Next Batch action separately
preserves canonical workbook bytes. Worksheet facts include native notification,
exact System_Key, incremented batch, cleared output/recall and unknown columns.

The shared guide harness now has an isolated Next Batch mode. It proves stopped
journals, exact published records and Event Detail, command/source-application
separation, explicit authoring, reader pairing, all three presentations, immutable
history and saved workbook bytes. This adds tests and evidence under existing D18;
no new contract, runtime change or product RED is claimed. Prior binding/observation
RED/GREEN remains below; comprehensive acceptance and tracking faults remain open.

Reproduce with `Test-Slice4beProductionRunLocal.ps1 -DeployRoot
deploy/validation-complete-activity-04 -Phase GREEN -CompleteBaselineOnly
-NextBaselineOnly -NextPathsOnly -NextPathMode Reusable|Worksheet`.

Reusable trial85/85: controller
`production-run-local-controller/158c699307a3471099bdafe485dfde12`, worker
`slice4be-production-next-paths-reusable/e6c15bb7e0b445479f3530c889c739cc/green.json`,
17:24:48.420--17:31:59.833 UTC. Five compiles, preservation, normal closure and
delayed zero Excel errors pass. Capture review found an inherited fixture heading
for editing instructions despite correct Next Batch steps. The authored heading/
scope was corrected and asserted in How-To; this was not an application defect. These
trial images do not establish the final visible guide proof.

Worksheet94/94: controller
`production-run-local-controller/9132df1e03d0470494fa983bd4cd415c`, worker
`slice4be-production-next-paths-worksheet/0a05c041b5ff4fe78ac814a8b1009a83/green.json`,
17:32:19.171--17:38:32.903 UTC. Five compiles, normal closure, settings/package
preservation and delayed zero Excel errors pass. All four guide/conclusion captures
are reviewed and readable, including the correct Next Batch heading, instructions,
distinct source/observed provenance and local-only diagnostic conclusion.

Final reusable86/86 retains all85 trial checks in exact order: controller
`production-run-local-controller/59e046d9907d4f3dbaba07b79f6b1477`, worker
`slice4be-production-next-paths-reusable/c43a179cdc024838893bd962662097e1/green.json`,
17:38:49.938--17:45:43.980 UTC. The corrected heading/scope assertion passes.
All four current captures are readable and reviewed. Five compiles, normal closure,
settings/package preservation and delayed zero Excel errors pass. No error5 occurred.

Static `next-paths-static-04b/` passes416 parses, three schemas and28 unchanged caps.
Runtime metrics remain310 components/6300 procedures/137871 lines,9 literal and45
unresolved dynamic calls,190 duplicate groups. Retain valid unchanged package,
layout/live-role and Complete Run evidence; the recorded full-chain native failure
remains unresolved. No deployment, Next Batch replay or broader A1/A2 acceptance.

## Recording candidate, 2026-10-04

Architecture D18 catalog26 registers `PRODUCTION_RUN_NEXT_BATCH`. The real form
handler now delegates to a typed guarded action: local STAGED requires owner
acknowledgment and refresh; incomplete batches are REJECTED, missing worksheets
are FAILED and denied attempts remain DENIED. Loading/busy/nested entries are
suppressed. Worksheet selection/keep rules and normal notifications are retained.

Frozen binding01 produces meaningful RED110 PASS/57 FAIL; candidate
`deploy/validation-production-next-activity-01` produces GREEN167/167.
Both close normally, restore settings, preserve packages and pass delayed native
audits. Identical ordered checks and all70 previous GREENs are verified;
Candidate smoke86 and layout18 size/page pairs pass with prior ordered checks,
normal closure/preservation and delayed zero native audits. Six layout captures
are reviewed; minimum/default share one clamped size. Full-chain fails at the
versioned Boxing form action with a native Excel crash (22:06:08.983 UTC,
ntdll.dll/c0000028, offset12d2f; harness HRESULT0x800706BE). This is not meaningful
product RED or desktop error5. Ten recovered fixture workbooks close without save;
saved bytes, settings, tracked reports and packages are preserved. The earlier
source-exit wait remains in place; crash causation is unproved. The standalone
ordered live-role comparison passes48/48 with all prior checks, normal closure,
preservation and delayed zero native events. This does not substitute for a passing
full chain; investigate the preceding setup context rather than repeat unchanged gates.
Use the earlier reproduction command with `-NextActivityOnly`; both runs retain
the original70 binding checks. The extension covers catalog history, actual owner
state, optional-policy off, exact actor/target and linked attempt/outcome records,
redaction/integrity, local-only completion, custom columns/formulas and saved bytes.

All five candidate packages compile;280/284 previous compiled components remain
unchanged. Core's catalog/Run codes and Operations' form/worksheet owner change;
Operations adds one typed action helper. Static:292 components/6178 procedures/
135605 lines,9 literal/45 unresolved Application.Run,190 duplicate candidates;
all28 oversized caps hold. Growth is one helper/procedure and50 net lines.
Three schemas and383 PowerShell parses pass. Two original-resolution captures
show ready and stale-context refusal; clipped long labels and the small output
region remain. This does not establish populated/multiline or human acceptance.

Earlier probe attempts exposed an unrecognized fixture palette name, a hash read
sharing violation and a native fixture-sheet deletion prompt. These are harness
issues, not product RED. Corrected setup and the owned notification observer yield
the final unassisted RED; the earlier assisted run is retained separately.
Tracking-fault/interruption, recording-to-guide evidence, remaining control
coverage and overall4be-A/B0 acceptance remain open.

Ignored receipts under `reports/runtime/`:

- RED controller `production-run-local-controller/bc452a3c80ed41baa968a7d80435a694`,
  report `slice4be-production-next-activity/5649dd246bbe4f2eb73f817ccd31db23/red.json`.
- GREEN controller `production-run-local-controller/140012873a394fb199a8c9b3707ba950`,
  report `slice4be-production-next-activity/3efa68f56a3a4d198521266edc6ef3c5/green.json`.
- Build `next-activity-build-01`; static `next-activity-static-01/ratchet-verification.json`.
- Verification `next-activity-focused-verification.json`; smoke
  `next-activity-regression/smoke-d16c3698ecb0439c92d93f6779d659f0/verification.json`;
  layout `next-activity-regression/layout-c479aea610c8429ab08b77a6b45d1eb1`.
- Failed chain `next-activity-regression/chain-d48688a2911448efa86efc11e8a61220`
  retains sanitized native failure and assisted-preservation receipts.
- Ordered live comparison `projection-live-control/77043e5b14084a159a6bf83d2ee567dc/scope-verification.json`.
- Assisted RED `production-run-local-controller/4067d70c06da4187899cc89460aa28f7`.

## Prior binding checkpoint

2026-10-04. Existing Architecture v4.11 D18 captured warehouse/session/workbook
rule; Operations `frmProduction.mBtnManagerNext_Click`. Observation integration
and full acceptance remain open. Candidate: `deploy/validation-production-next-binding-01`.

The actual packaged handler advanced a completed reusable batch after a target
change, replacement sign-in or sign-out. Its existing typed owner guard now
refuses those entries before changing owner state or projection. A current
context still advances once. No activity registration or business owner changed.

Reproduce using `tests/tooling/Test-Slice4beProductionRunLocal.ps1` with
`-CompleteBaselineOnly -NextBaselineOnly -Phase RED|GREEN -DeployRoot <candidate>`.
RED uses frozen `deploy/validation-production-complete-entry-02`; GREEN uses
the candidate above. Probes count real owner entry without replacing the handler.

| Evidence | Result |
|---|---|
| Focused RED | 58 PASS / 12 expected FAIL / 70; four binding failures for each stale context |
| Focused GREEN | 70/70, identical ordered checks and all 58 prior passes retained |
| Preservation | Current owner entry once; stale owner/projection unchanged; custom value/formula, decoy, saved operator bytes and other warehouse preserved |
| Build/compile | Five packages compile; 283/284 compiled components unchanged; only Operations/frmProduction differs |
| Packaged smoke | 86/86; prior ordered identities retained |
| Layout | 18 requested size/page pairs, six activated/maximized pages and five native transitions; geometry unchanged |
| Static | 291 components, 6177 procedures, 135555 lines; 9 literal/45 unresolved Application.Run; 28 oversized caps unchanged; three schemas and 382 PowerShell parses pass |
| Full chain | Corrected harness: chain32/live-role48/warehouse15 PASS; earlier native Shipping failure retained below |

Focused RED/GREEN, build, smoke and layout close normally, restore settings,
preserve package pins and pass delayed native-event audits. Two focused GREEN
captures were reviewed (current/target); this does not accept populated long-value
or multiline Production content. Six first-run layout images were reviewed;
minimum/default clamp to the same size and the small output region remains.
Two previews appeared blank, prompting a second layout run; original-resolution
review proves the original PNGs render correctly. No capture-tool defect is
established. Empty-list geometry is not populated/multiline user acceptance.

The first probe attempt failed instrumented compilation because insertion used
ProcStartLine's blank preamble. Anchoring the exact declaration corrected the
harness. That attempt is not product RED; its empty owned Excel needed assisted
cleanup. The full-chain failure likewise is not product RED or desktop error 5.
Its ten recovered test workbooks were closed without save; saved bytes, settings,
tracked reports and packages were preserved. The scoped investigation follows.

The existing ordered-live diagnostic (`-Cut Full -Setup None`) passes all 48 prior
ordered checks on this candidate, with normal closure, preserved settings/report/
packages and zero delayed native events. This isolates a passing workflow without
the preceding setup; it does not establish crash causation or full-chain acceptance.
The actual source-setup prefix then proves tooling RED4 PASS/1 expected FAIL:
source integration succeeds while its Excel process remains alive. Waiting with
the existing cleanup helper before ordered live work yields GREEN5/5. Both retain
identical checks, normal closure, settings/report/package preservation and zero
delayed native events. This changes orchestration only; runtime packages are unchanged.
The original Admin-boundary mode also retains4/4 after the test extension.

The corrected full chain passes32/live-role48/warehouse15 with every prior ordered
check retained, normal unassisted closure, restored settings/reports, pinned packages
and zero delayed native events. New static evidence retains all metrics/28 caps,
three schemas and 382 parses. The sequencing defect is proved; causation of the
earlier native crash is not. Do not repeat these unchanged passing gates.
Next Batch observation integration and the remaining 4be-A/B0 work stay open.

Ignored evidence roots (relative to `reports/runtime/`):

- Focused controllers: `production-run-local-controller/49aeffa130d64567a7efe13eafd73606`
  (RED), `production-run-local-controller/24113bca1515423ab878b7d2cfd84917` (GREEN).
- Focused reports: `slice4be-production-next-baseline/31c6f8916e954b648756941a38bc9fa8/red.json`,
  `slice4be-production-next-baseline/3b185675312f491f97e6c5dd3e85cf64/green.json`;
  `next-binding-focused-verification.json`.
- Build/static: `next-binding-build-01/verification.json`,
  `next-binding-green-static-01/ratchet-verification.json`.
- Smoke: `next-binding-regression/smoke-d13932b6ca72489786960599c4ac1dda/verification.json`.
- Layout: `next-binding-regression/layout-54c394ba707240c3a8fe63d31796b601`
  and `next-binding-regression/layout-4666c4d580384f79863110eed5e902d1`.
- Failed chain: `next-binding-regression/chain-676c70950cba483db488924629f5c7d4`
  (closure and assisted-recovery receipts); 21:00:14–21:04:08 UTC.
- Initial harness failure: `production-run-local-controller/b9f60819994a4d608d255ad3afa499b9`.
- Ordered live control: `projection-live-control/e6afea160c4a40e49147450b5e8277f7/scope-verification.json`.
- Source exit RED/GREEN: `chain-stage-exit/8095d6adb8d04a0a87ea5635d1cbc7d9/red.json`,
  `chain-stage-exit/663e8c054e8c44248328a4f39c93dfe3/green.json`;
  `next-binding-source-exit-verification.json`. Command: `Test-Slice4beChainStageExit.ps1`
  with `-Boundary Source -Phase RED|GREEN`, candidate DeployRoot and focused GREEN
  controller's `package-pins.json` supplied through `-PackagePinsPath`.
- Corrected chain: `next-binding-regression/chain-e482188f1f294a5abefc86c0a41e0759/verification.json`.
- Tooling static: `next-binding-source-exit-static-01/ratchet-verification.json`.
- Retained Admin boundary: `chain-stage-exit/fef8d3616c72493aa8ca39c1b2024bb8/green.json`.
