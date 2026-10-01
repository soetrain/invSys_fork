# Slice4be Ingredients Assignment observations

## Governing contract and scope

Architecture v4.11 D18, **Ingredients Assignment observations**, is the normative
discovered-control refinement, synchronized with Plan022, controls8.3 and the
Production tracking audit. The approved semantic-inheritance rule covers these
observations while preserving the existing algorithms. Any different editing,
read, validation or save behavior needs a separate architecture decision.

Catalog23 specifies seven buttons and two deliberate list selections, all owned by
PRODUCTION_ASSIGNMENT. The baseline was catalog22 with109 IDs and48/68 observed
Production buttons; specification and test adapters alone do not increase coverage.
The preserved baseline is `deploy/validation-process-worksheet-activity-03`.
Its completed worksheet regression checkpoint is recorded separately in
`plan022_slice4be_process_worksheet_activity_results.md` and must remain GREEN.

The protecting baseline invokes all nine existing handlers through disposable
Operations adapters. A released Process with a requirement drives actual selection
and a saved new Process version. Controlled empty/malformed read responses retain
the parser/population path; no activity records or replacement handlers are injected.
The shared harness still verifies actual Settings/UOM callbacks and compiles all
five instrumented packages. RED used unchanged saved packages; the subsequent
implementation candidate and its verification are recorded below.

Protecting files:

- `tests/tooling/Test-Slice4beProductionAssignment.ps1`
- `tests/tooling/Slice4beProductionAssignmentProbe.ps1`
- `tests/tooling/Slice4beProductionAssignmentActivity.ps1`
- companion `-Safety` gate in `Slice4beProductionAssignmentSafety.ps1`
- policy, Save guards and native closure in `Slice4beProductionAssignmentPolicy.ps1`
- opt-in `CheckProductionAssignment` in `Test-Slice4beConfigCommands.ps1`

The read fixture/probe is reused for released definitions, controlled source reads
and existing catalog test boundaries. Setup is excluded from the measured user
action intervals. Navigation collection is explicitly enabled by an authorized
disposable policy; production defaults remain unchanged.

## Excluded harness attempt

Controller `reports/runtime/production-assignment-controller/cc6346d26b6640b08e582b13b08f4e9d`,
2026-10-01 00:10:50.8004267--00:11:00.1404579 UTC, stopped during fixture installation:
the released-fixture source anchor did not match. Result
`reports/runtime/slice4be-production-assignment/f86bf63c303a468e8ea2cc79de9a6d89/red.json`
contains one package-identity PASS and one harness FAIL; no compile or behavioral
RED is claimed. Normal closure restored settings and preserved all package pins.
The adapter now tolerates VBA capitalization/outer whitespace while still requiring
one matching statement. The following attempt uses that corrected fixture.

The second attempt, controller
`reports/runtime/production-assignment-controller/e9dd2c8e7068465bbd183692e080cde4`,
00:11:27.6893541--00:13:03.8767506 UTC, compiled all five packages but failed its
expanded released-design setup. Result
`reports/runtime/slice4be-production-assignment/becdc6fa07884d1da7e13c2396cfd24a/red.json`
is39 PASS/2 FAIL and remains excluded from behavioral RED. The shared read setup
also creates/releases a Recipe; Assignment requires only a released Process.
The corrected adapter prepares that Process directly and retains the same required
requirement/output/save/release checks. The second attempt's exact failing setup
stage was not captured, so removal of the unnecessary Recipe is not proof of that
attempt's root cause. Settings/packages and normal closure were preserved.

## Initial packaged behavioral RED

Controller `reports/runtime/production-assignment-controller/c4e881eacbed418aaed64141d388ae81`,
2026-10-01 00:14:14.2920314--00:16:25.9743374 UTC; result
`reports/runtime/slice4be-production-assignment/fc53d62867f64fcfa2f846350175ff65/red.json`.
All349 check names are unique:103 PASS/246 FAIL. All42 shared checks and all nine
existing-owner-behavior checks pass, including a saved new DRAFT Process version
with the expected alternative. The failures separate into11 catalog assertions,
168 missing-observation assertions and67 context/re-entrancy assertions. There is
no harness exception in this result. The initial RED therefore identifies missing
observation and guard behavior without claiming an existing editing/save defect.

Refused actions, duplicate/no-match behavior, valid-empty Process presentation,
quiet no-op preservation, authority preservation before Save and after refusals,
unknown workbook values/formula/bytes and older activity records retain their
checks. Five instrumented compiles, normal cleanup, package/settings preservation
and a delayed Application1000/1001/1002 audit with zero Excel failures pass.
The controller's ignored `verification.json` records these counts and boundaries.
Desktop monitoring reports no error5; no runtime implementation or package changes
were made. This is initial RED, not the complete protecting suite or acceptance.

Regenerated static evidence: `reports/runtime/assignment-baseline-static` retains
277 components,6137 procedures,134745 lines,9 literal/45 unresolved dynamic calls,
190 duplicate bodies and28 non-growing oversized caps. All three schemas validate;
all350 PowerShell scripts parse. Runtime metrics match the prior worksheet
regression baseline. Runtime reports, fixtures and captures remain ignored.

## Save owner-boundary companion

The separate `-Safety` gate reuses the established lifecycle writer/processor fault
seams and invokes the actual Assignment Save handler. It distinguishes a confirmed
save, failure before append, an uncertain acknowledgement after append, processing
failure after submission and failure before the form refresh. Each case uses its
own queue root inside the disposable fixture, so pending events cannot be consumed
by later cases. The expected source reference is compared with the writer's exact
observed EventId and the owning Designs publication-source rows; status text is
not parsed for identity or application evidence. The initial349 gate is retained
separately and is not replaced by this companion.

First companion attempt is excluded: controller
`reports/runtime/production-assignment-controller/173b1ff018fb476483a4e070bb440fbf`,
result `reports/runtime/slice4be-production-assignment-safety/948870a582974b9ea888fc7fc8caf037/red.json`,
one package-identity PASS and one harness compile FAIL. The lifecycle installer
already installs its fault seams; calling that installer again duplicated Core
declarations. The redundant installation is removed. No behavioral RED is claimed.
Its empty owned Excel process remained after harness cleanup and an additional
Quit. After verifying the same process identity and zero workbooks, explicit
termination allowed settings restoration to finish. Both intervention records are
retained in the controller. This is not normal closure or a desktop-error5 event;
settings and all five package hashes were verified before the corrected run.

Corrected companion RED: controller
`reports/runtime/production-assignment-controller/29b53fb3f96d42cb974c666117d58b1e`,
2026-10-01 00:24:47.1408550--00:26:24.9741873 UTC; result
`reports/runtime/slice4be-production-assignment-safety/4e392b8ec4d945c9921dc14bba8b5c11/red.json`.
All89 names are unique:59 PASS/30 FAIL. The42 shared checks retain exact baseline
order and pass; all five actual owner boundaries and their actual write/application
facts pass. The30 failures are the six observation/reference/correlation/effect/
redaction/terminal assertions in each mode. There are no harness exceptions.
This protects CONFIRMED with the exact submitted/applied owning ID, FAILED without
an unused pre-append ID, FAILED with an uncertain written ID, and PENDING with
Submitted references after processing or refresh failure. Applied source evidence
after the refresh failure does not turn its command outcome into CONFIRMED.

Five instrumented compiles, normal cleanup without intervention, package/settings
preservation, prior-record and unknown-workbook preservation, and delayed Excel
Application1000/1001/1002 audit pass. The controller's `verification.json` records
those checks. No runtime/package change or human acceptance is claimed.
Regenerated `reports/runtime/assignment-safety-static` retains277 components,
6137 procedures,134745 lines,9/45 dynamic calls,190 duplicate bodies and28
non-growing oversized caps. Three schemas validate and351 scripts parse.

## Policy and Save-guard expansion

`Slice4beProductionAssignmentPolicy.ps1` expands the companion through the real
nine handlers: capability denial, disabled/older/unavailable policy, optional
Navigation defaults and programmatic selection suppression. Save adds loading,
busy and actual nested-handler entry; failure after local preparation; empty-load
validation; changed target/session, sign-out and closed workbook. A disposable
writer seam signs out only after a real successful append, protecting the original
context and preventing further refresh or an outcome under a replacement session.
No observations, replacement actions or submission results are fabricated.

First expanded attempt remains excluded: controller
`reports/runtime/production-assignment-controller/12df26b37a8a445484f462dd51695fa0`,
2026-10-01 00:31:23.9240279--00:37:18.9145437 UTC; result
`reports/runtime/slice4be-production-assignment-safety/4a4468770c7c4dadbd9a7214d54db77a/red.json`
contains145 PASS/88 FAIL, including a terminal harness exception. All five packages
compiled, but invoking the old form reference after closing its captured workbook
raised VBA automation80010007. The owned dialog was inspected; an attempted End
post returned a PowerShell binder error, so successful delivery is not established.
The exact verified disposable Excel process was terminated to release the worker
and restore settings. Package/settings preservation passes; this was assisted
cleanup, not normal closure, product RED or Windows desktop error5.

The corrected fixture follows the already established worksheet closure test:
measure the real visible form before and after workbook closure. If it survives,
an outer trapped adapter must enter the actual Save handler and preserve state.
If it is dismissed, assert no save/submission/activity and supplement that measured
native boundary with the existing typed binding check. A dismissed surface is
explicitly recorded as no handler invocation. No native-lifetime repair is claimed.

Corrected expanded RED: controller
`reports/runtime/production-assignment-controller/adb8f7068fbe4ac7bbbff89d8b8abafb`,
2026-10-01 00:38:02.9766351--00:41:34.8432166 UTC; result
`reports/runtime/slice4be-production-assignment-safety/79661a1e0e554ef190e84ab68391e62f/red.json`.
All245 names are unique:158 PASS/87 FAIL. All89 prior companion results and their
relative order are exact, including the42 shared GREEN checks. The156 additions
pass99/fail57:25 denial,9 missing unavailable notices,2 Navigation defaults and21
Save observation/guard assertions fail. The prior30 owner-observation failures
remain; no harness exception or compile failure is counted as behavioral RED.

All authorized actions continue with disabled/older/unavailable tracking; the two
Navigation actions also continue with collection off. Programmatic setup emits
no user activity. The real nested Save handler is entered, and the post-append
sign-out boundary observes one actual successful append. Current missing guards
permit a second nested write and later reads after sign-out. Both partial-load
cases preserve their existing local replacement effects without owner writes or
new Designs source events; restored guards pass. These observations protect the
normative distinctions without interpreting a missing activity pair as rollback.

The native closure record shows the visible form dismissed, workbook count5->4,
captured workbook closed and the independent workbook still open. HandlerInvoked
is false; all six closure assertions pass, including typed rejection, no write,
no redirected activity and saved bytes preserved. Five compiles, normal closure
without intervention, settings/package pins, unknown columns/formula/bytes, older
records and delayed Excel Application1000/1001/1002 audit pass. The controller's
`verification.json` records the boundaries and exact prior-result comparison.

Regenerated `reports/runtime/assignment-policy-static` retains277 components,
6137 procedures,134745 lines,9 literal/45 unresolved calls,190 duplicate bodies
and28 non-growing caps. Three schemas validate and352 scripts parse. No runtime
source, frozen package, completed coverage or human-acceptance change is claimed.

## Core matrix and continuation RED

The packaged Core outcome/reference matrix supplements actual-handler coverage.
Its schema fixtures are never presented as real activity or owning application
evidence. It covers each exact positive outcome, unsupported outcomes, completion,
wrong owner/catalog, empty versus required source references, Submitted/Unknown
states, wrong kind/warehouse, malformed and duplicate IDs, and field shape.

Controller `reports/runtime/production-assignment-controller/5bf6b39b3c154c2fb29571423a3a5676`,
2026-10-01 00:43:56.7865296--00:47:30.5355386 UTC; result
`reports/runtime/slice4be-production-assignment-safety/7d8071061d654d7aaf4953d64e0e3c02/red.json`:
532 PASS/228 FAIL/760. All245 prior results/order remain exact. The515 additions
pass374/fail141 because the specified catalog23 mappings are absent.

The next expansion adds sign-out after real owning read returns for Refresh,
Select Process, Processes selection and Save, plus a processor-continuation check
after the already protected actual queue append. Controller
`reports/runtime/production-assignment-controller/2b62d4b3b1024c4fb71a974d4ff2939e`,
00:47:47.9284514--00:51:43.3541811 UTC; result
`reports/runtime/slice4be-production-assignment-safety/bee717f098ae4f20ba2a234d52d1531a/red.json`:
546 PASS/243 FAIL/789. All760 prior results/order remain exact;29 additions
pass14/fail15. Real read/append boundaries are reached. Failures expose subsequent
reads, draft changes, missing refusal/original-attempt evidence and processor
continuation after sign-out. The loaded frozen worksheet candidate remains
unchanged; Core implementation began after its760-check RED, and Operations
implementation followed the protecting boundary failures.

Both runs compile five instrumented packages, close normally without intervention,
preserve settings/packages and pass delayed Excel Application1000/1001/1002 audits.
The closure records retain measured native dismissal with no handler invocation.
`Slice4beProductionAssignmentYield.ps1` contains the isolated read-boundary seams.

## First implementation candidate

`deploy/validation-production-assignment-01` is unpromoted. Core adds the explicit
catalog23 controls, fixed outcomes, Save-only Designs references and exact terminal
mapping. Operations wraps all nine actual handlers in typed observation/permission/
binding coordination. Shared local alternative-list algorithms move into
`modProductionAssignmentDraft`; identities, comparison rules and saved payloads
remain unchanged. `modProductionAssignmentActions` reuses the existing Operations
captured-action guard, which has no worksheet mutation authority.

A form-scoped continuation exists only while an Assignment action runs and clears
on normal/error exit. Shared reads check it before subsequent work. Save passes
typed owning facts and a bound continuation through queue, processing, projection
and refresh boundaries. Other callers leave that continuation unbound, preserving
their existing behavior. Original status/error presentation, quiet no-op behavior,
partial local preparation, optional tracking and submission uncertainty remain.
No source identity is extracted from status text or processor totals.

Companion GREEN: controller
`reports/runtime/production-assignment-controller/08763b11e54546ee8fd2ca63e921d4d9`,
2026-10-01 00:54:08.7529232--01:00:57.1159322 UTC; result
`reports/runtime/slice4be-production-assignment-safety/884c7493a9d748b59a2d81e305dd64d7/green.json`:
789/789, retaining the exact RED order and every earlier check. Five instrumented
compiles, normal closure without intervention, settings/package preservation and
delayed Application1000/1001/1002 audit pass. The native form is dismissed at
workbook closure; no closed-form Save invocation is claimed.

Original actual-handler baseline GREEN: controller
`reports/runtime/production-assignment-controller/7d9dcc62988d4eb6a87bd58d2ca527fd`,
2026-10-01 01:01:17.2729127--01:04:03.2810469 UTC; result
`reports/runtime/slice4be-production-assignment/ebbfc739e1c24e89839526a4c81ddbf6/green.json`:
349/349 with exact prior order, all nine normal action pairs and existing owner
behavior, frozen catalog definitions, refusals, local/saved state and binding
guards. Unlike the companion's visible native dismissal, this baseline's retained
form reference enters the eight closed-binding local handlers; all refuse and
preserve state. Five compiles, normal closure, settings/package preservation and
delayed Application1000/1001/1002 audit pass. Registered runtime coverage is now
118 IDs,55/68 constructed Production buttons and two newly observed nonbutton
handlers;13 buttons and28 nonbutton handlers retain coverage work. These are
focused observations, not completed Slice4be user acceptance.

Build log: `reports/runtime/assignment-build-01.log`. Independent cold-start and
five-package compile: `reports/runtime/assignment-compile-01`,2026-10-01
01:04:40.4097142--01:04:54.7890014 UTC. Settings, both candidate package sets,
normal closure and delayed Excel audit pass. Of273 compiled components,264 match
the prior270 by case-insensitive code and exact string-literal hashes (252 exact
code hashes,12 VBA identifier-case-only differences). Six existing components
change: Core catalog/references/evaluation and Operations form/lifecycle-facts/
reusable-design owner. The three new modules are Assignment codes, actions and
draft-list helpers. No Inventory Domain, Designs Domain or Admin behavior change
is inferred from packaging alone; their required regression gates remain open.

Existing worksheet Core regression retains188/188 exact ordered checks on this
candidate: controller
`reports/runtime/process-worksheet-activity-controller/bdd0e7a1c3c04773bbadb50aac12677c`,
2026-10-01 01:07:36.3938313--01:08:57.0012884 UTC; result
`reports/runtime/slice4be-process-worksheet-activity/e506069a6201418fae1181e31d9f9f15/green.json`.
Five instrumented compiles, normal closure, package/settings preservation and
delayed Excel audit pass. This supplements the new Assignment contract matrix;
it does not replace the full595-check worksheet workflow regression.

Regenerated `reports/runtime/assignment-implementation-static-01` has280 components,
6150 procedures and135006 lines. Dynamic calls remain9 literal/45 unresolved,
duplicate bodies remain190, and all28 oversized caps do not grow. The Production
form shrinks11599->11564 lines; total runtime growth is261 lines across the typed
coordination and explicit catalog. Three schemas validate;353 scripts parse.

## Required next evidence

Preserve the349 initial checks and789 expanded companion checks, including every
original89/245/760 result identity. The contract matrix and continuation RED above
are established. Do not promote the initial
source-envelope shape assertion into proof of exact owning EventId correlation.
The companion supplies that writer-boundary comparison for its five Save cases.

Require focused GREEN retaining all checks. Separate guide and observed recordings must prove publication, exact
Event Detail, authored intent, independent-reader How-To/Diagnostic/Compare and
Save's exact Designs applied/awaiting/incomplete distinctions. Build, compile,
layout, static-maintenance, live-role, full Release1 chain, reusable Production,
visible operator evidence and human acceptance remain required. No package
promotion, comprehensive control coverage or Slice4be acceptance is claimed here.
