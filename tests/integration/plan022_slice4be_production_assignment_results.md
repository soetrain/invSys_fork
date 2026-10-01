# Slice4be Ingredients Assignment observations

## Governing contract and scope

Architecture v4.11 D18, **Ingredients Assignment observations**, is the normative
discovered-control refinement, synchronized with Plan022, controls8.3 and the
Production tracking audit. The approved semantic-inheritance rule covers these
observations while preserving the existing algorithms. Any different editing,
read, validation or save behavior needs a separate architecture decision.

Catalog23 specifies seven buttons and two deliberate list selections, all owned by
PRODUCTION_ASSIGNMENT. Runtime is still catalog22 with109 IDs and48/68 observed
Production buttons; specification and test adapters do not increase coverage.
The frozen unpromoted candidate is `deploy/validation-process-worksheet-activity-03`.
Its completed worksheet regression checkpoint is recorded separately in
`plan022_slice4be_process_worksheet_activity_results.md` and must remain GREEN.

The protecting baseline invokes all nine existing handlers through disposable
Operations adapters. A released Process with a requirement drives actual selection
and a saved new Process version. Controlled empty/malformed read responses retain
the parser/population path; no activity records or replacement handlers are injected.
The shared harness still verifies actual Settings/UOM callbacks and compiles all
five instrumented packages. Saved packages and runtime source remain unchanged.

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

## Required next evidence

Preserve the349 initial checks and245 expanded companion checks, including the
original89 results. Before runtime edits, finish the explicit Core outcome/source
contract matrix for the nine new mappings: unsupported positive outcomes,
nonempty confirmed/pending exact Designs envelopes, failed submitted/uncertain
references, and rejection of local-action, wrong-kind, wrong-warehouse or malformed
references. The handler, owner-write, policy and binding RED cases above are now
established and must not be dropped or replaced. Do not promote the initial
source-envelope shape assertion into proof of exact owning EventId correlation.
The companion supplies that writer-boundary comparison for its five Save cases.

Complete the remaining focused RED cases before runtime edits, then require GREEN retaining
all checks. Separate guide and observed recordings must prove publication, exact
Event Detail, authored intent, independent-reader How-To/Diagnostic/Compare and
Save's exact Designs applied/awaiting/incomplete distinctions. Build, compile,
layout, static-maintenance, live-role, full Release1 chain, reusable Production,
visible operator evidence and human acceptance remain required. No package
promotion, completed control coverage or Slice4be acceptance is claimed here.
