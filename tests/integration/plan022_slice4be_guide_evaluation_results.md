# Plan 022 Slice 4be guide-to-run expectation binding

**Focused implementation GREEN: 204/204**, including all 166 preceding checks and
all 37 guide-binding checks. Report:
`reports/runtime/slice4be-viewer-published-read/d08d39397d2a43258459b505ed0d3147/green.json`.
UTC **22:29:46.2433071--22:41:53.8291893**, exit 0. Immediate closure is False;
Excel closes normally within the controller's bounded wait, without intervention.
The window has zero Application events 1000/1001/1002. Every RED identity remains.
Exact guide/run separation, actual apply/evaluate handlers, integrity/current-policy
guards, independent analysis scope and source non-mutation pass. Broader candidate
regressions remain in progress; this is not full Slice 4be acceptance.

**Visible packaged gate: 214/214**, preserving all focused 204 and preceding
visible 177 identities. Report:
`reports/runtime/slice4be-viewer-published-read/57bbeed83e4043fe96b6b1c7355ef546/green.json`.
UTC **22:54:30.9577355--23:04:24.2971573** on 2026-09-21, exit 0, normal Excel
closure without intervention, zero Application events 1000/1001/1002. All 15
images were directly reviewed: guide revisions, published reader, expectation
editor layouts/staging, four guide/run binding layouts, staged binding and the
separate evaluated result. The result displays exact guide/version intent beside
the different selected recording and original observations. Minimum, default,
larger and restored controls are readable. The staged reader capture shows its
generic published-guide notice after activation; the separate evaluation image
shows the guide expectation and result. No persistent staged-notice claim follows.

The guarded owned-caption input is exercised for the first capture; the remaining
14 need no caption input. Each image retains the strict foreground guard. The
helper validates the form's Excel ownership and visible/enabled state, hit-tests
its caption before input, guarantees mouse release, and restores its original
topmost flag in finally. It changes no runtime or global focus policy. This is
test capture support under existing visible-evidence requirements, not a new
product contract or substitute for physical human UAT.

Two visible attempts stop at **83 PASS / one harness failure**, before the new
binding captures: `51e475f8ce464a75abc36f21c844079f` fails the strict foreground
guard, and `e607a410d9f9441b97506d0d27c1d0f7` finds no uncovered owned caption point.
Both close Excel normally. The first log identifies Code as the foreground
process despite a successful activation return; it does not establish the cause
of the missing owned hit-test point. These are capture failures, not product RED
or regressions in the completed 204 checks. A test-only retry temporarily raises
the uniquely owned form, uses guarded native input, restores its topmost state,
and retains the strict foreground capture check. No runtime or workstation focus
policy changes follow from these observations.

**Focused behavioral RED verified, 2026-09-21: 171 PASS / 33 expected FAIL.**
Report: `reports/runtime/slice4be-viewer-published-read/465656eddaec49998b2c6d5a80df5a49/red.json`.
All 166 preceding guide/expectation identities pass; the saved-probe check also
passes. Of the 37 new checks, four non-mutation checks pass and 33 missing-control/
guide-binding checks fail. No harness failure or duplicate identity occurs.
The frozen overflow candidate has no guide-binding implementation.

UTC start **22:08:44.2478089**, terminal **22:20:00.5260915**. Immediate closure
is False; normal closure is verified at **22:20:26.5316097**, without Quit,
recovery action or forced termination. A read-only ROT attachment attempt during
shutdown returned unavailable; it did not establish a workbook count or change
Excel. The extended window has zero Application events 1000/1001/1002. Exact ignored
evidence uses prefix `reports/runtime/guide-binding-red`.

**Isolated implementation candidate:** `deploy/validation-guide-binding` builds
and all five packages compile. Against the frozen overflow candidate, 234 compiled
components become 235: six existing components change and private Core
`modGuideExpectation` is added; none is removed. Changes are confined to Core's
guide reader/store/intent/evaluation boundary and the two Operations library forms.
Core validates exact guide chains and current expectation visibility before use,
commit and saved display. Explicit form entry captures the observed run; Use stages
intent and clears prior displayed evaluation, while Evaluate remains separate.

Static evidence `reports/runtime/guide-binding-static` records 242 components,
5,970 procedures and 131,528 lines (+183 versus overflow). Duplicate groups stay
192; literal/unresolved dynamic calls stay 9/45. All 28 preceding module-size limits
pass. Scanner/reviewed candidates are 1,185/1,187, with no deletion approval.
The focused and visible packaged GREEN runs pass. Accepted deployment
and the seven preceding frozen candidate sets remain outside this build.

Architecture v4.11 D18's explicit guide-to-run refinement names the Published
guides `lblGuideObservedRun` and **Use for selected run** action. These implement
the already approved separation of exact guide intent from an explicitly selected
observed journal. Reading or applying a guide cannot choose its source run,
infer success, evaluate automatically, or grant maintenance/workflow capability.

`tests/tooling/Slice4beGuideEvaluation.ps1` prepares actual packaged handler tests.
Existing Admin recording and guide expectation/save controls supply two immutable
guide versions. A separate recording supplies the observed actions. The tests use
the existing generic form-control probe and ordinary reader access, then inspect
immutable evaluation output and byte preservation. Missing fixtures or existing
entry failures throw as harness problems; missing new controls/binding fail checks.

The 37-check draft covers exact observed-run provenance, explicit application,
browsing/refresh/close continuity, guide version/hash and expectation StepIds,
original observed ActivityIds, separate Evaluate, stale-run rejection, explicit
reopening, context loss and source/guide/config/activity preservation. It also
checks None intent, an empty observed run, analysis scope, corrupt guide evidence
before use/evaluation/saved display, and current-policy restriction. Disposable
corruption is restored byte-for-byte; intentional policy fixture commands are
separated from guide-action non-mutation checks. The opt-in `-CheckGuideEvaluation`
includes existing guide-expectation prerequisites and runs after those checks.
Behavioral RED has now run on the frozen candidate.
Retain all preceding guide/expectation identities and the overflow candidate's
independent tests. Implementation must turn these failures green without removing
the preceding checks.

Operations owns the controls. Core owns selected intent, exact immutable guide
validation and evaluation persistence through primitive cross-XLAM boundaries.
No schema change, authority fallback, workflow replay, deployment or full
How-To/Diagnostic/Compare acceptance follows from this bounded implementation.
Event Detail's visible scrolling remains separately open; its latest chain is an
assisted functional pass, with earlier native failures retained in its evidence.
