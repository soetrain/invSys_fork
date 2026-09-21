# Plan 022 Slice 4be guide-to-run expectation binding

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
evidence uses prefix `reports/runtime/guide-binding-red`. GREEN remains pending.

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
The focused packaged GREEN run and visible evidence are pending. Accepted deployment
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
No runtime change, schema change, authority fallback, workflow replay, deployment
or full How-To/Diagnostic/Compare acceptance is authorized by test preparation.
Event Detail's visible scrolling and full-chain failures remain separately open.
