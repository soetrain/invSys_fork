# Plan 022 Slice 4be guide-to-run expectation binding

**D13 test preparation only, 2026-09-21. No RED/GREEN or implementation claimed.**

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

The 31-check draft covers exact observed-run provenance, explicit application,
browsing/refresh/close continuity, guide version/hash and expectation StepIds,
original observed ActivityIds, separate Evaluate, stale-run rejection, explicit
reopening, context loss and source/guide/config/activity preservation. It also
checks None intent, an empty observed run, analysis scope, corrupt guide evidence
before use/evaluation/saved display, and current-policy restriction. Disposable
corruption is restored byte-for-byte; intentional policy fixture commands are
separated from guide-action non-mutation checks. Before implementation, connect
the opt-in gate and run behavioral RED on the frozen candidate.
Retain all preceding guide/expectation identities and the overflow candidate's
independent tests. This initial file has only passed PowerShell syntax validation.

Operations owns the controls. Core owns selected intent, exact immutable guide
validation and evaluation persistence through primitive cross-XLAM boundaries.
No runtime change, schema change, authority fallback, workflow replay, deployment
or full How-To/Diagnostic/Compare acceptance is authorized by test preparation.
Event Detail's visible scrolling and full-chain failures remain separately open.
