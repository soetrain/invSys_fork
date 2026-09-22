# Plan 022 Slice 4be.5 - reopen a published guide for editing

Status: focused behavioral RED verified; implementation pending. No runtime
implementation, deployment, visible acceptance or completion is claimed.

Architecture v4.11 D18's **published-guide editing refinement** governs this
discovered entry under the approved semantic-inheritance rule. It was recorded
with Plan 022 and controls in documentation commit `58e5f58` before this test.
The existing guide editing, ACTION_PATH_MAINT, immutable revision, exact identity,
current-policy, headless Core and captured-context rules remain binding. Neither
pending Event Detail nor D8-A proposal is approved by this work.

The actual packaged entry is Published guides `btnEditPublishedGuide`, **Edit
guide**, for the selected exact guide ID/version/hash. The test uses the existing
real recording/Create guide/Save guide/expectation fixtures and a different observed
run. It never fabricates a saved guide or treats the observed run as authoring
source. The protecting helper is `tests/tooling/Slice4bePublishedGuideEdit.ps1`,
called from the existing guide-evaluation gate after its presentation regressions.

Initial command, against the unchanged isolated paired-view candidate:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-guide-presentation -Phase RED -GuideDraftOnly -CheckPublishedGuideEdit -CheckViewerPublishedRead -CompileEvaluationProbesForTest -ViewerStartupPackageStateForTest SavedCopies
```

The expected behavioral RED is missing Edit guide entry/reopen/revision behavior.
The 39 new assertions protect entry/capability, exact restored authored/expected
identities, repeated-entry reuse, independent observed-run selection, four editor
sizes, immutable version 3/4 append, source and expectation preservation, reader
selection retention, Close/reopen text restoration, Cancel, real older-version
conflict, selection/parent/child lifecycle, corrupted guide bytes, current policy,
maintenance revocation, changed target and ordinary-reader read/Use access.
Five no-write/read-preservation assertions may already pass without the new entry;
missing functionality must not be disguised as a harness or fixture failure.

All 243 preceding focused GREEN identities must remain. The separate paired
visible 305/305, fresh Operations-only preference restart 27/27 and prior role,
evaluation and full-chain evidence remain baselines, not automatic acceptance of
the forthcoming changed runtime. Build/compile, focused GREEN, packaged layout and
visible evidence, maintenance metrics, relevant regressions, live roles and full
Release 1 chain are still required after implementation.

The broad attempt records **235 PASS / one harness failure**, reaching 235 of the
243 preceding identities, with eight not reached and **zero GuideEdit assertions
reached**. Report:
`reports/runtime/slice4be-viewer-published-read/3374e267f67d4b27b11b0c8985bf5c9c/red.json`.
UTC **2026-09-22 03:30:51.1312998--03:49:27.0124051**. The first rejected call is
the read-only `modWarehouseSync.PublishedReadPublishCallsForTest`, with inner
`0x80010001`; cleanup later reports `0x800AC472`. No owned foreground capture was
available. This is not product RED, and its cause is unproven.

The controller records ExcelClosed=False. The exact test-owned process later
exits without intervention; absence is verified at **03:50:52.3643771 UTC**.
Read-only attachment during cleanup returned MK_E_UNAVAILABLE, and native window
enumeration returned no windows. Neither observation proves an empty workbook
collection; no empty-instance Quit or forced termination was attempted. No
recovery data was changed. Zero matching Application events through that later
observation and unchanged running sources are verified. Retain the original
controller result separately from the later closure observation.

The focused retry uses `-PublishedGuideEditOnly` in the same command instead of
`-CheckPublishedGuideEdit`. Shared fixtures generate two guide revisions and their
explicit expectations through actual authoring/Save handlers, then a different
observed run. The same 39 edit assertions execute without repeating the entire
unchanged presentation chain. Eight existing selection/integrity/policy helpers
were extracted to `Slice4beGuideTestActions.ps1`; their function bodies are verified
identical to the preceding committed versions. No runtime fix or automatic retry
of a rejected command is introduced. Existing full regression evidence remains a
separate required baseline, not replaced by this focused gate.

The focused packaged RED is **17 PASS / 34 expected FAIL**, with 51 unique typed
results and no harness failure. Report:
`reports/runtime/slice4be-viewer-published-read/3b620d3700504fe0ae3e0bff9174fb80/red.json`.
UTC **2026-09-22 03:53:03.4243175--03:55:31.0701799**. All 39 edit assertions run:
five preservation/read assertions already pass, while the 34 failures identify
missing entry/reopen/revision/lifecycle behavior. All twelve compile/startup/actual
fixture assertions pass. Excel closes normally without intervention. Zero matching
Application events, unchanged running sources and unchanged frozen candidate
hashes are verified. This establishes D13 RED before runtime edits; it does not
claim that the earlier COM rejection was fixed or replace the full regression gate.
