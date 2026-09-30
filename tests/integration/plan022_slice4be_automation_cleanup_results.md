# Slice 4be isolated automation cleanup

Last verified:2026-09-30 UTC. This is test-harness work under unchanged Architecture
v4.11/Plan022, not an invSys runtime repair or product behavioral RED. The five
catalog16 packages and VBA source remain unchanged. All runtime paths below are
under ignored `reports/runtime/`; only sanitized findings are committed.

## Defect and protecting calibration

Smoke retains86/86 but its newly observed Initial and Final cleanup both request
termination after the existing1000ms wait. Controller
`production-components-regression/smoke-46bbf224c98e47aa809dc9100711423c`; receipt
`production-components-smoke-baseline-verification.json`.

The initial attempt to reuse completed-host reference release fails74/1 after
confirming zero workbooks; it is removed and safely restored. That attempt does
not identify the primary failure line. Its exact preservation/cleanup evidence
is retained in [component results](plan022_slice4be_production_component_results.md).

`Test-IsolatedAutomationCleanup.ps1 -IncludeExpiredNestedReferences` adds a closed,
released blank workbook to a list and dictionary alongside aliases, a matrix and
a cycle. The unchanged helper reproduces `InvalidComObjectException` at its
two-object `ReferenceEquals` call, line21; calibration
`isolated-automation-cleanup/27416b8e1f664590a7355febda2622dc`.
A self-comparison guard is insufficient: `247479d39ead4cebbca4d2f17166868b` fails
at the same comparison, shifted to line25. Neither failed approach is retained.

The helper now tracks visited object identities through .NET
[ObjectIDGenerator](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.serialization.objectidgenerator?view=netframework-4.8.1),
which identifies already-seen object references. No serialization or object IDs
are written. Packaged observation then identifies another invalid-wrapper failure
at the initial null comparison, line17. Fixed metadata receipt:
`packaged-smoke-closure/f1050cc3f1b34d949f7eb728ed793dd6/Initial-reference-failure.json`,
controller `production-components-regression/smoke-42b228fb06e04224bc64723f77ac4efe`.
The helper also skips `InvalidComObjectException` during retrieval/classification.
It retains alias/cycle handling, unique-reference accounting, repeated-release
safety and fixed counts only. It remains restricted to owned, completed hosts.

The final original calibration is **8/8**, root
`isolated-automation-cleanup/dbeaaf012cbb451a9970252612af50fd`; the expanded calibration
is **9/9**, root `isolated-automation-cleanup/ab92a4316e6c4cba9586065ff94011a5`.
Both exit unassisted with zero release failures. The original five-reference
expectation is unchanged. The added fixture identifies six references; its initial
five-reference expectation was corrected only for that opt-in case. The intermediate
`2307af8e2eaa4c65a9e9baf8e32b4f85` already exited normally but failed that stale count.

## Packaged smoke outcome

With all old workbooks closed and Quit returned, the smoke harness releases its
completed-host variables before the unchanged1000ms wait; replacement Excel is
created afterward. It records release counts, normal/forced exit and bounded
failure metadata (phase, exception type/HResult, source filename/line), never
fixture values, credentials or raw exception text in these sidecars.

Controller `production-components-regression/smoke-d3dde992de54428cbc3914a076ace473`,
01:04:32--01:04:53 UTC, retains **86/86** and every prior check identity. Both hosts
exit unassisted, with no termination request. Each releases19 references with zero
release failures. Settings, all five packages and the tracked report are preserved;
no Excel Application1000/1001 event occurs. Receipt
`production-components-smoke-corrected-verification.json`; shutdown evidence
`packaged-smoke-closure/e22e4091148649aaab409872bc3ad1fc`.

The preceding traversal-only packaged attempt
`production-components-regression/smoke-902db509280a455ea67d5b87ae585b12` fails74/1
before a primary-location receipt exists; the instrumented attempt above also
fails74/1. Both restore settings/packages/reports and leave no Excel process.
These are harness failures, not product RED or acceptance results.

The existing Settings restart consumer retains **202/202** and every preceding
identity, five instrumented compiles, unassisted internal restart/final exit and
settings/package preservation. Controller
`production-components-regression/settings-0d7b5cb67645400fb6378b45cc8d1e13`,
01:05:41--01:09:32 UTC; result
`slice4be-tracking-settings/98bd8df222e84d4faaeb779f6961e1b1/green.json`. Internal
cleanup releases10 references with zero failures and no termination request.
No Excel Application failure event occurs. Receipt
`production-components-settings-cleanup-regression.json`.

All310 tracked PowerShell scripts parse. The smoke workflow-preservation proof
retains every original validator statement and the1000ms wait; only completed-host
release/metadata are inserted (`packaged-smoke-cleanup-workflow-preservation.json`).
Normal reusable Production shutdown, earlier native crashes, comprehensive
coverage, guide transfer and human/NAS acceptance remain separate open requirements.
