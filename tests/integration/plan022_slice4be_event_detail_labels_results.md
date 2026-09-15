# Plan 022 Slice 4be.3 activity labels in Event Detail

## Contract and implementation

D18's contributing-line label refinement was synchronized in Architecture v4.11,
Plan 022 and controls before implementation. Under approved semantic inheritance,
the line picker uses each User activity record's published fixed caption and
recorded outcome, separated by ` - `. Missing values remain Unavailable. Other
lines keep their exact System_Key labels, including repeats. The prompt is
**Contributing lines - select a line to inspect its fields**.

Operations `cEventDetailController.LineLabels` uses cached fields only;
`frmEventDetail.Bind` uses that display collection. The existing `Keys` accessor,
source grouping, selected-line positions, published order and profile-permitted
detail values remain unchanged. No new identifier, authority, capability,
observation, canonical schema, workflow action or architecture exception exists.

## Focused RED/GREEN

The test uses actual Admin Settings Save, ordinary Admin publication, Viewer
selection and the visible detail controls. The unchanged Boxing candidate gives
**36 PASS / two expected FAIL**: both activity lines display Unavailable and the
prompt refers only to inventory identities. Exact RED:
`reports/runtime/slice4be-viewer-published-read/3c5bbbe23b4b41cc81a1e258d09bfa0a/red.json`.
The result-directory verifier initially wrapped its saved directory array; removing
that wrapper identified the single new report. No test was rerun or product RED
inferred from that verifier issue.

The corrected isolated `deploy/validation-event-detail-labels` candidate gives
**38/38 GREEN**, preserving every RED identity and all 36 preceding GREENs.
Exact GREEN:
`reports/runtime/slice4be-viewer-published-read/e06daaa074fa4bbd8e8245b2c9aff01a/green.json`.
The added Admin capture is inspected and shows Save Value - REQUESTED and
Save Value - COMPLETED together. Its selected detail remains a REQUESTED
observation, without inventing inventory application or a System_Key.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-event-detail-labels -Phase GREEN `
  -CaptureEvidence -CheckViewerPublishedRead -CompileViewerProbesForTest
```

The compile switch is an alias of the existing instrumented-project compile
option, now also accepted for the published Viewer route. All five instrumented
projects compile before forms in both focused runs. Standalone five-package
build/explicit compile and Operations cold start pass. All 225 compiled identities
remain; only `cEventDetailController` and `frmEventDetail` differ from the frozen
Boxing candidate. Excel closes after each focused run.

Fresh static evidence retains 232 components, 191 duplicate groups, 9/45 dynamic
calls, 1158 scanner candidates and 1160 reviewed candidates. The necessary cached
label method adds one procedure and 17 lines (5881/130098 total). Individual growth
limits all pass. The new candidate and frozen Boxing candidate retain all ten
package hashes; all 325 historical pins and protected source checks also pass.
Both unrelated user documents remain byte-identical. All static maintenance
candidate identities remain unchanged, with all 28 existing module limits met.

## Regression and visible evidence

The serial isolated queue completes these packaged gates:

| Gate | Result | Exact report under `reports/runtime/` |
|---|---|---|
| Event Detail | 34/34 | `slice4be-viewer-detail/a2ce252374724a7cb95b10d32bab689e/green.json` |
| Published Viewer, filters and Shipping state | 94/94 | `slice4be-viewer-published-read/9d628ec50c2440a8834412491b72938c/green.json` |
| Boxing and Shipping | 1707 PASS / seven known D8-A FAIL | `slice4be-shipping-activity/90024d0440a3432da8512d319697fd7f/boxing-activity-shipping-recording-green.json` |
| Saved Action Path evaluation | 376/376 | `slice4be-viewer-published-read/8cb4d5fea3a14f8283327aeb08f78b4d/green.json` |

Boxing retains all 1696 preceding identities and 1689 GREENs. All 698 Boxing
checks pass, including the 18 added prompt/activity-label/exact-key assertions;
all 18 owner-return records remain identical. The seven failures remain precisely
the `Shipping.Access.AuthUnavailable.<action>.MissingFileNotRecreated` checks
for Add, Update, Remove, Hold, Return, Stage and Send. D8-A remains unapproved.
Every preceding 34-check Detail identity and each preserved 56-check Shipping-state
and 59-check filter GREEN set is retained. The evaluation queue validates its
376 passing checks and retains all preceding 374 GREEN identities.

Inspected actual captures include the focused Admin detail, three Event Detail
default/native-maximize/restore views, all six published Boxing detail views and
three published Viewer/Shipping views. Activity labels distinguish actual outcomes;
business labels retain every exact key and repeated line. Published line order is
unchanged, including cases where a result precedes its request in that projection.
The diagnostic journal's separate observation order is not inferred from this picker.

Nineteen evaluation/editor captures are reviewed: thirteen directly inspected,
plus six restored/source views byte-identical to their inspected defaults. Pending
and partial application remain Awaiting published result; complete application
shows Conclusion observed. Each retains Stopped. Capture frozen. Minimum/default/
larger layouts and source views preserve the selected run and saved result.
These are diagnostic recordings and results, not implemented authored guides or
How-To/Compare evidence. Existing Shipping layout is not newly accepted by capture.

The Boxing process's immediate terminal snapshot still saw asynchronous Excel
cleanup. The next evaluation gate passed its enforced no-Excel preflight about
0.17 seconds later. No process was killed or restarted; the native audit must
include that cleanup interval. Evaluation then terminates with Excel closed.

## Full-chain interruption and retained evidence

The first full-chain attempt terminates unsuccessfully: 5 PASS / one harness
exception in its chain report, 32 PASS / one harness exception in its live-role
report, and 15/15 Create Warehouse. The live-role failure follows the successful
projection-delete checkpoint, in the `Delete and rebuild canonical inventory
projections` step, with RPC failure `0x800706BE`. The exact failing call and cause
are unproven. This is neither a meaningful product RED for the display change
nor a passing Release 1 chain.

The controller waits for a residual Excel process before restoring settings.
Read-only COM inspection confirms that process has zero open workbooks and matches
the observed window. Normal Quit reveals Document Recovery; the inspected
**Yes, I want to view these files later** choice retains recovery files without
opening, saving or deleting them. No forced termination occurs. The original
controller then exits, restores local settings and all three tracked reports,
and confirms Excel closure. This cleanup is not a native repair.

Exact first-run evidence is `reports/runtime/event-detail-labels-chain-` plus
`slice14_results.md`, `phase6_live_role_workflow_results.md`,
`create-warehouse-results.md` and `exit.json`; the prefix also retains the recovery
dialog captures and cleanup facts. The first run's exit is 1. The audited windows
for focused RED/GREEN, build, compile, Detail, Viewer, Boxing, evaluation and this
failed chain contain no Excel Event 1000; absence of that event does not negate
the RPC failure or prove clean native execution.

A fresh unchanged-candidate full-chain attempt starts only after the original
session is terminal and Excel is closed. Its evidence uses
`reports/runtime/event-detail-labels-chain-retry-` prefixes. It completes **32/32**
chain checks, **48/48** live-role checks and **15/15** Create Warehouse, with exit 0.
Its cleanup also needs normal Quit of a verified empty residual and the inspected
retain-recovery-files choice. The controller restores settings and all three tracked
reports and exits with Excel closed. This is an assisted cleanup, not proof of
unattended recovery or repair of the first RPC failure.

The retry's verified Application-log window contains no events 1000, 1001 or 1002.
Final candidate/frozen-package, compiled scope, static candidate/growth-limit and
unrelated-change preservation pass again after the run. Exact retry reports use the
same three report suffixes above; `exit.json` and `native.json` record terminal
state and audited scope. All original failure and cleanup evidence remains intact.

## Checkpoint scope and next action

The isolated activity-label correction has focused RED/GREEN, packaged regression,
build/compile, visual/layout, static and functional full-chain evidence. The seven
D8-A failures and unexplained RPC/recovery behavior remain open; no general native
repair, unattended recovery, deployed acceptance or complete Release 1 is claimed.
Continue Slice 4be.5 with actual guide authoring and How-To/Diagnostic/Compare
handlers, preserving this candidate as the baseline for meaningful behavioral RED.

Ignored evidence uses `reports/runtime/event-detail-labels-` prefixes. Earlier
native faults and unapproved D8-A remain separate. Comprehensive remaining control
coverage, Boxing-specific policy/detail proof, combined recordings, authored guide
lifecycle/search/version/export/import, How-To/Diagnostic/Compare, physical deployment
and human Release 1 acceptance remain required. No accepted deployment or operational
workbook was rebuilt; automated captures are not human acceptance.
