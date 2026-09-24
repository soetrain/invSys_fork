# Slice 4be paired-view layout stability

Last verified: 2026-09-24 UTC. **Acceptance incomplete.** Architecture v4.11 D18
requires current visibility and exact integrity on every view action and
activation. Repeated control-generated layout notifications at unchanged form
dimensions are not additional operator resizes. This correction preserves all
existing validation on opening, activation, method changes, Refresh and resizing;
it introduces no policy cache, new authority or contract amendment.

## Diagnosis and protecting RED

Disposable Operations/Core copies were instrumented with fixed procedure labels
and clock readings, without arguments or operational values. Five projects
compiled. Controller `guide-resource-diagnostic/039c0051a17240d598749463e65718f2`
and packaged report `slice4be-viewer-published-read/cb0ead5a24d24a02977fa2a1dca97ddd`
under `reports/runtime/` record **34/34**, normal unassisted closure, restored
settings and unchanged frozen packages, 00:58:08--01:04:03 UTC.

There are 28 `frmActionPathView.UserForm_Layout` entries consuming 166.801 seconds
inclusively. The fresh selected-method boundary takes 103.003 seconds, including
97.688 seconds before Excel reports ready. Repeated layout callbacks perform
complete guide/policy reads during that wait. Inclusive procedure totals overlap
and must not be summed. Instrumented timing is diagnostic, not an operator SLA.
All three images were individually reviewed; a Task Manager thumbnail obscures
the lower-right form area, so none establishes full-control visible acceptance.

The focused test calls the actual packaged `UserForm_Layout` handler twice at
unchanged dimensions, then changes and restores width through that same handler.
It counts actual `RefreshView` entries without altering reader behavior.
Controller `guide-resource-diagnostic/79120c3b1c974e1db2390a521fad561e`, packaged
report `slice4be-viewer-published-read/56bb3dfbd37a44b7b3237bd5d470e38f`, records
**31 PASS / one expected behavioral FAIL / one capture exception**,
01:04:59--01:09:31 UTC. Unchanged dimensions produce **two** unwanted refreshes;
the real-resize check passes with **two** refreshes. The final screenshot fails
because Windows rejects cursor positioning. Four later preservation assertions
are unreached. Excel closes normally and original settings/packages are preserved.
The capture exception is not product RED.

## Correction and build isolation

The four-line runtime change remembers dimensions after a successful layout and
ignores subsequent layout notifications at those same dimensions. The existing
loading/layout reentrancy guards and every real validation call remain intact.

An initial archive-based build is rejected before GREEN: its CRLF-only Production
source fails the builder's test-marker count regex, retaining test-only regions,
and the archive omits the uncommitted Event Detail measurement already present in
the frozen package. The working source's mixed line endings do trigger stripping.
This is a build-input calibration finding, not an accepted Production change.
The rejected candidate remains `deploy/validation-guide-layout`; the subsequent
[test-first builder correction](plan022_crlf_build_regions_results.md) protects
CRLF marker handling independently of this runtime change.

The corrected isolated source copy preserves the verified working bytes, including
the frozen Event Detail behavior, without editing or committing the user's file.
Candidate `deploy/validation-guide-layout-normalized` builds all five XLAMs and
passes explicit compile and Operations cold-start dependency checks. Comparison
of all **243 compiled components**, including separate string-literal hashes,
finds only **frmActionPathView** changed. The accepted deployment is untouched.

An attempted capture adjustment moved the pointer to a verified owned caption even
when the form already had focus, with a desktop paint interval. A short calibration
still fails at cursor positioning (`capture-foreground-calibration/caa6311967e345b79c3a1fd1d5c96d00`),
closes normally, and proves no product behavior. The unverified adjustment is
reverted; the previously accepted capture helper is preserved.
Input-desktop access alone does not prove cursor/capture access. Visible retries
remain suspended pending a changed desktop condition.

## Focused GREEN and remaining gates

With Excel visible and captures explicitly disabled, controller
`guide-resource-diagnostic/68e0d4e7d0a84bc5b569f16bfd71a63a` and packaged report
`slice4be-viewer-published-read/748cd4fab8bc422c8414fae78839fa80` pass **36/36**,
01:15:08--01:19:31 UTC. All **34 preceding identities** are retained. The unchanged
dimensions check now observes **zero** refreshes; real resizing still observes
**two**. Both Excel sessions close normally without assistance, original settings
are restored, saved workbook/training/Config/probe bytes remain protected, and
the 01:19:48 UTC audit finds zero Application events 1000/1001/1002. All five new
and five frozen package hashes hold; only the intended form changes among the
299 source pins. Receipt: `reports/runtime/guide-layout-green-verification.json`.

The fresh selected-method boundary takes **8.355 seconds**, including **8.313**
before ready. This run has no screenshot activation and cannot establish complete
operator responsiveness or a timing-equivalent comparison with the earlier trace.
The deterministic handler assertion establishes the redundant-read correction.
No current screenshots are accepted and the separate cursor failure remains open.

Regenerated maintenance evidence is `reports/runtime/guide-layout-static`:
250 components, 6,045 procedures, 132,897 lines (**+4** for this explicit guard),
unchanged 9 literal/45 unresolved dynamic calls and 193 duplicate groups. All 28
existing size limits pass. No new component, procedure, dependency or dynamic call
is introduced; the four supporting lines are the complete runtime change.
All three report schemas validate and all 260 PowerShell files parse.

The new candidate's ordered Release 1 chain passes **32/32**, live roles **48/48**
and Create Warehouse **15/15**, retaining every prior identity. Interval:
01:20:16--01:26:39 UTC; independent audit: 01:27:06 UTC, zero Application failure
events. Excel closes normally without assistance; local settings and all three
tracked reports are restored. All five candidate package hashes and 260 tooling
pins hold at that verification. Receipts:
`reports/runtime/guide-layout-chain-verification.json` and the three copied
`guide-layout-chain-*.md` reports. The later four-size layout-test extension changes
only test tooling and does not require another unchanged full-chain run.

The four-size extension then passes **40/40**, retaining every preceding 36
identity and validating control bounds/non-overlap at Minimum, Default, Larger
and Restored sizes through the existing packaged form probe. Controller:
`guide-resource-diagnostic/87d2980cd0cd450c9e1fc7285ea3b161`; packaged report:
`slice4be-viewer-published-read/8090707c419743cea67a514ee26999e7`. Interval:
01:27:44--01:31:54 UTC. Five instrumented projects compile; both sessions close
normally; all preservation checks pass. The 01:32:18 UTC audit finds zero
Application failure events. The fresh selected-view observation takes **6.879
seconds**, including **6.841** before ready. Captures remain disabled. Receipt:
`reports/runtime/guide-layout-size-green-verification.json`. All 260 PowerShell
files parse after the test extension. No runtime change follows the verified chain.

Reproduce the focused gate on the isolated candidate with
`Test-Slice4beGuideResourceDiagnostic.ps1 -DeployRoot deploy/validation-guide-layout-normalized -SavedWorkbook -TraceViewCalls -CheckLayoutStability -SkipCapture -Phase GREEN`.
The read-only trace analyzer is `Read-Slice4beGuideViewTrace.ps1 -ReportRoot <report>`.
Omit `-SkipCapture` only after desktop cursor access is independently restored;
the current capture failure is retained, not reclassified as a passing image.

Full visible and remaining applicable Viewer/guide/comparison regressions remain
required. The broader
[Slice 4be checklist](plan022_slice4be_remaining_acceptance.md)
and the earlier [resource investigation](plan022_slice4be_gui_resource_results.md)
retain their existing failures and scope.

## Desktop access recheck, 2026-09-24 UTC

At 03:01:20 UTC, GetCursorPos succeeds with error zero; the thread and input
desktop are both Default on WinSta0. The test shell and Excel both run at Medium
integrity without elevation. The user reports using both the physical console
and RDP. Session switching is a hypothesis, not a proven cause of earlier error5.
No Windows permission, elevation, desktop ACL or application change was made.
Microsoft documents the current-input-desktop and window-station access
requirements for [GetCursorPos](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-getcursorpos).

Changed desktop access justifies one attempt on the current palette candidate.
Controller `guide-resource-diagnostic/ce214dc1673f4660b83d42812cd201ce`, report
`slice4be-viewer-published-read/1b8abd874a3940a4a54dcdbca9d6e15d`, records **10 PASS /
one harness exception**, 03:02:34--03:03:45 UTC. Five instrumented projects compile
and actual publication/source-journal fixtures pass, then Create guide dispatch
is unavailable during Initialize-GuideRestartFixture. Restart/layout/capture checks
are unreached. This is neither meaningful product RED nor visible acceptance.
Excel closes without assistance; settings and five package bytes are preserved.
Do not repeat the broad attempt unchanged; inspect the exact Create guide return
and its captured context/control state before designing a focused regression.

The independent existing `Test-Slice4beCaptureForeground.ps1` calibration then
captures **all three cases** (hidden-first, visible, hidden-restored),
03:04:30--03:04:35 UTC, using only a blank workbook and disposable form. All three
images were individually inspected: the fixture is readable and unobscured.
Excel closes normally without termination. This proves current capture access
without elevation; it does not prove sustained access or invSys visible acceptance.
Report: `capture-foreground-calibration/41584aa806624030b24630a7e888a297`.

Receipt `reports/runtime/desktop-access-restored-verification.json` verifies zero
Application 1000/1001/1002 events at 03:05:30 UTC, closed Excel, unchanged candidate
packages and only the three previously recorded changes among 299 runtime pins.
No runtime, test helper, contract or accepted test predicate changes in this check.
Keep one desktop connection active during visible tests and recheck input access
after connection changes. Permanent resolution of the intermittent error remains
unproven; administrator access is not required by the successful calibration.
