# Plan 022 Slice 4be visible-capture diagnosis

This is test-tool diagnosis, not a product contract change or D13 product RED.
The guide-expectation implementation remains code **827a427**. The two packaged
capture failures are retained in the
[guide-expectation evidence](plan022_slice4be_guide_expectation_results.md).
Their foreground observations identify VS Code in front of the intended Excel
form despite successful AppActivate results; they do not establish the cause.

`Test-Slice4beCaptureForeground.ps1` imports the existing three capture helpers
from the root test's parsed functions, creates a blank disposable workbook/form,
and compares hidden, visible and restored-hidden Excel. It opens no invSys
package, changes no package or operational workbook, and leaves the existing
unique-owner and actual-foreground guards unchanged. No input injection, Windows
focus-policy changes, forced termination or recovery-file deletion is used.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beCaptureForeground.ps1 -RepoRoot .
```

Observed 2026-09-15, UTC **09:45:25.3747551--09:45:30.4588641**:

| Fixture state | Capture helper | Direct image review |
|---|---|---|
| Hidden first | Captured; exact owned form in foreground | Full blank form and calibration label readable. |
| Visible | Captured; exact owned form in foreground | Full blank form and calibration label readable. |
| Hidden restored | Captured; exact owned form in foreground | Title bar visible, form body/label absent. Not complete visible evidence. |

All three PNGs are 720 by 293 pixels. A successful native capture and expected
dimensions therefore do not establish complete rendering. All three images were
directly reviewed. The first hidden capture moved foreground from VS Code to the
fixture successfully: hidden Excel alone does not reproduce the packaged failure.
Normal Quit returns and the owned Excel process closes. The measured window has
zero Application events 1000/1001/1002. No product acceptance follows from this
blank fixture.

Ignored evidence root:
`reports/runtime/capture-foreground-calibration/af1803bef2f04edeb944aadd931c0fc6/`.
It contains `observations.json`, `verification.json` and the three PNGs. Window
identities remain ignored; the maintained record contains only sanitized findings.

The packaged discriminator is `-GuideCaptureVisibleExcelForTest`, requiring
`-GuideDraftOnly -CaptureGuideEvidence`. It verifies that Excel is visible from
startup and never toggles visibility between images. Default harness behavior,
capture functions and their guards remain unchanged. This is an explicit fixture
setup comparison on the same frozen candidate, not an implemented foreground
fix. The new setup check adds one harness identity; a complete run would contain
176 checks, including all 166 protecting behaviors and nine captures.

The visible-startup diagnostic instead stops **9 PASS / one harness FAIL** at
Viewer entry, before the guide workflow or any guide image. Its owned Visual
Basic dialog reports **run-time error 440, Automation error**. The dialog image
is directly reviewed; selecting its inspected Debug control exposes an active
Operations project in break mode. Read-only code-pane metadata locates
`modInventoryViewer.OpenInventoryViewer`, instrumented line 209, at
`mInventoryViewer.SetWarehouse warehouseId` (the source's ordinary SetWarehouse
call). This identifies the stopped call, not its underlying cause.

The inspected VBE Standard **Reset** control ends this disposable execution;
the waiting caller then returns `0x800A9C68`. That post-Reset HRESULT is not
substituted for the original 440 dialog. The controller exits 1 with Excel closed,
normal harness cleanup and no forced process termination. No recovery files are
deleted. UTC **2026-09-15 09:47:14.7141703--09:57:27.9017095** has zero Application
events 1000/1001/1002. Report:
`reports/runtime/slice4be-viewer-published-read/5cada388ce034b26b4ca0a771ef9990e/green.json`.
The nine passing checks include explicit visible startup, five instrumented
compiles, probe ordering and two existing published-data fixture checks.

Ignored `reports/runtime/guide-expectation-visible-excel-*` evidence retains the
dialog image, sanitized window/dialog observations, debug location/statement,
reviewed Reset identity/action and diagnostic verification. No VBE source image
or raw source literals are captured. Project mode is read from ActiveVBProject;
an unavailable VBE-level Mode property is not a valid state observation.
All 30 package hashes and both unrelated user documents are preserved. This is
neither guide-expectation RED nor visible acceptance. Investigate the exact
Viewer startup call under this fixture before another capture attempt or any
runtime correction; unchanged hidden-Excel 166/166 and 204/204 retain their scope.

[Microsoft's SetForegroundWindow documentation](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setforegroundwindow)
describes restrictions on foreground activation; it does not identify which
condition caused this failure. Do not weaken the guards or alter workstation
focus settings based on the current evidence.

The full-chain Receiving failure remains separate. Five diagnostic stage labels
now distinguish form-action invocation, outcome read, status read, capability
revalidation and post-action projection inspection. Calls, arguments and existing
assertions are unchanged. The script parses; a later full-chain execution must
establish any new failure location or acceptance. Reporting-only labels require
no manufactured product RED.

Checkpoint verification: six relevant PowerShell scripts parse, 91 local links
resolve, all 30 frozen package hashes and both unrelated user-document hashes
match, and runtime source is unchanged. Invalid visible-Excel option combinations
are rejected before Excel setup. The diagnostic tools and reporting-only labels
change no runtime or architectural contract; no new product RED is manufactured.
