# Plan 022 Slice 4be visible-capture diagnosis

**Measured visibility scope:** The published-Viewer test assigns Excel.Visible=True
before the callback, but the focused default-startup trace reads back Boolean
False. Its visible modeless Viewer opens/reuses successfully (15/15). The
early-visible trace reads Boolean True and fails at the SetWarehouse caption
assignment (14 PASS / two startup FAIL). Both have zero ordinary workbooks.
The saved/reopened-workbook comparison also reads Boolean True and fails at the
same boundary (16 PASS / two startup FAIL), preserving the active saved workbook
and its closed-file hash. Application visibility and form visibility are distinct.
The older 166/166 and 204/204 did not measure application visibility at callback
entry; they retain their default-startup scope without a retrospective visibility
claim. The underlying cause remains unresolved.

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
`-GuideDraftOnly -CaptureGuideEvidence` or explicit startup tracing. It verifies that Excel is visible from
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
runtime correction; unchanged default-startup 166/166 and 204/204 retain their scope.

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

Preceding checkpoint verification: six relevant PowerShell scripts parse, 91 local links
resolve, all 30 frozen package hashes and both unrelated user-document hashes
match, and runtime source is unchanged. Invalid visible-Excel option combinations
are rejected before Excel setup. The diagnostic tools and reporting-only labels
change no runtime or architectural contract; no new product RED is manufactured.

**Focused startup tracing:** `-TraceViewerStartupForTest` retains the same
guide-expectation/published-Viewer probes, valid Admin Settings/publication fixture
and ordinary reader context. It compiles disposable copies, invokes the actual
public OpenInventoryViewer callback through a diagnostic error boundary, and
then stops before the wider Viewer/recording/guide workflow. Fixed trace labels
identify existing launcher and SetWarehouse statements; only numeric errors and
fixed stage identifiers are emitted. The wrapper catches errors without changing
the underlying callback, authorizing a fallback or suppressing a product warning.
It checks measured application visibility, opening/reuse, captured context, source-file
preservation, publication/Shipping-authority counters and normal form cleanup.
Trace insertion or fixture/compile failure is a harness failure, not product RED.
Trace success alone is not proof that the uninstrumented failure is fixed.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-guide-expectation -Phase RED `
  -GuideDraftOnly -CheckGuideExpectation -CaptureGuideEvidence `
  -GuideCaptureVisibleExcelForTest -TraceViewerStartupForTest `
  -CheckViewerPublishedRead -CompileViewerProbesForTest
```

Focused reports under `reports/runtime/slice4be-viewer-published-read/`:

| Setup | Report directory / red.json | Result |
|---|---|---|
| Default startup, no ordinary workbook | `f369254c559040169019f7c1f204cef3` | 15/15; Boolean False application visibility; callback opens/reuses. |
| Early visible, no ordinary workbook | `9459de5b9fec4918b487d15884372710` | 14 PASS / two startup FAIL; Boolean True; `ERROR\|440\|Warehouse.Caption`. |
| Early visible, saved/reopened workbook | `27eec2d411ab47a6ad9e2304046ea3a7` | 16 PASS / two startup FAIL; Boolean True; same caption boundary. |
| Saved workbook, extended form-event trace | `ee3e04518ffa46cca999cfa6a6df35e3` | 16 PASS / two startup FAIL; initialization returns; no Layout or Activate entry before caption failure. |
| Saved workbook, Viewer probes without recording/guide probes | `c8fb3e1d398443e48a488eef44116782` | 16 PASS / two startup FAIL; same event trace and caption failure. |

All five disposable instrumented packages compile in these three comparisons.
The real Admin Settings/publication fixture passes. Captured context, warehouse
file bytes/count, publication/Shipping-authority counters and normal form cleanup
pass. The saved-workbook case additionally verifies exact active workbook identity,
Boolean Saved=True and its hash after normal closure. Reuse is marked failed
without attempting a second opening after the first opening fails. These are
startup observations, not guide-expectation RED or full release acceptance.

Earlier setup attempts remain non-product failures: `1bc4c33c661f49e799d4752ea2375c12`
could not locate a unique trace insertion; `e5904a5400dc4673af9955fc876f0a02`
incorrectly required visible application readback before entry; and
`910d617961094a76a17e8c18d1255e92` tried to hash the reopened, Excel-locked fixture.
The corrected trace matches inside its procedure, accepts a typed visibility
observation, and hashes the saved workbook only while closed. Earlier successful
visible trace `2bec01bbb6ed44a4bbdd826482dbeda2` independently reports 14 PASS /
two startup FAIL at the same caption boundary.

The extended disposable trace records bounded fixed stages for Initialize,
Activate and Layout as well as launcher/SetWarehouse entry and return. It ends
at Warehouse.Caption after UserFormInitialize.Return, with no Layout or Activate
entry. Removing recording/guide probes reproduces the same result. Those probes
and these form handlers are therefore not required to reproduce this observed
failure. No runtime statement is removed, deferred or bypassed; the native setter's
underlying cause and any contributing host state remain unresolved.

The extended trace window is **2026-09-15 10:26:16.0674555--10:27:29.7122380 UTC**;
the reduced-probe window is **10:29:02.1245810--10:30:17.2829149 UTC**. All nine
recorded startup attempts, including the setup failures, exit with Excel closed
and have zero Application events 1000/1001/1002 in their measured windows. The
ignored `viewer-startup-comparison-verification.json` records exact windows and
check identities. Eight scripts parse, 91 local links resolve, and all 30 package
hashes plus both unrelated documents remain unchanged. Runtime is unchanged.
Four invalid option combinations are rejected before Excel setup: early visibility
without diagnostic scope, a saved startup workbook without tracing, tracing in
GREEN mode, and tracing without instrumented compilation.

Reduced-probe command (retains actual Settings/publication fixture and callback):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 `
  -RepoRoot . -DeployRoot deploy/validation-guide-expectation -Phase RED `
  -GuideCaptureVisibleExcelForTest -TraceViewerStartupForTest `
  -ViewerStartupSavedWorkbookForTest -CheckViewerPublishedRead -CompileViewerProbesForTest
```

Next discriminator: exercise a blank form's caption setter in this same prepared
Excel session before Viewer entry, retaining the saved workbook and all existing
checks. The earlier standalone blank-form calibration did not use this prepared
session. Do not infer an invSys caption-lifecycle repair from the current trace.
