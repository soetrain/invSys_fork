# Plan 022 Slice 4be visible-capture diagnosis

**Latest package-state comparison (2026-09-21):** Writable temporary package
copies with unsaved probes reproduce startup error 440 (**18 PASS / two FAIL**).
Compiling and saving the same probe set in writable temporary copies before any
fixture/form activity passes **20/20**, including actual public Viewer opening,
generation reuse, measured visible Excel, captured context, loaded package paths,
warehouse bytes/counters and the saved/reopened workbook. Neither run performs
blank-form calibration or caption-probing warm-ups. Runtime source is unchanged;
this is a harness-state comparison, not an implemented Viewer or Core fix.
Earlier caption failures must be interpreted with their unsaved instrumentation
state. Full visible acceptance remains open: saved copies clear startup but the
first guide capture still fails, as recorded below.

Two full-gate setup attempts each reach **69 PASS / one harness FAIL**, after
Viewer and guide-draft behavior. `facb63b9c42347958673e05b359f775f/green.json`
tries to open the writable, Excel-locked Core copy as a ZIP to obtain expected
publisher identity. `ab704929229f4fad8766893e2be92a97/green.json` then encounters
a null COM document-property access in an added metadata cross-check. Neither is
product RED. The corrected harness reads expected identity from the immutable
input candidate; the existing saved-guide provenance assertion still requires
Core's published output to retain that exact package version and build identity.
No metadata assertion is relaxed. Both controllers finish with Excel closed and
zero Application events 1000/1001/1002 in their measured windows.

The corrected saved-copy run `f8f8a329b3d34a93a8af309afaed12cb/green.json`
passes **83 checks**, including saved-guide provenance, then fails the first
foreground capture. All three observations still identify VS Code ahead of the
owned guide form despite successful AppActivate. It produces no PNG, exits 1 and
closes Excel, with zero Application events 1000/1001/1002 in its measured window.
Saving probes resolves the observed startup failure; it does not by itself resolve
this capture failure.

The saved/reopened-workbook comparison
`31f8e064c28d4b9da35408c1cd082b11/green.json` reports **84 PASS / two FAIL**.
Capture still rejects the foreground; its three observations independently show
Boolean ExcelVisible=True, an enabled/non-minimized form and one ordinary workbook.
The combined saved-identity/Saved-state check also fails. Normal closure verifies
the exact fixture path and unchanged closed-file hash, narrowing the second failure
to the in-memory Saved-state assertion, without establishing its cause. Its actual
value/type was not recorded, so Boolean False is not separately proven. No PNG is produced.
The controller exits 1 with Excel closed. Do not claim complete workbook
preservation, visible acceptance or a runtime defect from this diagnostic alone.
`-GuideCaptureSavedWorkbookForTest` requires the visible saved-copy guide gate;
it preserves all existing foreground guards and records only fixed labels/counts/
booleans. The blank capture calibration explicitly disables these packaged-only
observations. Full-chain regression is the next gate; capture and in-memory
workbook-state diagnosis remain open.

`-ViewerStartupPackageStateForTest` defaults to `OriginalReadOnly`, preserving the
original reproducer. `WritableCopies` and `SavedCopies` copy only the five input
XLAMs into the disposable run directory, verify reference/path ownership and open
the copies writable. SavedCopies additionally verifies each save before execution.
Only startup diagnosis may use WritableCopies; SavedCopies may also run the
compiled guide-expectation gate. Frozen input XLAMs are never saved or replaced.

`-ViewerStartupCalibrationForTest` defaults to `None`. Other values explicitly
select the blank-form/caption/status/configuration comparisons below. They can
change the state encountered by the following Viewer callback, so a later opening
success in those modes is not an untreated startup result or product GREEN.
All emitted observations are fixed stage labels, booleans or numeric errors.

Reports under `reports/runtime/slice4be-viewer-published-read/`:

| Comparison | Directory / red.json | Observed result |
|---|---|---|
| Blank caption only | `baeb4340325e47d8a33195b360e6a5ec` | 17 PASS / two startup FAIL; blank caption works. |
| Viewer caption boundaries | `49588dbf56044b1a978c4df506389950` | 17 PASS / two startup FAIL; caption access works before/after layout and before generation; post-context probe does not return. |
| Blank across recording status | `5824841d097f4b96b60ae35f744ce3cd` | 19 PASS / one calibration FAIL; subsequent caption call returns -2147418105; later Viewer opens/reuses. |
| Blank across configuration loading, then status | `8fd5db118bbf4d0383d12c35e2d80582` | 20 PASS / one calibration FAIL after configuration loading; later status and Viewer succeed. |
| Show blank before configuration loading | `1317a51537514c49b64844ca5996ee72` | 17 PASS / three FAIL; showing first does not preserve blank caption access or Viewer startup. |
| Open/hide/close configuration separately | `9bebe08e5d5a47fa98ef3aea562c6d03` | 20 PASS / one calibration FAIL; caption works after open and hide, fails after close; package bindings preserved. |
| Close configuration without hiding | `d4cdc7d2803448ba8208708e44b2a254` | 20 PASS / one calibration FAIL after close; omitting hide does not fix it. |
| Writable unsaved copies; no calibration | `ab17fb9254cd43a48acbf26a4e3d123c` | 18 PASS / two startup FAIL; `ERROR\|440\|Warehouse.Caption`. |
| Compiled saved copies; no calibration | `212be00330df4d1ba7eb4a58235fc112` | 20/20; `OK\|1`; opening and reuse pass. |

The unsaved/saved copy comparison retains every original saved-workbook startup
check and adds a package-state assertion plus loaded-path preservation. No guide
expectation, caption timing, workbook visibility, configuration caching, public
callback or runtime error suppression is changed to obtain the passing result.
The test-only diagnostic wrapper still invokes the real public callback.

All 18 completed startup observations have verified report counts, distinct
Boolean check results and zero Application events 1000/1001/1002 in their measured
windows. The close-without-hide case's immediate ExcelClosed value is False;
its audit extends through the saved-probe runner's verified no-Excel preflight.
This is subsequent closure evidence, not an immediate-cleanup pass. Eight invalid
diagnostic option combinations are rejected before starting Excel, including
calibration without tracing, package-copy modes outside their declared scopes and
the guide workbook outside its visible saved-copy gate.
Nine scripts parse, 91 local links resolve, and all 30 frozen packages plus both
unrelated documents retain their hashes. Runtime source is unchanged.

The first shown-form attempt (`48702c2576c74d338fd100a37ea0a33a`) was interrupted
during fixture preparation. Its process handle disappeared; later inspection found
no Excel or runner and no final report. It is neither RED nor GREEN. The normal
cleanup/restoration block is not proven to have run: four temporary fixture
references remained in local invSys settings, and its in-memory original snapshot
was unavailable. Subsequent runs preserve their own starting settings; they do
not prove restoration of the pre-interruption selection. No previous setting is
guessed, and no operational workbook or recovery file is deleted.

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

The blank-form and package-state comparisons above supersede the former next
step. SavedCopies clears startup but does not complete visible acceptance.
Continue the full-chain gate and retain the unresolved foreground/Saved-state
findings; no caption-order or transient-workbook-visibility workaround is justified.
