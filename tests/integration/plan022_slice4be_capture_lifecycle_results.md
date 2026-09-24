# Slice 4be caption capture lifecycle investigation

Last reviewed: 2026-09-24 UTC. This is developer-tooling evidence, not a new
product contract or Slice 4be acceptance. Architecture v4.11 D18, D13 and the
saved-workbook operator deployment model remain binding. Runtime source and
the frozen `validation-auth-read-separated` packages are unchanged.

## Protecting calibration and helper correction

The complete Shipping/Boxing route twice stopped at the unavailable-store Make
capture. The isolated submission route passed 794/794 with normal shutdown.
Neither result established why the full route's caption was unreachable.

`Test-Slice4beCaptureForeground.ps1` now extends its original three disposable
capture cases through closing the previous form, retaining its project as an
add-in, creating/closing 40 blank workbooks, and reopening with/without a workbook.
It opens no invSys package or operational workbook. All six valid cases capture.
After closing the visible workbook, requesting Visible=False reads back True;
the diagnostic records that mismatch without treating it as proof of native form
unavailability. The original three strict visibility checks remain unchanged.

An offscreen-form negative calibration exposed a separate helper defect. In
`capture-foreground-calibration/660ef8bc0aa04dd8be874dc8ab658d67` under runtime,
physical bounds remain offscreen before and after raising the form. WindowFromPoint
nevertheless returns that form, SetCursorPos reports success, and the old helper
returns without the expected rejection while foreground remains the workbook.
This is a failing developer-tooling assertion, not product D13 RED.

The helper now rejects candidates outside all monitors before cursor movement or
input, and verifies the actual cursor coordinates after SetCursorPos before any
button input. Ownership and post-position hit-test checks remain. It does not
replay a business action, accept background captures or weaken visible acceptance.
Failure diagnostics contain window/process identities, classes, physical bounds,
hit targets, z-order, cloaking and resource counts; no titles or workbook content.

Corrected calibration report:
`reports/runtime/capture-foreground-calibration/41528362d09a4e0283093eb62922a25a/observations.json`.
UTC **19:34:36--19:34:46**. Six successful captures, all directly reviewed, plus
the expected offscreen rejection with three off-monitor candidates, no hit/input
path entered, and complete failure diagnostics. Normal Quit and process closure;
no termination. This corrects the calibration defect without claiming resolution
of the real full-route caption failure.

Final calibration adds direct native caption activation and independent topmost
restoration checks for every case. It passes **eight cases**: seven reviewed
captures and the expected offscreen rejection, with normal closure. Report:
`capture-foreground-calibration/9b894feab5814173bdebcb68f9d32d6b` under runtime;
UTC **19:59:21--19:59:31**. Receipt:
`reports/runtime/caption-state-calibration-verification.json`.

Earlier calibration attempts are retained: `e27b5fe764f94953b4d45a8607593638`
stops at the overly strict add-in-only visibility assumption; `9a80052470364fd2979ceb5699f52d53`
captures six lifecycle cases; `9e1e74b0266e4430a22bc84a90532a1a` and
`72e4a27538664c6d8503b1e990b4e7a7` do not produce the intended offscreen rejection.
The final pre-correction negative run above records the physical cause of that
calibration mismatch. These are not silently replaced with GREEN evidence.

## Packaged diagnostic

The unchanged complete Shipping/Boxing sequence records **1632 PASS / one capture
exception**, retaining every reached identity/outcome from the preceding partial
run. All seven Auth-recreation checks and both corrected read-only assertions
pass; 82 full-route checks remain unreached. Five instrumented compiles pass.
Sixteen images are produced but not promoted to new visual acceptance; this
diagnostic does not replace the earlier separately reviewed captures. Report:
`reports/runtime/slice4be-shipping-activity/f16817b7fdff4c4a9dde4ddb061606ee`.
Controller prefix:
`reports/runtime/caption-state-diagnostic-`; the original settings snapshot is
retained until Excel and its worker close. No package is rebuilt or modified.

At the failure, all three physical caption candidates are on a monitor. The
owned form is visible, enabled, non-minimized and not cloaked. SetWindowPos
returns success, but its topmost state remains False and it stays below VS Code;
all three hit tests resolve to VS Code. The helper correctly sends no input to
those points. GDI use is **816** at failure, with process peak **930**; 107 passive
samples reach at most 892 current GDI objects and 15 XLMAIN windows. This run's
failure is not explained by offscreen geometry, cloaking or the separate earlier
10,000-GDI exhaustion. Why the z-order request has no effect remains unresolved.

No alternate activation/window-message flag is implemented here. Any such helper
change needs focused calibration, ownership/hit-test protection and original
topmost restoration before a broad retry. See Microsoft's
[SetWindowPos contract](https://learn.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-setwindowpos)
for the activation and window-position message flags; their relevance to this
particular MSForms failure is not yet proven. The submission/outcome probes were
already installed before forms and five compiles; the later install call is
guarded, so do not assume it performs a late VBA edit without further evidence.

The terminal report is written before assisted host/Excel termination. Excel is
briefly still enumerable with HasExited=True; subsequent observation confirms it
is absent. The passive sampler exits without termination. The original controller
restores settings, preserves five packages/217 tooling pins, and the Application
1000/1001/1002 audit is clear. Normal shutdown is not accepted. UTC **19:35:28--19:56:32**.
Receipt: `reports/runtime/caption-state-diagnostic-boxing-verification.json`.
Static metrics remain unchanged (251 components, 6047 procedures, 132972 lines,
nine literal/45 unresolved calls, 191 duplicate groups); three schemas pass.

## Multiline follow-up

Source review found an independent original-value risk in
`frmInventoryViewer.ViewerUnescape`: sequential replacement of escaped newlines
before escaped backslashes can transform literal backslash sequences. The Core
publisher escapes backslashes first, so this decoder is not its exact inverse.
Three synthetic algorithm controls fail preservation in
`reports/runtime/detail-escape-static-control.json`. This is not packaged D13 RED.
The next multiline fixture must include both literal escape sequences and real
line breaks through actual publication, Viewer selection and Event Detail before
any runtime correction. No decoder or rendering behavior changes in this checkpoint.
This focused product work can proceed independently of the unresolved broader
caption failure; the complete Shipping/Boxing and normal-shutdown gates stay open.
