# Plan 022 Slice 4be.3: multiline field rendering

## Contract and implementation

Architecture v4.11 D18 already requires complete permitted text reachability.
Before implementation, the specification, Plan022 and controls catalog named
the read-only multiline area as a display refinement under semantic inheritance.
It follows the existing contributing-line selection and introduces no editor,
field-selection dependency, authority access, tracking event or persisted state.

`frmEventDetail` retains its caption/value list and adds
`lblDetailMultiline` / `fraDetailMultiline` beneath it. Caption/value Labels show
only multiline fields already permitted by the loaded profile, in that order.
The area has native scrollbars; changing the selected line rebuilds its content.
Existing context invalidation clears its values and scroll positions. The minimum
form remains 820 by 640 points; the list, preview, status and Close have distinct
regions. The owning controller and all Core/Domain services remain unchanged.

Office Label captions normalize lone CR/LF to CRLF. A disposable local control
probe proves this and preserves tabs, ampersands and Unicode exactly. The spec
clarifies visual line-break interpretation while retaining the exact original
field string in the loaded projection/list; no normalized text is written back.
The tests separately require binary original-value equality, exact rendered text
apart from display separators, and sufficient height for every logical line.
Labels use an empty Accelerator. See the
[Microsoft Forms Accelerator contract](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/accelerator-property).

## D13 attempts — 2026-09-24

Protecting tests: `Slice4beViewerEventDetail.ps1` and
`Slice4beDetailMultiline.ps1`. The existing Admin-generated fixture publishes
synthetic values through Core and selects them through the actual Viewer and
Detail handlers. Cases include CRLF/LF/CR, literal escapes, tabs, Unicode,
ampersands, many lines and an over-wide final line. The existing 56 checks remain;
three native-input checks run only with captures. The new native reachability
checks also require captures. No unrun native check is reported as passing.

Private report roots below are under `reports/runtime/slice4be-viewer-detail/`.

| Attempt | Result and interpretation | Report root |
|---|---|---|
| Initial visible RED | 29 PASS / one capture exception: cursor positioning denied before the new checks. Not meaningful product RED. | `e73abf3c85ad48d4bb8b3edfacbf017e` |
| Initial form-level RED | 53 previous non-capture checks PASS / 23 expected multiline FAIL, without a harness exception. No runtime source had changed. | `a97b9070c24745b1b06816ea598d0bdf` |
| First implementation attempt | Five compiles PASS, then runtime error438 opening Detail. Unsupported Label.UseMnemonic; the verified error dialog was ended so the original controller could clean up. Not GREEN. | `548fd30a31b24b19a7001aab0d20dc1e` |
| Corrected Label API | 37 PASS / two failures. One asserted exact Label separator encoding; the other compared nested scrolling content with outer-form bounds. Incomplete result. | `e6f887fecebf4752ba90d76dc7483163` |
| Corrected display/container probes | 76/76 form-level checks, including profile hiding and sign-out clearing. No further runtime edit. | `f56298cf40cc4747b90322e37c647cbb` |
| Target-clearing predecessor RED | 54 PASS / 25 expected multiline FAIL, retaining all 53 earlier non-capture checks and proving existing target invalidation still works. Adds a positive-content warehouse-switch scenario. | `0e3cd87a38954eaabeeffb4a958e5dfa` |
| Target-clearing form-level GREEN | 79/79, exact RED identities; all 53 prior non-capture checks plus profile/sign-out/target-switch protections. | `fd4475f395834da6a274180bda401ecc` |
| Multiple-field predecessor RED | 54 PASS / 27 expected multiline FAIL. Adds two permitted multiline fields and preserves all previous checks. | `43135f73c4da4a52825d7644979916fb` |
| Final multiple-field GREEN | 81/81, exact RED identities. Both roles preserve both field values, profile order and separate geometry. Seven native-input checks remain unrun. | `73a6135dfd184fc78e8c9e44ae93f287` |
| First captured predecessor RED | 57 PASS / 31 expected multiline FAIL. Existing native input and capture succeed after desktop access returns. | `2ab9224c6f1d4f27a31030e456d03ecb` |
| First captured candidate | 87 PASS / one detector FAIL. All new multiline native checks pass; the preceding list-scroll detector compares logical coordinates with physical image pixels. This attempt remains incomplete. | `257c09fe7f6b499f9326c05f54d33a52` |
| Corrected-detector captured RED | 57 PASS / 31 expected multiline FAIL; all previous 56 identities retained. | `647086b064b143f1b6331eeb03d9f62e` |
| Final captured GREEN | **88/88**, exact RED identities, all 56 preceding checks and all seven native-input checks. Fourteen principal images reviewed and accepted within this automated scope. | `cdb9d52f7ebf433b9413b42559deea0e` |

The two probe corrections follow disposable observed behavior. UserForm.Controls
includes nested controls, so every control is checked against its actual parent:
outer form viewport or inner scroll extent. No geometry check is removed. Label
display equivalence does not replace the separate byte-exact original-field tests.

All attempts restore settings and preserve their pinned packages/tooling. The
failed first implementation required error-dialog assistance; it is not described
as an unassisted success. The original candidate `validation-detail-original-text`
and rejected `validation-detail-multiline` remain preserved. Corrected candidate:
`deploy/validation-detail-multiline-labels`.

Target-clearing RED UTC21:01:42--21:03:17; GREEN UTC21:04:21--21:05:55. Both close Excel
normally, restore local settings, preserve package/tooling hashes and pass five
instrumented compiles. The combined interval has zero Application1000/1001/1002
events. Receipt: `reports/runtime/detail-multiline-final-verification.json`.

Final multiple-field RED UTC21:11:36--21:13:38; GREEN UTC21:13:59--21:16:00.
Both close normally and restore settings; package/tooling hashes and all previous
non-capture check identities hold. Five instrumented compiles per run and zero
Application failures. Receipt: `reports/runtime/detail-multiline-multiple-verification.json`.
The extra case is necessary to cover every permitted field rather than proving
only one field. The native vertical-end probe now measures the last content in
the entire area; a separate native check must reach the last of multiple fields.

The captured detector failure is isolated with a meaningful tooling RED/GREEN:
`Test-Slice4beDetailScrollEvidence.ps1` changes from 4 PASS / four FAIL to 8/8.
Synthetic scrollbar images at 100%, 150% and 200% scaling require movement,
reject stationary images and reject unrelated content changes outside the bar.
The detector now maps whole-window logical bounds to physical image dimensions;
its two-half movement threshold and border/arrow/input-point exclusions remain.
No screenshot is modified. Re-reading the failed candidate's original pair
detects 15107 changed scrollbar pixels, including both halves. The preserved
locked predecessor pair still has zero changed pixels and correctly fails
movement. This offline correction does not relabel the 87/1 run as GREEN.
Private detector roots: `detail-scroll-detector/8553362e4a9a431b8a87fe93ff3c4731`
and `detail-scroll-detector/f47a9097ab774750baf36d2495793da5`.

Final captured RED UTC21:30:00--21:32:07; GREEN UTC21:32:15--21:34:24.
Both close normally, restore settings, retain every package/tooling hash and
pass five instrumented compiles. The combined interval has zero Application
failures. Receipt: `reports/runtime/detail-multiline-dpi-visible-verification.json`.

## Gates and remaining acceptance

Both candidate builds and five-package compiles pass, including Operations cold
start. The corrected candidate differs from the preceding 244 compiled components
only in Operations/frmEventDetail. Source layout checks pass 8/8 and 7/7.
Static maintenance: 251 components, 6050 procedures, 133033 lines, nine literal/
45 unresolved calls, 191 duplicate groups and all 28 prior oversized limits
preserved; three schemas pass. Three small helpers add 47 lines to the existing
form; no new component, dynamic call or duplicate body is introduced. All 278
PowerShell files parse; the native multiline input helper compiles without sending
input and subsequently passes actual native interaction in the final gate.
Final maintenance receipt: `reports/runtime/detail-multiline-dpi-static-verification.json`.

Desktop access was intermittent: a read-only probe at 20:43 UTC and again at
21:01 UTC gets Windows error5 from GetCursorPos. Opening the input desktop succeeds,
so this is not proof that elevation or general access is missing. Cursor access
returns successfully at21:17:26 and again at21:23:07; captured gates then succeed.
The original error5 cause remains unproven.

Scoped automated visible acceptance is now **GREEN**. Fourteen principal captures
show the default/restored and maximized layouts, original list horizontal movement,
mixed separators/tab/ampersand/Unicode, two fields in profile order, the second
field's last line, literal escapes and the long final line through its `END`.
Native input reaches both vertical and horizontal limits; original strings remain
unchanged and native typing cannot edit fields. The existing field list, multiline
area, status and Close do not overlap. Larger geometry also passes the packaged
handler checks. Three intermediate list-input frames are supplemental, not counted
as separate accepted principal images. Review receipt:
`reports/runtime/detail-multiline-visible-image-review.json`.
This does not substitute for human Release1 acceptance.

The earlier Viewer/filter/Shipping-state non-capture regression passes **97/97**;
`EventFilters.VisibleCapture` is unrun in that attempt. UTC21:06:55--21:10:22. Report:
`slice4be-viewer-published-read/8503dbf072df43e496f2f1f8db15746c` under runtime;
receipt `detail-multiline-labels-viewer-verification.json`.

The final captured Viewer regression passes **98/98**, retaining every preceding
identity, with all three images reviewed (filter surface, activity detail labels,
and the actual Shipping held state). UTC21:35:06--21:38:08. Both runs close normally,
restore settings, preserve their package/tooling hashes and have zero Application
failures. Final report: `slice4be-viewer-published-read/40fe65ef603e48fcba2bba291135baed`;
receipt `reports/runtime/detail-multiline-viewer-visible-verification.json`.

Packaged XLAM smoke passes **86/86**, retaining the exact preceding 86 check
identities, UTC21:38:45--21:39:06. Excel closes normally; local settings and the
tracked report are restored, package hashes are unchanged and Application
failures are zero. Receipt: `reports/runtime/detail-multiline-packaged-verification.json`.

Current-candidate Release1 chain passes **32/32**, live-role **48/48**, and
Create Warehouse **15/15**, retaining every previous identity. UTC21:16:16--21:21:48,
normal unassisted closure, restored settings and three tracked reports, unchanged
packages/source during the gate, and zero Application failures. Receipt:
`reports/runtime/detail-multiline-labels-chain-verification.json`.
Broader Shipping/Boxing capture/shutdown,
comprehensive tracking, guide transfer/comparison and human/NAS acceptance remain
open in the maintained checklist. This is not completed Slice4be acceptance.
