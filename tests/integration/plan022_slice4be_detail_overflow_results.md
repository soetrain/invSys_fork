# Plan 022 Slice 4be Event Detail horizontal overflow

## Boxing probe reconciliation, 2026-09-24

The broader D8-A run exposed two stale Make/Unbox `ReadOnlyFields` assertions:
their shared probe equated read-only cells with `mFields.Locked=True`. Approved
D18 instead requires non-editable ListBox cells with selection/scrolling permitted.
The probe now checks the actual ListBox type and `Not mFields.Locked`. No runtime,
form handler or package changes. Both broader check identities are retained.

The focused Viewer/Detail gate now installs that same Boxing probe before opening
Detail through the actual Viewer handler. Before reconciliation it records
**49 PASS / one expected probe FAIL**; afterward **50/50**. All 44 preceding
checks remain and pass, including native scrolling, typing/value preservation,
captured context and source bytes; five instrumented compiles are added. This
is test-contract reconciliation evidence, not a newly introduced product RED.

Frozen candidate: `deploy/validation-auth-read-separated`. RED report:
`reports/runtime/slice4be-viewer-detail/f9df336decba4a21b69a04c1a7dfccc8/red.json`;
GREEN report: `reports/runtime/slice4be-viewer-detail/2183d0160a1d4b48a3bb44f342a4ae43/green.json`.
Both runs close normally and restore settings, with unchanged packages/tooling
during each run and zero Application events 1000/1001/1002. Four principal GREEN
images are directly reviewed: default, scrolled, maximized and restored. The
scrolled image exposes the complete Coverage ending. Receipt:
`reports/runtime/reconciled-boxing-detail-verification.json`.

The disposable foreground calibration passes hidden-first, visible and
hidden-restored cases and closes normally (`capture-foreground-calibration/a4402c86bd4546488f857502739b2d21`
under runtime). This does not resolve the earlier long-run caption failure or
establish sustained capture access. The existing submission-only route now admits
Boxing's required probes while still rejecting incomplete standalone recording
and owner-completion combinations before Excel starts. Default full-route behavior
is unchanged. This diagnostic routing changes no product contract and needs no
manufactured product RED.

The isolated route passes **794/794**, including **all 82 previously unreached
checks**, five instrumented compiles and seven directly reviewed captures. The
previously failing unavailable-store Make capture succeeds; Make/Unbox both show
successful owner outcomes with separate tracking notices. Older-policy notices
and valid disabled-policy behavior are visible. Normal delayed Excel/host closure,
restored settings, preserved packages/tooling and zero Application events
1000/1001/1002. Report:
`reports/runtime/slice4be-shipping-activity/27b3133c979c43c0bfe71eb7d25bf891/diagnostic-submission-boxing-activity-shipping-recording-green.json`;
receipt: `reports/runtime/boxing-submission-isolation-verification.json`.
The complete corrected Shipping/Boxing route subsequently reaches **1632 PASS /
one capture exception**, with 82 previous checks unreached. Both reconciled
Make/Unbox read-only assertions and all seven D8-A Auth-recreation checks pass.
The same `boxing.tracking.unavailablestore.make.png` caption failure reproduces:
the owned form is enabled/non-minimized, but foreground is VS Code and the helper
finds no uncovered caption point. Sixteen captures are directly reviewed. The
isolated result does not replace full recording coverage or explain this failure.
The original controller writes the terminal report after termination of only the
disposable Excel process; the test host is not terminated. It restores settings,
preserves all five packages/217 tooling pins and records zero Application events
1000/1001/1002. Normal shutdown is not accepted. Exactly 1625 prior GREENs are
retained; the other seven passes are the corrected Auth assertions. Report directory:
`reports/runtime/slice4be-shipping-activity/c527145cf74349daa227962018d7ae5b`.
UTC **19:02:03--19:23:22**; receipt:
`reports/runtime/d8-auth-read-reconciled-regression-boxing-verification.json`.
The later read-only window inspection occurred after failure cleanup had removed
Excel's windows, so it does not establish the caption failure's cause. Capture
diagnostics must observe physical hit-test targets, z-order/cloaking and resources
at the failure point before another full retry. Do not replay the business action.
Runtime maintenance metrics remain unchanged; three schemas validate in
`reports/runtime/reconciled-boxing-detail-static`; all 277 PowerShell files parse.
Multiline remains open.

## Approved scrolling implementation, 2026-09-24

The user explicitly approved the D18 read-only ListBox amendment on 2026-09-24;
Architecture v4.11, Plan 022 and controls record that decision. Historical pending
wording below describes earlier checkpoints and does not reopen approval.

Fresh focused RED on frozen `validation-production-diagnostics` completes
**43 PASS / one expected FAIL**. Only `EventDetail.NativeHorizontalScrollMoves`
fails; all previous 43 identities remain and a new native typing check passes.
Report: `reports/runtime/slice4be-viewer-detail/ee46ed5683934737a9ccd76e2f3edc68/red.json`.
UTC: **16:42:31--16:43:55**. Reviewed before/after images show zero changed
scrollbar-interior pixels: native input reaches the locked list but cannot move it.

The implementation changes only `lstEventFields.Locked` to False, preserving
the pre-existing uncommitted font-measurement helper verbatim. This permits native
selection/scrolling; the ListBox has no field-edit or workflow handler. The geometry
probe now tests text capacity independently of the superseded lock requirement.
The new owned-foreground keyboard probe types letters, Backspace and Delete into
the actual focused fields list and compares every original caption/value exactly.
No temporary test unlocking is enabled in either run.

The pending 18-line helper belongs to this same overflow correction, as its
earlier RED/GREEN below demonstrates. It is included unchanged with the approved
lock change so the committed form reproduces the tested candidate. Its previous
classification as unrelated applied while other workstreams were active; this
does not absorb the unrelated handoff067 or critique023 document changes.

Candidate `deploy/validation-approved-detail-scroll` builds and passes all five
packaged compiles and Operations cold start. Comparison with the Production
diagnostic candidate retains 244 components and changes only
`invSys.Operations.xlam/frmEventDetail`. Focused GREEN is **44/44**, retaining all
RED identities, including native movement, non-editability, profile/context,
read counters, exact contributing lines and authority/unknown-column bytes.
Report: `reports/runtime/slice4be-viewer-detail/9fe62bcdc0094b419c09629dba03abbf/green.json`.
UTC: **16:45:29--16:46:55**. Default, horizontal-scroll, maximized and restored
captures were individually reviewed: the scrollbar reaches the right edge and
the Coverage value's previously clipped ending is readable. Four-size geometry
checks retain complete single-line text capacity. This is scoped single-line
evidence, not complete multiline acceptance.

Both runs close Excel without assistance, restore local settings and preserve
all five package and tooling hashes; the combined interval has zero Excel
Application events 1000/1001/1002. Private receipts use the
`reports/runtime/approved-detail-scroll-` prefix. Broader candidate regression,
static maintenance, full-chain and live-role gates are recorded separately as
they complete. Accepted deployment is unchanged; Slice 4be remains incomplete.

Regenerated static metrics match the preceding candidate exactly: 251 components,
6,050 procedures, 133,025 source lines, nine literal and 45 unresolved dynamic
calls, 193 duplicate-body groups and 28 unchanged oversized-module limits. The
source layout contracts pass 8/8 and 7/7. The first layout invocation stopped
before testing because its default RepoRoot parameter evaluated to an empty
path; rerunning with explicit `-RepoRoot .` passes. Preserve that harness error
separately; it is not product RED.

The independent Viewer/filter/Shipping-state regression passes **98/98**, retaining
all 94 historical identities plus the existing reader-boundary additions. Report:
`slice4be-viewer-published-read/36b62502483849da8b6827df5578e4a6/green.json`
under runtime. All three captures are individually reviewed. Excel closes
immediately without assistance; settings are restored and the five frozen packages
remain unchanged. The native-command dispatcher calibration passes 36/36 after
initializing its stale trace flags (the initial 7 PASS / 29 fixture FAIL is retained).
Those fixture corrections and failure-only capture diagnostics change no product
behavior or command-retry policy.

Owner-command regression passes **460/460**, retaining every preceding identity,
with five instrumented compiles, normal delayed unassisted closure, restored
settings, preserved package/tooling hashes and zero Application failure events.
Report: `slice4be-shipping-activity/adcee7ce375a41edafc1a3e28ba59061/owner-completion-boxing-activity-shipping-recording-green.json`.
All 23 images were individually reviewed: 16 are clean, four have a taskbar
preview over the lower form, and three reset-confirmation captures are small
Settings crops rather than the intended dialogs. The latter seven are not clean
visual acceptance. Behavioral capture checks do not replace this image review.

The current candidate full chain passes **32/32**, live-role workflows **48/48**
and Create Warehouse **15/15**, retaining the exact preceding check identities.
UTC **17:18:30--17:23:41**, 2026-09-24. Excel closes without assistance; settings
and all three tracked reports are restored. Five package hashes and 273 tooling
hashes are preserved, with zero Application events 1000/1001/1002. Receipt:
`reports/runtime/approved-detail-scroll-chain-verification.json`. This does not
resolve earlier crashes or complete multiline, D8-A, Production capture or human
acceptance requirements.

The subsequent scoped Production capture gate also completes **390/390** with
all prior identities, five instrumented compiles, ten individually reviewed images,
normal delayed unassisted closure, restored settings, preserved packages and zero
Application failures. See [Production path evidence](plan022_slice4be_production_paths_results.md)
for the two preserved partial attempts and tested readiness coordination change.

## Earlier evidence

**Focused RED: 37 PASS / four expected FAIL; behavioral GREEN: 41/41,
2026-09-21. Visible scrolling and full-chain acceptance remain open.**
Report:
`reports/runtime/slice4be-viewer-detail/3c5ead5970b642788985a9615ea2e769/red.json`.
All 34 preceding detail identities pass. The four failures are
`EventDetail.CompleteText.FitDefault`, `.FitLarger`, `.FitDefault.Restored` and
`.NativeRestore`. Native-maximized text fits, matching prior visual evidence;
both new layout read counters remain zero. Exit 1, Excel closed, no harness
failure or duplicate check. UTC **20:54:30.5787780--20:55:24.4153773** has zero
Application events 1000/1001/1002. Test/runtime hashes are preserved across RED;
all 30 frozen packages and both unrelated user documents remain unchanged.
Exact audit: `reports/runtime/detail-overflow-red-verification.json`.

The current packaged
detail and published-activity captures show long Coverage text clipped at the
default/restored size. Architecture v4.11 D18 now explicitly constrains complete
single-line field reachability in the existing locked caption/value list. The
profile order, captions, original values, exact contributing lines and captured
context remain binding. Horizontal scrolling is excluded window mechanics, not
a new tracked selection. Multiline rendering needs separate evidence.

The focused test extends `tests/tooling/Slice4beViewerEventDetail.ps1`. It opens
the actual packaged Viewer and selects a generated published fixture through
the existing form handler. Five new checks compare each displayed caption/value
against its column's rendered font capacity (including spare visible space for
the last column) at default, larger, restored,
native-maximized and native-restored sizes. Two checks retain zero policy/data
reloads throughout layout. Every preceding detail identity remains required,
including read-only/profile/context behavior, repeated exact keys, unlike units,
unknown-column preservation and generated authority bytes.

Expected RED is the existing fixed column width failing full single-line text
capacity; missing fixtures, probe errors and capture failures are not that RED.
The RED gate ran after the expectation-candidate evaluation completed 376/376,
with Excel closure verified before it started. No form code changed before RED.

Optional GREEN evidence retains the existing default/maximize/restore captures
and adds a native horizontal-scroll image when columns exceed the viewport.
Input is limited to the verified foreground owned detail form. A capture or
delivered click alone is not proof of readable text; direct image review remains
required. Microsoft documents horizontal scrolling when configured columns
exceed the list width: [ColumnWidths property](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/columnwidths-property).

The correction adds one private font-measurement helper and 18 source lines to
`frmEventDetail`: column widths accommodate the original permitted values after
each field refresh. The list stays locked; no authority, profile, activity or
selection contract changes. The isolated `deploy/validation-event-detail-overflow`
candidate builds and all five packages compile. Compiled-source comparison keeps
234 components and changes only `invSys.Operations.xlam/frmEventDetail`, with no
added or removed component. Accepted deployment remains unchanged.

Behavioral GREEN is `c7dabe7f0e7445a58a202389e3c295e2/green.json` under the same
detail report parent. It passes all 41 identities, including all preceding 34,
with normal closure and zero Application events 1000/1001/1002. UTC:
**21:03:27.6079447--21:04:22.9066821**. The later focused-window and field-bounds
capture runs also pass 41/41, with normal closure and zero matching events:
`9289875110d24d00bccd26888fd8fbb1` and `072cba249cca4674ac4fe2fc312b925a`.
Their four images each were directly reviewed: default/restored controls and
the scrollbar are visible, but neither scroll image proves movement. Maximized
Coverage is readable. These are behavioral passes, not complete visible acceptance.

Static evidence in `reports/runtime/detail-overflow-static` records 241 components,
5,959 procedures and 131,345 lines; the increase is one procedure and 18 lines.
Duplicate groups stay 192 and literal/unresolved dynamic calls stay 9/45. All 28
preceding oversized-module limits remain binding; no scanner deletion is approved.
There are 1,184 scanner and 1,186 reviewed candidates.

Candidate Viewer/filter/Shipping-state regression passes **94/94**, preserving all
preceding 94 identities, with normal closure and zero matching Application events.
Report: `reports/runtime/slice4be-viewer-published-read/ad01fe65be7c4c1da8cca0d476c2f8ef/green.json`;
UTC **21:32:24.8961545--21:34:48.8933562**. All three images were directly reviewed:
activity labels and filters are readable; Shipping feedback is fully painted.
Adjacent System Key/Alternative headings and complete horizontal detail reachability
remain separate visible limitations.

Candidate Boxing/Shipping retains all 1,714 identities: **1,707 PASS / seven known
D8-A FAIL**, with no new GREEN regression. Report:
`reports/runtime/slice4be-shipping-activity/b7555d0f61a341d6b52c4b25ac86ca24/boxing-activity-shipping-recording-green.json`.
UTC **21:34:49.0050885--21:49:54.8504706** has zero matching Application events.
The immediate terminal closure observation is False; Excel then closes normally
before the next gate's strict no-Excel guard and **21:49:55.0214369** start marker.
No Quit assistance or recovery selection was performed in this gate. All 22 images
were directly reviewed: three recording layouts, six published activity/business
details, four accepted/zero-quantity actions, six tracking-policy/store cases,
Settings Save, permission denial and stale-session feedback. The known detail and
Shipping heading limitations remain. The seven failures are the unchanged
`Shipping.Access.AuthUnavailable.<action>.MissingFileNotRecreated` checks for Add,
Update, Remove, Hold, Return, Stage and Send; they do not approve D8-A.

Capture diagnosis retains failures separately from product RED: the first capture
run passed 25 checks before a current-state fixture COM exception; subsequent
runs stopped at foreground acquisition (20/1 twice), unavailable standard native
scrollbar metadata (25/1), and unavailable UI Automation ScrollPattern (23/1).
The latter two probes were discarded. Native focus identified the form's client
surface rather than the windowless fields list. The retained input harness checks
the actual list bounds, owned foreground form and unobstructed target, and uses
a bounded press with guaranteed release. Input delivery alone is never acceptance.

The bounded-press capture run `c5535b27fd114f138a5a334ce97f0b08` also passes
41/41 and closes normally. All four images were directly reviewed; they still
do not establish horizontal movement. An immediate post-input capture is the next
bounded harness diagnostic, before repaint or reactivation calls.

Candidate evaluation regression passes **376/376**, retaining all preceding 376
identities without duplicates. Report:
`reports/runtime/slice4be-viewer-published-read/c166af14586c42c29ad0e1292dd1d195/green.json`;
UTC **21:49:55.0214369--22:07:39.7785775**. Excel closes normally and the matching
Application event count is zero. All 19 captures were directly reviewed: fifteen
Pending/Partial/Applied diagnostic views and four expectation-editor views.
Pending/Partial remain awaiting; only the complete Applied fixture shows conclusion
observed. Source statuses, minimum-size scrolling and editor controls are readable.

The first candidate chain stops at 5 PASS / one harness FAIL, with live roles
36 PASS / one harness FAIL and Create Warehouse 15/15. Investigation found a
missing separator in `New-OrderedLiveValidator`: the Production result command
absorbed the following Boxing stage assignment. Joining sections with explicit
newlines restores that assignment. Before/after AST checks preserve all 59 macro
calls and their exact aggregate hash; malformed result commands drop from one
to zero. This is a reporting-harness correction, not product RED or an RPC fix.

The corrected-stage run remains failed: chain 5/1, live roles 32/1 and Create
Warehouse 15/15. Its RPC failure occurs during canonical projection rebuild.
Both chain windows have zero Application events 1000/1001/1002; neither proves
the native failure cause. Both required normal Quit of a separately verified
empty test Excel process, with recovery files explicitly retained. All three
tracked reports were restored. Exact ignored evidence prefixes are
`detail-overflow-chain` and `detail-overflow-chain-stage-corrected`.

Before a further chain diagnostic, the live-role macro wrapper now reports only
the validated module/procedure name and deepest HRESULT on failure. Workbook names,
argument values and exception text are excluded; the original exception remains
an in-memory inner exception. Nine fake-boundary dispatch cases retain exact
arguments, ordering and successful return values. Both failure identity/redaction
cases pass without running Excel. The generated chain still has exactly the same
59 macro calls and aggregate hash. This proportional harness check changes no
runtime contract and is not a product RED or native-failure fix. Evidence:
`reports/runtime/detail-overflow-macro-diagnostic-verification.json` and
`detail-overflow-stage-diagnostic.json`.

The packaged run with that diagnostic passes **32/32 full-chain, 48/48 live-role
and 15/15 Create Warehouse** checks, preserving every preceding identity. Ignored
report prefix: `reports/runtime/detail-overflow-chain-macro-diagnostic-`.
UTC **22:21:31.1037189--22:28:41.9132318** has zero Application events
1000/1001/1002. This is an assisted functional pass: one verified empty test Excel
process closed before a guarded Quit could run; the final separately verified
empty process required normal Quit and a directly reviewed **Yes, view later**
recovery selection. No forced termination or recovery deletion occurred. Excel
is closed and all three tracked reports are restored. This does not establish the
cause of the preceding RPC failures. The chain's static stage saw the concurrently
prepared guide-binding source draft; its runtime used the unchanged overflow XLAMs.

The immediate post-input detail diagnostic also passes 41/41 with normal closure
(`914cf33d03fb403d88dbcd351931b221`). Direct review of its first and last immediate
captures still shows no scrolling, so repaint/reactivation is not established as
the cause. A visible-host input diagnostic is pending; no visual acceptance is claimed.
The first visibility-only setup attempt (`2dda92d91c2144608b9ee63e8589d28a`)
stops at **zero product checks / one harness failure**: requested application
visibility was not verified. It closes normally. A disposable blank-workbook
capture host is the next bounded input fixture; it must preserve the same detail
checks and source bytes and is never an operational workbook or runtime fix.

The visible-host run on the frozen guide-binding candidate passes **41/41**,
retaining all preceding 34 identities. Report
`reports/runtime/slice4be-viewer-detail/ec5cd9c2f7ee47a5a4a36366b1acd783/green.json`;
UTC **23:38:29.6745872--23:39:45.1650280**, normal closure, zero matching Application
events. The disposable blank workbook verifies application visibility as a typed
Boolean True and closes unsaved. All seven images were reviewed individually.
The recorded list/form geometry places the input on the visible horizontal track,
but three inputs still show no horizontal movement. Maximized text fits. These
facts do not establish default-size visual acceptance or prove the cause.

The next opt-in `-DetailScrollLockDiagnostic` compares the same native input with
the disposable list temporarily unlocked, then restores its verified original
lock state in finally and checks unchanged field values. It changes no runtime
property or architectural contract. Its images require direct review; neither
the 41 checks nor successful input delivery alone proves scrolling.

The lock comparison completes at **41/41**, preserving all 34 preceding checks,
with normal closure and zero Application events 1000/1001/1002. Report:
`reports/runtime/slice4be-viewer-detail/13e494fd2dea4d3f8695863f76a70098/green.json`;
UTC **2026-09-22 00:04:22.4893293--00:05:41.8372982**. Of its 12 captures, four
comparison images were directly reviewed: `event-detail-horizontal-scroll.png`
shows the locked list stationary at the left; `event-detail-unlocked-before.png`
starts at the same position; `event-detail-unlocked-input-2.png` shows the right
edge after identical input; `event-detail-lock-restored.png` retains that position
after restoration. Typed probe results verify original/restored Locked=True and
unchanged field values. Existing byte/read checks remain green. This identifies
the lock property as the interaction blocker without changing runtime behavior.

D18, Plan 022 and the catalog now record an explicit pending proposal to permit
an unlocked, non-editable ListBox for selection/scrolling. Approval was requested
before runtime implementation because the current normative rule explicitly says
locked. The proposal is not yet approved; its future native-scroll RED/GREEN must
protect actual interaction, not merely the existing font/column geometry.

The opt-in native movement test now establishes focused behavioral RED on the
unchanged locked guide-binding candidate: **42 PASS / one expected FAIL**, with
all 41 preceding identities still passing. Report:
`reports/runtime/slice4be-viewer-detail/f9349cebae8f48e0b6a81d017fae5e50/red.json`;
UTC **2026-09-22 00:27:13.4997139--00:28:29.0877190**, normal closure and zero
matching Application events. Only `EventDetail.NativeHorizontalScrollMoves` fails;
`EventDetail.NativeScrollPreservesFieldValues` passes. No temporary unlocking is
allowed in this gate.

`Slice4beDetailScrollEvidence.ps1` compares the native scrollbar interior using
the actual owned control/form geometry, excluding borders, arrows and the input
point. The previously reviewed stationary pair changes zero pixels; the reviewed
moving pair changes 3,794, including both halves. The new locked RED again changes
zero of 8,856 inspected pixels. Images remain unmodified. This bounded detector
protects movement together with the existing complete-text geometry checks; it
does not replace direct image review or establish multiline reachability.

Required next gates: actual visible scrolling,
synchronized catalog/plan and reviewed Git
status. Multiline detail rendering, broader Slice 4be and human Release 1 acceptance
remain open. This entry does not authorize the separate D8-A Auth contract change.
