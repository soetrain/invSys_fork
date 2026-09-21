# Plan 022 Slice 4be Event Detail horizontal overflow

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

Capture diagnosis retains failures separately from product RED: the first capture
run passed 25 checks before a current-state fixture COM exception; subsequent
runs stopped at foreground acquisition (20/1 twice), unavailable standard native
scrollbar metadata (25/1), and unavailable UI Automation ScrollPattern (23/1).
The latter two probes were discarded. Native focus identified the form's client
surface rather than the windowless fields list. The retained input harness checks
the actual list bounds, owned foreground form and unobstructed target, and uses
a bounded press with guaranteed release. Input delivery alone is never acceptance.

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

Required next gates: actual visible scrolling and
a successful live-role/full-chain run, synchronized catalog/plan and reviewed Git
status. Multiline detail rendering, broader Slice 4be and human Release 1 acceptance
remain open. This entry does not authorize the separate D8-A Auth contract change.
