# Plan 022 Slice 4be Event Detail horizontal overflow

**Focused RED: 37 PASS / four expected FAIL, 2026-09-21; runtime unchanged.**
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
with Excel closure verified before it started. No form code has changed.

Optional GREEN evidence retains the existing default/maximize/restore captures
and adds a native horizontal-scroll image when columns exceed the viewport.
Input is limited to the verified foreground owned detail form. A capture or
delivered click alone is not proof of readable text; direct image review remains
required. Microsoft documents horizontal scrolling when configured columns
exceed the list width: [ColumnWidths property](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/columnwidths-property).

Required next gates: minimal form correction; focused
GREEN and visible scroll review; isolated XLAM build/compile/layout, refreshed
static-maintenance evidence, applicable regressions/live-role/full-chain checks;
synchronized catalog/plan and reviewed Git status. This entry does not complete
Slice 4be or authorize the separate D8-A Auth contract change.
