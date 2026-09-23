# Slice 4be restart shutdown: header-scan reference lifetime

Last verified: 2026-09-23 UTC. This repairs Release 1 verification tooling under
the existing Architecture v4.11 contract; it changes no runtime or XLAM behavior.
Focused shutdown and header compatibility pass. The full-chain retry fails
earlier in projection rebuilding, before the corrected restart stage. Slice 4be
remains incomplete.

## Reproduced failure and isolation

The passive full chain retained 32/48/15 chain/live-role/Create Warehouse checks
but its original restart Excel process crashed in `combase.dll` with `c0000005`,
followed by `OFFICE_MODULE_VERSION_MISMATCH`. No external post-Quit COM attachment
or second Quit occurred. The focused control extracts the actual restart
procedure and helpers, adds only lifecycle observations, and uses an isolated
copy of the generated live-test warehouse. It never opens operational inventory.
Full restart modes retain all 12 original reconciliation check identities;
phase-cut controls are explicitly diagnostic and skip later assertions.

| Control | Result and implication |
|---|---|
| Empty Excel; five packages; plain Excel table with reference collection | Normal closure and zero Application failure events. Loading packages or using a plain table alone does not reproduce the fault. |
| Stop after configuration, snapshot, or three operator refreshes | Normal closure, zero events. These phase cuts do not establish full restart acceptance. |
| Stop immediately before `NoRowHeaders` enumeration | Both preceding checks pass; normal closure, zero events. |
| Include header enumeration, or later reconciliation reads | Checks pass, but original Excel crashes during deferred cleanup. Processor replay is not needed to reproduce the fault. |
| Full original restart, worker kept alive after Quit | All 12 checks pass; original Excel survives the 210-second observation and crashes after worker release. A new empty Excel instance is closed under the user's standing authorization; settings are restored. |
| Full original restart, collect .NET references after return | All 12 checks pass; the same original-process crash occurs immediately. Collection alone is not a fix. |
| Replace nested COM enumeration with indexed reads only | Still fails shutdown; retained as a failed approach. |

Event 1000 identifies the original worker-owned Excel process in each audited
failed original control. Event 1001 retains the Office mismatch report. Audit
bounds use local time explicitly; a process disappearing is not sufficient proof
of successful shutdown.

## Correction and focused evidence

`tools/validate_release1_full_chain.ps1` now delegates its unchanged
`NoRowHeaders` assertion to `Test-NoRowHeaders`. The helper scans every supplied
workbook, worksheet, table and column with indexed access and releases each
locally acquired COM reference before returning. It reads `ListColumns.Name`,
preserving the prior trimmed, case-insensitive comparison and hidden-header
behavior. It neither closes the caller's workbooks nor changes their contents.

The first explicit-release variant read header cells. Its shutdown controls
passed, but compatibility testing exposed hidden-header failures. The corrected
name-based implementation passes **8/8** checks: allowed headers across workbooks,
the last table's mixed-case forbidden header, both hidden-header cases, unchanged
header visibility, preserved unknown header/value, normal closure and zero events.
The protecting RED is **6 PASS / two expected FAIL**. An earlier calibration
mistakenly renamed a column while headers were hidden and raised 0x800A03EC;
that incomplete fixture attempt is retained and is not the protecting RED.

The final full restart control retains **12/12** original checks, exits while its
worker remains alive, restores local settings, preserves all five package hashes
and has zero Application failure events. This is focused teardown RED/GREEN for
a verification failure, not a new product behavioral RED or full release proof.

## Exact evidence

All roots below are ignored local evidence, not committed runtime reports.

- Original full failure: `reports/runtime/slice4be-shutdown-control/0b1aa39cbefc4280ab0acd6a6e7faa1b`.
- Focused header failure: `slice4be-shutdown-control/51c58927af1148d0b2ff3b8e5a233ba9`.
- Indexed-only failure: `slice4be-shutdown-control/a51ce575f5a843c685d8a2266d8739e6`.
- Final full focused GREEN: `slice4be-shutdown-control/b933852678f8445985bb946faede2f7f`.
- Header compatibility RED/GREEN: `header-scan-contract/8ddb59f4b95c4e1aaa4286dfe8198f81`
  and `header-scan-contract/017c19f1145d499791044ab62c869333`; initial incomplete
  calibration `header-scan-contract/aff1cb63dd884f2b8bb113d2163c07d5`.
- Original fault correlation: `reports/runtime/shutdown-control-failed-event-facts.json`.
- Corrected full-chain prefix: `reports/runtime/settings-diagnostic-header-release-chain`;
  terminal exit 1; acceptance fails as detailed below.

## Full-chain retry and static verification

The actual corrected chain runs 22:59:10--23:05:19 UTC on 2026-09-23. It records
**5 PASS / one harness exception** in the chain, **32 PASS / one harness exception**
in live roles, and **15/15** Create Warehouse checks. Respectively 27 and 16 prior
chain/live-role checks are unreached. The failure occurs at **Delete and rebuild
canonical inventory projections**, before the changed restart scan. Excel crashes
in `ntdll.dll` with `c0000028`; both Application Error and Office mismatch events
are retained. The cause of this earlier failure is unresolved; this attempt does
not invalidate the focused result or establish a successful full chain.
The projection-deletion assertion passes. The subsequent packaged
`modProcessor.RunBatchReportForAutomation` call fails with `0x800706BE` while
rebuilding missing projections; deletion itself is not the observed failure.

The recovered Excel instance contains ten saved generated-test workbooks. They
are closed without saving under the user's standing authorization, then that
instance receives its first Quit. This is assisted recovery, not normal test
shutdown. Pre-closure file hashing fails because files are locked; byte
preservation across recovery closure is therefore **unproven**, not a detected
file change. Local settings and all three tracked reports are restored; Excel
is closed. Do not repeat the broad chain unchanged. Isolate the projection
deletion/rebuild failure on disposable generated fixtures first.

`reports/runtime/shutdown-header-evidence-verification.json` verifies all 299
runtime, 252 tooling and five package pins. Regenerated static evidence retains
250 components, 6,045 procedures, 132,893 lines, 9 literal/45 unresolved dynamic
calls and 193 duplicate groups. Three schemas, 28 size limits and parsing of all
252 PowerShell files pass. No rebuild or new product compile is claimed because
runtime/package bytes are unchanged. The focused 12 prior check identities and
eight compatibility checks are independently verified. Full-chain and full-slice
acceptance remain false.

## Projection boundary control

`Test-Slice4beShutdownControl.ps1 -Case RestartProjectionReplay` reuses the actual
restart loader/cleanup and extracts the live validator's exact projection-delete
block and helpers. Only a new copy of the failed generated warehouse is changed.
The first three calibration attempts stop at a fixture precondition: the saved
warehouse retains its pending trigger and six log/applied rows, but both
projection tables are present. These are fixture failures, not product RED.
Roots are `824f73fb592d4a3a9c301c24ae642b29`,
`955dca5d806d4df8a00792d0c7cbf58b`, and
`abadd558c0094f5d8de6f0a89d6346a9` under
`reports/runtime/slice4be-shutdown-control/`. All restore settings and close
normally with zero Application failures. Added sanitized stack/line evidence
identifies the precondition failure without recording row values or credentials.

After recreating the original deletion on the copy, root
`6248968ddfac470594907614703db211` passes **6/6**: missing-projection precondition,
both rebuilt tables, exactly one applied/log append, exact EventID/System_Key,
processed status, and replay without duplicate authority records. The same
packaged processor returns normally in this fresh session. This narrows the
failure but does not reproduce the preceding live role actions or prove their
full-chain context safe. No runtime fix or product RED/GREEN is claimed.

The control's strict lifecycle result is **FAIL**: Excel remains through the
210-second post-Quit observation and exits after the worker ends. No assistance
is used. A delayed audit at 23:21:50 UTC confirms Excel closed and zero Application
failure events. Settings and all five package hashes are preserved. Source
verification confirms all 299 runtime hashes unchanged, the diagnostic script as
the only change among 252 tooling files, and 252 successful PowerShell parses.
Receipts: `projection-replay-delayed-event-audit.json` and
`projection-replay-source-verification.json` under `reports/runtime/`.

Next isolate the processor boundary after the exact preceding role handlers in
one Excel session, with a diagnostic phase cut before later chain stages. Do not
repeat the unchanged broad chain or infer the earlier crash's cause from this
fresh-session success.

The [remaining acceptance checklist](plan022_slice4be_remaining_acceptance.md)
continues to govern comprehensive coverage and all other gates. This correction
does not resolve the Applied comparison label-read failure or approve pending
contract amendments.
