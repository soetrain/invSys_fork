# Plan 022 Slice 4be: explicit owner command completion

Architecture v4.11 D18's UOM/Boxing terminal-map refinement governs this change.
The three Admin UOM controls use their existing `COMPLETED`/`UNCHANGED` owner
facts; Boxing Make/Unbox use `CONFIRMED`. A successful command conclusion must
still say **Command completed; Domain application not asserted**. Rejection,
cancellation, requests and pending submissions do not become positive command
completions. Original observations and SourceEventsApplied authority are unchanged.

The specification, Plan 022 and controls catalog specified these mappings before
implementation (documentation commit `207ab8d`). The protecting test records real
Admin and Boxing form actions, independently verifies their owning effects,
stops complete recordings, publishes through Admin, chooses exact outcomes in the
packaged expectation editor and invokes the actual Evaluate control. It does not
inject completed observations or use service return values as conclusions.

## Focused RED

The unchanged `deploy/validation-admin-uom-expectation` candidate records
**452 PASS / eight expected FAIL**, 2026-09-22
**15:24:28.5754411--15:34:30.0752796 UTC**. The report is
`reports/runtime/slice4be-shipping-activity/d1fc90f292f940a18903ca34034dc236/owner-completion-boxing-activity-shipping-recording-red.json`.

| Exact recorded terminal fact | Expected RED failures |
| --- | ---: |
| Admin UOM Add: COMPLETED and UNCHANGED | 2 |
| Admin UOM Remove: COMPLETED and UNCHANGED | 2 |
| Admin UOM Reset: COMPLETED and UNCHANGED | 2 |
| Boxing Make and Unbox: CONFIRMED | 2 |

All eight observations match the authored expected step, but the existing
explicit Core map reports Failed. The four negative evaluations (Reset
CANCELLED, Add REJECTED, Make REJECTED and Unbox REJECTED) correctly remain
Failed. All 92 focused check identities execute: 84 pass and only the eight
positive conclusion checks fail. All 368 supporting check identities belong to
the previously verified Boxing/Shipping baseline and pass. This isolated route
does not claim the full Shipping regression.

The 42 shared D5/setup check identities are retained. The standalone UOM route's
`Harness.AdminUomProbesInstalledBeforeForms` is covered here by the existing
`Harness.RecordingProbesInstalledBeforeForms`: UOM and evaluator probes install
before the five instrumented package compiles and the same typed zero-loaded-
Admin-forms check. This is an explicit harness-name mapping, not a claim that
the standalone check name ran. The first external verification expected the
standalone name; its correction changes no executed test or product result.

All 23 screenshots are directly reviewed and hashed: Settings, four actual
Boxing actions, three Action Paths sizes, three painted native Reset questions
and twelve visible evaluation results. All original fixture/activity/journal
bytes are preserved; only twelve separate evaluation results are appended.
Generated configuration bytes are restored exactly. Excel closes normally and
immediately without assistance. Runtime, test and five package hashes remain
unchanged, and no matching Application 1000/1001/1002 events occur. The ignored
`verify-owner-completion-visible-workbook-red.ps1` verifies these facts.

## Test-host investigation before valid RED

Earlier combined/isolated attempts are incomplete and are not the complete RED
gate. The first encounters an exclusive file-hash read on a generated workbook
already open in Excel; shared-read hashing corrects that fixture issue. Later
attempts encounter GUI resource exhaustion or stop at the resource guard. Their
partial evaluator failures do not substitute for the complete run above.

Native measurements show accumulating empty Excel main windows, not additional
worksheet windows or forms. Test-host garbage collection does not release them.
A separate disposable Office experiment, without invSys packages, reproduces
one extra main window per transient open/close when only its hidden add-in
workbook remains. Normal close, hide/close, restore/close and ScreenUpdating
variants all reproduce it, with and without a modeless form. With a visible
disposable workbook, the main-window count remains two across all 64 reads;
the hidden-only phase adds 64 windows. Read-only fixture bytes remain unchanged.
This is evidence about this Excel test-host configuration, not an invSys
contract change or a general Office-version guarantee.

The corrected test retains its existing unrelated disposable sentinel workbook
until owner evaluation completes, after closing the Shipping launcher and
stopping its informational-dialog observer. It closes that workbook in finally.
All 33 native resource samples in the complete RED retain nine main windows;
maximum GDI/USER counts are 802/668. No product caching, hidden-window repair,
forced termination or command replay is introduced.

Some earlier native question captures are blank and are explicitly rejected in
their ignored review records. The capture helper now requires the exact owned
foreground question, copies its painted screen rectangle, checks prompt ink and
retries only capture before delivering the actual Yes/No once. These changes
affect test evidence only; the complete RED has no rejected images.

## Implementation and verification

The runtime correction adds the five exact IDs to the existing Core completion
map in `modEvaluationMatches.CommandCompleted` (+4 physical lines). It adds no
runtime component, procedure, dynamic call or generic success classifier.
The isolated `deploy/validation-owner-command-completion` candidate builds all
five XLAMs, passes Operations cold start and all five packaged compiles. A
component-by-component comparison of all 242 packaged components finds exactly
one code hash change: `invSys.Core.xlam/modEvaluationMatches`.

Regenerated static evidence passes all three JSON schemas and all 28 existing
oversized-module limits. It retains 249 components, 6,036 procedures, nine
literal Application.Run calls, 45 unresolved calls and 192 duplicate-body
candidate groups. Physical lines increase exactly four, to 132,555; every other
component's line count is unchanged. Unrelated user document pins are preserved.
The ignored `owner-command-completion-static` reports and
`verify-owner-command-completion-static.ps1` record the evidence. The first
offline schema-audit wrapper incorrectly expected a native-process exit code
from a directly invoked PowerShell script; invoking the validator as a child
process corrects that audit-only error. No runtime or test result changes.

The same packaged recording/editor/Evaluate route passes **460/460 GREEN**,
2026-09-22 **15:36:27.5479426--15:46:37.9532895 UTC**, report
`reports/runtime/slice4be-shipping-activity/f9bb412633dc46e3b56628689979c7b0/owner-completion-boxing-activity-shipping-recording-green.json`.
All 460 RED check identities are retained, including all 92 focused checks and
all four correctly negative outcomes. The eight formerly failing positive
conclusions now show the required Domain-application caveat. All original bytes
remain unchanged, exactly twelve separate evaluation results are appended and
generated configuration is restored. All 33 resource samples retain nine main
windows (maximum GDI/USER 802/669). Excel closes normally and immediately without
assistance; source/test/package pins remain intact and no matching Application
1000/1001/1002 events occur.

All 23 GREEN captures are directly reviewed and hashed. The twelve evaluation
results, three native questions, three Action Paths sizes and four Boxing actions
are accepted. `settings-save.png` has unpainted General Settings controls and is
explicitly rejected: **22 accepted / one rejected**. Fresh current-candidate
General Settings proof is required during regression; the rejected image is not
relabelled as accepted. `verify-owner-command-completion-green.ps1` records this
open capture gate separately from behavioral GREEN. Its first invocation used
the wrong test-pin filename; correcting the ignored verifier path changes no
executed test or source pin.

The first current-candidate full-chain attempt is **FAILED: 5 PASS / one harness
FAIL**; its ordered live-role subprocess has **14 PASS / one harness FAIL** at
**Invoke Receiving ConfirmWrites form action**, HRESULT `0x80020009`.
Create Warehouse remains 15/15. The failure is at that actual call boundary,
not established as startup failure or a completion-map defect. Its cause remains
unproven, and earlier records contain similar intermittent failures at this
boundary. No matching Application 1000/1001/1002 events occur. The runtime and
five package pins remain unchanged. Prefix `reports/runtime/owner-completion-chain`
preserves the failed reports and metadata; all three tracked reports are restored.

Cleanup is assisted normal cleanup: read-only native/COM inspection establishes
an exact process/creation/window identity and typed WorkbookCount zero, rechecked
before Quit. The resulting Recovery question is captured; **Yes, I want to view
these files later** is selected, captured, reviewed and confirmed. Recovery files
are retained, Excel exits, and local settings restore. No forced termination or
operational workbook mutation occurs. A separate fresh-fixture retry uses the
unchanged candidate; this failed attempt is not overwritten or counted as passed.

The fresh chain retry passes **32/32**, ordered live roles **48/48** and Create
Warehouse **15/15**, 2026-09-22 **15:55:42.6391739--16:02:34.4926551 UTC**.
Prefix `reports/runtime/owner-completion-chain-retry` preserves the reports and
verification. All preceding 32 chain and 48 live check identities remain GREEN;
the three tracked reports are restored exactly, package/test pins remain unchanged
at that gate, and no matching Application errors occur. Cleanup again requires
verified-empty normal Quit and captured, reviewed recovery retention; it is
assisted normal cleanup, not unattended closure. The successful retry neither
overwrites the first failure nor proves its cause.

The first UOM regression attempt records **84 PASS / one harness FAIL** because
the loaded General Settings form does not acquire foreground for capture. It
stops before the UOM commands, closes Excel normally and preserves source/package
pins. This is not product RED. Its prefix is `reports/runtime/owner-completion-uom`.
The capture-only correction delegates General capture to the existing unique
owned-form helper, which retries capture without replaying a command. The existing
disposable foreground calibration passes all three scenarios; all three images
are reviewed. This is the only test-source change after focused GREEN and the
chain retry; the focused route does not execute that UOM activity helper.

| Current-candidate regression | Result | Reviewed captures | Report root below reports/runtime |
| --- | --- | ---: | --- |
| Admin UOM retry | 228/228; every prior identity retained | 11 | slice4be-admin-uom/d920c028cd744a8aab2efc9722076a3f |
| Full Boxing/Shipping | 1,707 PASS; same seven D8-A FAIL | 22 | slice4be-shipping-activity/30163a07ab4d48e4af1b2ed14bbdc98c |
| How-To/Diagnostic/Compare | 421/421; every prior identity retained | 31 | slice4be-viewer-published-read/f589667dbe844a26bfcabf47b5b6324e |

The UOM retry runs **16:06:54.6149582--16:09:30.2166128 UTC**. Its fully painted
`admin-uom-general-loaded.png` and `settings-save.png` satisfy the fresh General
Settings capture gate; the earlier rejected focused image remains rejected.
All nine native Reset questions are painted. The full Boxing/Shipping gate runs
**16:09:32.1320506--16:25:10.0732401 UTC** and retains all 1,707 previous passes.
Its seven failures remain exactly
`Shipping.Access.AuthUnavailable.<Add|Update|Remove|Hold|Return|Stage|Send>.MissingFileNotRecreated`;
the gate is not fully GREEN and D8-A is not approved by this result.

Both regressions preserve runtime, current test and five package hashes, with no
matching Application 1000/1001/1002 events. UOM closes Excel immediately. Boxing
closes normally after a delay without assistance. Read-only inspection during
that delay cannot attach a typed Excel object; it establishes no workbook count,
and no additional Quit or forced termination is attempted. The process exits
naturally before the next gate. `verify-owner-completion-regression.ps1` checks
the exact baseline identities, image hashes, source/package pins and terminal
metadata; its early Boxing invocation before terminal metadata exists produces
no verification claim and changes no test result.

The comparison gate runs **16:25:11.8091078--16:55:38.9619186 UTC**, retaining all
421 prior checks, including the preceding 376-check evaluator baseline. The same
explicitly paired guide/run remains visible in How-To, Diagnostic and Compare
across pending, partial and fully applied evidence. Only full application
concludes; failure, restriction, stale evidence, cancellation, retries, immutable
expectations and changing selections retain their protecting checks. All 31
captures are reviewed and hashed. Excel closes normally and immediately without
assistance; current source/test/package pins remain intact and no matching
Application errors occur. The three regression gates supply 64 accepted captures.

The completion-map correction's applicable isolated gates are verified and the
specification, Plan 022 and controls acceptance records are synchronized. The
inactive Settings drafts are uncompiled, uninstalled preparation, not RED evidence
or implementation. Slice 4be and Release 1 acceptance are not claimed. The separate
D8-A decision and Event Detail proposal remain unapproved; the 26-control Settings
coverage refinement remains specified but unimplemented.
