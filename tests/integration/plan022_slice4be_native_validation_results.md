# Plan 022 Slice 4be native validation investigation

Last verified: 2026-09-12 PDT. Runtime source remains **0aa0084**; this is a
diagnostic continuation of the [Shipping timer checkpoint](plan022_slice4be_shipping_context_results.md),
not a runtime repair, architecture change or Release 1 acceptance claim.

## Evidence boundary

The ordinary live-role gate failed at Production Complete Run after39 passing
checks on both `validation-shipping-timer-context` and the preceding pinned
`validation-shipping-context-guard` set. The initial unsaved fixed-phase trace
also failed39/1, ending at P08 before quiet-UI entry, typed completion and quiet-UI
exit. It never reached P09 before output restoration. Those runs and native
faults remain in the linked checkpoint; a traced pass cannot replace them.

New diagnostics use the same existing live-role script and disposable fixtures.
They retain all48 business assertions, actual owner calls and authorization.
Only unsaved VBA inspection/instrumentation changes; no package, Production
source, accepted deployment or operational/NAS workbook is saved. Trace output
contains fixed phase markers only, never payloads, credentials or inventory values.
All raw reports/scripts below remain ignored under
`reports/runtime/slice4be-shipping-activity/`.

| Variant | Observed result | Interpretation |
|---|---|---|
| Expanded trace, initial setup | 0 PASS /1 harness failure: service source anchor not matched | No workflow RED. The diagnostic was corrected for VBE identifier casing. |
| Expanded trace, calibrated | 48/48; reaches P01-P11, Q01-Q03, E01-E05 and K01-K08 in actual call order | Queueing, persistence, processing/refresh, quiet-UI exit and output restoration all execute in this variant. It does not prove an uninstrumented repair. |
| Source-only service reload | CodeModule text remains identical apart from trailing line breaks; no markers;39 PASS /1 native/RPC harness failure at Complete Run | Reloading that service source alone does not reproduce the traced pass. Timing or broader compilation effects remain unresolved. |

The expanded trace's Q01 follows quiet-UI entry; Q02 precedes quiet-UI exit and
Q03 follows it. E markers surround service queue/persistence/processor/result
boundaries. K markers surround consume and completion queue returns and session
updates. Repeated E02 represents separate persistence calls, not duplicate trace
delivery. No marker alone proves business completion; the48 retained assertions
supply the traced variant's workflow evidence.

Exact new evidence:

- `timer-native-detail-run.log` and `timer-native-detail-setup-failed-results.md`:
  initial0/1 setup failure.
- `timer-native-detail-calibrated-run.log` and
  `timer-native-detail-calibrated-results.md`:48/48.
- `timer-native-detail-trace.txt`, `timer-native-detail-map.json` and
  `prepare-native-trace-detail.ps1`: fixed markers, phase map and exact injection.
- `timer-native-source-reload.log`, `timer-native-source-reload-results.md`,
  `timer-native-source-reload.json` and `prepare-native-source-reload.ps1`:
  identical-source comparison and failing39/1 run.

## Next controlled comparison

Source inspection found a mouse-scroll hook implementation but did not establish
its use in this workflow. No hook is disabled, removed or blamed from source
existence alone. No sleeps, retries, error suppression or business changes are
introduced as a runtime workaround.

A fresh isolated five-package build from unchanged0aa0084 source is available at
`deploy/validation-shipping-timer-rebuild`. All five build and compile checks plus
Operations cold start pass. All170 exported component source hashes match the
earlier timer candidate exactly (`timer-rebuild-source-comparison.json`).
The ordinary uninstrumented live-role gate passes48/48 on this rebuilt set;
`timer-rebuild-live-role.log` and `timer-rebuild-live-role-results.md` record it.
This is artifact evidence, not proof of the native fault's source cause.

The first ordinary reusable Production run on the rebuilt set signs in but
fails0/1 with native/RPC failure at the batch-scale contract. Its report/log use
`timer-rebuild-production`. That contract first calls the real Production
launcher, then checks numeric bounds; the coarse progress marker cannot say
which part failed.

The separate unsaved launcher trace wraps those two steps with B01-B04 and
records the existing launcher's L/S stage names. It reaches all four B markers
and passes the full2/2 reusable Production gate. Its clean-restart session loads
the original packages without that first-session instrumentation. This is still
a diagnostic variant, not an ordinary2/2 replacement. Evidence:
`timer-production-launch-diagnostic/production-reusable-production.md`,
`timer-production-launch-diagnostic.log`, `timer-production-launch-trace.txt`,
`timer-production-launch-map.json` and `prepare-production-launch-trace.ps1`.
The first preparation attempt rejected a duplicate source anchor before Excel
started; the unique configure/sign-in anchor resolves that diagnostic setup issue.

The ordinary uninstrumented full-settings comparison passes **2/2**, including
saved-workbook clean restart, in `timer-rebuild-production-uninstrumented`.
No reduced Production flags or unsaved diagnostic markers are used in this run.
Do not combine the earlier timer candidate's Shipping/chain results with this
rebuilt set's live-role result to claim one fully accepted candidate.

## Test-maintenance corrections

The earlier reusable Production fixture failure reported `SignedIn=False` yet
continued into callbacks and waited on a sign-in notice. The validator now throws
a fixed, sanitized harness error before workflow callbacks when fixture sign-in
fails. It does not retry sign-in, bypass authorization or change a valid fixture's
workflow. The prior0/2 remains failed setup, not meaningful Production RED.

The source-contract checker also retained a pre-extraction inline-constructor
pattern for Shipping reuse. It reports23 PASS /1 stale source assertion against
the already tested context-helper implementation. The checker now follows typed
`CanReuse`, `HasCurrentContext` and `CreateBoundForm`/captured-workbook binding;
all24 source checks pass. Packaged context/reuse tests remain the stronger evidence.
Logs: `timer-launcher-source-contracts.log` and
`timer-launcher-source-contracts-current.log`.

These are test-infrastructure maintenance changes under unchanged D13/D18 rules,
not new runtime behavior, architectural exceptions or manufactured behavioral RED.
Parser/diff checks pass. Static maintenance is regenerated for the updated test
references:177 components,5510 procedures,123755 lines,8 literal and45 unresolved
Application.Run calls,195 duplicate-body candidates and28 module limits remain
unchanged. The source-checker correction adds three supporting test references;
it authorizes no deletion or runtime exception. No Production, Core, Domain,
form or launcher runtime source changes.

## Rebuilt candidate gate continuation

The fresh Shipping run is **322 PASS /46 missing-activity FAIL** across the same
368 checks. Comparison proves zero lost GREENs, missing checks, duplicate check
identities or non-activity failures. All30 interruption/timer checks remain GREEN.
Exact evidence is `123de3dea09b4f94876d20d1f72e6f01/green.json`,
`timer-rebuild-shipping-green.log` and `timer-rebuild-shipping-comparison.json`.
The capture in that run, `shipping-stale-session.png`, was inspected: the notice
is legible; the prior key-editor/button overlap and adjacent-header crowding
remain. This is automated visible evidence, not human layout acceptance.

The remaining gate queue was confirmed absent, with no remaining gate logs and
no Excel process. Its first launch failed parsing a trailing comma before any
test or Excel startup. Removing that comma from the ignored queue helper allows
the ordinary gates to run. This setup failure is not D13 behavioral RED.
Gate results below must all identify `validation-shipping-timer-rebuild`; earlier
candidate passes cannot fill a missing gate on this set.

| Ordinary rebuilt-set gate | Result | Exact ignored evidence |
|---|---|---|
| Packaged smoke | 86/86 | `timer-rebuild-smoke.log` |
| Live-role | 48/48 | `timer-rebuild-live-role-results.md` |
| Full reusable Production and restart | 2/2 | `timer-rebuild-production-uninstrumented/production-reusable-production.md` |
| Ordered full Release1 chain | 30/30 | `timer-rebuild-full-chain-results.md`, `timer-rebuild-full-chain.log` |
| Viewer | PASS | `timer-rebuild-viewer-results.md`, `timer-rebuild-viewer.log` |
| Combined launchers, first run | 1 PASS /1 harness failure at Production batch-scale callback | `timer-rebuild-launchers/packaged-launcher-noeligible.md`, its `progress.txt`, `timer-rebuild-launchers.log` |
| Shipping layout | 1/1 | `timer-rebuild-shipping-layout/shipping-layout.md`, `timer-rebuild-shipping-layout.log` |

The combined launchers fixture is signed in; the failure is RPC0x800706BE with
ntdll.dll/c0000028 at2026-09-13T01:52:01Z. The passing full chain also has a
combase.dll/c0000005 fault at01:49:25Z. Neither is cleared by the separate ordinary
Production2/2 or traced passes. The sanitized event list
`timer-native-followup-faults.json` retains these plus the source-only diagnostic
ntdll failure at01:05:44Z and first rebuilt Production failure at01:15:00Z.
Native cause and acceptance remain unresolved; no speculative runtime fix is made.

The tracked full-chain and Viewer result files were copied into ignored evidence
before restoring their original worktree versions. No machine-specific raw report
is included in this checkpoint.

## Actual workbook-close coverage

`Slice4beShippingWorkbookClose.ps1` extends the preserved packaged Shipping suite
through the real workbook-close event, followed by the registered timer callback.
Its unsaved facade returns only whether the launcher/callback retains a form and
the loaded Shipping form count. It never calls the close handler directly.
Saved staging/key/unknown-column values, authority bytes and an unrelated active
workbook remain independently checked; workbook shutdown is not a user Close click.

The first run stops **319 PASS /48 FAIL**,367 checks, before workbook closure.
The shared COM fixture deliberately sets EnableEvents=False at startup, so the
test's normal-events prerequisite fails and its harness exception follows. The
46 missing-activity failures remain; no executed prior GREEN is lost, but four
tail checks are not reached. This is setup failure, not product RED. Exact run:
`f5d82bb828d445d38f6f093a4269ea4d/green.json` and `timer-workbook-close-green.log`.

The corrected fixture invokes normal `modShippingInit.ShippingPackageAutoOpen`
to establish its event hook (COM package loading does not invoke Auto_Open), then
enables normal events for actual close/reopen and restores the prior harness
setting in finally. It neither suppresses that close event nor changes runtime
source. Its first corrected run stops67 PASS /1 native/RPC harness failure before
Shipping, at02:00:20Z with ntdll.dll/c0000028; evidence is
`be0f9b149403452cb2c24c2ec5b99024/green.json` and
`timer-workbook-close-events-green.log`. The next launch is safely refused because
an Excel recovery process remains (`timer-workbook-close-events-retry.log`); no
test executes. That exact recovery process is checked twice with zero workbooks
and closed before the calibrated retry. Neither failure is behavioral RED.

The unchanged corrected suite then passes **333 checks with46 missing-activity
failures**,379 checks total, including **11/11 actual workbook-close checks**.
Every original368 check and all322 prior GREENs remain, with zero duplicate,
missing, newly failing prior GREEN or non-activity failure. The normal close
releases the form and callback binding; late callback dispatch does not reopen
either, enter the staging/submission owner, fabricate activity, or change saved
workbook/authority/unrelated-workbook bytes. Read-only reopen preserves all active
and held staging values, exact keys and unknown columns. Exact evidence:
`572687e4648540f698383a8ccbc12360/green.json`,
`timer-workbook-close-events-calibrated.log`, `timer-workbook-close-comparison.json`.
This ordinary lifecycle case does not claim coverage of a separately retained
stale form, capability loss or missing activity outcomes. No runtime repair was
needed for the conforming close behavior.

All test handles are terminal and Excel is closed. Final verification matches45
package pins across nine preserved sets and eight current Receiving/Shipping
source pins (`timer-rebuild-final-pin-verification.json`). Static regeneration
adds only test-reference metadata and timestamps; runtime metrics and all28 limits
remain unchanged. Accepted deployment, operational/NAS workbooks and unrelated
user changes remain untouched. The next D13 work is Shipping capability-loss and
pending/uncertain owner outcomes, followed by the synchronized catalog refinement
and activity implementation. Full4be.1-4be.6 and Release1 acceptance remain open.
