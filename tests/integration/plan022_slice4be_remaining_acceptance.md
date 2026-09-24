# Slice 4be remaining acceptance checklist

Last reviewed: 2026-09-24 UTC. **Incomplete.** This is an evidence index for the
existing contract, not a new scope, architecture decision, or percentage estimate.

Execution status: **resumed**, 2026-09-24. The user approved D8-A and Event Detail
scrolling, and desktop input/capture succeeds after host settings changed; the
earlier error5 cause remains unproven. Approved single-line scrolling passes
44/44 with four reviewed images. The current candidate retains Viewer 98/98,
owner 460/460 and full-chain/live-role/Create Warehouse 32/48/15, with normal
unassisted closure. Owner visual review rejects seven captures as clean acceptance.
Production subsequently passes 390/390 with ten reviewed images and normal closure.
D8-A is subsequently implemented on `validation-auth-read-separated`: focused
RED 60 PASS / 22 expected FAIL becomes GREEN 82/82; the updated unit passes 1/1
and full-chain/live-role/Create Warehouse retain 32/48/15. Five package compiles
pass; only Core/modAuth changes. Broader Shipping/Boxing reaches 1630 PASS / three
FAIL: all seven Auth-recreation checks pass, but two obsolete lock assertions and
a caption-capture exception remain; 82 previous checks are unreached. Sixteen
images are reviewed; host/Excel cleanup is assisted and settings are restored.
The stale Boxing probe is subsequently reconciled with approved D18: focused
49 PASS / one expected probe FAIL becomes 50/50 with all 44 earlier checks retained.
An isolated submission route passes 794/794, recovers all 82 unreached checks,
reviews seven images and closes normally. The complete corrected route reaches
1632 PASS / one repeated caption-capture exception, leaving 82 checks unreached;
both read-only assertions now pass. Sixteen images are reviewed. Only disposable
Excel termination is required; the original host writes the result and restores
settings, preserving packages/tooling with zero Application failures. Normal
shutdown is not accepted. The capture/cleanup cause remains unproven.
Multiline and the other requirements below remain open. See
[D8-A evidence](plan022_slice4be_auth_read_results.md).

Authority: Architecture v4.11 D18 and Plan 022's six-sub-slice acceptance table.
The maintained controls catalog remains the control-level acceptance record.
The subsequent [caption lifecycle investigation](plan022_slice4be_capture_lifecycle_results.md)
corrects an offscreen-input helper defect and passes eight disposable cases.
The full diagnostic remains 1632 PASS / one failure: SetWindowPos succeeds but
does not raise the on-screen, non-cloaked Shipping form; all caption hits belong
to VS Code, with only 816 GDI objects at failure. Assisted host/Excel cleanup is
recorded. This narrows the capture failure without closing it. Source review also
finds a literal-escape decoding risk requiring packaged RED as part of original-value
and multiline acceptance; no runtime behavior is changed by the diagnosis.
Existing passing results retain their recorded candidate and scope; they do not
prove a newer candidate or an untested control family.

The subsequent [original-text correction](plan022_slice4be_original_text_results.md)
restores the existing D18 decoder contract on `validation-detail-original-text`:
packaged RED 52 PASS / four expected FAIL becomes GREEN 56/56 through Operations
and Admin, retaining the previous 50. Viewer regression retains 98/98. Seven
images are reviewed across those two GREEN gates; both close normally and restore
settings. Five compiles and comparison of 244 components isolate the Viewer form
change. The unchanged-candidate chain retry passes 32/48/15 with exact previous
identities, normal closure, restored settings/reports and no Application failures.
The first attempt stops with an Excel/ntdll crash (chain 5 PASS / one harness
FAIL, live-role 32 PASS / one harness FAIL, Create Warehouse 15/15); verified
recovery-process termination allows the original controller to restore settings.
Its cause remains unproven and the failed attempt remains recorded.
That original-text gate alone does not establish complete visual reachability;
the subsequent multiline evidence is recorded below.

The subsequent [multiline candidate](plan022_slice4be_multiline_results.md)
implements a read-only scrolling area tied to the same selected line. Final
captured RED 57 PASS / 31 expected FAIL becomes **88/88**, preserving all 56 prior
checks and all seven native-input checks. Fourteen principal images establish
scoped visible reachability through the last of multiple fields and the wide
final line. Profile/sign-out/target clearing and exact original values remain
protected. The DPI detector correction has separate 4 PASS / four FAIL RED to 8/8 GREEN and
retains the locked negative case. Desktop input currently works; the earlier
error5 cause remains unproven. Current-candidate chain/live-role/warehouse creation
pass 32/32, 48/48, 15/15. Viewer regression passes 98/98 with all prior identities
and three reviewed images. This closes scoped automated multiline visibility, not human or
broader Slice4be acceptance.

The reviewed [Production form census](../../../invSys_docs/0%20plan%20docs/xlam_invSys/invSys-Production-Tracking-Coverage-v1.md)
accounts for 68 constructed buttons (seven registered tracking IDs, 61 pending),
14 unconstructed legacy buttons and 34 non-button handlers requiring explicit
reachability/semantic classification. It reconciles all 82 button constructions
and declarations. This closes a source-accounting gap for that form; it does not
implement pending tracking or establish complete Operations/Admin coverage.

The subsequent [Production header correction](plan022_production_header_results.md)
on `validation-production-header` passes focused20/20 after16 PASS/four expected
caption-fit failures. Four reviewed size captures show the full Committed / Used
heading, adjacent headings and eight palette rows without overlays. Five builds/
compiles, cold start, layout8/8+7/7, unchanged static metrics and current-candidate
chain/live-role/Create Warehouse32/48/15 pass. This closes the discovered heading
defect; human and full-workflow visual acceptance remain open.

| Workstream | Verified foundation | Required to close |
|---|---|---|
| 4be.1 Comprehensive control coverage | Catalog 12 and 68 registered controls; packaged Receiving, Shipping, Boxing, Settings and UOM evidence plus the first six Production draft handlers. Positive Recipe validation uses an actual released Process (244/244); both designer captures are reviewed. | Account for every reachable Operations/Admin action or explicit exclusion. Finish the six draft controls' remaining visible evidence and applicable regressions, then remaining Process/Recipe workflows. Finish shared session/target, inventory-management, remaining role navigation and Admin maintenance coverage through actual callbacks and owner facts. Reconcile carrier authority with D5 before changing it. |
| 4be.2 Settings and profiles | Dedicated Event Tracking surfaces, policies/profiles, personal views and Operations access without Admin. Settings diagnostic candidate: 780 focused behavioral checks, 202 Settings regression checks. | Finish remaining General/lifecycle/navigation observations, preserve restart/isolation and context guards, and close the focused shutdown limitation. Treat Windows-user connection preferences separately from warehouse Config authority. |
| 4be.3 Comprehensive Events | Published projection, paging/filtering, contributing lines, selected detail and policy-aware reads have recorded tests. Event Detail now has 88/88 focused GREEN and 14 reviewed principal captures, including complete multiline scrolling. | Retain complete family/control coverage as 4be.1 expands. Finish broader candidate gates and human acceptance. Preserve read-only authority and all prior Viewer identities. |
| 4be.4 Recording and conclusions | Immutable recordings, correlation, interruption and explicit diagnostic owner-outcome mappings have focused evidence. The current comparison attempt's strengthened PolicyChangeClosesIncomplete assertion passes. | Exercise the complete required Operations/Admin coverage, including Production, and complete current-candidate comparison acceptance. No staging or submission may imply saved settings or Domain application. |
| 4be.5 How-To and comparison | Recorded and directly curated guides, immutable versions, expectation authoring/binding, and paired views have recorded tests. | Implement approved export/import requirements after specifying the exact compatible wire/provenance contract. Embedded foreign observations remain origin-only evidence. Preserve authoring/search/edit/restart and current-policy behavior. Validate both presentations and an incomplete example for Operations and Admin. |
| 4be.6 Release and user acceptance | Existing focused builds/compiles, static checks and multiple passing role regressions remain preserved. | Finish the current-candidate gates below; then run final applicable five-package build/compile/initialization, layout/header, static, launcher/inventory/reusable Production, restart/binding, live-role and full-chain gates. Obtain and record human comparison and applicable NAS acceptance separately from automated images. |

## Current candidate gate ledger

Base candidate: `deploy/validation-settings-diagnostic`; later candidates are
identified per gate. No accepted deployment is changed by these tests.
Exact Settings receipts and failed attempts are recorded in
[Settings diagnostics evidence](plan022_slice4be_settings_editor_activity_results.md).

| Gate | Current evidence | Remaining action |
|---|---|---|
| Production draft observations | Approved Detail candidate: 390/390 retains every prior identity, with five instrumented compiles, ten reviewed images, normal delayed unassisted closure, restored settings, unchanged packages and zero Application failures. Earlier RED and incomplete attempts remain preserved. | Scoped six-control path/visible gate is GREEN; remaining comprehensive Production coverage is open. See [path evidence](plan022_slice4be_production_paths_results.md). Earlier desktop and harness failure causes remain unproven. |
| Focused Settings observations/diagnostics | 780/780 behavior, 54 reviewed images; shutdown limitation remains. | Resolve closure without reclassifying behavioral success as clean shutdown. |
| Settings regression | Corrected Production-caption candidate: 202/202 retains exact prior checks; ten accepted images, immediate unassisted closure, restored settings and unchanged five package/211 test hashes. All captures succeed without elevation. | Preserve; repeat only for a relevant change or new concern. Earlier error5 cause remains unproven. |
| Owner-command regression | Approved Detail candidate: 460/460 retains every prior identity; normal delayed unassisted closure, restored settings, unchanged packages/tooling, zero Application failures. All 23 images reviewed: 16 clean, four with a taskbar overlay, three incorrect dialog crops. | Preserve behavioral GREEN. Correct the seven limited/rejected captures through focused evidence work; do not infer clean visual acceptance from the automated capture predicates. Earlier 23 accepted images retain their historical scope. |
| Admin UOM regression | 228/228; 11 accepted images; normal unassisted closure. | Preserve; repeat only for a relevant change or new concern. |
| Broader Boxing/Shipping | D8-A candidate with reconciled probe: 1632 PASS / one repeated caption-capture exception; 82 previous checks unreached, 1625 prior GREENs retained. Both read-only assertions and all seven Auth-recreation checks pass. Sixteen scoped images reviewed; packages/217 tooling hashes preserved, zero Application failures. Disposable Excel termination is required; original host writes the result and restores settings. Isolated submission route separately passes 794/794 with all 82 unreached checks and seven images, normal closure. | Observe hit-test/z-order/cloaking/resource state at the caption failure before another full retry; obtain broader GREEN and clean shutdown. The isolated result does not replace full-route coverage. Historical partials remain preserved. See [probe/capture evidence](plan022_slice4be_detail_overflow_results.md). |
| Comparison | Follow-up attempt: 333 PASS / one harness exception, 96 prior checks unreached; 23 reviewed images, delayed unassisted closure, no Application failures and unchanged hashes. Both preservation failures and all 15 diagnostic-library captures pass. Pending and Partial paired views pass. Earlier failed attempts remain retained. | Isolate the new 0x800AC472 Applied pair-label read before another broad retry. Applied paired views and later expectation checks remain unverified. Earlier candidate's 421/421 and 31 images remain historical evidence. |
| Curation | Current candidate 77/77; all prior identities, 18 reviewed captures, five instrumented compiles, normal immediate closure, zero Application failures and unchanged hashes. | Preserve; library captures establish layout/direct-curation entry with unavailable recordings, not populated-library acceptance. |
| Presentation restart | New `validation-guide-layout-normalized` candidate: 40/40 retains all 36 preceding identities and adds four-size bounds/non-overlap checks after the focused unchanged-size layout RED. Zero redundant refreshes, two real-resize validations, five compiles, normal closure, preserved bytes and zero Application failures. Exactly one of 243 compiled components changes. Latest capture-disabled selected-view observation is 6.879 seconds. Earlier add-in-only, overlay and desktop failures remain recorded. | Obtain unobscured current captures and complete applicable regressions. Full responsiveness, native visible acceptance and add-in-only behavior remain unaccepted. See [layout correction](plan022_slice4be_layout_stability_results.md) and [earlier resource evidence](plan022_slice4be_gui_resource_results.md). |
| Viewer/guide | Multiline candidate Viewer: 98/98 retains all prior identities; three reviewed images, normal closure, restored settings and unchanged packages. Earlier guide candidate: 351 checks. | Preserve Viewer evidence; verify guide on final applicable candidate with expected 355 checks after reader-boundary additions. |
| Production completion boundary | `validation-production-quiet-boundary`: packaged 50 PASS/two expected FAIL RED, GREEN 52/52, original audit 14/14, five explicit compiles. Only mProduction differs among 243 compiled components. Both batches enter the primitive bridge with the captured workbook name and restore UI state. | Preserve this existing-contract correction. See [boundary evidence](plan022_production_quiet_boundary_results.md). Broader Production activity coverage remains open. |
| Reusable Production palette/header | Header candidate: focused20/20 retains all16 preceding checks after four expected caption-fit failures. Four reviewed sizes show the complete Committed / Used heading, clear neighbors and eight display-only rows without overlays; normal focused closure holds. Earlier broader reusable gate retains67 Boolean observations. | Preserve scoped visibility and unchanged captions/identities. Human/full-workflow visual acceptance remains open; broader reusable cleanup is separately unproven. See [header evidence](plan022_production_header_results.md) and [palette/fixture evidence](plan022_production_palette_results.md). |
| Full chain/live roles/Create Warehouse | Multiline candidate: 32/48/15 passes with exact prior identities, normal unassisted closure, zero Application failures, settings/three reports restored and package/tooling hashes preserved during the gate. Packaged smoke also passes 86/86. Earlier candidates retain their scope. | Preserve; repeat only after a relevant change or new concern. Earlier crashes remain unexplained. See [multiline evidence](plan022_slice4be_multiline_results.md). Other workstreams and human acceptance remain open. |

## Decisions and execution limits

The subsequent [CRLF build correction](plan022_crlf_build_regions_results.md)
passes 12/12 and produces 243 compiled components identical to the validated
layout candidate. Its regression audit also finds a pre-existing D12 Production
boundary failure (13/14, reproduced on the predecessor): completion passes a
Workbook to Core `BeginQuietUi` instead of the existing primitive bridge. A
packaged completion-boundary RED was required before correction; the passing chain
did not waive the architectural requirement. This was a discovered Release 1 blocker,
not an approved contract exception.
The subsequent [Production boundary correction](plan022_production_quiet_boundary_results.md)
has a packaged two-batch RED (50 PASS/two expected FAIL), GREEN 52/52, original
audit 14/14 and five explicit compiles. Both completions use the captured-name
bridge and restore UI state; only mProduction differs among 243 compiled
components. Current-candidate regression gates remain recorded separately in
that evidence file. This restores D12 without changing its contract.

A fresh-session projection control retains its 6/6 behavior and failed strict
worker-lifetime closure result. A later same-session control passes 35/35 and
the complete current-candidate chain passes 32/48/15 with normal closure. Earlier
crashes remain unexplained; preserve all attempts without invalidating the scope
of the later verified result.

- **D8-A approved 2026-09-24:** ordinary Auth reads must not create/repair
  authority; explicit authorized setup retains provisioning. Architecture,
  Plan 022 and controls record the effective decision. Focused Core packaged RED
  is 60 PASS / 22 expected FAIL, followed by 82/82 GREEN; implementation preserves
  explicit provisioning and the D2 credential-authorized station transition.
  Current candidate chain/live-role/Create Warehouse pass 32/48/15. All seven
  Shipping Auth-recreation checks now pass in a partial broader run; two obsolete
  lock assertions, capture failure and assisted cleanup prevent broader GREEN.
- **Event Detail approved 2026-09-24:** allow Locked=False only for
  selection/scrolling in a non-editable ListBox. Preserve unrelated user edits;
  fresh native-scroll RED is 43 PASS / one expected movement FAIL, followed by
  44/44 GREEN and four reviewed captures on validation-approved-detail-scroll.
  All prior checks remain; native typing cannot edit values. Five compiles and
  the single-form component delta pass. Subsequent multiline native/visible
  evidence is 88/88 with 14 reviewed principal images; broader and human acceptance
  remain open.
- **Carrier authority:** Windows-user persistence conflicts with D5's warehouse
  Config authority. No hybrid, implied UOM permission, or silent migration.
- Guide transfer details are still a proposal. D18 already requires export/import;
  do not substitute local-only guides for that accepted outcome.
- A new contract amendment requires normative approval first. New discovered
  controls that inherit D18 require exact catalog/plan/control entries and D13,
  not repeated approval of the already accepted synthesized contract.
- Run one Excel gate at a time on disposable fixtures. Never rebuild/deploy with
  relevant workbooks or add-ins open. Preserve operational files and unrelated edits.
- When desktop access fails, retain the terminal attempt and stop visible retries
  until access changes. A successful short calibration does not prove sustained
  access; the cause of the observed intermittent Windows error 5 is unresolved.
- Existing GREENs are not rerun merely because time elapsed. Harness failure is
  not product RED. Require a changed precondition, fix or new evidence before retry.
- Revalidate approval, desktop/process state and the candidate before dependent
  work. Do not interpret automatic continuation as architecture approval.

Latest desktop checkpoint (03:05:30 UTC, 2026-09-24): cursor access succeeds and
the blank-form capture calibration passes all three cases without elevation.
The current palette candidate's presentation retry stops earlier at Create guide
fixture dispatch (10 PASS/one harness exception); no product screenshots or
restart/layout acceptance follow. Both attempts close Excel normally. Earlier
error5 remains unexplained; the user uses console and RDP, so session switching
is a hypothesis only. Inspect that exact fixture boundary before another broad
retry. The [layout evidence](plan022_slice4be_layout_stability_results.md) records
the preserved attempts and verification receipt; this supersedes only the stale
desktop-access observation in the ledger, not any product acceptance status.
