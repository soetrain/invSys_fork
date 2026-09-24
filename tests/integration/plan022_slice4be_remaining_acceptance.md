# Slice 4be remaining acceptance checklist

Last reviewed: 2026-09-24 UTC. **Incomplete.** This is an evidence index for the
existing contract, not a new scope, architecture decision, or percentage estimate.

Authority: Architecture v4.11 D18 and Plan 022's six-sub-slice acceptance table.
The maintained controls catalog remains the control-level acceptance record.
Existing passing results retain their recorded candidate and scope; they do not
prove a newer candidate or an untested control family.

| Workstream | Verified foundation | Required to close |
|---|---|---|
| 4be.1 Comprehensive control coverage | Catalog 11 and 62 registered controls; packaged Receiving, Shipping, Boxing, Settings and UOM evidence. | Account for every reachable Operations/Admin action or explicit exclusion. Production currently registers only UOM Retrieve; its Process/Recipe workflow coverage remains open. Finish shared session/target, inventory-management, remaining role navigation and Admin maintenance coverage through actual callbacks and owner facts. Reconcile carrier authority with D5 before changing it. |
| 4be.2 Settings and profiles | Dedicated Event Tracking surfaces, policies/profiles, personal views and Operations access without Admin. Current candidate: 780 focused behavioral checks, 202 Settings regression checks. | Finish remaining General/lifecycle/navigation observations, preserve restart/isolation and context guards, and close the focused shutdown limitation. Treat Windows-user connection preferences separately from warehouse Config authority. |
| 4be.3 Comprehensive Events | Published projection, paging/filtering, contributing lines, selected detail and policy-aware reads have recorded tests. | Retain complete family/control coverage as 4be.1 expands. Resolve the pending Event Detail lock amendment; prove native horizontal scrolling and complete multiline rendering. Preserve read-only authority and all prior Viewer identities. |
| 4be.4 Recording and conclusions | Immutable recordings, correlation, interruption and explicit diagnostic owner-outcome mappings have focused evidence. The current comparison attempt's strengthened PolicyChangeClosesIncomplete assertion passes. | Exercise the complete required Operations/Admin coverage, including Production, and complete current-candidate comparison acceptance. No staging or submission may imply saved settings or Domain application. |
| 4be.5 How-To and comparison | Recorded and directly curated guides, immutable versions, expectation authoring/binding, and paired views have recorded tests. | Implement approved export/import requirements after specifying the exact compatible wire/provenance contract. Embedded foreign observations remain origin-only evidence. Preserve authoring/search/edit/restart and current-policy behavior. Validate both presentations and an incomplete example for Operations and Admin. |
| 4be.6 Release and user acceptance | Existing focused builds/compiles, static checks and multiple passing role regressions remain preserved. | Finish the current-candidate gates below; then run final applicable five-package build/compile/initialization, layout/header, static, launcher/inventory/reusable Production, restart/binding, live-role and full-chain gates. Obtain and record human comparison and applicable NAS acceptance separately from automated images. |

## Current candidate gate ledger

Candidate: `deploy/validation-settings-diagnostic`; no accepted deployment is
changed by these tests. Exact receipts and failed attempts are recorded in
[Settings diagnostics evidence](plan022_slice4be_settings_editor_activity_results.md).

| Gate | Current evidence | Remaining action |
|---|---|---|
| Focused Settings observations/diagnostics | 780/780 behavior, 54 reviewed images; shutdown limitation remains. | Resolve closure without reclassifying behavioral success as clean shutdown. |
| Settings regression | 202/202; 10 accepted images; normal unassisted closure. | Preserve; repeat only for a relevant change or new concern. |
| Owner-command regression | 460/460; 23 accepted images; normal unassisted closure. | Preserve; repeat only for a relevant change or new concern. |
| Admin UOM regression | 228/228; 11 accepted images; normal unassisted closure. | Preserve; repeat only for a relevant change or new concern. |
| Broader Boxing/Shipping | Current candidate retains all 1,707 prior passes and seven pending D8-A failures across 1,714 checks; all 22 captures reviewed and source/package hashes preserved. The completed test host needed teardown cleanup; strict closure verification fails. | Preserve behavioral evidence and resolve clean shutdown without repeating unchanged successful behavior merely to obtain another report. The seven D8-A failures remain failures. |
| Comparison | Follow-up attempt: 333 PASS / one harness exception, 96 prior checks unreached; 23 reviewed images, delayed unassisted closure, no Application failures and unchanged hashes. Both preservation failures and all 15 diagnostic-library captures pass. Pending and Partial paired views pass. Earlier failed attempts remain retained. | Isolate the new 0x800AC472 Applied pair-label read before another broad retry. Applied paired views and later expectation checks remain unverified. Earlier candidate's 421/421 and 31 images remain historical evidence. |
| Curation | Current candidate 77/77; all prior identities, 18 reviewed captures, five instrumented compiles, normal immediate closure, zero Application failures and unchanged hashes. | Preserve; library captures establish layout/direct-curation entry with unavailable recordings, not populated-library acceptance. |
| Presentation restart | New `validation-guide-layout-normalized` candidate: 40/40 retains all 36 preceding identities and adds four-size bounds/non-overlap checks after the focused unchanged-size layout RED. Zero redundant refreshes, two real-resize validations, five compiles, normal closure, preserved bytes and zero Application failures. Exactly one of 243 compiled components changes. Latest capture-disabled selected-view observation is 6.879 seconds. Earlier add-in-only, overlay and desktop failures remain recorded. | Obtain unobscured current captures and complete applicable regressions. Full responsiveness, native visible acceptance and add-in-only behavior remain unaccepted. See [layout correction](plan022_slice4be_layout_stability_results.md) and [earlier resource evidence](plan022_slice4be_gui_resource_results.md). |
| Viewer/guide | Earlier candidates: 94/351 checks respectively. | Verify on final applicable candidate; preserve exact prior identities and required visible evidence. Prepared current expected totals are 98/355 after reader-boundary additions. |
| Production completion boundary | `validation-production-quiet-boundary`: packaged 50 PASS/two expected FAIL RED, GREEN 52/52, original audit 14/14, five explicit compiles. Only mProduction differs among 243 compiled components. Both batches enter the primitive bridge with the captured workbook name and restore UI state. | Preserve this existing-contract correction. See [boundary evidence](plan022_production_quiet_boundary_results.md). Broader Production activity coverage remains open. |
| Reusable Production palette | The focused reusable gate fails one aggregate check on both the boundary candidate and its predecessor: EightPaletteRows=False. All 67 Boolean observations match; reusable and Chai completion envelopes report OK. | Preserve the accepted eight-row requirement. Measure default/minimum, expanded and restored palette geometry before correcting it; IntegralHeight setup order is an unconfirmed hypothesis. No unchanged broad retry. See [candidate/control evidence](plan022_production_quiet_boundary_results.md). |
| Full chain/live roles/Create Warehouse | Production boundary candidate: 32/48/15 passes with exact prior identities, normal unassisted closure, zero Application failures, settings/three reports restored and five package/263 tooling hashes preserved. Earlier layout/frozen candidate gates retain their scope. | Preserve; repeat only after a relevant change or new concern. Earlier crashes remain unexplained. See [current chain evidence](plan022_production_quiet_boundary_results.md), [layout candidate](plan022_slice4be_layout_stability_results.md) and [earlier projection/header evidence](plan022_slice4be_shutdown_header_results.md). Other workstreams and human acceptance remain open. |

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

- **Pending D8-A:** ordinary Auth reads must not create/repair authority, while
  explicit authorized setup retains provisioning. The proposal is not effective
  until the user approves and Architecture, Plan 022 and controls record it.
- **Pending Event Detail:** allow Locked=False only for selection/scrolling in a
  non-editable ListBox. Preserve current Locked=True and unrelated user edits
  until approval; native-scroll RED already exists.
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
