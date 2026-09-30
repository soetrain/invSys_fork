# Slice 4be Production Close observations

Last verified: 2026-09-30 UTC. Runtime is unchanged; focused RED diagnosis is open.
No deployment promotion, completed control coverage or human acceptance.

Architecture v4.11's Production Close clarification, Plan022 and controls1.324
were committed as documentation5da0371 before tests2dc77af. Catalog21 reserves
PRODUCTION_CLOSE under D18 semantic inheritance. It observes existing button/native
dismissal without adding a permission gate, saving or posting work, retaining
unloaded form drafts, or attributing internal unload/workbook shutdown to a user.
The runtime baseline is catalog20 in frozen `deploy/validation-detail-columns`.

## Initial packaged run

`Test-Slice4beProductionClose.ps1 -DeployRoot deploy/validation-detail-columns -Phase RED`

93 PASS /139 FAIL across232 unique checks,2026-09-30
13:40:31.0605129--13:43:25.5068535 UTC. All42 existing common command/activity
checks retain their exact relative order and remain GREEN. Five instrumented
compiles pass. Settings/package pins are preserved, Excel closes unassisted and
the delayed Application audit finds zero Excel failures.

Controller: `reports/runtime/production-close-controller/abf6294100744153ba601d572be06ff5`.
Result: `reports/runtime/slice4be-production-close/f9f373a7876b4b14b136590f804a03f7/red.json`.

137 failures are expected missing-contract evidence: catalog21 is unavailable
(including105 definition comparisons at the unsupported new version), Close
metadata/outcomes/terminal classification are missing, four actual dismissal
routes lack correlated records, and unavailable optional tracking produces no
Close notice. Existing catalog20 definitions are not claimed to be broken.
Button/native dismissal, internal-unload exclusion, stale-context disposal,
actual permission-loss disposal, disabled-tracking disposal, UOM staging/custom
column retention on public reopen and prior-record preservation pass.

Two additional failures prevent treating this as the clean protecting RED:
`ProductionClose.Public.WorkbookShutdownDisposes` and
`ProductionClose.SavedAuthorityPreserved`. Their causes are initially unresolved.
The harness starts with Excel events disabled; this is a concrete event-delivery
lead, not yet proof of the observed disposal failure. Authority comparison does
not retain changed row values or identify a causal command in the first run.
Neither failure authorizes a runtime repair or a weakened preservation assertion.

Four images are directly reviewed and hashed in `visible-review.json`. Three show
the visible Production form, readable Close button/native X and public reopening;
they do not themselves prove disposal. The shared Settings image distinguishes
successful configuration save from unavailable optional tracking. This is bounded
agent review, not human acceptance or full Production visual acceptance.

## Diagnosis and next action

Test-only commit778fdde adds unsaved workbook-close owner-entry observations,
public event/binding/lifetime metadata, and per-stage authority hash comparisons
with fixed file categories. It changes no business handler or runtime package.
Its rerun retains the same232 ordered checks and93 PASS/139 FAIL, all42 existing
checks GREEN, five compiles, preservation, normal closure and zero delayed Excel
Application failures,13:46:17.9750382--13:49:13.9224884 UTC. Controller:
`reports/runtime/production-close-controller/8614df150eda4fe99d5eda07b3da15e5`;
result `reports/runtime/slice4be-production-close/b039113592144c6d8ef70be5df3f2d4c/red.json`.
`close-public-lifetime.json` establishes EventsAtEntry=False,
EventsBeforeClose=False, OwnerBoundBeforeClose=True, WorkbookCloseEntries=0 and
LoadedFormsAfterClose=1. This identifies the disposal failure as a harness
event-delivery problem; it does not establish a runtime close defect.

`close-authority-checkpoints.json` shows all14 pinned generated workbooks unchanged
after permission restoration, policy commands and immediately before the public
launcher. After public launch/reopen, one Inventory workbook hash differs; the
other13 remain unchanged. The precise operation and workbook contents changed
are unresolved. Runtime remains unchanged; do not reset the pin to hide this.

Test-only461f900 enables/restores Excel events for the public native-close case
and records file-container part hashes at each public step. Its rerun is94 PASS /
138 FAIL across the same232 ordered checks, all42 prior checks GREEN, five compiles,
preservation, normal closure and zero delayed Excel Application failures:
13:50:55.0146999--13:53:49.3423852 UTC. Controller:
`reports/runtime/production-close-controller/eeb00f6031eb4606b2685c1aef856260`;
result `reports/runtime/slice4be-production-close/f8676a61a6fe4ceb8e2460a329b652e8/red.json`.
The public lifetime trace now shows events enabled, the original owner bound,
one workbook-close owner entry and zero loaded forms afterward. That resolves
the harness disposal failure without a runtime change.

The additional preservation failure remains. Container-part hashes are identical
until the first public Production launch. That launch changes the canonical
`WHx.invSys.Data.Inventory.xlsb` hash and six parts: `docProps/core.xml`,
`xl/styles.bin`, `xl/tables/table3.bin`, `xl/workbook.bin`,
`xl/worksheets/sheet4.bin` and `xl/worksheets/sheet9.bin`. Its file hash then stays
identical through UOM workbench setup, button dismissal, reopening and workbook
shutdown. This proves an initial-launch write beyond document-property metadata;
it does not identify changed business values or a supplying procedure yet.

This is a discovered Release1 blocker under Architecture D3's explicit read-only
Domain query rule and D9/D10's snapshot/read split. Core's
`modInventoryDomainBridge.OpenOrCreateCanonicalInventoryWorkbookLocal` is a
candidate path: the read bridge can invoke schema ensure and save a writable
authority workbook. Static reachability is not yet the supplying-call proof.
Do not reset the preservation pin, classify this as optional activity, or treat
an unchanged repeat as a repair. Trace the initial public callback into the
query/resolver/schema/save boundary next; retain its real owner behavior until
the protecting evidence identifies the correction. Close runtime implementation,
independent Action Paths and the remaining acceptance gates are still pending.
