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

## Supplying-call evidence (2026-09-30 UTC)

Test-only61aa9c4 observes fixed procedure tags and Saved/ReadOnly booleans in the
unsaved Core project around the real public launcher. Its first harness attempt
stops before behavior at a case-sensitive VBIDE identifier match; this is not
product RED (`production-close-controller/8f8c773fe2014b60b653e3dff02019f7`).
Case-insensitive matching retains unique anchors. The subsequent run gives the
same232 ordered checks,94 PASS/138 FAIL, five compiles and42 prior GREEN checks,
normal closure/preservation and zero delayed Excel Application failures:
14:01:29.5668347--14:04:27.2354102 UTC, controller
`reports/runtime/production-close-controller/4a1459ecc49a481c82aa60b754528955`,
result `reports/runtime/slice4be-production-close/e1686754a4114bf984a7905cf25bd164/red.json`.
That run's trace was mistakenly stored under the disposable fixture root and
removed by normal cleanup. It is not retained supplying-call evidence.

Test-only0e6840e corrects the report destination. The unchanged frozen candidate
again yields94 PASS/138 FAIL, retaining232 checks and all42 prior checks, with
five compiles, normal closure, package/settings preservation and zero delayed
Excel Application failures:
14:05:07.2190753--14:08:06.9203764 UTC. Controller:
`reports/runtime/production-close-controller/dd9eba7fd43b4a99a5958daaa2c85c42`;
result `reports/runtime/slice4be-production-close/7242c356545a4d9c9325126ae66177a0/red.json`.
Its retained `close-public-inventory-trace.tsv` has this exact sequence:
Picker.Query -> Resolver.Enter -> OpenOrCreate.Enter ->
Schema.Before(Saved=True,ReadOnly=False) -> Schema.Unprotect ->
Schema.After(Saved=False,ReadOnly=False) -> Save.Before(False,False) ->
Save.After(True,False). No add-column, blank-row deletion, ROW-header deletion,
or projection-rebuild tag occurs in this bounded trace. These are observed
boundaries, not a claim that all possible internal changes were instrumented.

This identifies the supplying read path and an actual schema-unprotect/save
sequence during initial launch. Static ownership is
`frmProduction.EnsureRunInventoryCache` ->
`mProduction.LoadProductionRunInventoryPickerItems` ->
`modInventoryDomainBridge.ListInventoryPickerItemsBridge` ->
`ResolveInventoryWorkbookBridge` -> `OpenOrCreateCanonicalInventoryWorkbookLocal`
-> `EnsureInventorySchemaLocal` / `EnsureWorksheetEditableLocal` -> save.
The existing D3 read-only contract therefore requires a query-specific resolution
and lifetime correction, not a Close workaround or a new write permission.
Normative/docs1db3949 clarify this before runtime changes. Supplemental seeded
query preservation/result tests are the immediate test-first step. Runtime still
matches0b72dff; catalog21/Close and whole-release acceptance remain pending.

Supplemental test39b3144 adds65 checks for four seeded Core query bridges:
nonempty exact Domain results, byte preservation, transient cleanup, implicit
and explicitly supplied dirty callers, protected sheets/custom values and missing
stores. Its initial run fails in setup before any supplemental check and misses
the final three shared checks (91 PASS/139 FAIL,230 total); it is not product RED.
Controller `reports/runtime/production-close-controller/fc7863ee2d8945888808321975dd696f`,
result `reports/runtime/slice4be-production-close/3794ec7b4c81488babc6f1fb869abd68/red.json`,
14:08:19.8936258--14:11:32.9334456 UTC. Five compiles pass, closure is normal,
settings/packages are preserved and the delayed Excel Application audit is zero.
The setup exception is a workbook-open failure; its exact failing statement was
not retained. Do not infer a product cause from this failure.

Test79f0194 restores the original target after supplemental tests. Test8e720eb
uses the existing second Admin-generated fixture after all Close preservation
checks, seeds it explicitly, and adds fixed setup-stage metadata. This avoids
requiring another Generate call in the active public-form session. It does not
claim to diagnose that failed Generate/setup path or change runtime behavior.

The revised fixture repeats the setup exception at the copied reference-workbook
open, before supplemental checks:91 PASS/139 FAIL,14:12:41.6309900--14:15:48.0555926
UTC. Controller `reports/runtime/production-close-controller/00da7a0624254706a6b093b590936fa0`;
result `reports/runtime/slice4be-production-close/576a69eeab3c44b79929d8e44fd785e9/red.json`.
`inventory-query-setup.jsonl` proves Seed finishes with the real source open;
the source then closes, reopens, accepts fixture setup and saves successfully.
The last stage is BeforeReferenceOpen. No new query assertion is reached. Five
compiles, normal closure/preservation and zero delayed Excel failures hold.
The exact reason Excel rejects the copied reference is unresolved, not a runtime
defect attribution. Teste5e034a instead captures the four expected Domain results
in memory from the seeded source before closing it. The pristine file copy remains
only a byte baseline/restore source; no second Excel reference workbook is needed.

The continued fixture calibration, valid80-case query RED and isolated Core
candidate are recorded in `plan022_slice4be_inventory_query_results.md`. The
query correction does not implement or accept catalog21/Production Close.

That candidate now passes all80 query cases and the public authority-preservation
assertion, retaining232 Close checks plus80 supplemental checks:175 PASS/137
expected missing-Close failures. Five compiles, static ratchets, smoke86, normal
closure/preservation and four reviewed images pass. The initial-launch write is
resolved in focused evidence. The relevant query regressions now also pass:
chain32/live48/Create15, three-size/five-page layout, full reusable171 observations,
lifecycle615 and draft/paths390, with exact prior checks, normal closure and
preservation. The query record documents the changed-call-site coverage and
the suites retained as frozen evidence rather than rerun. Close test-first work
may resume from this baseline. Exact roots, native-fixture exclusions and counts are in
the linked query record; this does not claim Close or whole-release acceptance.
