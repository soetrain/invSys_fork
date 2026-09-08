# Slice 4be.1 Receiving navigation and selection

Architecture v4.11 D18's Receiving navigation/selection clarification governs
this checkpoint. Architecture, Plan 022 and controls v1.77 were synchronized and
pushed as docs **c4cbde6** before runtime implementation. This inherits approved
comprehensive coverage, optional-navigation defaults, owner-fact semantics and
captured context; it creates no business permission or authority.

The full Goal, comprehensive 4be coverage and human acceptance remain incomplete.
The accepted preceding checkpoint is
[Receiving lifecycle](plan022_slice4be_receiving_lifecycle_results.md), code
**61d401f**, candidate `deploy/validation-receiving-lifecycle-dismissal`, 596/596
plus its recorded complete technical gates. Accepted deployment/NAS is unchanged.

## D13 entry, 2026-09-08

Thirteen reserved controls cover Receiving/Returns/Purchasing tabs, item,
aggregate, history and staged lists on both operational pages, receipt Condition
and disposition kind. Catalog 6 defines Navigation class, fixed product captions,
RECEIVING_NAVIGATION ownership, existing RECEIVE_POST eligibility and default-off
collection. A current whole-policy Collect flag can enable them. The capture
flag alone does not collect outside explicit recording; recording integration
remains separately pending. Older saved policies cannot implicitly enable them.

One actual committed selection produces correlated REQUESTED/Info/Unknown and
SELECTED/Info/Unchanged observations, with empty source references. These facts
describe UI selection/detail only. No selected values, list positions, physical
keys/buttons, staging or inventory completion are inferred. Programmatic changes,
automatic selection, initialization, dependent fills and internal refresh/detail
calls are not user actions. Stale context must reject before new detail-owner
entry; optional tracking failure must remain visible without blocking otherwise
valid selection or retrying the owner.

`Test-Slice4beConfigCommands.ps1 -CheckReceivingNavigationActivity` extends the
preserved packaged harness. Unsaved seams expose controls/state and trace only
fixed event/owner names. Native window messages target the isolated Excel
process's focused control or a verified same-process mouse point. The test
requires both the expected event and resulting selection before asserting
activity; input/fixture failure is a harness exception, never behavioral RED.
These are automated native-handler tests, not physical input or human UAT.

The fixture starts with Admin Generate Warehouse/Seed and actual Receiving Add
handlers. Read-only list projections may receive fixed synthetic choices to
exercise otherwise empty list selection without inventing canonical history.
Staging keys/unknown values and saved operator/unrelated/authority bytes are
independently compared. Activity payloads and fixture inputs stay only in ignored
runtime/disposable fixtures; curated evidence contains no row-level values.

## First meaningful focused RED

The unchanged catalog-5 candidate produces **113 PASS / 69 FAIL**, with no harness
exception, in the separate diagnostic navigation run. All thirteen native
keyboard selections and one native item-list mouse selection reached their
intended controls. Thirteen assertions fail for absent catalog definitions;
56 fail for missing activity pairs and their dependent correlation/caption,
redaction and validated-read assertions. Missing payloads do not establish a
payload leak. Default-off behavior, programmatic/internal exclusions and data
preservation pass. This diagnostic run does not replace the full 596-check gate.

Ignored evidence:
`reports/runtime/slice4be-receiving-activity/navigation-first-focused-red.json`.
Catalog/default-off source implementation began only after
this focused RED; UI event handling was still unchanged at that point.

The first full run completes **632 PASS / 74 FAIL**, all **596** prior GREEN
check identities retained. The additional five failures prove missing tracking
notices for older/malformed policy and failed storage, plus stale-session detail
owner entry and absent visible rejection. This is
`navigation-first-full-red.json`. Expanded native mouse coverage reaches all
thirteen controls and completes **118 PASS / 122 FAIL** in the separate focused
report `navigation-mouse-expanded-red.json`; fixed event-name traces are retained
in `navigation-mouse-expanded-trace.txt`. No harness exception occurs in either.

## First candidate, 2026-09-08

`deploy/validation-receiving-navigation` builds all five packages and passes
explicit VBE compilation/cold-start reference validation. Its separate focused
run is **253/253 GREEN**, including all thirteen mouse and keyboard selections,
same-choice mouse input followed by programmatic change, programmatic/internal
exclusions, stale context, disabled capture, older/malformed policy and failed
optional storage. See `navigation-first-focused-green.json`.

The complete packaged activity/form regression is **767/767 GREEN**, retaining
every prior 596-check identity without duplicate checks. Evidence is
`navigation-first-full-green.json`. Four additional closed-captured-workbook
guards then pass in the separate **257/257** focused run,
`diagnostic-navigation-green.json`: no detail-owner entry, visible reopen notice,
no activity and no redirection to ActiveWorkbook. Candidate source and all five
packages are unchanged between these runs. This is 767 full plus 257 focused,
not a claimed 771-check full run. Release gates are recorded separately below.

The form retains the existing selection/detail owners and adds a shared typed
navigation boundary. Form-local `cReceivingSelectionInput` instances distinguish
input from property assignment and release their event bindings on close. Native
dropdown selection occurs after MouseUp; opening/closing DropButtonClick clears
the pending input token, including cancellation/same-choice closure. Fixed
captions/tab presentation move to `modReceivingNavigation` without geometry or
business changes; the form shrinks from 1,218 to 1,200 lines. Core owns catalog-6
definitions and default-off policy; no observation is a business command.

The full packaged harness command uses `Test-Slice4beConfigCommands.ps1` with
`-RepoRoot . -DeployRoot deploy/validation-receiving-navigation -Phase GREEN
-CaptureEvidence -CheckActivityEvidence -CheckActivityFoundation
-CheckReceivingActivity -CheckReceivingStagingActivity -CheckReceivingLocalActivity
-CheckReceivingLifecycleActivity -CheckReceivingNavigationActivity`. Add
`-ReceivingNavigationOnly` only for the separate focused report. The final four
closed-workbook guards were added after the recorded 767-check full run and are
included in the 257-check focused run against the same candidate.

### Release gates

All gates use `deploy/validation-receiving-navigation`. Automated disposable
fixtures and native-handler captures do not constitute human UAT.

| Gate | Verified result |
|---|---|
| Five-package build; explicit VBE compile; cold-start references | PASS |
| Complete activity/form regression | 767/767; all prior 596 identities retained |
| Latest focused navigation/binding | 257/257 |
| `validate_phase6_packaged_xlams.ps1` | 86/86 |
| `validate_phase6_live_role_workflows.ps1` | 48/48 |
| First ordered Release 1 chain | 29/30; Admin Seed duplicate identity blocker |
| `validate_inventory_viewer.ps1` | PASS |
| Layout and launcher/Production gates | Pending |

The first chain fails SeedDemoInventoryThroughAdmin with DUPLICATE_SYSTEM_KEY
after a successful Admin Generate. All other 29 checks pass, including final
reconciliation. The failure report is retained in ignored
`navigation-chain-first-results.md`. Do not retry to hide this failure or treat
the checkpoint as complete. The rejected identity has the non-hexadecimal format
of Core's existing per-call Randomize/Rnd fallback. A focused packaged identity
probe now tests that generator, supplementing the actual Admin seed boundary.
This is a discovered blocker under existing D14, not a proposed identity migration.

The unchanged packaged generator then produces **35,017 unique / 14,983 duplicate /
zero blank** values in 50,000 normal calls. The separate ambient-RNG reset test
passes 50,000/50,000; it is not claimed as RED. Package bytes stay unchanged.
`reports/runtime/slice4be-system-key/red.json` records two passing checks and one
behavioral failure. No identities or inventory rows are exported by this probe.

After that RED, `modSystemIdentity.NewId` extracts the existing native GUID code
from `modTrainingWire`. Training preserves its lowercase representation and
failure wording; `modRoleEventWriter.CreateEventIdRole` preserves uppercase for
new entity/event IDs and the archive-name collision suffix. Reviewed source
reachability found the archive caller as well as the creation path; both now use
the typed helper. The obsolete Scriptlet/Randomize fallback and its private
normalizer are removed with no remaining callers. Existing saved identities,
ownership and Domain duplicate rejection are unchanged; native generation failure
raises an error instead of manufacturing an identity.

[Microsoft CoCreateGuid](https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateguid)
documents the distributed persistent-identifier guarantee;
[StringFromGUID2](https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-stringfromguid2)
documents buffer sizing and its terminator-inclusive result. This reuses Core's
existing Windows API dependency rather than adding a package boundary.

Replacement candidate `deploy/validation-receiving-navigation-identity` builds and
compiles all five projects with cold-start references intact. The same focused
identity probe is **3/3 GREEN**, with 100,000 total creations and zero duplicates
or blanks across the two cases. Its source/package hashes and count-only reports
are under `reports/runtime/slice4be-system-key`. Seed source checks are 12/12;
packaged WAN/HQ System_Key source checks are 7/7. The replacement candidate still
requires its own complete activity and release gates before checkpoint completion.

The first replacement-chain attempt pauses in the legacy source Create Warehouse
harness: its explicit module list omits modSystemIdentity, producing Variable not
defined in modRoleEventWriter. This is a harness dependency error, not product
RED; packaged five-project compile already passes. Recovery verifies the exact
chain/child process and the source fixture plus its three disposable runtime
workbooks, stops only that chain parent, and resets the identified test VBA
project so the child closes normally. No Excel process is killed and no
operational workbook is changed. All sixteen explicit source-harness import lists
gain the single shared dependency and pass parsing. The full chain restarts from
Admin Generate against unchanged replacement XLAMs; the interrupted attempt is
not counted as a completed gate.

Receiving capture `coverage-navigation-receiving.png` was inspected: the three
tabs, item/history/staging/aggregate lists, selected-reference detail, condition
choice and action/status controls remain visible. Captions/geometry are preserved.
Runtime evidence remains ignored under `reports/runtime/slice4be-receiving-activity`.

| Exact candidate package | SHA-256 |
|---|---|
| invSys.Admin.xlam | 5027921c0fffbd39a28116b9344c2b96d6c4f58b5b91032c8f563d53b23cdbfb |
| invSys.Core.xlam | 38aa14cc8a90dc432f896984898602d7e3864b20349e46e35050cda9a1cd6802 |
| invSys.Designs.Domain.xlam | 04f300f953b3a707e157e55fbe24b7620cd729337cb9334810435fa60c8efcfd |
| invSys.Inventory.Domain.xlam | ce8ca0c850781a979eec85eddd363cc30d9590a517c168327199c24964455260 |
| invSys.Operations.xlam | e966f95f87637f0396ad580c5edda5118e4b5afbc8a46ebcfcde17363986e754 |

### Reviewed maintenance exception

The new control coverage adds seventeen native event roots and three normalized
duplicate-body groups: five one-line input-reset callbacks, two one-line mouse
input callbacks, and six one-line form dispatch callbacks with distinct control
names. Keep these distinct MSForms entry points and their single typed
controller. This is an explicit exception for these three small adapter groups,
not for duplicated business logic, new dynamic calls or growth of existing
oversized modules. Raw scanner counts remain visible: 192 -> 195 duplicate groups
and 1,077 -> 1,097 total candidates, including required retained event roots.

The class-event registry omitted DropButtonClick and initially marked its proven
private callback REMOVE. Classification was verified false, then corrected to
include that existing MSForms event. Native dropdown GREEN proves its role in
input cleanup; no callback is deleted. Tooling tests remain 62/62. Initial navigation
evidence reports 174 components, 5,501 procedures, 1,097 scanner candidates and
1,099 reviewed candidates. REMOVE remains 101; RETAIN increases 435 -> 452.
All 28 existing size limits hold; dynamic calls remain 45 unresolved/eight literal
targets. The Receiving form is 1,200 lines; modTS_Received remains 1,575.
The identity fix's final regeneration is 175 components/5,500 procedures, with
all other counts unchanged. modRoleEventWriter shrinks from 3,104 to 3,064 lines.

## Harness corrections, not product RED

The first expanded run stopped at **79 PASS / 1 harness exception** before
Receiving staging could execute. An injected parameter used VBA's reserved word
Tab; the compiler reported Expected: identifier. Corrected to pageIndex. The
TabStrip trace declarations also require ByVal Index to match the actual event
interface. The saved first report is `navigation-first-harness.json` beside the
focused evidence. Neither failure proves a product defect.

The isolated Excel instance remained in compiler dialogs after ordinary cleanup.
Before terminating that exact test process, zero workbooks and only the five
candidate-path VBA projects were verified, along with its captured start time.
No operational workbook was saved or modified; all five candidate hashes still
matched. No unidentified Excel process was terminated.

Microsoft's [Forms Click documentation](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/click-event)
describes value-selection events, while its
[key-event documentation](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/keydown-keyup-events)
describes input dispatch and ordering. Native fixture traces, not an assumption
that Click proves user origin, govern the protecting tests. Programmatic
Value/ListIndex changes are explicitly tested separately.

## Replacement candidate verification

All final gates use `deploy/validation-receiving-navigation-identity`.
The preceding candidate and failed/interrupted evidence remain preserved.

| Gate | Verified result |
|---|---|
| Five-package build, explicit VBE compilation and cold-start dependencies | PASS |
| Packaged identity burst and ambient-RNG tests | 3/3; each 50,000-key case has zero duplicates/blanks; package unchanged |
| Full packaged activity/form/Ribbon suite | 771/771; every prior 596/767 check identity retained, no duplicate check names |
| Packaged smoke | 86/86 |
| Live-role workflows | 48/48 |
| Ordered Release 1 chain and restart reconciliation | 30/30, including the previously failing Admin Seed |
| Create Warehouse source integration | 15/15 |
| Viewer | PASS; events, filters, refresh, export and read-only authority preserved |
| Production layout and native transitions | Three requested sizes/five pages PASS; minimum clamps to the existing default; captures inspected |
| Public NoEligible launchers | 3/3 |
| Full ProductionReusable and clean restart | 2/2; workbench/edit/export/import, lifecycle/run, Chai fork/convergence and saved-workbook restart |
| Relevant source checks | Receiving 4o 5/5, 4p 8/8, 4q 6/6, 4au 7/7, 4v 4/4, Slice 5 13/13, Slice 10 10/10; launcher 24/24, restart 6/6; tooling 62/62; Seed 12/12 and WAN/HQ identity 7/7 |
| Static maintenance | 175 components, 5,500 procedures, 1,097 candidates, 195 duplicate groups, 45 unresolved calls/eight literal targets; all 28 existing caps hold; adapter exception above |

Final count-only results, compile/source and package hashes are in ignored
`reports/runtime/slice4be-system-key`. The complete activity command is the one
above with the replacement DeployRoot; no diagnostic-only switch is used.
Launchers use `validate_plan022_packaged_launchers.ps1 -WorkbookState NoEligible`;
Production uses `-CallbackFilter Production -WorkbookState ProductionReusable`
without either reduced-workflow flag. Other gate script names are recorded in
the earlier gate table and the retained logs.

The serial gate wrapper stopped before Production because Excel was still exiting
after successful launcher cleanup. Excel subsequently exited normally; Production
was launched separately only after the closed-process guard passed. No concurrent
Excel validator or extra launcher retry was started.

The minimum/default layout images initially appeared blank when displayed together.
Pixel/hash inspection and individual viewing confirm valid identical saved images:
the requested minimum clamps to the default geometry. One complete layout
recapture also passes and matches; no runtime or screenshot-harness change was
made. Both `layout` and `layout-recapture` are retained. Receiving's replacement
capture and all three Production captures were inspected; this is agent evidence,
not human UAT.

| Exact replacement package | SHA-256 |
|---|---|
| invSys.Admin.xlam | ffead950428b02100ac7c068c692e4fe87926ce0a2b0e814ab7fd1b69e0188ca |
| invSys.Core.xlam | 9568f12d8aac3f02e2711cfd877bc05c272aee56f05a468126dbf5a4a73fc26c |
| invSys.Designs.Domain.xlam | b4d53fea523223317b22e0571a15b6f018d22862994f28724f707d89b54c41db |
| invSys.Inventory.Domain.xlam | 1f13a51ee34c1cdd621fd22f44cf8fc5252be02940fcc65e25a0705a3a534aca |
| invSys.Operations.xlam | 25764ace24245a666d233f28ef2982e398e8c351cda4ac58d41650966785aed7 |

Comprehensive 4be.1 coverage remains incomplete: worksheet Confirm reachability,
launcher denial observation, other Operations/Admin controls and shared
publication still require D13 evidence. Event Tracking Settings/profiles/personal
preferences, comprehensive Viewer, recording/conclusions, guide management,
How-To/Diagnostic/Compare and physical multi-station/NAS/user comparison are also
pending. No accepted deployment or operational workbook is replaced by this work.

All technical gates for this navigation/identity checkpoint are GREEN. Final
replacement, first navigation and accepted catalog-5 package hashes still match;
accepted catalog-4 hashes were also reverified unchanged. Excel is closed, and
the final Production interval has zero Excel Application Error events. Git review
preserves unrelated handoff 067 (3 additions/3 deletions) and untracked critique
023; runtime evidence remains ignored. This completes this technical checkpoint,
not comprehensive 4be.1 or Release 1 user acceptance.
