# Slice 4be Production Complete Run

## Current full gate, 2026-10-05 UTC

`validation-complete-submission-01` passes **289/289**, controller
`30c5740aa1564e69b08508a011b6be3f`, worker
`slice4be-production-complete-baseline/abd75a40c9874c46b988cf0255f018fc/green.json`,
08:24:05.699--08:37:25.597 UTC. Run `Test-Slice4beProductionRunLocal.ps1` with
`-DeployRoot deploy/validation-complete-submission-01 -Phase GREEN
-CompleteBaselineOnly -CompleteVisibleHostForTest` and no diagnostic switch.
All258 prior checks remain in order;27 post-consume/audit and four host-visibility
checks pass. Five instrumented compiles,42 shared checks, exact package pins,
settings restoration, unassisted closure and the delayed zero-error Excel audit
pass. Peak GDI is869. The visibility transition is slow but recovers without
intervention. No permissions, runtime package hashes or quotas are changed.

This verifies the existing post-consume continuation correction from its four
behavioral REDs. Retain same-candidate155/155 for the surviving-form entry click;
the full visible-host run covers native dismissal instead. Both are required
evidence, as detailed below. Hidden-host failures remain documented; this does
not prove a native Excel repair or full Slice4be acceptance. Affected guide58 and
packaged smoke86 pass below; broader gates, worksheet/later submissions and
Complete Run recording remain open.

Affected published-guide regression passes58/58 with exact prior ordered checks:
controller `execution-profile-guide-regression/47d6b481471f4f01bee4087e39f00cf2`,
worker `slice4be-viewer-published-read/aedfa3ec13464da2afd4ec9e736cbbf6`,
08:37:53.825--08:44:51.057 UTC. Five instrumented compiles, normal closure,
settings/package preservation and delayed zero Excel errors pass.
Packaged smoke passes86/86 with exact prior order, both unassisted Excel exits,
canonical pins, restored settings/tracked reports and delayed zero Excel errors:
`complete-submission01-regression/smoke-af8a2382957b45359ac3c94bdd7dfca0`,
08:45:27.368--08:45:51.112 UTC. All408 repository PowerShell scripts parse.
Runtime build/static evidence remains `complete-submission-build-01/` and
`complete-submission-static-01/`; these verification-tool changes do not alter VBA
sources or package hashes. No candidate is promoted.

## Original loading mode comparison, 2026-10-05 UTC

The corrected full gate also exhausts GDI with `OriginalReadOnly` package loading:
controller `ba2e35e04142479eb4c3655e0e3b4dfa`, worker
`slice4be-production-complete-baseline/48262189d8f34ba785bed64e5c4e980b/green.json`,
07:36:11.596--07:49:39.334 UTC, **255 PASS/1 harness FAIL**. All27 new
post-consume/audit checks and228 prior checks pass;30 prior checks are unreached.
The first closure case already grows to peak3977; the third reaches10001 with
bounded native windows. Read-only loading does not resolve the failure.
Excel exits naturally; the completed worker requires termination, after which
the parent restores settings and verifies package hashes. No memory dialog is
dismissed. This assisted result is not acceptance. Next narrow the existing
Check In handler using fixed-label native counters in disposable probes only.

The bounded prior-sequence run with Check In phase counters closes normally:
controller `abc6224316f14b54b55de82d3c9045c3`, worker
`slice4be-production-complete-preparation/639124ec2f264f40a5ff7bb45a9d0a2d`,
07:49:51.228--07:59:32.677 UTC,202 PASS/1 diagnostic-bound failure, peak3777.
Earlier Check In phases remain below750 GDI; the closure-case Check In grows
934->3754. List fills remain flat; repeated continuation/read phases commonly
add130, and action begin/finish also grow. This localizes the next investigation
to authorization/configuration reads, without proving their native cause or
permitting cached authorization as a workaround. Settings/packages are preserved.
Next trace Core's Config/Auth resolve/hide/read/close boundaries with fixed labels
and native counts only. No runtime or acceptance change.
The first Core-trace setup, controller `b1cc463d73804650b4ec02c2ecb5be0f`, stops1/1:
the helper was incorrectly sought in Core although it belongs to Admin. It closes
normally and preserves settings/packages. A dedicated disposable Core module fixes
the setup; this is not product RED.

Core boundary trace: controller `0cfe67ada8fc4f43b929b91b03682324`, worker
`slice4be-production-complete-preparation/b5e1b127b4d449ddb58649eb6c8da53f`,
08:01:09.062--08:10:53.861 UTC,202 PASS/1 diagnostic stop, peak3787 at preparation.
All202 prior diagnostic passes remain in order; closure/settings/packages pass
without intervention. In the failing Check In, each Config and Auth read retains
about65 GDI after closing its transient workbook, accounting for130 per capability
check. Earlier reads do not show that sustained retention. The native cause is
unproved. Next compare the same sequence with the isolated Excel host explicitly
visible before each closure case; retain all handlers, assertions and live rights.
The isolated option records its mode/visible checks. Subsequent
native markers include process ID to distinguish any separate Excel instances.

The visible-host diagnostic passes262/262: controller
`89f0537613234a8188c57b62f9ceb6bc`, worker
`slice4be-production-complete-preparation/6043eebb7f0e4926b64961e49a797952`,
08:11:09.501--08:22:39.197 UTC. All258 prior checks remain in order, with four
explicit host-visibility checks. Peak GDI is873; all four closure cases and normal
unassisted cleanup pass, preserving settings/packages. Making the isolated Excel
host visible avoids the resource growth in this sequence; native cause is not
proved. The visibility transition temporarily stops responding, then recovers
without intervention. A guarded cleanup attempt rejects the now-advanced state;
no termination or dialog dismissal occurs. Its scoped blank-window capture is
reviewed; no operational content is present. All native markers share one process ID.

At closure entry this run dismisses the form, so no click is delivered there.
Do not call that surviving-form coverage: the same five package hashes separately
pass155/155 in controller `83fb61c085f54a0e9a7b5e72768d0a78`, worker
`35ba899c1fc04ac1bee165ef697759f1`, with a surviving form and exactly one actual
refusal click. Retain both native outcomes. The full completion gate is next with
`-CompleteVisibleHostForTest`, no preparation diagnostic switch: expected289 checks
(prior258 +27 post-consume/audit +4 host checks). No runtime package is changed or
promoted by this fixture comparison; broader acceptance remains open.

## Audit inspection correction, 2026-10-05 UTC

The new audit inspector unconditionally reopened/closed an authority workbook
already owned by Core. Focused tooling RED is52/2: the workbook-lifetime assertion
fails and the harness deliberately stops; byte preservation passes. This is a
test defect, not another runtime D13 RED. The inspector now borrows an existing
workbook and closes only its own transient read-only open. Both interruption
cases preserve workbook lifetime and bytes; all122 previous diagnostic checks
remain in order in126/0 GREEN. Runtime packages are unchanged.

| Diagnostic / controller | Worker under `slice4be-production-complete-preparation/` | Result / GDI peak |
|---|---|---|
| Cold closure: `a33fade33b90436e9a3def0220fc44be` | `f902b0cdd4e0445588a6e5dae214afc0` | 99/0;842 |
| After submission: `ca67a1e680b14e4d83489d191906c08f` | `4d22256e8773401286211168e1fdd885` | 122/0;901 |
| Inspection RED: `84412b0b0820450fb9a5269c59b11c4e` | `9b56975b48a6468ca46423d07945956c` | 52/2;stopped before closure |
| Corrected inspection: `6d871876499f48a68698e9d4318939b1` | `e35c0c26b5dd464690078cb741398c99` | 126/0;935 |
| Entry guards: `83fb61c085f54a0e9a7b5e72768d0a78` | `35ba899c1fc04ac1bee165ef697759f1` | 155/0;852 |
| Read/yield interruptions: `be84c558c2bd4d77b037c9e064279936` | `ebff4992600b4ecabefd917a4cdc338d` | 160/0;941 |
| Initial completion/context: `0e10b6ef1cdf4c84b1820b37e557b6f9` | `3565b9789af54753965fa80648697f15` | 143/0;1036 |

All seven controllers preserve settings/packages and close without intervention.
The diagnostic-only route skips unrelated completion scenarios and cannot accept
the full gate. Its initial inner markers were incorrectly ordered; that trace is
excluded from phase attribution. Statement-bracket replacement/readback corrects
later traces. Outer native counts remain valid. Cold, submission, entry and
read/yield-interruption routes do not reproduce the full-workload exhaustion;
its cause is not yet proved. The corrected full gate still fails: controller
`bfe9cca945f44dfeab13f638bd7d6d61`, worker
`slice4be-production-complete-baseline/8e37abb77edc41f28c2fd9bd275b77c6/green.json`,
06:04:20.394--06:17:34.399 UTC, 257 PASS/1 harness FAIL. All27 post-consume/audit
checks pass; 30 prior checks remain unreached. GDI rises from723 before the first
closure case to peak10001. Final target selection reports status7. Excel exits
naturally; the completed worker requires termination before settings/package
restoration. No memory-dialog click is posted. This assisted closure is not
acceptance, and the inspector fix does not resolve native resource exhaustion.
One unchanged template is copied/hash-verified from user-policy02 into candidate01
while Excel is closed; receipt: `complete-submission-build-01/template-copy.json`.

The initial completion/multi-Process/context group also passes in isolation,
retaining nine/ten native Excel windows during closure. Its read-only trace adds
native XLMAIN and saved/visible workbook counts without logging names or
contents. Diagnostic-only resource bounds stop at3000 GDI objects or80 XLMAIN
windows and retain failure/cleanup; those bounds do not accept any product gate.
Prior Excel-only saved-workbook controls are in
`plan022_slice4be_gui_resource_results.md`; similarity is a hypothesis to test,
not proof of this failure's cause. The full comparison now offers an explicit
saved/reopened decoy fixture with its byte-preservation check; every existing
handler/assertion remains, and the unsaved-decoy failures are retained.
All408 PowerShell scripts parse.

Saved-decoy full comparison: controller `98320a80e75a4282b8e841178d412b90`, worker
`slice4be-production-complete-baseline/c8d531d3a07749b48dd50afd2614c4f4/green.json`,
06:35:09.891--06:48:22.033 UTC:259 PASS/1 harness FAIL, GDI peak10001. The two
saved-decoy checks pass, but the same30 prior checks remain unreached. XLMAIN
stays between nine and eleven during closure; at least one saved visible workbook remains.
This control does not prevent growth and does not match the older retained-window
pattern. Excel exits naturally; the completed worker needs termination. No memory
dialog click or later controller-collection command runs. Settings/packages restore;
exit-1 and assisted cleanup do not accept the gate.

Bulk Value2 reads also fail to prevent growth. Controller
`fc7be40567ba4735848fee291c9aaafb`, worker `55f084bddf944c7d846eea88557c4071`,
uses the original unsaved decoy. Its partial log records257 PASS/1 harness FAIL;
GDI peaks at10001 with bounded XLMAIN counts. The exact audit assertions pass.
Excel exits naturally, but the worker stalls before final JSON and is terminated
at07:00:53.595 UTC. The controller retains the source/pins and explicitly incomplete
partial-log receipt; no synthetic green.json is created. This speculative reader
change is reverted. The parent closes at07:02:14.263 UTC with exit-1 and verified
settings/package restoration. No runtime, workbook authority, quota or contract
changes. The07:02:29.916 UTC desktop probe succeeds; no error5 is observed.

The bounded prior combined sequence (all earlier completion/context/interruption/
entry cases, without post-consume cases) also fails on candidate01: controller
`627437f5823a489dad4be2a2c70d8dba`, worker
`slice4be-production-complete-preparation/e6f18c5dd96b474a99ddadcc91a85bbc`,
07:04:15.634--07:13:53.915 UTC:204 PASS/1 diagnostic-bound failure. The first
closure case remains at940 GDI through staging, then its actual Check In handler
raises usage to3760 (observed peak3783). Cleanup closes normally and restores
settings/packages. This is not product RED or acceptance; later closure checks
are unreached. Added post-consume cases are not necessary to reproduce growth.

The prior258 GREEN is specifically `validation-production-complete-entry-02`,
controller `9f7cf60d4d554fa8a4239d77737b42fe`, 2026-10-02 03:10:09--03:24:41 UTC;
it is not a completed gate on user-policy02. The identical bounded prior sequence
on frozen `validation-user-policy-02` also fails204/1, with the same ordered205
checks/results and Check In GDI940->3760 (observed peak3783). Controller
`e4049e5bd5764623ac53eb896a5e9172`, worker
`slice4be-production-complete-preparation/82349d4d043d4027bd77a7f8fa541ea2`,
07:15:11.427--07:25:09.619 UTC, closes normally with settings/packages preserved.
This failure predates the continuation guard. Next compare complete-entry02 using
the same bounded harness; tracing remains matched. No acceptance is inferred.
Historical test source at `50241047` and its controller receipt establish that
the original258 used `OriginalReadOnly` package loading. Current bounded comparisons
use saved writable probe copies. Preserve this distinction: a saved-copy failure
alone does not establish regression of the original loading mode.
The older complete-entry02 package also fails204/1 with the saved-copy bounded
harness: controller `28ac0ea972874e04bf32215f9f7080ed`, worker
`slice4be-production-complete-preparation/c257f101f71240a2a57d062880d1c268`.
Its first closure Check In increases GDI941->3761 (observed peak3784). It runs
07:26:12.843--07:35:46.026 UTC and closes without intervention, restoring settings
and package hashes. Next run the full current candidate in the original read-only loading mode, including
post-consume checks. The saved-copy failure remains recorded, not waived.

## Post-consume continuation, 2026-10-05 UTC — under validation

Existing D18 captured-context/permission continuation applies after consumption
as well as before it. An applied consume remains applied; interruption must stop
the later completion submission. This changes no authority, permission, rollback
or observation contract. Complete Run recording remains a separate open A1 task.

The actual packaged handler on frozen `validation-user-policy-02` records four
behavioral REDs: after real consume processing, sign-out and permission loss each
allow another queue attempt and miss the specific interruption feedback. Exact
input consumption, absent completed output, partial owner/projection state and
custom/unrelated workbook preservation pass. Core still refuses the later write;
this evidence does not show unauthorized inventory creation.

Controller `production-run-local-controller/39262e7eb4c9499eb33841f156beee60` runs
04:31:19.341--04:42:30.601 UTC; worker
`slice4be-production-complete-baseline/5ff823b4004348dbae03595f6a1fda52/red.json`
records245 PASS/7 FAIL. Besides the four behavioral failures, two new audit checks
incorrectly assumed quantity1; the corrected test reads the exact staged entity/
quantity before invoking the handler. A later existing closure case cannot select
its fixture target, leaving prior checks unreached; its cause remains unproved.
The shared harness now reports only the sanitized numeric/ERROR selection code.
Excel exits without intervention after a delay; settings/packages are preserved.
No desktop error5 is observed. An attempted shutdown inspection fails before
producing usable resource counts and proves no resource diagnosis.

Two earlier setup attempts (`5715f26ad28f4cda8943b8e178e11b90` and
`734baf15d430481fa7d0c1e045dc0a00`) each stop1/1 before behavior. The second proves
one case-insensitive match for the stock fixture's unchanged statement. Its probe
now tolerates VBA identifier recasing while retaining the exact numeric literal.
These are harness failures, not product RED.

Candidate `validation-complete-submission-01` adds the existing continuation guard
after recording the consume processor return and before completion submission.
The unchanged event-list append helper moves to the typed action module: owner
size2707->2703 lines. Cold Operations startup/five compiles pass, settings and the
previous candidate are preserved, and source comparison retains300/302 components
with only the two intended Operations modules changed. Build receipts:
`complete-submission-build-01/`. Static `complete-submission-static-01/` retains
all28 oversized caps,9 literal/45 unresolved calls and190 duplicate groups:
309 components,6296 procedures,137769 lines (+2), three schemas/405 script parses.

First candidate run: controller `9d2e9ff6a2a94c62ae5c2be7c229a17b`, worker
`8a9b761c951041a49fde5d1539c8dbe1/green.json`,04:45:06.235--04:56:42.594 UTC,
251 PASS/1 harness FAIL. All23 new post-consume checks pass;228 of258 prior checks
are reached. Target selection returns status7 (Auth unavailable) before the other30.
Read-only native counts show GDI peak10001 against quota10000. This establishes
resource exhaustion, not its cause. Excel exits unassisted; preservation passes.
Both new refusal captures were reviewed: existing reopen/context or permission
guidance is visible. Long field clipping remains outside this scoped proof.

A saved-disposable-package comparison, controller
`8eb2b298f3e140c1844116ea33ee5db0`, worker `129110dbb4534cf1a6a023acc460a1db`,
reaches282 PASS/1 metadata-read FAIL across283 checks. It also reaches GDI peak10001
and displays Excel's insufficient-memory dialog. The dialog is inspected and
dismissed at05:11:39 UTC; residual test Excel/worker cleanup is assisted. Thus
saved probes do not fix resource exhaustion and this is not clean GREEN acceptance.
Interventions and resource samples remain in that controller directory. Exact
closure-stage resource tracing is added to locate the trigger before another fix.
The controller closes at05:15:57.264 UTC with settings/packages restored. Its
exit code is-1 after assistance, not a successful gate. All407 repository PowerShell
scripts subsequently parse without errors; the runtime static ratchets are unchanged.

Independent metadata tooling reproducer `package-identity-sharing/`:
`6cf2923498474787bc282cb578f3e238/red.json` is3/1; default ZIP sharing conflicts
with a writable-open package. Read-only FileStream with ReadWrite sharing retains
the same metadata validation and passes4/4 in `0e1aad5293344795afaaa672e36f1705`.
Source/copy bytes remain exact. An intermediate2/2 run (`eb6d68f7a21045be893cef2233a8890a`)
exposed a missing Compression assembly load, corrected before that GREEN.
Full unassisted regression, broader gates and Complete Run observations remain open.

Closure trace comparison: controller `f6383263c4044de4b74ba81ad04f2e84`, worker
`57af4823ac8b408495308235bb68369a`, reaches283/0 but requires dismissal of the
same memory dialog at05:26:12 UTC. All258 prior checks remain in order. At the
CompletePending case, GDI grows4685->7845 during preparation and stays8019 across
SafeClose. The next preparation starts8414 and reaches10001. Thus cleanup alone
is not the measured growth interval. Read-only samples and fixed-stage JSONL
retain the diagnosis; no native root-cause repair or clean acceptance is claimed.
Controller `10595d3f340c41928ee5d32d74b55e62`, worker
`48f98426d81c4be9a66cfd30d876d890`, tests explicit client garbage collection
between cases. It leaves GDI4286 and8015 unchanged, reaches10001 again and requires
dialog dismissal at05:40:30 UTC. Final assertions are253/1 with target-selection
status7;23 new checks pass,228 prior checks are reached and30 remain unreached.
The ineffective collection change is removed. Do not repeat these full runs or
increase quotas as a proposed fix: isolate real case preparation (fixture staging,
Check In and balance reads) with native counts and retain the ordinary handlers.
Clean GREEN, shutdown, affected guide/package regressions and broader acceptance
remain open. No new architecture or permission approval is required for this
existing-contract correction; future conflicting behavior still needs approval.
Last verified05:45 UTC: all three diagnostic controllers finish with original
settings/package hashes restored and Excel closed; the two later controllers end
05:30:57.181 and05:45:08.111 UTC with exit-1 after assisted cleanup. The last also
uses managed collection in its owned controller to aid shutdown; this is not
runtime acceptance. No desktop error5 was observed. No candidate is promoted.

## D18 completion entry guard correction

`modProductionCompleteActions.Execute` now wraps the actual form handler with
typed calls. It suppresses loading/busy entry before owner work, holds busy through
the original completion, restores prior flags and rethrows captured errors. The
form's completion method becomes Public for this same-project typed call; its
algorithm and form line count are unchanged. No observation/catalog or inventory
contract is added.

New unpromoted `validation-production-complete-entry-02` was built under
`complete-entry-build-02`,2026-10-02 01:38:54.9399549-01:39:34.6645530 UTC. Cold
Operations startup/five compiles pass. Of283 old components,282 remain identical
and only Operations/frmProduction changes; the new guard module makes284 total.
Settings/frozen continuation01 are preserved, with normal cleanup/delayed zero audit.

The first attempt, `complete-entry-build-01`,01:37:13.3416033-01:37:28.1681218 UTC,
failed before Operations import when Excel crashed in VBE7.DLL (c0000005,
offset275b57; WER reported OFFICE_MODULE_VERSION_MISMATCH). Its two native event
records are failure evidence, not product RED. Excel closed and settings/prior
candidate were preserved. The incomplete entry01 candidate remains untouched.
The fresh entry02 build passed; this does not establish a root-cause repair for
that native crash. Desktop error5 was not observed.

Focused GREEN is185/185, retaining every RED identity and all169 prior passes:
controller `production-run-local-controller/44800f286a65445c9019d51478a09419`, result
`slice4be-production-complete-baseline/b5a6d01fc24c4c28bc9c11577e47f61c/green.json`,
01:40:03.2229820-01:51:00.2420470 UTC. Loading/busy entry preserves flags,
staging/projection/message and exact inputs without owner entry. Nested entry
preserves the active action's state/message, while the outer action enters its
owner once, consumes exact inputs once, creates the expected fresh output and
retains success. Shared42/five instrumented compiles, package/settings preservation,
unassisted closure and delayed zero native audit pass. All three new captures were
individually reviewed; other regenerated captures were not additionally reviewed.
Long field clipping and populated-list/multiline acceptance remain unresolved.

`complete-entry-green-static-01` records291 components/6177 procedures/135555 lines:
exactly one small module/procedure and25 lines above the RED baseline. All28
oversized caps,9 literal/45 unresolved calls and190 duplicate candidates are
unchanged. Three schemas/379 parses pass. This accounts for the new entry helper;
there is no oversized-cap or dynamic-call exception.

The subsequent error/recovery expansion and new-candidate Check In regression
are recorded below. Same-candidate smoke/layout pass below; live-role/full-chain
also pass after the focused stage-exit tooling correction below. Earlier native
failures remain preserved; their cause is not established by the later success.
Reusable closed-workbook coverage passes below. Worksheet interruption, partial-submission facts and
observations remain open; the tested cases do not prove every exit or full
Slice4be acceptance.

## Entry error restoration and explicit recovery

On unchanged entry02, a test-only expansion raises a fixed error at the real
reusable pending UI yield through the actual Complete Run handler. The caller
checks all five original error fields and observes restored loading/busy flags
before fixture cleanup. Owner state, displayed projection and message are compared
at that boundary; prior staging is not claimed to roll back. No completion owner
is entered and exact input balances remain unchanged. A separate actual-handler
click then completes once, with fresh output and success, without reopening the
form or resetting production flags before delivery. This adds 21 checks without
changing runtime or claiming a new product RED.

GREEN is206/206, preserving all185 prior ordered PASS results: controller
`production-run-local-controller/8edff1e8e31345cbab31712a58a986a1`, result
`slice4be-production-complete-baseline/31cd319dcee44bb2abca813ebecfe8f0/green.json`,
2026-10-02 02:09:33.7330157-02:21:27.6784245 UTC. Shared42/five compiles, package/
settings preservation, normal unassisted closure and delayed zero native audit
pass. Both new captures were individually reviewed: the injected fault retains
READY rows and the pending saving message; the subsequent click shows COMPLETE
rows and success. The test adapter catches the original error; this does not
prove new operator-facing error presentation. Other regenerated captures were
not additionally reviewed. Long field clipping remains unresolved.

`complete-entry-fault-static-01` preserves291 components/6177 procedures/135555
lines,9 literal/45 unresolved calls,190 duplicate candidates and28 unchanged caps;
three schemas and379 PowerShell parses pass. There is no runtime growth.

Check In also passes629/629 on entry02, preserving all629 prior ordered PASS
results: controller `production-run-local-controller/f2e362fb8a8b47ffba0b1f4b1034a5d0`,
result `slice4be-production-check-in-activity/790cdede229940e299bb9c62e35e08fb/green.json`,
01:52:03.3196997-02:09:07.4274392 UTC. Shared42/five compiles, canonical pins,
settings restoration, normal unassisted closure and delayed zero native audit
pass. Its captures were regenerated but not additionally reviewed. Controllers
were retained through slow responsive Excel cleanup; desktop error5 was not
observed. Same-candidate smoke/layout pass below; the chain subsequently failed
with a native Excel crash. Full acceptance remains pending.

## Entry02 packaged regressions

Packaged smoke passes86/86 with exact prior ordered results, both unassisted
Excel exits, canonical package pins, restored settings/tracked reports and a
delayed zero native audit: `complete-entry02-regression/smoke-49febf96c67144ccb8d3b87ff5eed634`,
2026-10-02 02:21:50.5854110-02:22:13.1182464 UTC.

Layout retains all18 requested size/page pairs across six activated/maximized
pages, five native transitions and two actual sizes, with unchanged representative
geometry: `complete-entry02-regression/layout-a3d52668f9bc44ea8c2f1a702a9c0383`,
02:23:02.3954611-02:23:19.2601835 UTC. All six PNGs are byte-identical to the
previously reviewed continuation01 captures; `visual-comparison.json` records
the hash linkage. Preservation, natural shutdown and delayed zero audit pass.
These empty Run List/Settings images do not establish populated-list or multiline
acceptance.

The chain failed, not GREEN: `complete-entry02-regression/chain-220ba60d60da4ccba3da2e56b7cbdad1`,
02:23:41.9299813-02:32:14.8377131 UTC. Warehouse creation passes15/15; live-role
records32 PASS/1 Harness.Exception and the chain records5 PASS/1 Harness.Exception.
Production's form Check In and completion checks pass before the failure at
`Delete and rebuild canonical inventory projections`,
`modProcessor.RunBatchReportForAutomation`, HRESULT `0x800706BE`.

Application1000 records an Excel crash at02:25:24.3295609 UTC in `ntdll.dll`,
exception `c0000028`, offset `0000000000012d2f`; Application1001 at02:25:31.6699393
reports `OFFICE_MODULE_VERSION_MISMATCH`. These are two records for native failure,
not product RED or desktop error5. Excel restarted and reopened ten test-fixture
workbooks. After verifying their process identity and fixture-only paths, they
were closed without save; their saved bytes remained identical. The parent
controllers were retained and restored settings and tracked reports. All five
package pins are preserved and Excel is closed. The delayed native audit confirms
two events; `failure-verification.json` and `assisted-recovery-closure.json`
explicitly reject unassisted/full-chain acceptance. The source of the native crash
is unresolved; no blind retry or claim that the completion helper caused it is
made. Investigate this boundary before a new chain attempt. Prior continuation01
chain results do not establish entry02 acceptance.

## Focused investigation of the projection crash

The compiled-source manifests for continuation01 and entry02 retain identical
Core and Inventory Domain components. Only Operations/frmProduction changes and
modProductionCompleteActions is added (282 old components unchanged). This rules
out a source delta in those owning components; it does not prove the native cause.
The established investigation in `plan022_slice4be_shutdown_header_results.md`
documents the same projection boundary/signature and the focused control to use
before any further full-chain attempt.

The existing `Test-Slice4beProjectionLiveControl.ps1 -Cut AfterProjection
-TraceBoundaries` runs the actual preceding ordered role sequence and projection
rebuild against entry02, then cuts later chain stages. No runtime, test or package
file is changed. Its unsaved instrumentation uses fixed stage identifiers only.
Root `projection-live-control/1bf808489592489286c1667f6609c580` passes35/35,
retaining the exact ordered first35 PASS identities from the prior complete
live-role result, including all four projection assertions. Four traced packages
compile; 37 markers are installed and33 distinct markers are observed. The real
processor reaches its return and the diagnostic cut. Start02:35:52.0989368 UTC,
delayed audit02:37:22.0865834 UTC on2026-10-02. Normal unassisted closure, settings/
tracked-report/package preservation and zero Application failure events pass.
`source-comparison.json` and `scope-verification.json` preserve the comparison and
ordered checks. This is focused diagnostic evidence, not product RED/GREEN,
native-cause repair or complete chain acceptance. It supports one subsequent
uninstrumented full-chain verification on the unchanged candidate.

That second chain also fails: `complete-entry02-regression/chain-8d85e98b462346b48eeae55fbbc7d659`,
02:37:54.1620794-02:40:19.2625079 UTC. Counts remain chain5 PASS/1 harness failure,
live-role32 PASS/1 harness failure, warehouse15 PASS. The failed call and HRESULT
match the first attempt. Application1000 at02:39:34.1293030 UTC again records
`ntdll.dll/c0000028/0000000000012d2f`; Application1001 at02:39:40.8139857 reports
the Office mismatch. Ten verified recovered fixture workbooks are closed without
save, preserving their saved bytes; both controllers restore settings/reports and
all package pins. The delayed audit retains two native events. No third unchanged
broad chain is justified by these results.

The untraced AfterProjection control also passes35/35:
`projection-live-control/c0dc992b439d4475ae69bfb9510d4681`, start02:40:49.4895977,
audit02:42:16.4575751 UTC. It retains the exact prior ordered checks, returns from
the real processor, closes normally and preserves settings/report/packages with
zero native events. Tracing/recompilation is therefore not required for the
observed focused success; this does not establish any crash cause.

The diagnostic harness gains a `-Cut Full` option that keeps the entire generated
ordered live-role body and inserts only its completion marker before the existing
catch/finally. Existing cuts retain their behavior. No runtime/XLAM or architectural
contract changes; this is diagnostic expansion rather than product RED/GREEN.
Untraced Full passes48/48, retaining every prior ordered PASS identity:
`projection-live-control/aeacb58738ef46898f529fd7729fdf77`, start02:43:06.8507818,
audit02:45:15.4534042 UTC. All later role actions run; normal unassisted closure,
settings/tracked-report/package preservation and zero delayed native events pass.
The scope receipts explicitly keep full-chain acceptance false.

The remaining comparison is the enclosing chain's preceding packaged Admin entry
and source Create Warehouse setup versus these standalone ordered-live controls.
That context is a hypothesis to test with the actual helpers and process/lifecycle
observations, not an established root cause. Other diagnostic differences remain:
ephemeral fixture credential generation, report redaction, lifecycle markers and
child launch/output capture. Standalone48 alone does not isolate Admin setup.
Do not change product authority,
delete assertions or infer full-chain success from the standalone48 result.

`projection-context-static-01` retains291 components/6177 procedures/135555 lines,
9 literal/45 unresolved calls,190 duplicate candidates and28 unchanged oversized
caps. Three schemas and all379 PowerShell parses pass. The only diagnostic code
change is the Full option; no runtime package rebuild or product fix is claimed.

## Chain setup lifetime and focused stage-exit correction

The diagnostic gains optional None/Admin/Source/Chain setup using the actual
chain helpers. It records fixed lifecycle/process metadata and check identities,
omits helper result details, preserves the source report bytes, and retains the
private settings snapshot through cleanup. The first immediate-exit calibration,
`projection-live-control/d3b046a427bd4ed68537d75fe23eb7e6`, passes all three Admin
checks but stops before source/live stages because Admin Excel remains just after
Quit. It exits naturally during cleanup; settings/packages and delayed zero native
audit pass. This calibration is not product RED or evidence of a stuck process.

The diagnostic then observes up to30 seconds of natural stage exit. With both
actual setup stages, `projection-live-control/d60f570110ac45279e2fe3bfd78d9949`
passes Admin3, source15 and all48 prior ordered live-role checks. Start02:53:21.1969840,
audit02:56:09.2722742 UTC on2026-10-02. Admin's first residual-process observation
is02:53:33.9338532; exit is observed02:53:39.7687459. The source-stage residual
exits between02:54:01.1354027 and02:54:01.6604680. No Excel remains before live
actions. All source/live report bytes, settings and package pins are preserved;
normal shutdown and delayed zero native audit pass. The diagnostic's waits differ
from the original chain; this is not an established native-crash cause.

`Test-Slice4beChainStageExit.ps1` protects the actual canonical chain prefix up to
the source-stage assignment. It extracts the unchanged Admin helper and original
top-level statements, runs the real packaged entry, then observes process exit
before any source integration begins. On the unmodified chain, root
`chain-stage-exit/ca1b6de2b4f64374a9808c98f7cd066d` records tooling RED3 PASS/1
expected FAIL,02:57:39.1641275-02:58:08.9217051 UTC. Generation, seeding and canonical
inventory creation pass; `AdminEntry.ExcelExitedBeforeSource` fails with one
remaining Excel process. It exits naturally during cleanup, and settings/packages
and the delayed native audit are clean. This is meaningful stage-lifetime evidence,
not a product behavioral RED or a claimed reproduction of the native crash.

The canonical chain now calls the existing `Wait-RecordingCleanup` immediately
after `Invoke-AdminEntryGate`, before source integration. No product, package,
business phase, assertion, permission or architectural contract changes. Focused
GREEN retains the same four identities and passes4/4:
`chain-stage-exit/9c31850834294664b1523d31ddd26288`,02:58:30.9933277-
02:59:00.5136621 UTC. No Excel remains at the observed boundary; normal closure,
settings/package preservation and delayed zero native audit pass. The changed
chain's full validation is recorded below; the earlier native failures remain
preserved and unexplained.

`chain-stage-exit-static-01` retains291 components/6177 procedures/135555 lines,
9 literal/45 unresolved calls,190 duplicate candidates and28 unchanged caps.
Three schemas and380 PowerShell parses pass, including the new focused harness.
There is no runtime-source growth or package rebuild in this tooling correction.

The unchanged entry02 packages then pass the corrected full chain:
`complete-entry02-regression/chain-a98044117e04412d98d6d30a403bc752`,
2026-10-02 02:59:19.1108119-03:04:59.7629941 UTC. All32 chain,48 live-role and15
warehouse checks pass with exact prior ordered identities preserved. The later
restart/runtime-evidence/static stages are reached. Excel closes without
assistance, settings and all three tracked reports are restored, all five package
pins remain unchanged, and the delayed native audit is zero. This verifies the
current candidate's chain gate after the tooling correction; it does not prove
that stage overlap caused either earlier native crash or establish Slice4be
acceptance. Resume the remaining Complete Run closure/interruption/submission
and observation coverage rather than repeating this passing chain unchanged.

## Native captured-workbook closure verification

The unchanged entry02 packages pass258/258, retaining all206 prior ordered checks,
including shared42 and five instrumented compiles. Controller
`production-run-local-controller/9f7cf60d4d554fa8a4239d77737b42fe` runs
2026-10-02 03:10:09.7916836-03:24:41.7370957 UTC; result root is
`slice4be-production-complete-baseline/12d6b22d19c04b7ab87d6828cc3689af`.
The52 new assertions exercise native close at entry, CompletePending,
AvailableQuantity and EntityKind through the actual Complete Run Click handler.

At entry, the native form survives closure; one real click returns the existing
captured-context refusal. Its new capture was reviewed. In all three in-action
closures, the native form is dismissed after the already-running handler returns.
Each case has exactly one handler entry, zero later reads and zero form
reinitializations. Dismissed controls are not queried or recreated. Owner state,
exact input balances, saved operator bytes and the unrelated workbook are
preserved, with no new or redirected activity. Surviving-control assertions are
conditional on the native window remaining available; they are not claims about
dismissed controls. Populated field clipping and multiline acceptance remain open.

Settings and all five package pins are preserved, Excel closes without assistance,
and the delayed native audit is zero. Both overlapping desktop monitors covering
the run finish with398 samples each and no cursor/desktop/capture failures. The
usage interruption delayed independent receipt verification until14:17 UTC; it
did not require restarting the test or forcing cleanup.

`complete-closed-static-01` preserves291 components/6177 procedures/135555 lines,
9 literal/45 unresolved calls,190 duplicate candidates and28 unchanged caps.
Three schemas and381 PowerShell parses pass. This is test-only evidence under
existing D18; no product RED, runtime fix, package rebuild or new contract is
claimed. Worksheet completion/interruption, partial-submission facts and Complete
Run observations remain open. Existing same-package smoke/layout/chain evidence
is retained without repeating unchanged gates.

## D18 loading and nested completion entry RED

Architecture's Complete Run action-entry clarification constrains existing D18
semantics: loading/busy/nested callbacks are not independent completion attempts.
The test invokes the actual `mBtnManagerApplyOutput_Click` with loading/busy flags
and recursively at the real reusable pending UI yield. Unsaved probes count actual
handler entries and completion-owner entries before any guard, snapshot the state/
projection/message around nested entry, and read real exact inventory balances.
No owner, writer, processor or inventory-read result is replaced.

On unchanged `validation-production-complete-continuation-01`, RED is169 PASS/
16 FAIL/185 unique checks, retaining all150 prior ordered PASS results. Controller
`production-run-local-controller/7788fc7a77614871ac8247a5dc9882c8`, result
`slice4be-production-complete-baseline/0113a18ce8614b518e09481b2efe9198/red.json`,
2026-10-02 01:21:43.1783392-01:34:15.1829016 UTC. Loading fails six checks: its
flag is not restored, the owner is entered, staging/projection/message change and
exact inputs are consumed. Busy entry fails the same five effect checks while
preserving its flag. The recursive case reaches two real handler entries and
two completion-owner entries; nested owner/projection/message preservation,
single owner entry and the final success message fail. Input is consumed once
and a fresh output with the expected quantity exists; duplicate consumption is
not claimed. The outer call overwrites completion success with a Check In refusal.

Shared42/five instrumented compiles, canonical pins, settings restoration, natural
unassisted closure and delayed zero native audit pass. Captured-workbook custom
value/formula, saved bytes, decoy and other warehouse are preserved. All three new
captures were individually reviewed: loading/busy visibly complete, while nested
entry shows completed rows/output and the contradictory Check In refusal. Other
regenerated captures were not additionally reviewed; long field clipping remains.

`complete-entry-red-static-01` retains290 components/6176 procedures/135530 lines,
9 literal/45 unresolved calls,190 duplicate candidates and28 unchanged oversized
caps. Three schemas and379 PowerShell parses pass. No runtime correction, GREEN
or observation integration is claimed. Next is a typed completion entry guard
that suppresses loading/busy/nested work, holds the outer busy state across yields
and restores prior flags without swallowing errors, followed by packaged GREEN.
The current new cases prove normal-return guard behavior only; forced-exception
restoration, closed-workbook entry and partial-submission facts still need evidence.

## D18 and D-NAS interruption correction

The correction enforces the existing captured-context and Core capability rules.
`cProductionWorksheetAction.BeginOwner` binds an owner continuation without
starting an observation; the existing tracked `Begin` is unchanged. The form
checks at entry and after both pending UI yields. The selected reusable owner
accepts the existing optional action, checks entry, passes it through real Check
In reads and checks again before generating output keys. Existing direct callers
without an action retain their optional-argument behavior. No new activity ID,
permission, authority, retry or rollback contract is introduced.

New unpromoted `validation-production-complete-continuation-01` was built under
`complete-continuation-build-01`,2026-10-02 00:34:17.1184213-00:34:54.7928211 UTC.
Cold Operations startup/five compiles pass. Of283 components, exactly the
Operations action class, form and reusable owner change;280 remain identical.
Frozen binding01/settings are preserved, with natural closure/delayed zero audit.

Final GREEN is150/150, retaining all150 RED identities and126 prior passes:
controller `production-run-local-controller/2756cf2f9b904bda97e9e30da4d582e3`,
result `slice4be-production-complete-baseline/513182fc040f44bb9f98522a5ccc15c6/green.json`,
00:44:41.8669611-00:53:36.4456737 UTC. All six interruptions preserve the owner
and displayed projection at the interrupted boundary, stop later reads and
preserve exact input balances. Pending-yield cases never enter the completion
owner; read-return cases enter once. Shared42/five instrumented compiles,
authorization/settings/package preservation, normal unassisted shutdown and
delayed zero native audit pass. All six new refusal captures were individually
reviewed; other regenerated captures were not additionally reviewed. Long field
clipping remains; this is not populated-list or multiline acceptance.

A preliminary150/150 GREEN used controller
`production-run-local-controller/53d44b38d8874ea2a2eefeeaac198a4a`, result
`slice4be-production-complete-baseline/4a524467b089444e9b6d47bdaae8cc09/green.json`,
00:35:12.9673336-00:44:03.5325714 UTC. Preservation/closure/audit passed, but its
entry counter followed the newly added owner guard. Both owner counters were
moved before the first `On Error` statement, ahead of all guards, and the entire
gate was rerun above. No runtime/package/assertion was changed for that rerun;
the earlier RED on binding01 remains valid. Preliminary captures were not reviewed.

`complete-continuation-static-01` records290 components/6176 procedures/135530 lines,
9 literal/45 unresolved calls,190 duplicate candidates and28 non-growing oversized
caps. The only growth is one method/nine lines in the existing small action class,
explicitly accounting for owner-only binding; form and reusable-owner line counts
are unchanged. Three schemas/378 PowerShell parses pass. This introduces no
oversized-cap or dynamic-call exception. Raw reports/packages/captures stay ignored.

On this unchanged candidate, Check In passes629/629, preserving every prior
ordered identity/PASS result: controller
`production-run-local-controller/c63ad36fa0f4454993b148872252e1ff`, result
`slice4be-production-check-in-activity/daa7ac04b75c4d8e934c29c14bd53a7b/green.json`,
00:54:29.8284785-01:11:15.0043695 UTC. Shared42/five compiles, canonical pins,
settings restoration, unassisted shutdown and delayed zero native audit pass.
The controller was retained through the slow Excel shutdown. Captures were
regenerated but not additionally reviewed in this regression.

Packaged smoke passes86/86 with every prior ordered PASS retained under
`complete-continuation01-regression/smoke-66c889c53cf94d2aa761a2fe61bfb1a4`,
01:11:42.3850940-01:12:04.7438626 UTC. Both unassisted exits, canonical pins,
settings/tracked-report restoration and delayed zero native audit pass.
Layout under `complete-continuation01-regression/layout-21746bdeeb664d35858f57e0f261734e`,
01:12:27.6954647-01:12:44.5549188 UTC, retains18 requested size/page pairs,
six activated/maximized pages, five native transitions and two actual sizes.
All six PNGs match the previously reviewed binding01/selection01 captures byte
for byte; the ignored capture-review record retains the hash linkage. Geometry,
settings/pins, natural closure and delayed zero native audit pass. These empty
Run List/Settings views do not establish populated scrolling or multiline acceptance.

Full-chain32/32, live-role48/48 and warehouse creation15/15 pass under
`complete-continuation01-regression/chain-755bd4cf22cc4dc3bb7ceb9572624592`,
01:13:02.8042493-01:18:13.6250998 UTC. Every prior ordered PASS result is retained;
settings/tracked reports are restored, candidate pins preserved, Excel closed
and the delayed native audit is zero. These are evidence-only regressions on
the frozen continuation01 candidate, with no additional runtime change or RED.

Closed-workbook entry, reentrancy, post-submission interruptions, exact partial-write facts and
Complete Run observations remain open. Worksheet yield checks are implemented,
but these six interruption cases exercise the reusable branch only. This focused
correction is not full Complete Run or Slice4be acceptance.

## D18 and D-NAS completion interruption RED

The initial binding correction below does not protect later yielding boundaries.
D18 requires captured-context integrity; D-NAS requires current Core capability
checks for operator writes. Test-only expansion on unchanged complete-binding01
adds six interruptions: sign-out or revocation of the fixture producer's PROD_POST
capability after the real reusable `ShowPersistencePending` return, or after real
`AvailableQuantity`/`EntityKind` Domain reads. Existing Check In interruption
adapters apply and verify each interruption; no completion owner, writer, processor
or read result is replaced. The permission fixture remains signed in with the same
captured context, isolating capability loss from sign-out.

RED is126 PASS/24 FAIL/150 unique checks, retaining all89 prior ordered PASS
results and adding61 checks. Controller
`production-run-local-controller/a8222d8687a14df09841cac385806c46`, result
`slice4be-production-complete-baseline/c3dad397684c435592a5e9d47440b1d9/red.json`,
2026-10-02 00:21:56.3359746-00:30:16.3881465 UTC. Each interruption reaches its
real boundary. At the pending UI yield, both interruption types still enter the
owner and perform later reads. After AvailableQuantity they still perform later
reads. All six cases alter owner staging and its displayed projection, then show
writer refusal instead of the boundary context/permission refusal. Exact input
balances remain unchanged in all six cases; no inventory-consumption failure is
claimed here. Disposable authorization bytes are restored, and operator bytes,
custom value/formula, decoy and other-warehouse pins are preserved.

Shared42/five instrumented compiles, canonical package pins, settings restoration,
unassisted closure and delayed zero native audit pass. All six new interruption
captures were individually reviewed; long field clipping remains. Other
regenerated captures were not additionally reviewed. `complete-yield-red-static-01`
retains290 components/6175 procedures/135521 lines,9 literal/45 unresolved calls,
190 duplicate candidates and28 non-growing caps; three schemas/378 PowerShell
parses pass. A test-comment reference was corrected from D5 to D-NAS without
changing test behavior; D5 governs Config commands specifically.

No runtime correction or GREEN is claimed for this expansion. Source review
identifies the missing continuation: the form yields before owner entry, and
`CompleteReusableProcess` calls `CheckInReusableProcess` without its existing
optional captured-action parameter. Next is to carry the existing context/Core
permission checks through those boundaries, preserving owner state at interruption.
This constrains existing rules and adds no activity ID, observation contract,
authority store, permission, automatic retry or rollback claim. Closed-workbook
entry, reentrancy, later partial submissions and complete tracking remain open.

## D18 captured binding at completion entry

D18 requires each action to check its captured warehouse, invSys session and role
workbook. The actual `mBtnManagerApplyOutput_Click` lacked that entry check.
The expanded packaged baseline invalidates the binding after real Check In by
changing target, replacing the session, or signing out. It retains all68 earlier
ordered checks and adds21. Owner-entry probes observe the original completion
owner; inventory reads, writes and processing remain real disposable-fixture work.

On unchanged complete-selection01, RED passes80/fails9 of89 unique checks:
controller `production-run-local-controller/4ddd87c3b66c438b96ea98ab7e32fbb9`,
result `slice4be-production-complete-baseline/a7811824e5c546c8ba158340ba0cd180/red.json`,
2026-10-02 00:01:23.0509757-00:06:54.1014669 UTC. All68 prior checks pass. Target
change still enters the owner and shows a stock refusal instead of a context
refusal. A replaced session permits completion and actual input consumption.
Sign-out enters the owner and changes staging before the writer rejects it.
Exact original input balances remain unchanged for target/sign-out; other-warehouse
pins, operator bytes, custom value/formula and decoy remain preserved.

The correction calls the existing `modProductionRunBinding.RequireCurrentContext`
at the start of `frmProduction.CompleteProductionRun`, before either completion
branch. It uses the existing context-refusal message and changes no contract or
tracking catalog. It does not establish protection across later yields.

New unpromoted `validation-production-complete-binding-01` was built and cold
compiled under `complete-binding-build-01`,00:07:37.1833200-00:08:14.7507983 UTC.
Cold Operations startup/all five compiles pass. Only Operations/frmProduction
changes among283 packaged components;282 remain identical and the old candidate
is preserved. GREEN passes89/89 with all ordered identities and prior passes
retained: controller `production-run-local-controller/358760a7ddb24e28994a3455170405bd`,
result `slice4be-production-complete-baseline/3de518bb39274a8a9b914748345df281/green.json`,
00:08:36.4675150-00:13:50.9942211 UTC. All three invalid bindings are refused before
owner entry; staging, exact input balances and unrelated surfaces are preserved.

Both focused gates retain shared42/five instrumented compiles, package/settings
preservation, unassisted closure and delayed zero native audits. The build audit
also passes. Three new RED and three new GREEN captures were individually reviewed;
GREEN's existing context message is legible. Other regenerated baseline captures
were not additionally reviewed. Long labels still clip. Ignored static evidence
`complete-binding-red-static-01` and `complete-binding-green-static-01` retains290
components/6175 procedures/135521 lines,9 literal/45 unresolved calls,190 duplicate
candidates and28 non-growing caps; three schemas/377 PowerShell parses pass.

New-hash smoke passes86/86 with exact prior ordered checks and both unassisted
exits: `complete-binding01-regression/smoke-02934cc44ac74a808bd143bdd42e6e59`,
00:14:20.8088609-00:14:54.5160064 UTC. New-hash layout passes18 requested size/page
pairs, six activated/maximized pages, five native transitions and two actual sizes:
`complete-binding01-regression/layout-425f3afc1ae24046abd0cc6c6c801ef9`,
00:15:30.9620903-00:15:47.8976385 UTC. All six PNGs are byte-identical to the six
individually reviewed complete-selection01 captures; the ignored review records
the exact hash linkage. Geometry remains unchanged. Both gates preserve packages,
restore settings/tracked reports, close Excel and pass delayed zero native audits.
These empty-form views do not prove populated scrolling or multiline acceptance.

New-hash full-chain also passes32/32, live-role48/48 and warehouse creation15/15,
retaining every prior ordered PASS result: controller
`complete-binding01-regression/chain-3c500425e8ee410c83e965fa4b1737a7`,
2026-10-02 00:16:12.1291833-00:21:37.0738826 UTC. Settings and tracked reports are
restored, packages preserved, Excel closed and delayed native failures zero.

New-hash Check In regression, closed-workbook entry and
changes during yields, partial submissions, observation integration, populated
scrolling and full human/NAS acceptance remain open. Earlier complete-selection01
evidence below applies to that older candidate; it is not substituted for new-hash evidence.

## Broader gates on the corrected package

The unchanged, unpromoted complete-selection01 candidate passes the following
gates on its canonical five-package pins. Each controller restored settings,
preserved packages and restored any tracked test reports before terminal exit.
Excel was closed and the delayed Application1000/1001/1002 audit found zero
Excel failures for each gate. Raw reports and captures remain ignored.

| Gate | Evidence under reports/runtime/complete-selection01-regression | Result |
| --- | --- | --- |
| Packaged smoke | `smoke-60f4b5a4d14d432f8013082fbad2a3e2` | 86/86; exact prior check order; initial/final unassisted closure verified |
| Packaged layout | `layout-2d3ca5d270224d49867a7b5d384db208` | 18 requested size/page pairs, six activated/maximized pages, five native transitions, two distinct actual sizes; representative geometry preserved |
| Full Release1 chain | `chain-625b549f08094ec786e94c0ccecaf6fe` | 32/32 chain,48/48 live-role,15/15 warehouse creation; exact prior ordered PASS results retained |

Smoke ran2026-10-01 23:50:38.8715725-23:51:01.4770427 UTC; layout ran
23:51:39.4175655-23:51:56.6063880 UTC; chain ran
23:52:43.1959440-23:58:07.1783001 UTC. The chain harness completed its ordered
warehouse/Seed/Receiving/Production/Boxing/Shipping/restart sequence.

All six new layout captures were individually reviewed and their hashes recorded
in the ignored capture-review record. No overlap or out-of-bounds regression was
found. These are empty Run List and Settings views; they do not establish populated
list scrolling, multiline rendering or human usability acceptance. No runtime,
test or contract change was needed for this evidence-only checkpoint; no new RED
is claimed. Owner interruption/partial-submission tests, Complete Run observations
and full human/NAS Release1 acceptance remain open.

## Selected Process with an unallocated second Process

Test-only expansion on unchanged complete-selection01 passes68/68, preserving
all59 prior ordered identities/PASS results and adding nine checks. Controller
`production-run-local-controller/a9cfe91e4add40be81c59084cba2e839`, result
`slice4be-production-complete-baseline/508474226bc24e3bb0689568bff23bc7/green.json`,
2026-10-01 23:44:56.7100635-23:48:34.2268026 UTC. The fixture creates two released
Processes and a released Recipe through the existing form workflow, allocates
and checks in only the selected Process, then invokes the actual Complete Run
handler. The selected owner alone is entered; its exact allocations are consumed
and its new output key has the expected quantity. The second Process remains
unallocated, visibly NEEDS ALLOCATION, without an output key or completed state.
The batch remains incomplete. Captured-workbook custom values/formula, decoy,
saved operator bytes and other warehouse are preserved.

Shared42, five instrumented compiles, canonical pins, restored settings,
unassisted closure and the delayed zero native audit pass. All three captures
were individually reviewed. Long fields still clip; the multi-Process capture
shows the second Production Output row only partly visible at the captured size.
Scrolling/populated-list usability remains unverified; this observation does not
establish a new minimum-row contract. `complete-multi-static-01` retains290
components/6175 procedures/135521 lines,9 literal/45 unresolved calls,190 duplicate
candidates and28 non-growing caps; three schemas/377 parses pass. There is no
runtime change or new product RED in this expansion. Complete Run tracking,
interruption/partial-submission evidence and full acceptance remain open; the
broader gates above have now passed.

## Check In regression on the corrected package

On unchanged complete-selection01, Check In passes629/629 with all629 prior
activity03 identities and PASS results retained in order. Controller
`production-run-local-controller/cb210925e13843ac8bc7e4fb003aa362`, result
`slice4be-production-check-in-activity/ab2d6ec127e44ffb82b6e5a2a5edbd1e/green.json`,
2026-10-01 23:28:03.3323285-23:44:38.8953672 UTC. Shared42, five instrumented
compiles, canonical new-candidate pins, settings restoration, unassisted shutdown
and the delayed native audit with zero Excel failures pass. Excel shutdown again
took several minutes; the controller was retained until natural exit. Captures
were regenerated but not additionally reviewed in this regression. No runtime
change or new product RED is claimed. The later multi-Process gate above adds
selected-only owner evidence; the broader new-candidate gates are recorded above.

## D15 selected-Process prerequisite

Architecture v4.11 D15 requires one selected Process at a time. Source review
found that `frmProduction.CompleteProductionRun` instead called
`CompleteReusableRun` when `ActiveRunProcess()` was empty. The test exercises the
actual `mBtnManagerApplyOutput_Click` through an unsaved adapter in the packaged
Operations XLAM. Its completion-owner counters observe entry without replacing
the owner, writer, processor or inventory reads.

On the unchanged `validation-production-check-in-activity-03` candidate,
the focused RED is55 PASS/4 FAIL/59 unique checks. Controller
`production-run-local-controller/47cf6f7264754377a97062c4fccf386c`, result
`slice4be-production-complete-baseline/418bd17ee6a2429fbb77a1a1fcc76834/red.json`,
2026-10-01 23:18:29.2097244-23:20:54.7677237 UTC. The selected positive case
completes, consumes the exact allocated entities and creates a fresh output key
with the expected quantity. With no Process selected, the actual handler enters
the whole-run owner, changes staging, consumes input inventory and presents batch
success. The four failures protect owner non-entry, unchanged staging, unchanged
exact input balances and the selection-required message. Shared42, five compiles,
package/settings preservation, unassisted closure and the delayed zero native
Excel failure audit pass. Both captures were individually reviewed.

An earlier attempt stopped because the test had not loaded `RestartPins`:
controller `207bf758a8414587ab15f7598c8fb954`, result
`slice4be-production-complete-baseline/f73e6f86b0a944ab91f0b5db674828de/red.json`.
Its39 PASS/1 harness failure is not product RED. Settings/packages were restored,
Excel closed unassisted and its delayed native audit was zero. The helper now
loads the established fixture utilities inside the test's fixture scope.

The correction follows existing D15: reject an empty Process selection before
batch-note synchronization, Actual Output staging or completion-owner entry,
using the existing owner wording `Choose one Process before Complete Run.`
The selected owner remains `CompleteReusableProcess`; the form's whole-run
fallback is removed. No new event catalog or tracking contract is introduced.

New unpromoted candidate `validation-production-complete-selection-01` was built
and cold compiled under `complete-selection-build-01`,23:21:53.8385555-
23:22:31.4543385 UTC. All five compiles and cold Operations startup pass. Exactly
`invSys.Operations.xlam/frmProduction` changes among283 compiled components;
282 remain identical. Frozen activity03, settings and normal shutdown are
preserved; the delayed native audit is zero.

Focused GREEN passes59/59 with all59 ordered identities and all55 prior passes
preserved. Controller `production-run-local-controller/0b52b368e4e441e18ef5b788d5065873`,
result `slice4be-production-complete-baseline/b97fc38d64124c58813d9ba37adb0fee/green.json`,
23:22:51.0809958-23:25:10.2943940 UTC. Shared42, five instrumented compiles, new
candidate pins, settings restoration, unassisted closure and the delayed zero
native audit pass. Both new captures were individually reviewed: the refusal
message is legible and selected completion remains visible. Long fields still clip.
Raw captures/reports stay ignored.

`complete-selection-static-01` retains290 components/6175 procedures,9 literal/
45 unresolved calls and190 duplicate candidates. Runtime lines decrease by one
to135521; all28 module caps are non-growing. Three schemas and377 PowerShell
parses pass. The RED baseline is `complete-baseline-static-01` (135522 lines).

Check In629 was subsequently verified on this candidate above. This
one-Process positive fixture does not establish multi-Process completion;
the later gate above adds that focused case. Packaged smoke/layout and live-role/
full-chain subsequently passed above. Interruption behavior, observations and
full human/NAS Release1 acceptance remain open.
