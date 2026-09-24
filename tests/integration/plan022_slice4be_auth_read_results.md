# Slice 4be D8-A: ordinary Auth reads and explicit provisioning

Last verified: 2026-09-24 UTC. Architecture v4.11 D8-A was explicitly approved
before this implementation. D2's credential-authorized station transition and
D8's cache/processor rules remain binding. This record does not claim full
Slice 4be or Release 1 acceptance.

## Protecting packaged tests

`Test-Slice4beAuthReadOnly.ps1` drives the existing packaged-caller harness and
`Slice4beAuthReadOnly.ps1`. Both authority fixtures originate in real Admin
Generate Warehouse. Ordinary calls use public Core LoadAuth/ReloadAuth, the real
sign-in validator, capability/processor gates and the actual Admin Settings Save
handler. Unsaved probes only observe the chosen Auth path and expire the in-memory
cache; they do not substitute authority, results, credentials or owner commands.

Coverage includes healthy closed/open authority, extra user columns inserted
before managed headers, an unsaved custom value, missing files/directories,
invalid required headers in both Auth tables, an exclusively locked selected file
with another warehouse open, wrong warehouse arguments, unselected discovery,
explicit headless runtime binding, valid TTL behavior, expired refreshes, denied
Settings writes and restored command regressions. Explicit station provisioning
and the existing S1 transition are protected separately: wrong credentials and
missing capabilities are denied; the valid same-user transition and original
grant remain; current-station denies survive.

The existing Core target getter suppresses a stale cached target when an explicit
transaction root differs. The new reader preserves that binding: use the exact
current target, or the explicitly bound Core root with matching validated Config
for bootstrap/headless callers. It does not discover authority from remembered,
default or arbitrary open workbooks. This implements approved read separation
without changing the established provisioning or station-transition contract.

## Focused RED

Frozen `deploy/validation-approved-detail-scroll`, final **60 PASS / 22 expected
FAIL**, all 74 preceding identities retained, five instrumented compiles and no
harness failure. Report:
`reports/runtime/slice4be-auth-read/8b5c27a0be124f46a46daccf42f343d6/red.json`.
Controller: `auth-read-controller/eb3ccf4b1f0b4f8383dd97bd94cda63d` under runtime.
UTC **17:55:48--17:58:12**. Normal delayed unassisted closure, restored settings,
unchanged packages/tooling and zero Application events 1000/1001/1002.

The expected failures prove creation/repair during ordinary Load/Reload, sign-in,
expired operator/processor capability reads and Settings Save; directory creation;
invalid-schema repair/save; unreadable-source fallback; and wrong/unselected
authority resolution. A missing Auth source even recreates processor capabilities
and allows that system capability check. Healthy byte/unknown-column, valid-TTL,
explicit provisioning and authorized station-transition checks pass.

Earlier complete RED checkpoints are **41 PASS / 16 expected FAIL** (57 checks,
report `a8478ac7855147c3ab67e06a97711829`) and **55 PASS / 19 expected FAIL**
(74 checks, report `4bf5a034d1554d60b2fc44af26efc1c8`) in the same runtime folder.
The final expansion adds headless binding/unselected-source checks; no preceding
check is removed or weakened. Receipt: `reports/runtime/d8-auth-read-red-verification.json`.

Three harness attempts are retained separately: 0 PASS/one anchor exception
(`46a9fe0b30b149d18f6811d061750a3b`), 15 PASS/one file-sharing exception
(`38179a9ed64242c0872f7e5f85baf4aa`), and 27 PASS/19 reached expected failures
plus one unsupported-argument exception (`f77dcc04b68943708c274de29d335ab8`).
The corrections honor VBA identifier casing, hash open files with read/write
sharing, and use the existing one-string provisioning bridge. These are harness
repairs, not behavioral RED. Every controller closes and restores settings.

## Implementation and GREEN

Only `src/Core/Modules/modAuth.bas` changes runtime behavior. Ordinary reads
resolve existing exact-context authority, open it read-only if not already open,
reject missing schema, and close only their own transient workbook without saves.
They no longer create, seed, repair or search open lookalikes. Existing callers'
open/dirty workbooks remain owned by those callers. Missing-context and invalid
schema diagnostics now describe refusal rather than self-repair.

Three private fallback-only helpers are removed after reviewing their former
resolver-only callers and checking source/dynamic-name references. Other modules'
similarly named private helpers are retained. This is reviewed related cleanup,
not automatic deletion from a scanner report.

Candidate `deploy/validation-auth-read-separated` builds all five packages and
passes all five explicit compiles plus Operations cold start. Exactly
`invSys.Core.xlam/modAuth` changes among 244 compiled components; source change
is 26 additions/79 removals. The first build preceded the related helper removal;
Core was rebuilt before final compilation or GREEN. Accepted deployment and the
frozen RED candidate remain unchanged.

Focused GREEN passes **82/82**, retaining every RED identity and correcting all
22 expected failures. Report:
`reports/runtime/slice4be-auth-read/ff56507c6f184f89b7fd2de5f22bf2e9/green.json`.
Controller: `auth-read-controller/8e2e1e54b29744abba22c9730d16f14e` under runtime;
UTC **18:02:37--18:05:12**. Five instrumented compiles, normal delayed unassisted
closure, restored settings, preserved packages/tooling and zero Application
events 1000/1001/1002. Receipt: `d8-auth-read-focused-verification.json`.

The superseded Phase 6 auto-bootstrap unit is reconciled as
`TestLoadAuth_RequiresExistingExplicitlyProvisionedAuthority`: missing authority
is refused without creation, then loads after explicit provisioning. Registered
test 96 passes **1/1** using the existing Phase 6 runner with private result paths;
normal closure and restored settings. This supplements packaged RED/GREEN.

The current candidate also passes full-chain/live-role/Create Warehouse
**32/48/15**, retaining exact prior identities. UTC **18:10:13--18:15:39**;
normal unassisted closure, restored settings and three tracked reports, preserved
five packages/275 tooling hashes, zero Application events 1000/1001/1002.
Receipt: `reports/runtime/d8-auth-read-chain-verification.json`.

Final maintenance generation validates all three JSON contracts: 251 components,
6047 procedures (down three), 132972 lines (down 53), nine literal/45 unresolved
dynamic calls (unchanged), 191 duplicate-body candidates (down two), and 28
oversized module ratchets. Output: `reports/runtime/d8-auth-read-static-final`.
The source-accounting reduction follows the reviewed fallback-helper removal.
All 277 discovered PowerShell files parse. Form/layout source is unchanged;
the preceding source-layout 8/8 and 7/7 results retain that scope. Receipt:
`reports/runtime/d8-auth-read-static-verification.json`.

## Broader Shipping/Boxing partial evidence

Current candidate report
`reports/runtime/slice4be-shipping-activity/1c0e716b3cfb4b0e81bde25c211fab6b/boxing-activity-shipping-recording-green.json`
records **1630 PASS / three FAIL**, with **82 preceding checks unreached**.
All seven real Shipping `AuthUnavailable.*.MissingFileNotRecreated` checks now
pass: Add, Update, Remove, Hold, Return, Stage and Send. This is partial evidence,
not a completed broad GREEN despite the harness's phase-based filename.

Two failures are `Boxing.PublishedRead.Make.ReadOnlyFields` and the corresponding
Unbox check. `Slice4beBoxingPublishedRead.ps1` still equates read-only fields with
`mFields.Locked=True`; the approved D18 amendment permits the non-editable
ListBox with `Locked=False` for native selection/scrolling. Reconcile this stale
test with the specification, preserving its identities and the separate native
typing/value-preservation test. Do not revert approved runtime behavior.

The third failure is a capture exception at
`boxing.tracking.unavailablestore.make.png`: no uncovered owned-form caption
point was found. The diagnostic records a visible, enabled, non-minimized owned
form and VS Code in the foreground. It does not establish why raising the form
failed. Sixteen preceding screenshots were individually reviewed and accepted
for their scoped status/control/published-detail presentation. Action Path text
continues outside its smaller viewports; these images do not prove complete text
reachability or multiline acceptance. Six later captures were not reached.

The script reached its terminal result but its host and Excel remained running.
The completed test host, then the verified residual disposable Excel process,
were terminated; the original outer controller remained alive and restored its
in-memory settings snapshot. Excel is now closed. **Normal shutdown is not
accepted.** Packages/tooling are unchanged and the Application event audit found
zero events 1000/1001/1002. Receipt:
`reports/runtime/d8-auth-read-resume0923-regression-boxing-verification.json`;
the same prefix's `host-cleanup.json` and `excel-cleanup.json` preserve assistance.
The run retains 1623 previous GREEN identities, corrects the seven Auth failures,
and leaves the two obsolete assertions plus 82 unreached prior checks unresolved.
Preserve this attempt and isolate the capture precondition before a broad retry.
