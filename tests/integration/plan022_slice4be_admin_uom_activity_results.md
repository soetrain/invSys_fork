# Plan 022 Slice 4be Admin UOM activity

Architecture v4.11 D18's Admin UOM command refinement governs the existing
Settings Add, Remove and Reset handlers. D5 retains Core configuration authority;
Admin observes actual entry and explicit owning outcomes. Catalog 10 and the
three controls are specified before implementation under semantic inheritance.
No business permission, authority store or Inventory event is introduced.

Focused packaged RED is **58 PASS / 68 expected FAIL**, root
`b252af03029142419489a8e573ea1618`, **2026-09-22
11:11:16.9200212--11:13:14.2024076 UTC**, on unchanged
`deploy/validation-guide-action-curation-integrity`.

All **43** baseline/setup checks pass, including explicit compilation of all five
instrumented packages before forms. The **83** Admin UOM checks exercise the real
Add/Remove/Reset handlers, including actual owned native Yes/No confirmations.
Eleven actual outcomes pass their owning-effect assertions: Add, duplicate Add,
invalid Add, Remove, no selection, all three denied commands, Reset No, Reset Yes
and unchanged Reset Yes. Each lacks its six required observation/correlation/
effect/privacy/integrity/reference checks (**66 failures**). Two further failures
prove a held Settings form can still change UOM configuration after a replacement
sign-in session or redirect the change to a different selected warehouse. These
are behavioral RED under existing D18; they are not compile/fixture failures.

Direct service calls create no extra observations, existing activity bytes remain,
and both generated Config fixtures are restored exactly. Test/runtime source and
all five package pins stay unchanged. Excel closes normally without assistance;
no matching Application 1000/1001/1002 events occur. Verification is
`reports/runtime/verify-admin-uom-commands-red.ps1`, run before later edits.

All three native dialog captures are directly inspected and hashed. No and
unchanged Yes visibly show the exact question; the first Yes capture has a blank
body and is **rejected as visible evidence**. The successful actual Yes command
and missing observations remain independently proven by the fixture assertions.
A test-only 300 ms paint opportunity before capture is added; it revalidates the
exact question and replays no command. New visible evidence remains required.

The isolated implementation now passes **126/126**, preserving every RED check
identity, on `deploy/validation-admin-uom-activity`, root
`ecfbcfdc4f75493c8f246c9a0b42c661`, **2026-09-22
11:21:03.5475342--11:23:08.7978035 UTC**. All five packages build and explicitly
compile, and Operations cold start passes. Of 242 compiled components, only
frmAdminSettings, modAdminSettingsAction, modActivityCatalog and modUomSettings
change. Captured context is checked before the command; Core propagates explicit
writer outcomes; catalog 10 retains earlier definitions. All three fresh native
Reset captures are fully painted, directly reviewed and hashed.

Build, compile and GREEN close Excel normally without assistance, retain source
and package pins, and have no matching Application 1000/1001/1002 events. Static
maintenance reports **249 components / 6,036 procedures / 132,551 lines**, dynamic
calls **9/45**, duplicate groups **192**; all 28 prior module limits pass. The
increase is one procedure and 59 lines. Runtime pin normalization preserves the
original captured file/hash vectors; it does not substitute new baseline hashes.
Verification is `reports/runtime/verify-admin-uom-activity-green.ps1`, passed
before the subsequent editor-test extension.

Source review then discovers that the expectation editor's candidate outcome
list omits CANCELLED despite Reset registering it. A separate actual-editor RED
passes **96 checks with exactly two expected failures**, root
`db5ea93ceae54003b236910881006bd5`, **2026-09-22
11:39:40.8223517--11:43:16.4257653 UTC**. The native Reset No action records
CANCELLED without changing Config. The actual editor cannot select that outcome
or retain an authored cancelled step after actual Stop. Catalog registration,
Add exclusion, authoring without business mutations, unchanged original activity
and all existing checks pass. All five instrumented compiles pass; six native
captures are directly reviewed and hashed. Sources/packages remain unchanged,
Excel closes normally without assistance and no matching Application errors occur.

After verified RED, modExpectationDraft adds CANCELLED to its candidates, still
filtered through each control's existing catalog definition. A cancelled Reset
step can precede a completed Save Value terminal; cancellation itself does not
claim command completion. The corrected candidate
`deploy/validation-admin-uom-expectation` passes **98/98**, root
`1e8d3b9d787c498698fa6f4a8b7438a7`, **2026-09-22
11:47:26.0639409--11:51:10.6642631 UTC**. All six new captures are directly
reviewed and hashed. All five packages build/compile, Operations cold start
passes, and only modExpectationDraft differs from the prior 126-GREEN package
set. The editor closes Excel normally without assistance. Static metrics stay
249/6,036/132,551, dynamic calls 9/45 and duplicates 192; all 28 module limits
and all three static report schemas pass.

The master's next activity invocation exits 2 before report/Excel creation, with
two PowerShell parameter-matching WER events. It is an incomplete harness run,
not product RED. An offline two-switch probe reproduces failure when one flag is
held as a scalar string and native-splatted, while a typed string array succeeds
with the intended flag only. The fresh activity runner uses explicit arguments;
no product behavior or assertion is relaxed.

The explicit run reaches **166 PASS / one harness exception**, root
`2b2d368e4c12416db5a9786fbd653229`; a counts-only repeat reaches the same boundary,
root `4c0ffb57dfff46c99ad81dc847fbd110`. Both close normally. The new publication
fixture incorrectly assumes eleven pairs; the real Reset setup also performs an
Add, so there are **twelve pairs / 24 records** (Add 10, Remove 6, Reset 8).
The corrected assertion requires every one of those twelve real actions, including
the setup Add. These are incomplete harness runs, not product RED. Runtime and
frozen packages remain unchanged.

The corrected expanded gate is **225/225**, retaining all 126 preceding check
identities, root `e2df987c476a4d8792a1b4d6299df0cb`, **2026-09-22
11:59:40.9999819--12:02:24.3204599 UTC**, on the same frozen expectation candidate.
It preserves all catalog-9 definitions, excludes new controls from versions 1-9,
checks catalog-10 metadata and Reset-only cancellation, publishes all twelve
actual actions with exact original identities/captions/outcomes, and preserves
the original activity bytes. Actual Add, Remove and Reset No/Yes retain their
owning results with unavailable storage, an older valid whole policy or explicit
collection off. Tracking notices match availability; reads and commands leave
policy rows/unknown columns unchanged; generated Config and prior activity bytes
are restored exactly. All nine fresh confirmation captures are directly reviewed
and hashed. Normal unassisted Excel closure, unchanged source/package pins and
zero matching Application errors pass. `verify-admin-uom-expanded.ps1` verifies
both 98-check editor and 225-check activity gates, their retained identities,
15 reviewed images and static limits.

The complete Settings regression is **191/191**, retaining every preceding
191 identity, root `8d870c592b19445586203540cb919848`, **2026-09-22
12:03:05.6717073--12:06:19.5050481 UTC**. It includes whole-policy saves,
detail profiles, preference restart/isolation, Operations-only Settings, native
maximize/restore, existing Production UOM behavior and Admin launcher/Close.
Sources remain unchanged, Excel closes normally and no matching Application
errors occur. All ten captures are directly reviewed: nine are accepted; the
initial General-tab image has unpainted controls and is **not accepted as UOM
layout evidence**. Fresh General-tab evidence remains required. Settings now
loads exactly 36 catalog-10 entries; earlier catalog-specific tests stay intact.

Fresh visible General coverage now passes **228/228**, retaining all 225 preceding
checks, root `c4f00c15aa664dd4817251a6edf9f382`, **2026-09-22
12:12:38.0882399--12:15:22.0399439 UTC**. Test-only probes verify loaded Config/UOM
lists and visible contained UOM controls, then capture the unique owned foreground
form after repaint. Viewing the surface changes no Config or activity. All eleven
current captures are directly reviewed and accepted, including the fully painted
General tab; the earlier rejected image remains rejected. All instrumented
packages compile before forms. Runtime/package hashes and static metrics remain
unchanged; Excel closes normally without assistance and no matching Application
errors occur. This supplementary evidence requires no runtime or contract change.

The frozen candidate also passes full Release 1 chain **32/32**, live roles
**48/48** and Create Warehouse **15/15**, **2026-09-22
12:17:16.2931347--12:25:54.9148365 UTC**. Every preceding chain/live-role identity
is retained; all 239 test/tool source pins and five package hashes remain intact,
the three tracked reports are restored exactly, and no matching Application
1000/1001/1002 events occur. Runtime source pins remain unchanged.

Cleanup is **assisted normal cleanup**, not unattended success. An earlier Excel
instance exits naturally before a read-only inspector attaches; its identity
guard refuses the stale process. A later exact native process/creation identity
has a typed WorkbookCount of zero, rechecked before normal Quit. The ensuing
Document Recovery question initially selects removal; the existing **Yes, I want
to view these files later** option is selected, captured, directly reviewed and
confirmed. Recovery files are retained and Excel exits; no forced termination
or operational workbook mutation occurs. `verify-admin-uom-chain.ps1` validates
the result and both recovery captures. Process IDs/handles in ignored helpers
are stale after this run and must not be reused.

The same frozen candidate retains three further supporting gates on 2026-09-22:

| Gate | Result / retained prior GREEN checks | Report identity | Reviewed captures |
| --- | --- | --- | --- |
| Direct curation / native layouts | 77/77 | `f0f763bfd964477887652b2e59c4a3ac` | 18 |
| Saved preference / new Excel process | 27/27 | `6b9f16f9ff694f68a971ec6d53e62c35` | 3 |
| Published Viewer / filters / Shipping state | 94/94 | `62264356052c46a0acb07cfc523cf119` | 3 |

Each gate preserves every preceding check identity and runtime/test/package
hash, closes Excel normally without assistance, and has no matching Application
1000/1001/1002 errors. Restart closes during the runner's bounded exit grace
period; it does not require intervention. All 24 captures are directly reviewed
and hashed. The ignored `verify-admin-uom-supporting-gate.ps1` verifies terminal
results, prior GREEN identities, source/package pins, image hashes and native
error audit. These runs span 12:26:45.8821751--12:36:45.8071131 UTC.

Full-guide, Operations comparison, Boxing/Shipping and human acceptance remain open. This is a
verified implementation checkpoint, not
completion of the runtime slice or Release 1.
These focused results do not establish comprehensive Admin or Slice 4be acceptance.
