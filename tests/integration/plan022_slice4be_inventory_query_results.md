# Slice 4be Inventory query ownership correction

Last verified:2026-09-30 UTC. Focused query/public-launch GREEN, static, smoke,
chain/live roles, layout, full reusable Production and lifecycle pass;
remaining applicable Production gates are in progress. No release acceptance, promotion or
native-crash repair is claimed.

## Contract and supplying path

Architecture v4.11 D3 already requires read-only Domain queries for UI display.
Documentation1db3949 clarifies Core's four Inventory query bridges and their
resolver/lifetime before runtime edits; Plan022 and controls1.328 are synchronized.
Queries must not create/repair/save stores, mutate caller-owned state or borrow
another target. Temporary sources open read-only and close without saving;
supplied or already-open exact sources retain ownership and dirty state.
Existing result envelopes, exact `System_Key` and custom columns are preserved.

The packaged public Production launcher initially changes canonical Inventory
bytes before any Close action. Retained trace proves Picker.Query -> resolver ->
write/create -> schema unprotect -> Saved=False -> save. The exact232-check
baseline, trace, container-part evidence and early trace calibration are in
`plan022_slice4be_production_close_results.md`. Close observations/catalog21
remain a separate unimplemented contract; this correction registers no control.

## Supplemental RED

Test548fadd uses Admin-generated/seeded disposable inventory and typed, unsaved
VBA adapters. Expected Domain values remain in memory and are compared exactly;
reports expose booleans/counts only. The pristine Admin Seed file is preserved
between cases. Custom-column/protection edits are staged without saving in the
caller-owned cases. A missing-store case temporarily moves only the verified
fixture source and restores it afterward. Domain failure uses the declared
Empty/zero envelope, not an unhandled cross-project VBA exception.

Command: `Test-Slice4beProductionClose.ps1 -DeployRoot deploy/validation-detail-columns
-Phase RED -CheckInventoryQueryReadOnly -InventoryQueriesOnly`.

Result:88 PASS/34 expected FAIL across122 unique checks, including80 query cases
and all42 prior GREEN checks in exact relative order. Five instrumented compiles
pass, settings/packages are preserved, closure is unassisted and the delayed
Application audit finds zero Excel failures. UTC interval:
14:38:51.9846184--14:40:24.0286798. Controller:
`reports/runtime/production-close-controller/9d2bab8bd090496bb288872ceac4b1b1`;
result `reports/runtime/slice4be-production-close/50a54cd2fe2e42aa899d8c96043f2e66/red.json`.

The34 failures are four queries opening writable sources and leaving them open
(8), implicit dirty callers losing dirty/protected state and changing saved bytes
(12), absent stores triggering Domain queries/create/open ownership (12), and
empty-result failure leaving a writable temporary source open (2). Nonempty exact
results, explicit caller preservation, existing Empty/zero shapes and reference
bytes pass. Cold seeded-source bytes also pass: that saved seed does not need
schema edits. This does not erase the separate unseeded public-launch write RED.

## Fixture calibration retained separately

Earlier supplemental runs are not clean contract RED. All roots below are under
`reports/runtime/production-close-controller/`; result GUIDs are under
`reports/runtime/slice4be-production-close/`. Their individual closure/evidence
files retain exact UTC intervals. All five compiles pass in these runs.

| Controller | Result | Observation |
|---|---|---|
| `fc7863ee2d8945888808321975dd696f` | `3794ec7b4c81488babc6f1fb869abd68` |91/139; setup open fails before new checks.|
| `00da7a0624254706a6b093b590936fa0` | `576a69eeab3c44b79929d8e44fd785e9` |91/139; stage trace localizes copied-reference open.|
| `e12ad307e6ab4831b754a8d49eaf0ea2` | `d24add78d19d48a283a4c7292a546706` |100/142; saved custom fixture cannot reliably reopen; incomplete.|
| `20d41f06de6a4c29800e89774c6e5072` | `186a3b1c8a6c48a09ec4779eee4f4c6d` |47/9; isolated diagnostic confirms valid XLSB container, not reliable Excel reopen.|
| `5d8fd822d26b47c49085da0317a0c57c` | `c32a5278edf545b98f63c9f7a201d4d5` |40/2; untouched Seed permits nonempty query; hash reader lacks file sharing.|
| `44ed2496a44942e5a86554f9797590e1` | `0cb26ba4a603490e8caf9fb82492a714` |47/9; shared hashing exposes8 cold-query failures; dirty setup index failure.|
| `926d2604a0d5461e9eae9d40411284f4` | `4949e2c8d6c349aa935bc4dc1e8a598d` |47/9; stage trace identifies PowerShell named-column lookup after rename.|
| `e02d55756327406a932e471f650e4a3a` | `d78a2f28907449ffbe4e7e030cef0eaa` |79/33; typed staging reaches cold/dirty/missing checks, then invalid failure injection stops.|

Typed VBA staging replaces the problematic COM column setup, expected values are
captured in memory, and shared-read hashing permits observation while Excel owns
the source. These are fixture corrections, not runtime repair claims. Copied or
saved custom reference reopening is not required by this fixture and its exact
failure cause remains unresolved. Do not reintroduce that setup or repin after
a query to hide a write.

The last failed run injected an unhandled `Err.Raise` across the Domain bridge,
which showed a four-button VBA dialog. An exact owned-fixture End action was
posted at14:36:46.8161692 UTC; Excel crashed at14:36:47.1928147 (Application1000,
EXCEL16.0.20326.20158, oleaut32.dll10.0.26100.9278, c0000005, offset33b41), followed
by1001 at14:36:54.1811960. This is temporal evidence, not a native supplying-stack
diagnosis. A new recovery process had zero workbooks and was quit explicitly at
14:38:25.5567473; outer settings/package restoration completed at14:38:28.3515742.
The run is assisted and incomplete regardless of its closure booleans. Sanitized
`application-failures.json`, `assisted-fixture-end-retry.json` and
`recovery-cleanup.json` are retained in its controller root. Earlier cleanup helper
compile/binding failures performed no verified dismissal. No desktop error5 was
observed. Do not repeat the unhandled exception or use this event as crash repair.

## Candidate and remaining gates

Only `modInventoryDomainBridge` changes: the four query wrappers share a typed
read-only resolution/lifetime helper; write/bootstrap resolution is unchanged.
The module has956 lines, below the1000-line module limit. No Domain query result,
control, catalog, schema or display contract changes.

Isolated package root: `deploy/validation-inventory-query-readonly`. All five
packages build, explicitly compile and pass the cold-start dependency check,
14:41:23.1277996--14:42:00.3915653 UTC, with preserved settings/frozen packages and
Excel closed. Evidence: `reports/runtime/inventory-query-build`. Comparison of262
compiled components finds only Core `modInventoryDomainBridge` changed; the other
261 retain normalized code and string-literal hashes.

Combined public-launch/query results retain all232 Close identities and all80
query identities in exact relative order:175 PASS/137 FAIL across312 checks,
14:42:27.4135985--14:45:50.7467962 UTC. All80 query cases are GREEN, as are all42
prior checks and the public launch/close/reopen authority-byte assertion. The
only137 failures are the still-unimplemented Close observations/catalog21; this
is D3 query GREEN, not Close GREEN. Five instrumented compiles pass. Settings and
packages are preserved, closure is unassisted and the delayed Application audit
finds zero Excel failures. Controller:
`reports/runtime/production-close-controller/931858bcc50e491b8dd1f1e2dbc167db`;
result `reports/runtime/slice4be-production-close/9ad558db450c43179ce406d6fb73e425/red.json`.
The retained trace now contains only Picker.Query; all14 authority pins remain
unchanged through the public chain. `query-green-verification.json` retains the
ordered comparisons, and `visible-review.json` hashes four directly reviewed
images. Close/native X remain readable, the public reopened form retains its
fixture owner/explicit Designs-disabled status, and Settings distinguishes saved
configuration from unavailable optional tracking. Images prove visible views and
reopening, not disposal itself or comprehensive Production visual acceptance.

Static evidence: `reports/runtime/inventory-query-static`,269 components,
6109 procedures,134383 lines;9 literal/45 unresolved Application.Run calls,
190 duplicate-body candidates,28 non-growing oversized caps,3 schemas and336
PowerShell files parsed without errors. Only one procedure and24 source lines
are added; no metric exception is needed.

Packaged smoke retains86/86 exact prior Detail-candidate check identities,
14:48:03.2615387--14:48:25.4999974 UTC. Controller:
`reports/runtime/inventory-query-regression/smoke-3bcb1f95c8214b7ab3be175ac947f45a`;
shutdown evidence `reports/runtime/packaged-smoke-closure/72c0b7e003cb4ae9820611297d4062ad`.
Initial/final shutdown is unassisted, no termination is requested, settings,
packages and the tracked report are preserved, and the delayed Excel audit is zero.

Full Release1 chain retains32/32, live roles48/48 and Create Warehouse15/15 in
exact prior Detail-candidate order,14:49:33.5238605--14:54:40.5219691 UTC.
Controller `reports/runtime/inventory-query-regression/chain-e62bd93bea7844c7aa0584d2f3d8d85d`
retains the three reports before restoring tracked originals. The worker and
Quit-only controller finish without intervention; settings/packages/reports are
preserved and the delayed Excel Application audit is zero. No promotion is implied.

Packaged Production layout retains the exact prior geometry report across three
sizes and five pages,14:55:14.3757816--14:55:28.9529096 UTC. Controller:
`reports/runtime/inventory-query-regression/layout-08d3e4f5cb5d4c4e9ec6b160909e2cdf`.
Settings/packages are preserved, closure is normal and the delayed Excel audit is
zero. Three captured Run List views are directly reviewed and hashed in
`visible-review.json`: sections, scrolling and the Close footer remain reachable.
These blank fixture images do not establish a populated Production run.

Full reusable Production retains both aggregate checks and all171 boolean
observations in exact catalog20 order,15:01:49.2770848--15:10:21.0937474 UTC.
Controller `reports/runtime/inventory-query-regression/reusablefull-9cb05c1c7820481e9037d62f83d977dc`.
Both restart and final shutdown are unassisted, with zero reference-release
failures or termination requests. Final workbook closure and reverse-order
package closure pass. Settings/packages are preserved, Excel is closed and the
delayed Application audit finds zero Excel failures. `verification.json` retains
the exact assertion comparison and cleanup checks. This does not establish a
repair for the historical intermittent native fault.

Production lifecycle retains615/615 exact ordered prior checks and five
instrumented compiles,15:10:49.8266356--15:16:39.2489747 UTC. Outer controller:
`reports/runtime/inventory-query-regression/lifecycle-8a82fe752df14f509e977068179d04d7`;
inner controller `reports/runtime/production-lifecycle-controller/99a38b54efcb40069241aea70e34b329`;
result `reports/runtime/slice4be-production-designer/e24c92f6a86547a5953eef34a53ab1e9/green.json`.
The gate completes without intervention, settings/packages are preserved,
Excel is closed and the delayed Application audit is zero. The inner
`verification.json` records the outer interval and ordered comparison.

Applicable remaining Production regressions, beginning with the390-check draft
and Action Path gate, remain pending. Production Close is still unimplemented.
No transport-exception or historical native-crash repair is asserted by the
Domain Empty-result test.
