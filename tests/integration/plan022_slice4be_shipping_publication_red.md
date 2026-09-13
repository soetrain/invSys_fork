# Slice 4be.3 Shipping publication RED

Last verified 2026-09-13. Release 1 and Slice 4be remain active. This checkpoint
extends the actual Admin publication test; it changes no runtime VBA or XLAM.
The Events artifact, comprehensive publication, Viewer paging, Action Paths and
physical/human acceptance remain unimplemented or unproven.

## Contract and fixture

Architecture v4.11 D18 names ShippingBOM and ShippingHolds as separate current-
state sources, with exact owner fields and Warehouse versus Station profile
coverage. Plan 022 and controls v1.114 carry that same test scope. This is a
semantic-inheritance refinement, not permission to invent historical events or
give Viewer a canonical-read fallback.

`Slice4beShippingPublicationFixture.ps1` reuses the existing form-action facades
from `Slice4beShippingActivity.ps1`, selecting their form declaration through
PowerShell AST. It does not install the latter suite's interruption/submission
probes. The existing public Shipping launcher establishes the captured workbook;
actual Box Designer/Box Maker handlers create ten boxes from two distinct
components. Actual Add and Hold handlers prepare a held line. The fixture starts
with Admin Generate Warehouse and Seed, and credentials stay in disposable
fixture memory/files under the existing harness cleanup.

The owner-created BOM rows and persisted hold record are calibrated before
publication. An unknown leading BOM column exercises field exclusion. The
generated Inventory handle is closed without saving, then reopened read-only to
verify the owner's box-build event was already durable. The harness never saves
unsaved Inventory merely to make this assertion pass. Publication subsequently
opens a transient read-only source, preserving the prior source-read test scope.

The controlled 5,001-group Inventory source is copied before Shipping setup can
publish its inventory snapshot. Activity counts come from all actual records
after owner setup, rather than assuming only the original Settings pair exists.
The Settings pair's exact-ID assertions remain unchanged.

## Packaged evidence

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-designs-publication-source -Phase RED -CheckViewerPublication -ViewerPublicationOnly -CaptureEvidence
```

The final run records **26 PASS / 21 FAIL**, 47 distinct check identities and
**no harness failure**. All 36 previous publication check identities remain;
every previous 21 GREEN remains GREEN. Designs retains 15/15. The four new
fixture checks prove real Box Designer/Box Maker, Add/Hold, persisted owner rows
and an already-durable box-build event. Shipping local-source byte preservation
also passes, as do the prior source-copy/canonical/activity preservation checks.

Six new assertions fail for the absent Events artifact: retained exact owner
lines, current-state classification without invented history, and available
coverage/scope, separately for BOM and Holds. The previous fifteen artifact
failures remain. A helper-only in-memory calibration accepts absent historical
fields and rejects an invented historical SourceId; that is harness validation,
not packaged acceptance. Publication-only diagnosis cannot claim GREEN.

The final Shipping capture was inspected. It shows the held line under Not
Shipped and a pending sync indication. It does not prove a completed shipment,
reservation application, diagnostic conclusion or human acceptance.

Four earlier setup attempts are retained, not counted as product RED:

| Run suffix | PASS / FAIL | Setup failure |
|---|---:|---|
| `a1d8f45bf1d14d4c8e8abfde05a33a3d` | 20 / 15 | Helper check output contaminated its returned object; owning setup also left Inventory open, invalidating transient-source preconditions. |
| `be5acccce5274e0baf6ccba94968cf64` | 18 / 1 | Misspelled capture-helper call. |
| `3df28614069c4364894f6c9bf68641b8` | 19 / 1 | Generated Inventory was unsaved; replaced the setup assumption with close-without-save plus a durable owner-event check. |
| `63acef375e814a84b04aea41d1aab9cb` | 23 / 16 | PowerShell unwrapped the one-row Hold expectation; explicit array capture corrects its shape. |

Each earlier attempt contains a harness failure. No runtime repair was inferred
from them. The successful fixture run cleans all four exact, absent-before local
file targets. Two generated local files from the first failed attempt remain:
automatic approval review rejected their proposed cleanup with only `blocked by
policy`. They are isolated fixture state, not an acceptance blocker; no further
removal of those two files or approval bypass was attempted after that rejection.

## Preservation and remaining gates

All **175 package pins** and **16 protected source pins** match. Excel is closed,
both changed PowerShell files parse, and document references/diffs are checked.
No runtime/build/layout/static source changed, so no new full build, compile,
static regeneration, live-role, Receiving or Release 1 chain GREEN is claimed.
Their last verified scope remains in the
[Designs source checkpoint](plan022_slice4be_designs_publication_source_results.md).
Accepted deployment/NAS and unrelated handoff 067/critique 023 remain untouched.

The next implementation must use a declared owning Shipping read boundary for
Core publication, retain full rows and truthful unavailable coverage, and create
the validated atomic Events artifact. The legacy flattened Viewer supplement
cannot satisfy these assertions. Complete publication and grouped Viewer tests,
focused GREEN and all required regression/visible gates remain necessary.

Ignored evidence is under `reports/runtime/slice4be-viewer-groups/` with final
run `21be2226017f40c6b3f5c138e7f0b4ac`: `diagnostic-publication-red.json`,
`publication-source-preservation.json`, `shipping-fixture-owned-files.json` and
`publication-shipping-held.png`. Fixed-result verification and preservation logs
are `reports/runtime/shipping-publication-fixture-verification.json` and
`reports/runtime/shipping-publication-fixture-preservation.log`. No runtime rows,
credential material or local source paths are included in this maintained record.
