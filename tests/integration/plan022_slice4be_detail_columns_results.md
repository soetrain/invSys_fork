# Slice 4be Admin Event Detail heading alignment

Last verified:2026-09-30 UTC. Focused RED/GREEN verified; broader gates pending.
No deployment promotion or human acceptance.

Architecture v4.11's 4be.2 heading-alignment clarification, Plan022 and controls
1.321 were committed as documentation0e97970 before tests12f56ce and runtime edits.
This applies D13 visual acceptance and D18 Settings semantics without changing
configuration ownership, permissions, schemas, tracking controls or activity.

The existing single space-padded heading is visibly offset from the actual list
columns in both catalog19 and catalog20 Settings images. The focused probe measures
rendered heading starts using the live font and current column widths. It accepts
three-point alignment tolerance, fitting text and vertical separation from the
list. It handles a combined caption as well as separate headings; it does not
require a particular implementation. Default, enlarged and restored sizes are
tested through the packaged form and the actual `modAdmin.Open_Settings` launcher.
The public modal activation adapter measures geometry and invokes the real Close
handler. All adapters remain unsaved in disposable packages.

## RED

`Test-Slice4beDetailColumns.ps1 -DeployRoot deploy/validation-production-regulation-final -Phase RED`

222 PASS/18 FAIL across240 unique checks,2026-09-30
12:59:57.3063956--13:04:14.5813792 UTC. All202 prior Settings checks retain their
exact relative order and remain GREEN. The only failures are Show/Order/Required
alignment at three sizes through each of the two entry routes. Field alignment,
hidden identity, staged-request and saved-configuration preservation pass.
Five instrumented compiles pass; there are no harness failures. Settings/package
pins are preserved, Excel closes normally and the delayed Application audit
finds zero Excel failures. Runtime still matches3559e2d at RED.

Controller: `reports/runtime/detail-columns-controller/80c27ff6725f4d8ea38b02598c616ec6`.
Result: `reports/runtime/slice4be-detail-columns/01147446c23348b782f2bd8ecadd375a/red.json`.
Three dedicated images are directly reviewed and hashed in `visible-review.json`.

## Candidate

Only `cAdminEventDetail` changes: four labels obtain their positions and widths
from the existing list columns. The field identity remains hidden; captions,
row values, field order, validation, profile actions and observations are unchanged.
No new catalog version or control action is introduced.

Isolated candidate: `deploy/validation-detail-columns`; build evidence:
`reports/runtime/detail-columns-build`. Build/explicit compile/cold load passes
for all five packages,13:05:21.6767902--13:06:09.3091323 UTC, with preserved Settings
and frozen candidates. All262 compiled components retain identity; only the Admin
`cAdminEventDetail` normalized code differs. The other261 components retain both
normalized code and string-literal hashes. This is not a completed GREEN claim.

## Focused GREEN and maintenance

The same test on `deploy/validation-detail-columns` passes240/240, retaining all
RED identities in exact order and all202 prior Settings checks. Interval:
2026-09-30 13:06:28.6782872--13:10:54.3534259 UTC. Controller:
`reports/runtime/detail-columns-controller/8e07b0b7adc145a3a3b6e599be4ed9d7`;
result `reports/runtime/slice4be-detail-columns/971d16012eb844309560b3125e00e300/green.json`.
Five instrumented compiles, settings/package preservation, unassisted preference
restart with all ten references released, normal final closure and zero delayed
Excel Application failures pass. Three corrected images are directly reviewed
and hashed: headings are readable above their columns at all three sizes; the
existing list-cell text inset is retained. This is bounded agent visual review,
not human acceptance or a claim about every Settings layout detail.

Static evidence: `reports/runtime/detail-columns-static`. Metrics are269 components,
6108 procedures and134359 lines (+one procedure/+12 lines); literal/unresolved
Application.Run remains9/45 and duplicate groups190. All28 oversized caps are
non-growing, three report schemas and332 PowerShell parses pass. The first
verification invocation preceded generation completion and found a missing
manifest; it is not product RED. Verification after generation passes.

Packaged smoke retains86/86 exact prior checks,13:11:27.5870608--13:11:50.0814969
UTC. Controller `reports/runtime/detail-columns-regression/smoke-4ff10b93e7404d9682a7658cc2dbf864`;
normal Initial/Final shutdown evidence is in
`reports/runtime/packaged-smoke-closure/516a9639b60a47af95d443c3e2c29c9b`.
Settings/packages/tracked report are preserved, no termination is requested and
the delayed Application audit finds zero Excel failures.

## Remaining gates

Complete Settings activity and Release1
chain/live-role regressions for the isolated candidate. Preserve all earlier
regulation evidence and unrelated user changes. Human/NAS acceptance and remaining
comprehensive control coverage remain open.
