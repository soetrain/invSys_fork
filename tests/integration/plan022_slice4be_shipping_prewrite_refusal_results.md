# Slice 4be.1 Shipping pre-write identity exclusion

The focused correction is packaged GREEN: actual Shipping Add excludes a Core-
allocated identity when neither submission route attempted a write. Release and
comprehensive Slice4be acceptance remain open. D8-A is still pending approval.

## Contract and D13

Architecture v4.11 D18 already prohibits preallocated-but-unsubmitted references.
Normative/Plan/controls commit `590d41a` records the primitive submission-entry
refinement before runtime implementation; it changes no permission, stored
reference state, business Boolean, event-ID allocation or fallback identity reuse.

The protecting test enters the actual packaged Add handler. Unsaved Core probes
refuse both routes after Core allocates the ID and before either writer. Counters
at the real server/local writer entries first calibrate against the four existing
acknowledgment scenarios. The refusal then proves two public route entries, two
post-allocation refusals, zero writer entries, no accepted route, and an actual
failed owner result retaining its allocated ID. Authority/submission files retain
their names and bytes; captured workbook, unknown headers and unrelated workbook
checks remain independent of activity assertions.

| Evidence | Result |
|---|---|
| Initial probe installation | 280 PASS / 1 harness failure; no behavioral RED |
| Guarded pre-fix candidate (`ec8110a`) | 401 PASS / 1 FAIL across402 checks |
| Corrected candidate | 402 PASS / 0 FAIL; every prior identity/GREEN retained, no duplicates |

The sole meaningful RED was
`Shipping.Submission.PrewriteRefusal.Activity.ExactAppliedSourceReferences`:
FAILED included an Unknown reference even though neither writer was entered.
The initial installation failure was corrected using guarded case-insensitive
VBE source anchors. Both old and extended local-writer calls remain supported by
the same probe; unavailable anchors fail setup explicitly rather than producing
false product evidence.

## Implementation

Core `QueuePayloadEventServer` and `QueuePayloadEventCurrent` expose optional
ByRef Boolean `writeAttemptedOut`, initialized False per call. Core sets it at
the inbox-row write or serialized local-row append, after preparatory rejection
paths. Local JSON serialization retains its original position after opening the
append file and before writing. Existing Boolean results and EventIds remain.

Shipping's private server-first route keeps separate server/fallback facts,
including exceptional exits. `QueueObservedShippingPayload` records accepted
identities as before, retains uncertain attempted writes, and excludes identities
when neither route accepted nor attempted a write. Its original error envelope
is preserved. No activity is added to direct service or automatic-sync calls.

Core's three existing JSON helpers move into `modRoleEventJson`; two existing
Shipping timing/report helpers move into `modShippingReportText`. All five bodies
match their predecessors exactly apart from required procedure visibility.
Calls remain typed inside their package; no Application.Run is added.

The extracted JSON helper is also required by the 16 source-built harnesses that
explicitly import `modRoleEventWriter`. Their fixed import lists now include
`modRoleEventJson`; all 16 scripts parse. The first full-chain attempt stalled in
Create Warehouse with an observed compile-error dialog because that dependency
was missing. Its four workbooks were verified as disposable harness fixtures
before resetting the interrupted test. This is a setup failure, not a behavioral
RED or evidence that the packaged candidate failed to compile. The initial
full-chain report retains three PASS rows and one `Harness.Exception`.
After repair, Create Warehouse produces a fresh15/15 result. The full-chain retry
then ends at four PASS rows and one `Harness.Exception` with RPC0x800706BE, before
`OrderedLiveProcessCompleted` is reached. It does not provide full-chain acceptance.

## Package and maintenance evidence

Isolated candidate: `deploy/validation-shipping-prewrite-facts`. All five builds
and explicit compiles pass, including Operations cold start. Compiled comparison
finds the intended Core writer and Shipping owner/report changes plus the new Core
JSON module. A separate read-only comparison explains the additional form hash:
old2928/new2924 code lines are identical after trimming trailing whitespace. Both
package files retain their bytes; no form statements changed in this correction.
Inventory/Designs Domain, Admin, Receiving and Production compiled source remains
unchanged.

Maintenance:183 components,5527 procedures. All28 previous oversized module
limits pass (Core writer3035 versus3064; Shipping main22386 versus22398).
Literal Application.Run8, unresolved calls45 and duplicate-body candidates195
remain unchanged; all three static JSON schemas pass. No Excel Application1000
fault is observed in the focused RED/GREEN log windows with three seconds' margin;
this does not establish a general native-crash repair.
Final source formatting removes six trailing blank lines from the Core writer.
The pre-edit source pin was verified; comparison proves equality after trimming
trailing whitespace. The updated source pin and proof are recorded explicitly in
`prewrite-facts-writer-whitespace-proof.json`; candidate package bytes remain
unchanged and its compiled executable statements remain the tested statements.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-owner-activity -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity -ShippingSubmissionOnly
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-prewrite-facts -Phase GREEN -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity -ShippingSubmissionOnly
```

## Exact evidence and open gates

Ignored evidence under `reports/runtime/slice4be-shipping-activity/`:

- `cc46083de87a4363a1595e3e34af3334/diagnostic-submission-red.json` and
  `1f5bdc8331d44be2b3971d8a501187ab/diagnostic-submission-green.json`;
  `prewrite-refusal-{red,green}-summary.json` retain scope/comparison.
- `prewrite-refusal-red.log` retains the anchor failure;
  `prewrite-refusal-guarded-red.log`, `prewrite-refusal-green.log` and
  `prewrite-refusal-native-windows.json` distinguish meaningful tests and faults.
- `prewrite-facts-{build,compile,static}.log`, compiled source/comparison,
  extraction proof, module limits, package/source hashes and form comparison.

The full prepared-fixture Shipping route completes **974 checks: 967 PASS /
7 FAIL**, exit1. All failures are the pending D8-A missing-Auth recreation
findings. Every check and GREEN from both890-check baselines and the953-check
expanded RED remains, with no duplicates. No Excel fault is observed in this run
window. This is the first completed expanded candidate route; it does not erase
the preceding candidates' native failures or treat D8-A as approved.

Exact result:
`e621c5eaf97e4b6cbf492f95f888ac11/diagnostic-prepared-fixtures-diagnostic-shipping-first-green.json`;
`prewrite-facts-full-shipping-comparison.json` and
`prewrite-facts-full-shipping-native-window.json` retain comparison/window scope.
Packaged smoke passes86/86; Viewer passes; independent reusable Production and
clean-process restart pass2/2; Shipping layout passes1/1. Separate live roles
stop at39 PASS/1 FAIL during Production Complete Run. Combined launchers stop
at1 PASS/1 harness RED after Receiving. Both failing windows contain an Excel
Application1000 fault (`ntdll.dll`, `c0000028`); they remain acceptance failures.
Both full-chain attempts and the import repair are described above. The repaired
full-chain RPC failure also coincides with an Excel Application1000 fault.

Receiving's full candidate attempt stops at621 checks:591 PASS/30 FAIL. The
preserved845/845 baseline comparison finds225 unexecuted identities,29 lost
GREENs and no duplicates; none is counted as preserved acceptance. All29 failed
existing checks are `SupportedCatalogRead`: the shared assertion still allowed
only catalogs3-7, while approved catalog8 is now emitted. The remaining failure
is a keyboard-input harness exception at `RECEIVING_PAGE_RECEIPTS`, index1.
The staging and navigation assertions now explicitly include8, preserving
minimum-version/field checks, unsupported-version rejection and actual Core
reads. No Receiving runtime changes. Focused navigation then stops at106 PASS/
1 harness failure at the same first keyboard action. Neither Receiving window
contains an observed Excel Application1000 fault. Catalog-read GREEN and full
845 preservation are still pending; the assertion correction is not evidence of
either. Next calibration: prove keyboard focus belongs to the intended form and
control after the unrelated workbook is activated. The existing native helper
checks Excel-process ownership only; this is a suspected targeting gap, not a
confirmed cause or an approved runtime workaround.

Ignored evidence: `prewrite-receiving-prior-green.json`,
`prewrite-receiving-before-catalog-repair.json`,
`prewrite-receiving-incomplete-comparison.json`,
`prewrite-receiving-navigation-catalog-eight.json`,
`prewrite-followup-exits.json`, and
`prewrite-facts-completed-gate-native-windows.json`.
Final integrity verification passes65 package pins and16 source pins with Excel
closed. Accepted deployments and unrelated user changes remain untouched.
Partial/mixed multi-source outcomes, additional failure/policy cases,
remaining Operations/Admin controls, comprehensive Viewer/Settings/Action Paths,
visible operator comparison and human acceptance remain required. The earlier
native faults and pending D8-A findings are not erased by this focused correction.
See [the owner checkpoint](plan022_slice4be_shipping_owner_activity_results.md)
for the prior candidate's gates and remaining D14 TSV compatibility defect.
