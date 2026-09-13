# Slice 4be.1 Shipping owner activity

Status: implementation candidate; expanded validation and Release 1 acceptance
remain in progress. Architecture v4.11 D18 governs this work. D8-A is still a
pending proposal; this implementation does not change Auth provisioning.

## Contract and implementation

The seven catalog-8 Shipping controls observe their real form handlers. A valid
captured context permits REQUESTED before validation/authorization; the existing
fresh permission and workbook/session guards still govern owner entry. Optional
tracking failure adds its fixed notice without changing authorized business work.
Direct service calls and automatic synchronization do not impersonate controls.

Operations-local `cShippingOwnerFacts` receives exact identities and acceptance
facts after actual Shipping submission boundaries. The per-action set retains
Submitted and Unknown identities separately. Known required-step failures cannot
be promoted from the owner's generic Boolean to a clean activity outcome. The
form-owned `cShippingActivity` uses the existing typed Core observation/reference
API. Core/Domain source, business Boolean contracts and submission behavior are
unchanged. Cross-package calls remain declared primitives; no Application.Run is
added. Hold/Return are local actions with no source references.

Changes are confined to the Shipping form, posting/owner modules, two small fact/
observation classes, and two helpers extracted from existing Shipping procedures.
All 16 extracted helper bodies match their predecessors exactly after changing
Private to Public. No scanner candidate is deleted on scanner authority.

## D13 and retained evidence

The pre-implementation protecting run is **890 checks: 810 PASS / 80 FAIL**,
recorded in [exact outcome evidence](plan022_slice4be_shipping_exact_outcomes_results.md).
It preserves all863 preceding identities and GREENs. Seventy-three failures
concern missing normal Shipping observations; seven concern pending D8-A.

The first owner candidate completed the same890 checks at **873 PASS / 17 FAIL**.
All73 missing-observation assertions passed. Ten old assertions failed because
they rejected a REQUESTED record written before later sign-out or rejected the
independently required tracking notice following the fixed access notice. The
original result remains evidence; it is not retrospectively called full GREEN.

Normative/Plan/controls commit `71d3faf` explicitly clarifies those existing D18
requirements. Tests retain the ten original check identities and add independent
proof: record contents captured immediately before actual sign-out must be
unchanged afterward, with no late/new-context record; Config failure must preserve
the exact primary access message and exact independent tracking notice.

The expanded route also exercises real Add with four acknowledgment failures,
including accepted local fallback and uncertain acceptance. It compares exact
activity references with independently observed owner/submission identities and
durable submission rows, and proves submission is not yet Domain application.
Physical optional-store failure exercises actual Add/Remove, their owner effects,
captured binding, unknown headers, unrelated bytes and prior activity preservation.

## Fixture recovery and expanded RED

Earlier expanded runs stopped at14 PASS/9 FAIL on the prior package and2 PASS/
1 harness failure on the candidate. A traced candidate failure stopped inside
bootstrap Seed before operator creation. These native failures are not product
RED, do not prove a runtime cause and remain unresolved evidence.

The explicit `-PrepareShippingFixturesBeforeProbesForTest` mode requires the full
Shipping-first route. It generates the same three disposable Shipping fixtures
through actual Admin bootstrap before Shipping VBE probes, then consumes each
once at its existing use site. Ordinary bootstrap Seed, later explicit Admin Seed
and every handler/assertion remain. The mode checks actual generation roots and
template preservation, restores the normal operator root, and reestablishes each
consumed fixture's Core data root. Root evidence distinguishes generation from
later use; result names identify the diagnostic ordering.

The prior catalog-8 package completed **953 checks: 828 PASS / 125 FAIL**, exit1.
All890 preceding check identities remain, with63 additions and no duplicates.
The only lost prior GREEN was the accepted-template hash assertion: preparation
used a contiguous SHA-256 string while its consumer used a hyphenated digest.
An independent same-file calibration proves identical digest bytes and different
formatting. Preparation now uses the same existing hash function as the consumer.
This is a harness correction, not product RED. The125 failures partition into117
missing activity assertions, seven pending D8-A findings and that one hash-format
assertion. The run is retained without changing its recorded counts.

The candidate's prepared-fixture run then stopped at **5 PASS / 1 harness
failure** in explicit Admin Seed. A bounded experiment moving that same Seed
before Shipping probes stopped at **1 PASS / 1 harness failure** in Seed. Both
windows record ntdll.dll/c0000028. The latter experiment was removed; it disproves
treating Shipping-probe order as a reliable cause or repair. The retained optional
mode is the one that completed953 prior-package checks. Both failures remain;
neither supplies expanded candidate GREEN. The first left an empty recovery
instance: zero workbooks and matching process/window were verified before Quit,
then the same process was stopped after its application disconnected but lingered.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-shipping-catalog-eight -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckShippingActivity -ShippingBeforeSharedFormsForTest -PrepareShippingFixturesBeforeProbesForTest
```

## Package and maintenance evidence

The isolated candidate is `deploy/validation-shipping-owner-activity`. All five
builds and explicit compiles pass, including Operations cold start. Exported
compiled-source comparison shows only the three intended Shipping components
changed and four Shipping helpers/classes added (171 to175 components, none
removed). A subsequent source-only trailing-blank-line trim is not represented as
a separate rebuild. Source/package pins retain the exact validation inputs.

Maintenance regeneration reports182 components,5527 procedures and124139 lines.
All28 preceding oversized-module limits pass: `frmShipmentsTally` is2938 lines
and `modTS_Shipments`22398. Literal Application.Run8, unresolved dynamic calls45
and duplicate-body candidates195 remain unchanged. All three static JSON schemas
pass. Static observations do not substitute for runtime or visible acceptance.

The candidate passes packaged smoke **86/86**, Viewer, combined launchers **3/3**,
independent reusable Production/restart **2/2**, and Shipping layout **1/1**.
Separate live-role validation stops **39/1** at Production Complete
Run with ntdll.dll/c0000028. The full-chain runner reports **30/30**, including
its ordered live-role report **48/48**, but an Excel combase.dll/c0000005 event
occurs in that run window. In addition, source inspection finds that the runner
does not check `$live.ExitCode` before accepting the ordered child's report.
The30 passing assertions are retained, but clean full-chain acceptance is **not
established**. Do not erase the native fault or infer its cause from counts.

`Test-Release1OrderedChildExit.ps1` executes the runner's actual ordered-child
decision statements and result parser, substituting only external process results.
Its passing-report case calibrates the fixture; nonzero/abnormal exits with the
same passing report reproduce the defect, and a missing-report case protects the
existing rejection. Before correction: **2 PASS / 2 FAIL**. The runner now records
`OrderedLiveProcessCompleted` from the child exit code; afterward: **4 PASS / 0 FAIL**.
No Excel, package, business or architecture change is involved in this gate fix.
The corrected packaged rerun completes **31 PASS / 0 FAIL**, exit0, retaining
all30 prior identities and adding only OrderedLiveProcessCompleted. Its window
again records Excel combase.dll/c0000005 (`2026-09-13T07:26:44.7401495Z`). The
child-exit assertion is now proven, but native-clean full-chain acceptance remains
open. This repeat does not justify changing business code or declaring a native
cause. Full Receiving845 and remaining expanded Shipping coverage were not rerun
in this checkpoint.

After all runs, Excel is closed and all60 package pins match, including the55
prior package pins and five candidate binaries. Four preserved Receiving/Shipping
source pins, three Core catalog pins and seven owner-implementation source pins
also match. The verifier explicitly replaces the old form/main-module expectations
with the already captured new source pins; it never rewrites mismatched hashes.

## Exact evidence and remaining work

Ignored evidence is under `reports/runtime/slice4be-shipping-activity/`:

- `owner-activity-build.log`, `owner-activity-compiled-source.json`,
  `owner-activity-compiled-comparison.json` and source/package hash sets.
- `68d61275da384f4494d3ac0453ce3a63/diagnostic-shipping-first-green.json`:
  original873/17 result; `owner-activity-initial-comparison.json` preserves its
  exact ten prior-check failures and73 newly passing assertions.
- `owner-activity-expanded-{red,green,traced-green}.log`: retained setup failures;
  `4dc1033b8ab94f3faaa2b6def1adc362/bootstrap-phases.log` ends at Seed.
- `6a166801c89d473487350c7ac63a5103/diagnostic-prepared-fixtures-diagnostic-shipping-first-red.json`:
  expanded828/125; `owner-activity-prepared-red-comparison.json` and
  `owner-activity-template-hash-calibration.json` distinguish the harness error.
- `owner-activity-extraction-proof.json`, `owner-activity-module-limits.json`
  and `owner-activity-prepared-static.log` retain proportional maintenance proof.
- `819bb563eb2c45498d6f21dfc23be77f/diagnostic-prepared-fixtures-diagnostic-shipping-first-green.json`
  records5/1; `97a682b96a6e4fb0b25e4f1fbe1101a1/diagnostic-prepared-fixtures-diagnostic-shipping-first-green.json`
  records the removed Seed-order experiment1/1. Logs are
  `owner-activity-prepared-green.log` and `owner-activity-prepared-seeded-green.log`.
- `owner-activity-prepared-native-windows.json`, `owner-activity-gate-exits.json`,
  `owner-activity-gate-native-windows.json` and individual gate logs/reports retain
  exact observed windows and results. Runtime reports remain ignored and are not
  copied into this sanitized record.
- `ordered-child-exit-red.json` and `ordered-child-exit-green.json` record the
  actual orchestration RED/GREEN; `owner-activity-ordered-live-checks.json` retains
  the original child48 passing check identities without runtime values.
- `owner-activity-checked-child-full-chain.log`, its `-results.md` and `-exit.json`
  companions, `owner-activity-checked-child-comparison.json` and
  `owner-activity-checked-child-native-window.json` record31/31 and the retained
  fault; `owner-activity-final-pin-verification.json` records60 package/14 source
  pins and closure. `owner-activity-checked-child-static.log` is final maintenance.

Expanded candidate GREEN remains pending. Partial/mixed multi-source failures,
multi-row owner failures, policy variations and remaining Shipping/Operations/
Admin control coverage still need focused handler evidence. Packaged smoke,
live-role, full Release 1 chain, Viewer/launcher/Receiving/reusable Production,
layout, saved-workbook restart and visible operator comparison/human acceptance
remain release gates. The reachable legacy Shipping TSV ROW compatibility path
also remains a separate D14 test-first correction. No full Slice4be or Release1
completion, native-crash repair, NAS rollout or human acceptance is claimed.

Source review also finds an unproven reference boundary: Core `QueueEventCore`
allocates `eventIdOut` before inbox resolution/read-only/schema checks and before
the first write. If server and fallback both fail before writing, the current
Shipping observer may label that allocated-only ID Unknown. D18 explicitly
forbids unsubmitted references. Next prove this through actual Add with both
routes refused before writes, preserving independent owner/storage evidence;
then distinguish owner-confirmed NotSubmitted from uncertain attempted writes.
Do not infer write entry merely from a nonempty identity or an outer queue call.
