# Plan 022 Slice 4be.1 Shipping captured-context repair

Last verified: 2026-09-12. This is a candidate under approved Architecture v4.11
D18, not complete Shipping activity, Slice4be or Release1 acceptance. Normative
clarification **089ab0d** was committed/pushed before runtime edits, with Plan022
and controls v1.92 synchronized. The clarification inherits captured-context and
workbook-preservation rules; it adds no permission, activity catalog or authority.

## Protecting test and RED

`Test-Slice4beConfigCommands.ps1 -CheckActivityEvidence -CheckActivityFoundation
-CheckShippingActivity` retains the real-owner eight-action sequence and its exact
key/source/application checks. `Slice4beShippingContext.ps1` then prepares genuine
active/held lines. A separate probe mode counts entry at existing owners and stops
there: each healthy handler must reach its owner once, while stale handlers must
stop earlier with a context notice. Probe results do not prove business effects;
the preserved normal sequence supplies that evidence. State and source-byte checks
verify the probe does not mutate business data.

- Initial matrix: **136 PASS / 50 FAIL**, including a harness failure. A picker
  threshold demanded five available units for a two-unit Add; this is not matrix RED.
- Corrected matrix: **242 PASS / 91 FAIL**. All seven healthy calibrations and all
  eight timing-status checks pass. Each of21 signed-out/reauthenticated/changed-
  target cases fails pre-owner rejection and its notice. Earlier49 failures remain.
- Explicit public relaunch: **246 PASS / 92 FAIL**. Valid reuse passes, but the
  stale form is also reused. Workbook/staging preservation and next-launch reuse pass.
- Candidate: **292 PASS / 46 FAIL**. All46 formerly failing context/relaunch
  assertions pass; no non-activity failure, lost previous GREEN check, missing check
  or duplicate exists. All46 remaining failures are missing Shipping activity.

The successful guard comparison is focused GREEN, not a passing full activity
suite. The suite retains exit1 and every unresolved activity assertion.

Ignored evidence root: `reports/runtime/slice4be-shipping-activity/`:

| Log | JSON run/report | Result |
|---|---|---|
| `context-matrix-red.log` | `9c82a4aedd264cffb5616e185e39d94f/red.json` | 136/50 harness failure |
| `context-matrix-calibrated-red.log` | `36747aa3361a4a39954ec43fe7421ca0/red.json` | 242/91 RED |
| `context-reopen-red.log` | `ce9fdc23667e403095d4857e1b85a3c2/red.json` | 246/92 RED |
| `context-guard-green.log` | `35cea1eb29e14fb59ba2a3e436bed660/green.json` | 292/46; guard GREEN |

## Candidate implementation and gates

`frmShipmentsTally` captures its trusted context once and guards all seven mutation
controls before owner entry, including after pending-status UI yields and between
multi-row Remove owner calls. Rejection stops automatic synchronization and shows
the fixed notice. `modShippingFormContext` validates the captured live workbook
and session; the explicit launcher replaces stale forms through its bounded factory,
preserving active/held staging on that same workbook. Valid reuse and ordinary Close
remain. Core authorization and business owner contracts are unchanged.

`cShippingActionTimer` receives the form's existing timing behavior, protecting the
form size limit while retaining all eight normal-action timing status checks.
The isolated candidate is `deploy/validation-shipping-context-guard`. All five
packages build; all five explicit compiles and cold-start Operations dependency
checks pass. Build/compile logs, exported source report and five SHA-256 pins are
`context-guard-build.log`, `context-guard-compile.log`,
`context-guard-compiled-source.json` and `context-guard-package-hashes.json`.

Static JSON contracts pass at177 components/5511 procedures/123758 lines. Literal
Application.Run8, unresolved dynamic calls45 and duplicate-body candidates195
are unchanged. All28 existing module limits hold: the form shrinks to3022 lines,
the launcher module to22443. The net73 runtime lines add the context/lifetime
helper and typed timing class; no size-limit exception is taken.

Candidate gates completed on the same pinned set:

- Packaged smoke **86/86** (`context-guard-smoke.log`).
- Live-role workflows **48/48** (`context-guard-live-role.log`).
- Ordered full Release1 chain **30/30**, including restart/reconciliation and
  source/static gates (`context-guard-full-chain.log`).
- Viewer launch/reuse/export/filter/read-only contract passes (`context-guard-viewer.log`).
- Combined public launchers **3/3** (`context-guard-launchers/packaged-launcher-noeligible.md`).
- Shipping layout/identity **1/1**, including fixed status anchor, grow/resize,
  headers, exact key, and no Boxing page overlap (`context-guard-shipping-layout/shipping-layout.md`).
- Full reusable Production and clean-Excel restart **2/2**, with no reduced test
  flags (`context-guard-production/production-reusable-production.md`).

Windows records Excel `c0000005` in `combase.dll` at2026-09-12T23:58:10Z, within
the full-chain window23:55:31Z-00:00:14Z. The30 passing business checks do not prove
native stability or repair the earlier unexplained native faults. The sanitized
module/code/time record is `context-guard-native-faults.json`; no causal claim is made.

The visible rerun also records292 PASS / 46 missing-activity FAIL, preserving
all338 check identities and all292 prior GREEN checks with no duplicates or other
failure. Its report is `0a955d0e608744df9358934fe91843d7/green.json`; the capture
`shipping-stale-session.png` is in that same run directory. The log is
`context-guard-visible.log` at the runtime evidence root.
The inspected capture shows the fixed session notice legibly. It also exposes
crowding at the System Key editor/Add button and adjacent table headers; the
layout code is unchanged, but baseline comparison is still required before
classifying that visible finding or claiming full form-layout/human acceptance.

All seven gate-queue stages and the visible run are terminal; Excel is closed.
All30 earlier package pins/four earlier source pins and the candidate's five
package/four Shipping source pins match. No accepted deployment or NAS runtime
is changed. Raw generated reports are retained only under ignored runtime paths.
Remaining activity
source/outcome definitions, capability-loss, during-yield/closed-workbook and automatic-sync probes,
broader Operations/Admin coverage, Settings, Viewer, recording, guides/comparison,
native stability and human/physical acceptance remain required.
