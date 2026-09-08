# Slice 4be.1 Production provenance diagnostics

This investigates the unresolved Production gate in the
[Receiving denial checkpoint](plan022_slice4be_receiving_launcher_denial_results.md).
Architecture v4.11 D12/D13/D18 still governs acceptance. No checked-in runtime source,
architecture, canonical data contract or accepted deployment changes in this
investigation. Diagnostic successes do not waive the standard gate or establish
that a native failure has been repaired.

## Controlled observations

The standard full `ProductionReusable` harness invokes the public Production
launcher and its existing form-action tests, followed by a clean Excel restart.
It loads Core, Inventory Domain, Designs Domain and Operations; it does not load
Admin. All diagnostic directories contain five package files, but Admin's binary
cannot be attributed a causal role in this particular harness.

The read-only provenance variant checks pinned file hashes before Excel opens,
then checks each loaded package's actual directory and every invSys project
reference's resolved directory/broken status after initial loading and restart.
It does not compile, edit VBA, save packages or inspect operational row values.
Reports contain fixed package/reference names and booleans. Hashes are checked
again after Excel closes. These are preloaded dependency checks, not a separate
cold-start role-only acceptance test.

| Diagnostic | Result | Interpretation limit |
|---|---|---|
| Candidate Core + Operations, prior Domain packages; read-only provenance | 2/2, initial/restart provenance valid | Does not prove either Domain package causes the original failures |
| All four executing candidate packages relocated; same provenance | 2/2, initial/restart provenance valid | Negative control prevents attributing the prior result to a Domain substitution |
| Original candidate files at original location; same provenance | 2/2, initial/restart provenance valid | Relocation alone is not a sufficient explanation; inspection/timing/intermittency remains unresolved |
| Original candidate, one-second post-load pause without VBA project access | 0/1; native failure at batch scale | This pause alone does not reproduce the inspection variant's success |
| Copied candidate, resolved references and explicit package saves; plain harness | 0/1; native failure at batch scale | All 168 extracted components remain identical and five compiles/cold start pass; saving did not establish a durable repair |
| Original candidate, initial native observer | 2/2; no target fault captured | Inconclusive: observer could swallow non-target exceptions; not acceptance evidence |
| Original candidate, corrected native observer, v2 | 0/1; native failure at variable-quantity handler | Target fault captured; observer preserves non-target exceptions and detaches normally |
| Original candidate, preceding-exception capture, v3 | 0/1; native failure at variable-quantity handler | Captures preceding `0xc0000005`, followed by `0xc0000028`; exact VBA trigger remains unknown |
| Original candidate, unsaved fixed-stage instrumentation | 0/1; released Process edit/export test | Last fixed stage is ReleaseSource, after SaveSourceDraft; source instrumentation alters execution conditions and does not identify the underlying defect |
| Fresh copied candidate, all five projects forcibly recompiled and saved; plain harness | 1 PASS / 1 failure | Initial complete workflow passes; clean restart fails with RPC `0x800706BE`; compilation persistence is not a verified repair |
| Same compile-before-save candidate, fixed restart markers only | 2/2 | Teardown, new-session setup, public launcher and restart actions finish; timing/intermittency remains unresolved |

The planned individual Domain substitutions were not run after the all-current
control passed. That control retains the original executing Core/Domain/Operations
hashes; the copied prior Admin package is present but is not loaded. It is not an
accepted five-package replacement or a proposal to combine release contracts.

An initial provenance attempt tried hashing an XLAM while Excel held it open.
It stopped before any Production callback with a file-sharing error. Moving
hash verification to the pre-load boundary corrected the harness issue. This
failure is neither a Product regression nor meaningful D13 behavioral RED.

## Native exception evidence and limits

The bounded observer attaches only to the diagnostic harness's own Excel process,
after checking its process name and creation identity. It uses the installed
Windows debugging engine, disables engine output and network symbol paths, and
captures only exception codes plus allowlisted module names and relative frame
offsets. Unknown modules/frames are redacted. No dumps, arbitrary memory reads,
stack arguments, operational rows or source values are persisted.

Synthetic calibration first proved 11 checks, then review found that the observer
could resume unrelated exceptions as handled. A preceding handled exception was
added to the fixture; the corrected observer passes 13/13, including preservation
of that exception and normal fixture survival/detachment. The first Excel run is
therefore inconclusive even though the harness completed 2/2.

The v2 observer captured `0xc0000028` through ntdll and VBE7 frames, with one
preceding exception. A further test-first diagnostic extension produced 13 PASS /
2 FAIL for absent preceding-code/frame evidence, then 15/15 after adding that
bounded capture. The v3 trace records `0xc0000005` first: its first instruction
frame is unresolved, followed by oleaut32, VBE7 and Excel frames. The subsequent
`0xc0000028` trace passes through ntdll and repeated VBE7 frames. Both corrected
observers detached without an engine error; both full harness runs failed at the
variable-quantity test boundary. These are diagnostic-tool calibration results,
not new invSys behavioral RED/GREEN or a repaired regression gate.

An unresolved first frame does not identify a VBA procedure or prove a particular
pointer, declaration, package or workflow caused the failure. Debugger attachment
and project inspection change execution conditions. The standard uninstrumented
gate remains required, and all failed runs remain retained.

The stage variant adds an unsaved diagnostic module and numeric markers before
fixed launcher/test-stage statements. Its transport calibration must write marker
0 before any test action. The full standard sequence then runs without reduced
workflow switches. The terminal trace reaches ReleaseSource in the released
Process edit/export test, immediately before the real Process Release handler.
No later ViewSource marker is written. This identifies a narrower failure interval
for that instrumented run only; it does not prove the release command itself is
the cause. Package source edits are never saved, and trace files contain numeric
IDs with a separate fixed source-statement map, not operational values.

The compile-before-save comparison differs from the earlier finalization trial:
each of all five copied projects is invalidated by inserting/removing one comment,
its exact source is checked, explicit compilation must finish, and only then is
that disposable package saved. A subsequent independent cold-start/five-compile
check passes; all 168 extracted components match the original candidate. The
unmodified full harness passes its initial workflow and fails during clean
restart. Neither this preparation nor the earlier save-only preparation is an
adopted build change. The restart probe adds 20 fixed PowerShell markers to
the existing full harness's teardown/reload/action interval, without accessing
VBProject or changing VBA; it does not reduce the workflow to obtain a pass.
It passes 2/2, including the full initial workflow and clean restart. Its trace
reaches the post-restart-action inspection marker. This does not locate the
earlier plain run's failure or establish that logging repaired anything.

Next diagnostic: establish whether the outgoing owned Excel process is terminal
before new Excel construction. The current harness requests Quit, releases its
automation reference, waits 750 ms, and stops a remaining owned process without
an explicit WaitForExit before constructing the next instance. An exit-order
race is a hypothesis, not a verified cause. Protect any harness correction with
focused process-identity/terminal-state evidence before rerunning the unchanged
full Production gate; do not weaken its workflow or substitute marker success.
The [subsequent exit/execute investigation](plan022_slice4be_production_exit_and_execute_results.md)
observed terminal ordering in three focused cases and a full 2/2 run, without
support for an exit-wait fix. It also classified the original native failure as
an execute access violation and verified the existing package-edit boundary.

## Reproduction and retained evidence

Ignored directory: `reports/runtime/slice4be-launcher-denial/`.

- `combined-probe/`: initial file-lock harness failure, preserved.
- `combined-probe-v2/`: Core+Operations comparison, including initial/restart
  `*-package-provenance.json`; `combined-package-preservation.json`.
- `domain-both/`: all-current executing-package control; the Domain-isolation
  runner stops here because this control passes.
- `original-location-probe/`: original-location reference-inspection comparison.
- `timing-probe/`: timing-only control, with no VBA project access.
- `finalized/`: package-save/reference checks, five compiles/cold start,
  source comparison, independent hashes and failed plain `production-1/` run.
- `native-stack-probe/`: inconclusive initial observer run, preserved.
- `native-stack-probe-v2/`, `native-stack-probe-v3/`: corrected observer runs;
  `native-frames.json` contains bounded code/module evidence only.
- `NativeFrameProbe.cs`, `native-exception-fixture.ps1`,
  `test-native-frame-probe.ps1`: exact ignored observer/calibration sources;
  each synthetic result is preserved in its own `native-calibration-*` directory.
- `native-calibration-helper-sha256.txt` pins the calibrated helper; the harness
  rejects a changed helper until it has been calibrated again.
- `fixed-stage-probe/`, `create-stage-probe.ps1` and
  `fixed-stage-production-probe.ps1`: unsaved stage-instrumented full sequence,
  numeric trace, source map and failed terminal result.
- `compiled/`: forced-compile-before-save preparation, independent compile/source
  checks and hashes; `production/` retains the plain 1 PASS / 1 failure result.
  `restart-stages/` retains the 2/2 marker variant and all restart markers.
- `compile-save-candidate-probe.ps1`, `restart-stage-production-probe.ps1`:
  exact ignored preparation and full-harness restart-marker variants.
- `combined-production-probe-v2.ps1`, `original-location-probe.ps1` and
  `timing-production-probe.ps1`: exact ignored harness variants.

Each variant retains `-RepoRoot . -CallbackFilter Production
-WorkbookState ProductionReusable`, without focused/reduced-workflow switches.
Original-location/timing variants use
`-DeployRoot deploy/validation-receiving-launcher-denial`.
The combined variant uses `deploy/validation-diagnostic-denial-core-operations`;
the relocated all-current executing set uses
`deploy/validation-diagnostic-denial-both`.
The saved-copy comparison uses `deploy/validation-receiving-denial-finalized`.
Native variants use the original candidate and the same complete workflow flags.
Compile-before-save and restart-marker variants use
`deploy/validation-receiving-denial-compiled`.

The saved-copy package hashes (SHA-256) are:

| Package | SHA-256 |
|---|---|
| Admin | `f8e29ae265e259826cfb2c2e6e7f52881f3273763612f977ab30560fdb8f3832` |
| Core | `65a7d59851a5c5946a99287a11dbc4f1eca7e4dc955203f7191f98385f5888a6` |
| Designs Domain | `70e0f15c45acbfbb5880110cd82eb8a1de715543a102e9c5f69bb6951d3a7eb6` |
| Inventory Domain | `89b86526780e569c625bcef85555ebefc439378b6f2e2bd76e64b3f47da89259` |
| Operations | `1fbe59cf4ccfb831a9bf763b32ee32a14e572f33fdf24d27163780c2c2355bae` |

The compile-before-save package hashes (SHA-256) are:

| Package | SHA-256 |
|---|---|
| Admin | `b0a8e0dd4917a4872dd14cdf868bbdb770e359f40916c98f10f0f2649c64c546` |
| Core | `e4e85375c96f41e30d60bcc0fb9f6b033407d4851cb6c12137784e057579e608` |
| Designs Domain | `e7c2df1995f548f4293127b5fdec2b841d4a3578701918ba6e34419060f1dc42` |
| Inventory Domain | `f86094e5f68b066b7ff398e2c7072072fab3638c8d289703bad2b0e7d266db8e` |
| Operations | `2b578d11a56f00fd7bdaf51dadb220bd539d4035737f1ce42e331872c687d952` |

The first checkpoint's four native failures remain recorded with their own
hashes and conditions. No identical rebuild, Operations-only recompilation,
speculative runtime edit or debugger dump was repeated. The broader Event Viewer,
Settings, recording, guide/comparison and human Release 1 requirements remain open.

Final preservation check, 2026-09-08: all 20 pinned packages across original,
clean rebuild, saved-copy and compile-before-save sets remain byte-identical to
their respective pre-test manifests. Excel is closed and all diagnostic sessions
are terminal. Accepted deployment, operational workbooks and NAS were unchanged.
Only this curated record and synchronized Plan/controls status are checked-in
changes; no runtime or architectural contract changed. D13 product RED/GREEN and
static regeneration are not newly claimed for documentation/diagnosis. Native
helper calibration, parser checks, local links, diffs and Git status were reviewed
proportionally; the existing runtime/static baseline remains the prior checkpoint.
