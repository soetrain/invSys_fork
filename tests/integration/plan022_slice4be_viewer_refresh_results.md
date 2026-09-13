# Slice 4be.3 Events Refresh failure evidence

Last verified 2026-09-13. This checkpoint enforces existing Architecture v4.11
D18: missing coverage is Unavailable; failed Refresh retains visibly Stale
content only in its valid captured context. It does not complete comprehensive
Events, publication, Action Paths, or human acceptance.

## Focused RED and implementation

`Test-Slice4beConfigCommands.ps1 -CheckViewerRefreshFailure` generates disposable
warehouses through Admin and enters packaged `OpenInventoryViewer`, the Events
tab, and the real Refresh/Search handlers. A test-only Core read-boundary hook
supplies synthetic populated, failed, valid-empty and empty-failure envelopes.
Handler bodies remain intact. Shipping entry is counted, not replaced. Values
and selection are compared inside VBA; reports contain fixed checks/Booleans.
Generated authority files are hashed before and after Viewer actions.

On `deploy/validation-operations-settings-close`, the initial focused RED is
**6 PASS / 7 FAIL**. A failed Refresh loses loaded rows/search/selection, lacks
Stale/Unavailable notices, reads Shipping supplements despite failed primary
coverage, and fails sign-out/target invalidation. Valid-empty projection and
authority byte checks already pass.

The first candidate `deploy/validation-viewer-refresh` passes **13/13**.
Expanded empty-envelope checks then produce **14 PASS / 2 FAIL**: Refresh raises
runtime error 9 and reads the projection twice. The dialog capture and the
initial interrupted log are retained separately. The test now catches raised
errors and asserts the action result; this does not change the runtime handler.

Operations now validates the primary response before reading supplements,
handles empty failures without indexing an empty Split result, and reads the
projection once per Refresh. Viewer preserves the current Events rows, search
and selection on read failure; first failure is Unavailable, successful Refresh
clears Stale, and invalid session/target clears loaded content. An Inventory
surface is never retained as stale Events. No Core/Domain/schema change or new
architecture decision is introduced.

## Validation status

Final candidate: `deploy/validation-viewer-refresh-final`.
Final focused GREEN is **16/16**, report
`slice4be-viewer-refresh/3cdce48d5cb545ebacddd641fbb4178b/green.json`.
All five packages compile/cold start. The 192-component compiled comparison
finds exactly the two intended Operations changes; the other 190 components
have identical compiled code hashes to the preceding Settings checkpoint.

Settings is **187/187**, report
`slice4be-tracking-settings/f17310920c214f8e9124d86b71c70b83/green.json`.
Every prior check identity/GREEN is retained and no duplicates are present.
The populated Viewer regression passes all existing facts, including Settings
preservation on all three tabs, launcher reuse, export, readable events, date
filters/remembered range and snapshot non-mutation. Packaged smoke is **86/86**.
Full Release 1 chain is **31/31**, including successful ordered-child exit,
restart/reconciliation and five-package runtime evidence; its ordered live-role
child is **48/48**. See `viewer-refresh-full-chain.log`,
`viewer-refresh-full-chain-result.md`, and the ignored raw
`viewer-refresh-live-role-result.md`. The maintained full-chain result is
`tests/integration/slice14_results.md`. No retry was needed for this candidate.

Static maintenance is regenerated: 199 components / 5638 procedures / 125766
lines; 8 literal / 45 unresolved dynamic calls; 189 duplicate-body groups.
All 28 oversized modules do not grow; no components are added. All three JSON
schemas validate. Exact records include `viewer-refresh-component-review.json`,
`viewer-refresh-settings-identity-review.json`, `viewer-refresh-static-ratchets.json`,
and the corresponding regression/static logs under the ignored evidence root.
Final preservation verifies **140 prior package pins, 10 candidate pins, and
16 protected source pins**. Excel is closed and all validation handles are
terminal. `viewer-refresh-preservation.json` records the result. Unrelated
handoff 067 (3 additions/3 deletions) and untracked critique 023 remain untouched.

The final candidate's inspected `viewer-stale-events.png` shows both synthetic
rows, retained selection/search, and the complete readable Stale guidance.
This is automated visible evidence, not human UAT.

Focused command (use `-Phase RED` with the recorded preceding candidate for RED):

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -RepoRoot . -DeployRoot deploy/validation-viewer-refresh-final -CheckViewerRefreshFailure -CaptureEvidence -Phase GREEN
```

Ignored evidence root is `reports/runtime/`. Logs:

- `viewer-refresh-red.log`: initial 6/7 RED.
- `viewer-refresh-green.log`: first 13/13 GREEN.
- `viewer-refresh-empty-red.log`: expanded 14/2 RED.
- `viewer-refresh-empty-error.log/.png`: runtime error 9 before the probe caught it.
- `viewer-refresh-final-build.log`, `viewer-refresh-final-compile.log`,
  `viewer-refresh-final-compiled.json`, `viewer-refresh-final-green.log`.
- GUID directories under `slice4be-viewer-refresh/` contain fixed-result JSON
  and optional synthetic captures. Final runs never overwrite older packages.

Initial instrumentation used `ProcStartLine`, whose range included a leading
blank line; the hook landed outside the procedure. The editor capture and
`viewer-refresh-probe-failed.log` record this setup failure, not behavioral RED.
The corrected hook uses `ProcBodyLine`. Test-owned Excel processes were stopped
after inspecting their ID/start time and the captured interruption.

## Remaining D18 scope

Successful Shipping supplements still read current-state sources; moving them
to owning publication boundaries remains required. Core's old Events reader
still lacks complete grouped identities/detail, Designs/activity coverage,
publication metadata, policy/profile rendering and 100-record paging. The
5,000-complete-group publisher, recording/conclusions, How-To/Diagnostic/Compare,
guide authoring/import/export, comprehensive control coverage and human
comparison remain open. This checkpoint must not be used as evidence for those
requirements. Accepted deployment, operational workbooks and NAS are untouched.
