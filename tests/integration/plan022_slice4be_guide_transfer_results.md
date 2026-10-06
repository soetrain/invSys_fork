# Slice 4be-A guide transfer

Architecture v4.11 D18 defines guide-transfer-v1, imported-guide-v2 and the
reserved executable-transfer-v2 interface. These refine the approved transfer
boundary; they add no execution or permission grant. Transfer implementation,
GREEN and acceptance remain open.

## Packaged entry test

**2026-10-06 UTC: RED18 PASS / 6 expected FAIL**, with no harness failure.
The controller ran05:38:17.5598746--05:41:04.7614925 UTC and closed normally.
Five instrumented compiles pass, settings and all five frozen package hashes are
preserved, and the delayed Application-event audit finds zero Excel failures.
All22 host observations retain saved identity, one sheet and zero names; closed
file bytes match. The reviewed Published guides capture shows the selected second
version, separate authored/observed content and the absent transfer controls.

Receipts under `reports/runtime/`:

- `guide-transfer-controller/6bc2b397656f4a9fb5651da16bdaaacd/{closure,verification}.json`
- `slice4be-guide-transfer/0e9c365155544235a488b32b0ad29e47/{red.json,transfer-host-state.jsonl,visible-review.json,guide-transfer-admin-entry.png}`
- `guide-transfer-entry-static.json`: all436 repository PowerShell scripts parse.

Runtime source remains identical to `89674af1`. The preceding
`receiving-replay-visible-static-01/ratchet-verification.json` remains applicable:
312 components/6307 procedures/138053 lines, dynamic calls9/45,190 duplicate groups,
28 size caps and three schemas. It is reused evidence, not a new maintenance scan.
The existing broad GREENs on this same frozen candidate remain applicable; no
new release-chain or transfer acceptance is claimed. Desktop probes observed no
Win32 error5 during this work.

`Test-Slice4beGuideTransfer.ps1 -Phase RED -CaptureEvidence -SavedHost` uses
the frozen `deploy/validation-print-recorded-01` packages. Disposable copies are
instrumented and compiled before forms load. Actual Admin Settings actions supply
the recording; real guide handlers save two versions and a separate observed run.
The test selects the exact published version and exercises the same form-action
probe used by the existing operator-handler tests.

Required failures are the two absent Export/Import controls, their Admin click
handlers and their Reader-disabled states. A future compiled file-selection seam
returns cancellation without replacing either actual handler. Successful file
transfer, cancellation handling and permission enforcement are not established by
missing controls. No runtime source or deployed package changes are part of this
checkpoint.

## Earlier trials and fixture diagnosis

All controller paths below are under ignored
`reports/runtime/guide-transfer-controller/`; their first worker-log line names
the corresponding `slice4be-guide-transfer/` evidence directory. Preserve these
results; none is a clean transfer baseline.

| Controller | Result | Qualification |
| --- | --- | --- |
| `2e3c829c31184fc8b162501b26b08f24` | 12 PASS / 1 FAIL | Missing scope-local recording helper; corrected by loading the existing helper. |
| `64c07ac70dcc4025a95e9a253c98ff6e` | 9 PASS / 1 FAIL | Existing Create guide handler returned DISP_E_EXCEPTION before transfer checks; original Excel exited and left a recovery child. |
| `9c342c93999740878afad22dd8b1715a` | 17 PASS / 7 FAIL | Six missing-transfer assertions plus a saved-host fixture failure. Saved bytes were preserved after close. |
| `0bdca0454cd045f0b5e5d0ff9b6bea15` | 8 PASS / 2 FAIL | Diagnostic tried to hash an Excel-open file; sharing denial also affected cleanup. Diagnostic now reads only in-memory host state; final bytes are checked after close. |
| `5f163bf22659408fb76d5b566572bfb6` | 9 PASS / 2 FAIL | Existing OpenInventoryViewer returned DISP_E_EXCEPTION before transfer checks; recovery child required closure. |
| `9d148e42d8e646c2869c7aba90e9b38f` | 17 PASS / 7 FAIL | Reproduced six transfer failures and identified the host change at fixture publication. |

The final diagnostic's `transfer-host-state.jsonl` shows a saved one-sheet host
before publication, then an unsaved five-sheet host immediately afterward.
`PublishReadFixtureForTest` supplied no Admin workbook to ordinary
`GenerateInventorySnapshot`; its existing audit path resolved the active host and
added Admin sheets. The transfer-only probe now supplies a separate disposable
Admin workbook to that same publisher. It neither resets the host's Saved flag
nor weakens the original saved-identity/closed-byte assertions.

Both recovery children were closed through their owned native windows, with
**retain recovery files** selected and reviewed. No secondary COM attachment or
process kill was used. Their `assisted-recovery-closure.json` receipts distinguish
assisted closure from ordinary shutdown. All six controllers ultimately restored
settings, preserved package pins and closed Excel. The native failures remain
unexplained; a later pass must not be represented as a general Excel repair.

## Remaining work

Extend the packaged test with exact-version cross-warehouse round trip, imported
editing/re-export, native compatibility, foreign-origin labels, long/multiline
text and the D18 malformed-file, size, cancellation, existing-destination,
visibility, capability, captured-context and empty-library guards before runtime
implementation. Then obtain GREEN and the affected regression, layout, visible,
static and release-chain evidence. Entry RED alone does not accept transfer or A.
