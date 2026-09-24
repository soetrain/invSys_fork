# Plan 022 Slice 4be.3: original event text preservation

## Contract and scope

Architecture v4.11 D18 requires original permitted values in the Operations-owned
Event Detail, opened from the captured Operations or Admin Viewer projection.
The existing Core publisher escapes backslashes before tabs and line breaks.
The Viewer previously replaced newline escapes before escaped backslashes,
corrupting literal text such as `\n`. The correction restores the existing wire
contract; it adds no authority, profile field, workflow or architectural decision.

`frmInventoryViewer.ViewerUnescape` now decodes each escape once. Literal
backslashes, unknown escapes, trailing backslashes, actual CRLF/LF/CR and tabs
remain distinguishable. No canonical data or publisher format changes.

## Packaged D13 evidence — 2026-09-24

The protecting test is `tests/tooling/Slice4beViewerEventDetail.ps1`. It starts
with the existing Admin-generated warehouse fixture, publishes three synthetic
permitted Item values through the real Core publisher, opens the public Viewer
and selects contributing lines through the actual form handler. It verifies
each fixture cell before publication and compares final field values by binary
equality. A missing fixture or failed publication is not the expected RED.

Six new checks cover literal escapes, actual line breaks/tab/Unicode and mixed
text separately for Operations and Admin. All preceding 50 checks remain.

- RED: **52 PASS / four expected FAIL**, on frozen
  `deploy/validation-auth-read-separated`; only LiteralEscapes and MixedText
  fail in each role. All prior 50 and both actual-line-break checks pass.
  UTC **20:07:30--20:09:00**.
- GREEN: **56/56**, on isolated `deploy/validation-detail-original-text`,
  with exactly the same identities. UTC **20:13:10--20:14:41**.
- Each run passes five instrumented compiles, native scrolling and typing
  protection, original source/unknown-column preservation, profile filtering,
  captured-context invalidation and owned-form lifecycle checks.
- Both runs close Excel normally, restore local settings and preserve all
  five package/tooling hashes. Application 1000/1001/1002 audit: zero events.
- Four GREEN images are individually reviewed: default, horizontal-scroll,
  maximized and restored. Literal escape text remains visibly literal; the
  long Coverage ending remains reachable. These are single-line reachability
  and original-value evidence, not complete multiline rendering acceptance.

Private report roots under `reports/runtime/slice4be-viewer-detail/`:
RED `940e50951ef245d4994acd7808e4dbe1`, GREEN
`eae33c2edcb94ee6b88bf61ba149bd24`. Receipts:
`reports/runtime/detail-original-text-red-verification.json` and
`reports/runtime/detail-original-text-verification.json`.

All five candidate packages build and compile; Operations cold-start references
resolve within the candidate. Comparing 244 compiled components against the
preceding candidate finds only `invSys.Operations.xlam/frmInventoryViewer`
changed. No component is added or removed. Accepted deployment is unchanged.

## Regression and maintenance gates

Viewer/filter/Shipping-state regression passes **98/98**, retaining the exact
preceding 98 identities. Three captures are individually reviewed: filters,
activity-line labels and the held Shipping projection. Excel closes normally,
settings restore and package/tooling hashes are unchanged. UTC
**20:15:08--20:18:18**. Report under runtime:
`slice4be-viewer-published-read/5636b4f5fa3c4f6eaadd9c38cbca6bd4`;
receipt `detail-original-text-viewer-verification.json`.

Static maintenance regenerates successfully with three schema checks: 251
components, 6,047 procedures, 132,986 source lines, nine literal/45 unresolved
dynamic calls and 191 duplicate groups. The decoder adds 14 lines; this is the
bounded implementation needed to decode each character once, with no new
procedure or duplicate helper. All 28 pre-existing oversized-module limits hold;
dynamic-call and duplicate-body metrics do not grow. Source layout checks pass
8/8 and 7/7; all 276 recursively enumerated tools/tooling PowerShell files parse.
Receipt: `reports/runtime/detail-original-text-static-verification.json`.

The first current-candidate chain attempt fails: **5 PASS / one harness FAIL**;
live-role stops at **32 PASS / one harness FAIL**, while Create Warehouse passes
**15/15**. The original Excel process crashes during canonical projection rebuild
(`modProcessor.RunBatchReportForAutomation`, RPC HRESULT `0x800706BE`). Application
events 1000/1001 identify Excel/`ntdll.dll`, exception `0xc0000028`; cause is
unproven. No decoder-specific failure is established. UTC
**20:18:39--20:25:46**. A verified post-crash recovery Excel process is terminated;
the original controller then restores settings and all three tracked reports.
Five package and tooling hashes remain unchanged. Normal shutdown and the chain
are not accepted from this attempt. Evidence prefix:
`reports/runtime/detail-original-text-chain`; receipt
`detail-original-text-chain-failure-verification.json`.

The unchanged-candidate retry passes **32/32 chain, 48/48 live-role and 15/15
Create Warehouse**, retaining every previous identity. UTC
**20:26:53--20:32:19**. Excel closes normally without assistance; settings and
all three tracked reports restore, five package/275 tooling hashes hold, and
the retry interval has zero Application failure events. Receipt:
`reports/runtime/detail-original-text-chain-retry-verification.json`. The retry
does not explain the earlier crash or erase that failed attempt.

Full multiline readability still needs its own protecting packaged test and visible acceptance.
The broader Shipping/Boxing caption failure and normal-shutdown gate remain
open with their previously recorded evidence. This checkpoint does not complete
Slice 4be or Release 1 acceptance.
