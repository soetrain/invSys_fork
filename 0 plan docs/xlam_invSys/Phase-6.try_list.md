# Phase 6 Robust LAN Try List

## 1. Objective and pass/fail bar

- [ ] Lock the immediate tactical goal. Pass: the next proving target is specifically S1 and S2 `Confirm Writes` working in real LAN use so that users on any participating computer with `invSys.Inventory.Domain.xlam` loaded get an accurate snapshot-fed `invSys` table overview. Evidence: paired S1/S2 `Confirm Writes` runs plus refreshed `invSys` table state on both machines.
- [ ] Lock the proving ladder before any new LAN fix work starts. Pass: work is executed in this order only: one-account saved workbook, real LAN S1/S2, LAN + WAN, central aggregation. Evidence: dated test log naming the current rung.
- [ ] Lock the primary LAN success bar. Pass: S1 and S2 both complete real-LAN `Confirm Writes -> processor run -> canonical apply -> snapshot rebuild -> operator refresh` without one station regressing the other, and both resulting operator `invSys` tables stay accurate. Evidence: paired processor logs and operator workbook screenshots/log extracts from both machines.
- [ ] Lock the XLAM session success bar. Pass: role commands remain reliable regardless of which invSys `.xlam` files are already loaded in the Excel session, as long as the required Core + role + domain set is present, and `invSys.Inventory.Domain.xlam` participants still see an accurate `invSys` overview after refresh. Evidence: add-in inventory dump and run results from at least two different session-load orders.
- [ ] Lock the throughput success bar. Pass: normal LAN use moves 10,000 entries across workbooks without correctness loss, silent skips, or unrecoverable lock state. Evidence: batch report, row counts before/after, elapsed time, and spot-check reconciliation.
- [ ] Lock the investigation bar for Excel scale. Pass: 1,000,000-entry behavior is explicitly stress-tested and recorded as either credible in Excel or the point where database investigation becomes mandatory. Evidence: stress result sheet with failure mode, timings, memory symptoms, and operator impact.
- [ ] Keep warehouse-scale implications in scope. Pass: every LAN try is evaluated for whether it still holds when scaling toward 15 warehouses, even if current proving uses one warehouse and two stations. Evidence: per-section “15-warehouse note” or explicit “no additional warehouse risk” mark.
- [ ] Keep this step documentation-only. Pass: this file changes without modifying VBA, PowerShell, or deploy artifacts. Evidence: git diff limited to this markdown file.
- [ ] Lock the in-scope contracts for all future LAN fixes. Pass: all tries explicitly reference existing interfaces only: `tblStationConfig.PathInboxRoot`, `tblWarehouseConfig.PathDataRoot`, `setup_lan_station.ps1`, `modProcessor.RunBatchReportForAutomation`, `SkipInboxTargetInvalidPath`, `FF_AutoSnapshot`, `LastRefreshUTC`, `SnapshotId`, `SourceType`, `IsStale`. Evidence: each contract appears somewhere in this checklist or the execution record.

## 2. Non-negotiable invariants

- [ ] Treat the shared runtime root as warehouse-owned only. Pass: `WHx.invSys.Config.xlsb`, `WHx.invSys.Auth.xlsb`, `WHx.invSys.Data.Inventory.xlsb`, `WHx.invSys.Snapshot.Inventory.xlsb`, and `WHx.Outbox.Events.xlsb` live in one warehouse-authoritative shared root. Evidence: path inventory from the host machine and one station.
- [ ] Treat station inbox roots as station-owned only. Pass: each station writes only to its own inbox workbook under its own inbox root. Evidence: station config row plus physical file path listing.
- [ ] Keep `PathInboxRoot` processor-reachable. Pass: `tblStationConfig.PathInboxRoot` resolves to a processor-reachable path, preferably UNC, never a mapped-drive assumption. Evidence: raw cell value, expanded path, and shell/Excel reachability result.
- [ ] Separate shell reachability from Excel reachability. Pass: `Test-Path` and `Get-ChildItem` success is not counted as LAN success unless Excel `Workbooks.Open` also succeeds. Evidence: paired shell output and Excel open result for the same path.
- [ ] Require shared auth provisioning for LAN completion. Pass: station user auth rows exist and validate for the intended role before role-ready is claimed. Evidence: `setup_lan_station.ps1` output and auth validation result.
- [ ] Require config bootstrap for LAN completion. Pass: shared config row and local config copy both exist and agree on station identity, role default, and inbox root. Evidence: row dumps from both workbooks.
- [ ] Require snapshot refresh for LAN completion. Pass: role-visible inventory appears only through snapshot-fed refresh, never by direct local mutation of the operator `invSys` table. Evidence: refresh metadata and unchanged local staging tables.
- [ ] Treat the operator `invSys` table as the visible LAN success surface. Pass: the shared-runtime and processor flow is only counted as working when the local operator workbook's snapshot-fed `InventoryManagement!invSys` table shows the correct overview on the participating machines. Evidence: operator workbook table snapshot after refresh.
- [ ] Require processor locking for LAN completion. Pass: simultaneous use never relies on luck; processor serialization and retry-after-release behavior are proven. Evidence: lock logs, contention run result, and no corruption in inventory/outbox/inbox tables.
- [ ] Do not count simulated-only success as LAN success. Pass: automation harnesses may guide diagnosis, but real-PC LAN evidence is captured before claiming completion. Evidence: separate labels for simulated vs real-LAN results.

## 3. Try matrix by lane

### Config truth and path serialization

- [ ] Read the raw `tblStationConfig.PathInboxRoot` value for S1 and S2 directly from the live config workbook. Pass: raw text matches intended station root exactly. Evidence: Immediate Window or table dump with raw value.
- [ ] Record `Len(PathInboxRoot)` for each station row. Pass: length is consistent with a real UNC/local path and does not show doubled escaping or trailing garbage. Evidence: Immediate Window output with `Len=`.
- [ ] Expand `{WarehouseId}` and `{StationId}` placeholders exactly as the runtime does. Pass: expanded value points to the expected real station root. Evidence: raw value plus expanded value side by side.
- [ ] Normalize slash direction and trailing slash behavior. Pass: serialized config value survives normalization without changing meaning or producing duplicate separators. Evidence: before/after path text.
- [ ] Confirm `PathInboxRoot` is station-specific and `PathDataRoot` is warehouse-shared. Pass: no station row points back at the canonical warehouse runtime unless intentionally falling back for a single-box test. Evidence: shared/local config row dump.
- [ ] Confirm no mapped-drive path leaks into station config. Pass: all real LAN station rows use paths valid from the processor host, preferably UNC. Evidence: config export and processor-host shell check.
- [ ] Confirm UNC share root existence independently of invSys. Pass: the parent share exists before the inbox workbook path is evaluated. Evidence: `Test-Path` / `Get-ChildItem` against the share root.
- [ ] Compare local config copy vs shared config for each station. Pass: `StationId`, `WarehouseId`, `StationName`, `PathInboxRoot`, and `RoleDefault` match where they are supposed to. Evidence: row-by-row parity log.
- [ ] Try intentionally bad path text to prove diagnostics are honest. Pass: malformed path reproduces a clear failure such as `SkipInboxTargetInvalidPath` and does not silently route to the wrong inbox. Evidence: processor log excerpt.
- [ ] Try a missing share target. Pass: missing share fails clearly and does not mark role-ready or processed success. Evidence: setup output, processor log, and workbook-open result.

### Station bootstrap and shared-auth provisioning

- [ ] Run `setup_lan_station.ps1` for S1 and S2 from a clean starting point. Pass: each run completes with clear output for shared runtime root, local config path, inbox path, auth provisioning, and role readiness. Evidence: saved script output per station.
- [ ] Verify shared config row creation from bootstrap. Pass: station row exists in shared `WHx.invSys.Config.xlsb` with the intended inbox root and role. Evidence: shared config row dump after setup.
- [ ] Verify local config copy creation from bootstrap. Pass: local config copy exists and preserves the same station bootstrap values needed by the operator workbook. Evidence: local config workbook row dump.
- [ ] Verify shared auth row creation from bootstrap. Pass: station user receives the expected capability for the configured role and warehouse/station scope. Evidence: auth provisioning and validation output.
- [ ] Verify setup fails clearly when auth provisioning fails. Pass: no false `RoleReady=True` when auth rows are missing or invalid. Evidence: captured failing setup output.
- [ ] Verify setup fails clearly when inbox bootstrap fails. Pass: no role-ready result if the inbox workbook cannot be created or opened. Evidence: setup output and inbox path existence check.
- [ ] Verify setup fails clearly when snapshot shell or Excel openability fails. Pass: bootstrap does not mask inaccessible snapshot paths. Evidence: `SnapshotShellAccessible=` and `SnapshotExcelOpenable=` lines.
- [ ] Re-run bootstrap over an existing station. Pass: repeated setup is idempotent and does not drift config/auth state or create duplicate station rows. Evidence: before/after table counts and setup output.

### Add-in/build parity and stale XLAM detection

- [ ] Dump `Application.AddIns` for S1 and S2 in the failing scenario. Pass: both machines show the same intended invSys add-in set and file paths. Evidence: Immediate Window add-in inventory from both PCs.
- [ ] Compare `FullName` paths for every installed invSys add-in. Pass: no machine points at stale archive/build paths when another machine points at `deploy/current` or its intended equivalent. Evidence: side-by-side path comparison.
- [ ] Confirm build parity for Core, role, and domain XLAMs. Pass: both stations are running the same build set for the same test. Evidence: file timestamps, hashes, or version stamp note.
- [ ] Try alternate load order scenarios. Pass: required commands still work when unrelated invSys XLAMs are already loaded in session. Evidence: session inventory and command result for each load order.
- [ ] Try a deliberately stale add-in on one station. Pass: the failure becomes explicit and reproducible rather than silently flipping S1 vs S2 behavior. Evidence: captured error or regression log.
- [ ] Verify no suspected signature drift is hidden. Pass: current callers can still run exposed macros such as `modProcessor.RunBatchReportForAutomation` without wrong-argument or no-op behavior. Evidence: direct macro invocation result on both stations.

### Session state and ghost workbook detection

- [ ] Enumerate all open workbooks in the active Excel session on S1 and S2. Pass: transient runtime artifacts are visible in diagnostics even when hidden. Evidence: workbook name/full path dump.
- [ ] Enumerate hidden workbook windows. Pass: hidden transient workbooks can be identified by name/path and correlated with lock symptoms. Evidence: Immediate Window dump including `Visible=` status.
- [ ] Force-close known transient artifacts after a failed run. Pass: next run changes behavior only if the prior failure was caused by ghost-open state. Evidence: forced-close log plus retest result.
- [ ] Repeat the same command without restarting Excel. Pass: the second run does not accumulate extra hidden workbooks or change outcome unexpectedly. Evidence: before/after workbook inventory.
- [ ] Repeat after full Excel restart. Pass: restart either clears the symptom or proves the issue is not session residue. Evidence: restart note, workbook inventory, and retest result.
- [ ] Check for double-open attempts against the same inbox/snapshot/outbox/runtime workbook. Pass: either reuse occurs correctly or failure is explicit; no silent read-only fallback is accepted. Evidence: workbook inventory and processor/role logs.

### Shell-vs-Excel path accessibility

- [ ] Test shell access to every critical path from the processor host: shared runtime root, station share root, inbox workbook, config workbook, auth workbook, snapshot workbook. Pass: shell can see every required file/folder. Evidence: `Test-Path` and `Get-ChildItem` output.
- [ ] Test Excel open access to the same critical paths from the same machine. Pass: `Workbooks.Open` succeeds for every required workbook that shell can see. Evidence: Excel macro or COM harness result.
- [ ] Capture mismatches between shell success and Excel failure. Pass: every mismatch is recorded as its own defect class, not dismissed as “path is good.” Evidence: paired shell/Excel matrix.
- [ ] Test access from both the processor host and the opposite station. Pass: processor reachability and operator reachability are both proven where required. Evidence: two-machine access matrix.
- [ ] Test access using hostname UNC and raw IP UNC where applicable. Pass: the chosen production path style is the one that stays reliable under Excel open semantics. Evidence: comparison result.
- [ ] Test access when credentials are cold after reboot/logoff. Pass: Excel still reaches the target or fails loudly enough to block false success. Evidence: cold-session result note.

### Processor inbox target discovery and misrouting

- [ ] Run `modProcessor.RunBatchReportForAutomation` with known queued work for S1 and S2. Pass: expected inbox target count matches actual configured stations and roles. Evidence: batch report plus processor log.
- [ ] Review processor diagnostics for every run. Pass: no `SkipInboxTargetInvalidPath`, unexpected target omissions, or wrong-station processing occurs in a passing run. Evidence: saved log excerpt.
- [ ] Confirm each station row expands to the correct inbox workbook name for Receiving, Shipping, and Production. Pass: no station ever points at another station’s inbox file. Evidence: resolved path matrix.
- [ ] Queue work on S1 only and prove only S1 inbox work is consumed. Pass: no wrong-target pickup. Evidence: inbox row status and applied log.
- [ ] Queue work on S2 only and prove only S2 inbox work is consumed. Pass: no wrong-target pickup. Evidence: inbox row status and applied log.
- [ ] Queue work on S1 and S2 simultaneously and prove both are discoverable in one warehouse run. Pass: both stations’ queued rows are processed without target loss. Evidence: batch report and inbox statuses.
- [ ] Remove one station row temporarily or corrupt one station path and re-run. Pass: the processor still discovers the healthy targets and reports the unhealthy one explicitly. Evidence: target count and diagnostics.
- [ ] Confirm no hidden fallback to local-only processing when cross-station access fails. Pass: a cross-station failure remains visible and does not masquerade as success because the local inbox still processed. Evidence: processor target count and station-specific row statuses.

### Workbook open/reuse/close semantics and read-only fallback

- [ ] Prove whether inbox workbooks are reused when already open in-session. Pass: same workbook is reused cleanly instead of reopened into lock trouble. Evidence: workbook inventory before/after command.
- [ ] Prove whether canonical runtime workbooks are reused when already open. Pass: open reuse does not flip later opens into read-only or hidden lock failure. Evidence: workbook inventory and run result.
- [ ] Try opening transient workbooks in one session, then running from another. Pass: read-only state is either blocked clearly or retried safely; no silent no-op. Evidence: read-only/open result and processor log.
- [ ] Check for silent `DisplayAlerts=False` masking. Pass: any open failure that would have shown a lock/read-only dialog surfaces somewhere deterministic in logs or reports. Evidence: failure capture path.
- [ ] Verify close-after-use behavior for processor-opened workbooks. Pass: workbooks opened by the processor do not remain orphaned after the run. Evidence: workbook inventory after run.
- [ ] Verify close-after-use behavior for sync/snapshot workbooks. Pass: snapshot or outbox access does not leave hidden handles that poison later runs. Evidence: workbook inventory and follow-up run result.
- [ ] Try repeated open/process/close loops at short cadence. Pass: no handle leak or eventual read-only drift appears across many cycles. Evidence: loop count and final workbook inventory.

### Lock contention and simultaneous station use

- [ ] Hold the canonical inventory workbook open in one Excel session and post/process from another. Pass: contention is denied or deferred clearly, with no corruption and no false processed count. Evidence: lock-denied run result.
- [ ] Release the hold and retry immediately. Pass: retry-after-release succeeds cleanly without manual cleanup beyond releasing the lock. Evidence: second run result.
- [ ] Post from S1 and S2 at nearly the same time. Pass: both posts persist to their own inboxes without cross-contamination or overwrite. Evidence: inbox row counts and event IDs.
- [ ] Process while one station is actively posting again. Pass: serialization remains correct and no station loses rows. Evidence: row reconciliation and processor log.
- [ ] Repeat simultaneous posting across Receiving, Shipping, and Production inboxes. Pass: mixed-role use does not create different lock behavior than receiving-only tests. Evidence: three-role run matrix.
- [ ] Verify lock expiry and heartbeat behavior under longer runs. Pass: no premature lock break, no stranded held lock after completion, and no hidden overlap between processors. Evidence: lock table/log timestamps.
- [ ] Verify a crash or forced Excel close mid-run does not permanently poison the warehouse. Pass: break-lock/retry recovery returns the system to a clean state without duplicate apply. Evidence: recovery drill result.

### Operator read-model refresh and stale-state visibility

- [ ] Verify `FF_AutoSnapshot=True` open refresh on a saved operator workbook. Pass: `invSys` refreshes on open without mutating `ReceivedTally`, shipping staging, production staging, or local logs. Evidence: before/after workbook table snapshot and metadata.
- [ ] Verify post-write refresh after a successful command. Pass: operator inventory changes only after post, processor apply, snapshot rebuild, and refresh. Evidence: ordered timestamps and operator workbook state.
- [ ] Verify the core tactical view outcome after `Confirm Writes`. Pass: after S1 or S2 `Confirm Writes`, any participating machine with `invSys.Inventory.Domain.xlam` loaded and refreshed against the shared snapshot path sees the accurate `invSys` overview expected by the model. Evidence: post-refresh `invSys` table snapshot from more than one machine.
- [ ] Verify cadence refresh. Pass: configured interval refreshes the read model without damaging local staging tables. Evidence: elapsed time log and before/after table comparison.
- [ ] Verify visible stale-state signaling. Pass: missing or stale snapshot sets `IsStale=True` visibly and does not silently present stale inventory as current. Evidence: operator workbook metadata and UI capture.
- [ ] Verify metadata correctness. Pass: `LastRefreshUTC`, `SnapshotId`, `SourceType`, and `IsStale` all update coherently for local, share, and cached cases. Evidence: metadata table dump.
- [ ] Verify missing snapshot does not block posting. Pass: operator can still queue inbox events while stale state is shown honestly. Evidence: stale operator state plus successful inbox post.
- [ ] Verify stale snapshot does not mutate local workflow surfaces on refresh failure. Pass: failed refresh leaves staging intact and only affects freshness metadata/state. Evidence: before/after local staging comparison.

### Saved-workbook reopen/restart behavior

- [ ] Use saved `.xlsb` or `.xlsm` operator workbooks only for proving, not `Book1`. Pass: every LAN claim is backed by saved-workbook evidence. Evidence: workbook paths in logs/screenshots.
- [ ] Close and reopen the operator workbook in the same Windows account/session. Pass: workbook resumes without identity drift, missing surfaces, or stale-runtime confusion. Evidence: reopen run result.
- [ ] Close all Excel windows, reopen from scratch, and retest. Pass: the same saved workbook still works after cold start. Evidence: cold-start run result.
- [ ] Reopen after changing which invSys add-ins are preloaded in the account session. Pass: required commands still work if the required set is present. Evidence: add-in inventory and command result.
- [ ] Reopen after a prior failure that left hidden workbooks. Pass: restart clears only session residue, not true config/runtime defects. Evidence: before/after result classification.
- [ ] Reopen on both S1 and S2 with their own saved operator workbooks. Pass: the same operational pattern works on both PCs, not just one golden machine. Evidence: paired reopen test records.

### Throughput, batching, and memory pressure

- [ ] Establish a baseline at 1 event. Pass: correctness and diagnostics are clean before scaling. Evidence: single-event run log.
- [ ] Run 100-event batches. Pass: no new lock/open/path failure class appears. Evidence: batch report and elapsed time.
- [ ] Run 1,000-event batches. Pass: correctness still reconciles and operator refresh remains credible. Evidence: counts, timings, and sample operator read-model check.
- [ ] Run 10,000-event batches. Pass: this is the required Excel success bar; correctness, recovery, and operator usability remain acceptable. Evidence: reconciliation sheet and timing.
- [ ] Run 100,000-event batches. Pass: failure mode, if any, is characterized precisely rather than hand-waved as “too slow.” Evidence: timings, memory symptoms, and recovery result.
- [ ] Run 1,000,000-event stress. Pass: record whether Excel remains credible, partially credible only with operational constraints, or no longer credible. Evidence: stress report with outcome classification.
- [ ] Try different `RunBatch` chunk sizes instead of assuming one batch size. Pass: chunking behavior is measured as a technique, not assumed as the answer. Evidence: chunk-size comparison table.
- [ ] Measure workbook growth and open/save cost as event counts rise. Pass: storage growth and reopen behavior stay within acceptable operational limits or are recorded as blockers. Evidence: file sizes and open/save timings.
- [ ] Measure repeated processor cycles under load. Pass: later runs do not degrade into read-only, stale locks, or hidden workbook buildup. Evidence: multi-cycle stress log.

### Multi-warehouse publish/aggregate implications

- [ ] Verify the local warehouse LAN fix does not break publish-to-share behavior. Pass: outbox/snapshot publication still works after LAN-focused changes are later tried. Evidence: publish result and artifact presence.
- [ ] Verify local warehouse LAN fix does not assume a single warehouse forever. Pass: path conventions and diagnostics stay warehouse-qualified. Evidence: config/path examples using different warehouse IDs.
- [ ] Run at least a small multi-warehouse publish simulation after local LAN proving. Pass: snapshots/outboxes from two warehouse roots remain distinguishable and aggregatable. Evidence: publish and aggregation result.
- [ ] Check warehouse-count pressure toward 15 warehouses. Pass: any technique chosen for S1/S2 is reviewed for whether it would become operationally fragile with many warehouse roots, snapshots, or published artifacts. Evidence: short warehouse-scale note.
- [ ] Verify global snapshot remains advisory only. Pass: no LAN fix causes warehouse-local authoritative balances to be overridden by global data. Evidence: aggregation test note and UI/output check.

### Failure injection and recovery drills

- [ ] Inject bad UNC text into `PathInboxRoot`. Pass: failure is explicit, diagnosable, and reversible. Evidence: config diff and processor/setup output.
- [ ] Inject unreachable share path. Pass: the system does not silently process local-only and pretend LAN is healthy. Evidence: shell/Excel mismatch log and processor result.
- [ ] Inject stale XLAM state on one machine. Pass: failure is attributable to stale deployment and does not look like random LAN drift. Evidence: add-in path inventory and failure output.
- [ ] Inject ghost lock/ghost workbook state. Pass: detection and cleanup steps are reproducible and the system recovers cleanly afterward. Evidence: hidden workbook dump and retest result.
- [ ] Crash or forcibly close Excel mid-processing. Pass: recovery path is documented, idempotent, and free of duplicate apply. Evidence: recovery drill log.
- [ ] Delay snapshot publication or refresh. Pass: stale-state handling is honest and posting remains available. Evidence: operator metadata and posting result.
- [ ] Delay WAN/share synchronization after local processing. Pass: eventual publication and aggregation behavior remain coherent once connectivity returns. Evidence: delayed publish/aggregate result.
- [ ] Remove or corrupt a snapshot workbook. Pass: stale-state signaling and recovery are explicit. Evidence: refresh result and restored-good retest.

### Logging/diagnostics gaps

- [ ] Capture the minimum diagnostic bundle for every LAN try. Pass: each run saves add-in inventory, open workbook inventory, shell path checks, Excel open checks, setup output, batch report, and processor log excerpt. Evidence: one bundle folder per try.
- [ ] Confirm `SkipInboxTargetInvalidPath` is sufficient when path text is wrong. Pass: if not sufficient, note the exact missing detail needed for future fixes. Evidence: path-failure log review.
- [ ] Confirm read-only/open failures surface somewhere deterministic. Pass: if not, write down the exact blind spot. Evidence: failure case and missing signal note.
- [ ] Confirm target-count diagnostics are enough to detect “local succeeded, remote skipped.” Pass: if not, note the missing target discovery metrics. Evidence: mixed-station run review.
- [ ] Confirm snapshot refresh diagnostics are enough to distinguish missing snapshot, stale snapshot, and refresh failure. Pass: if not, note the missing signal. Evidence: refresh-failure matrix.
- [ ] Confirm add-in/build diagnostics are enough to prove parity across machines. Pass: if not, note the minimum additional stamp needed. Evidence: parity review.
- [ ] Close this section only after every known failure class from `Phase-6.notes.md` maps to either an existing diagnostic or a named logging gap. Evidence: failure-class mapping table.

## 4. Scale and soak proving

- [ ] Run a short soak with S1 and S2 alternating commands under real LAN conditions. Pass: no flip where one station starts working only because the other stopped. Evidence: time-sequenced activity log.
- [ ] Run a same-role soak with both stations posting repeatedly. Pass: repeated receiving-only or shipping-only use stays stable across many cycles. Evidence: soak summary and final reconciliation.
- [ ] Run a mixed-role soak across Receiving, Shipping, and Production. Pass: role diversity does not introduce new session or path instability. Evidence: mixed-role soak log.
- [ ] Run a soak with Excel left open for an extended period. Pass: long-lived sessions do not accumulate hidden workbooks, stale handles, or wrong add-in state. Evidence: periodic workbook/add-in inventory dumps.
- [ ] Run a soak with periodic Excel restarts. Pass: restart cadence does not become an implicit crutch for correctness. Evidence: restart schedule and results.
- [ ] Run a soak with snapshot auto-refresh enabled. Pass: cadence refresh remains non-destructive and honest about stale state. Evidence: metadata timeline and staging-table diff checks.
- [ ] Run a soak approaching operational volume for 10,000 moved entries. Pass: throughput remains reliable over time, not just in one clean burst. Evidence: aggregate counts, timings, and reconciliation.
- [ ] Record any condition that requires manual cleanup, Excel restart, hidden workbook closure, or rebuild redeploy. Pass: these conditions are treated as failures until eliminated or formally accepted as an operational constraint. Evidence: incident log.

## 5. Exit criteria

- [ ] Exit one-account proving only after saved-workbook reopen/restart behavior is stable. Evidence: passing saved-workbook record.
- [ ] Exit LAN proving only after S1 and S2 both pass `Confirm Writes` in real LAN use with the same build, without mutual regression, and with accurate refreshed `invSys` tables on the participating machines. Evidence: paired pass bundle from both PCs.
- [ ] Exit LAN proving only after shell access and Excel open access are both green for all critical paths. Evidence: path matrix.
- [ ] Exit LAN proving only after shared auth, config bootstrap, processor discovery, lock contention handling, and operator refresh are all green. Evidence: completed try matrix sections.
- [ ] Exit LAN proving only after 10,000-entry movement is reconciled successfully. Evidence: throughput proof bundle.
- [ ] Exit LAN proving only after at least one soak run passes without ghost-workbook/manual-cleanup incidents. Evidence: soak summary.
- [ ] Exit LAN + WAN proving only after delayed sync and stale snapshot scenarios are green. Evidence: WAN-delay test bundle.
- [ ] Exit central aggregation proving only after published snapshots from more than one warehouse can be aggregated without violating advisory-only rules. Evidence: aggregation result bundle.
- [ ] Do not call Phase 6 LAN complete while any checklist item remains red in config truth, stale XLAM parity, hidden workbook state, or simultaneous-station processing. Evidence: final checklist review.

### Execution lanes

- [ ] Route Excel/COM, deployment, config/auth/bootstrap, runtime workbook, and packaged XLAM fixes to `runtime-packaging`. Evidence: future work item tagged with that lane.
- [ ] Route processor, locks, inbox/outbox discovery, idempotency, and inventory-apply semantics to `core-event-inventory`. Evidence: future work item tagged with that lane.
- [ ] Route reproducibility, fixtures, and validation-script coverage to `test-harness`. Evidence: future work item tagged with that lane.

## 6. Database pivot gate

- [ ] Stay on Excel only if real-LAN reliability is credible after exhausting the try matrix. Pass: no unresolved class remains in pathing, stale add-ins, workbook-open semantics, locking, refresh honesty, or restart behavior. Evidence: final red/green matrix.
- [ ] Stay on Excel only if 10,000-entry movement is reliable and operable. Pass: operators can run it without hidden cleanup rituals, silent skips, or fragile session-state tricks. Evidence: throughput and soak bundles.
- [ ] Stay on Excel only if the chosen approach still looks sane when projected toward 15 warehouses. Pass: warehouse-scale review does not uncover an obvious operational cliff. Evidence: short scale assessment.
- [ ] Open formal database investigation if 1,000,000-entry stress shows Excel is no longer credible for reliability, recoverability, or operability. Evidence: stress report with named pivot reason.
- [ ] Open formal database investigation if LAN correctness depends on manual Excel restarts, force-closing hidden workbooks, rebuilding add-ins, or hand-repairing config/auth state as a normal operating procedure. Evidence: incident pattern log.
- [ ] Open formal database investigation if simultaneous-station reliability cannot be made deterministic for S1 and S2 on real LAN hardware. Evidence: repeated failing contention/recovery record.
- [ ] Open formal database investigation if projected 15-warehouse operation requires an unacceptable volume of shared workbook opens, sync churn, or operator-facing stale-state exceptions. Evidence: scale assessment note.
- [ ] Do not start database work as a parallel default. Pass: pivot happens only after this checklist is exhausted and failure reasons are documented concretely. Evidence: signed-off try matrix and pivot memo.
