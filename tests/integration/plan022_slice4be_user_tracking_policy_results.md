# Slice 4be-A / A1: per-user recording policy

Last verified: 2026-10-05 UTC. Approved D18-REPLAY-01 boundary7 now has a concrete
policy-v2 wire definition in Architecture v4.11; Plan022/Controls1.452 agree.
This checkpoint changes specification/tests only. **52 PASS/18 FAIL** on frozen
`deploy/validation-receiving-run-05` is meaningful RED; no implementation or
acceptance is claimed.

## Contract and protecting route

The versioned policy adds sparse `{UserId, Record}` overrides, enabled by default,
and a counted append-only user table. Preserve v1 reads/history, reject downlevel
writes that would erase v2 flags, preserve unknown columns and fail closed on
malformed policy. Admin stages/saves through its existing D5 authority; Core owns
optional collection/capture decisions. Visibility, rights, history, required audit
and ordinary business work remain independent. Removed-user flags are retained;
roster reads must not repair/write Auth or expose credentials.

`Test-Slice4beUserTrackingPolicy.ps1 -Phase RED` uses the existing compiled Settings
route with all five packaged projects compiled before forms. The new live-control
probe looks up the real named ListBox/CheckBox/buttons and changes their values;
Save/Reload/Reset use their actual handlers. Missing controls return MISSING,
never a constructed policy. Supplemental public Core recording Start/Status tests
and real General Settings saves check actor behavior and earlier record visibility.
No activity or recording outcome is fabricated.

## Evidence

- Initial59-check gate:45 PASS/14 FAIL; controller
  `user-tracking-policy-controller/26a1d3862ee9443a824978c2ca53ae21/`, worker
  `slice4be-tracking-settings/9cdf5308ede04a5da4d1eb050bbeeb0c/red.json`,
  02:34:35.687--02:35:53.891 UTC.
- Expanded70-check gate:52 PASS/18 FAIL; controller
  `user-tracking-policy-controller/9c0acc29300f48aeaf1d7671e26ccf6c/`, worker
  `slice4be-tracking-settings/4f5d396072bc4457a3ab43f2ea49eb69/red.json`,
  02:37:10.612--02:38:51.917 UTC. All18 failures are UserPolicy checks; no duplicates.
- Each controller preserves `worker.log`, `package-pins.json` and `closure.json`;
  the expanded controller also has `verification.json`. All paths are under ignored
  `reports/runtime/`; raw runtime values are not committed.

Failures isolate missing v2 projection/user controls, staging/save/reload/reset/close
behavior and protection against downlevel loss. Actor checks reproduce the intended
disabled user being able to start capture and emit optional activity. Enabled
recording, real Settings work, earlier activity visibility and policy-change partial
closure pass. Malformed-request negatives currently pass because v2 is unsupported;
they must remain passing once v2 is accepted. The first rejected harness option
(`ba79a76f159f446194f27ea5332af8f1`) never opened Excel and is not product RED.

Both completed controllers restore settings, preserve the five frozen packages and
close Excel normally. No runtime source changed; runner93, guide58, Receiving854
and candidate05 compiles remain applicable. Regenerated `user-policy-static-red-01/`
retains307 components/6279 procedures/137389 lines, dynamic calls9/45 and190
duplicate groups; all28 size limits and three schemas pass. All399 repository
PowerShell scripts parse. The expanded gate retains all45 initial passes.
Native audit02:34:35--02:40:22 UTC finds zero Excel1000/1002 events; desktop probe
02:40:23.133 UTC (Oct4 19:40 PDT) succeeds. No desktop error5 was observed.
This gate does not prove Receiving
business audit preservation, replay capture-loss stopping, roster privacy, malformed
v2 storage, v1-history retention or comprehensive observation/visual acceptance.

## Next implementation and tests

Implement the defined v2 policy through `modTrackingPolicyModel`,
`modTrackingPolicyCommand`, `modActivityPolicy`, `modTrackingPolicySettings` and
`cAdminTrackingPolicy`, factoring read-only roster/user-row handling into a small
Core module. Avoid growing the capped Auth module or exposing its credential cache.
Project only the caller's effective user flag to Operations personal Settings.
Register the selector/toggle observations without admitting them to older catalogs.

Retain all52 passes; add stored-row/count corruption, v1 history, missing/dirty Auth,
role/context/stale-version/cancelled-save and unknown-column guards, then actual
Receiving/replay capture-loss proof. Check new layout and visible controls and run
affected policy/preference/observation regressions before accepting A1. Broader
4be-A and the previously recorded full-chain native failure remain open.
