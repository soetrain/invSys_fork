# Slice 4be.1 Receiving staging and disposition activity

Architecture v4.11 D18 governs this extension of the approved shared activity
contract. Candidate implementation and release acceptance are not complete.

## Focused RED, 2026-09-07

Unchanged runtime code `1689bff` and the five package hashes recorded in
[Receiving confirmation evidence](plan022_slice4be_receiving_activity_results.md)
were used for both runs. Tests invoke actual packaged `mBtnAdd_Click` and
`mBtnConfirm_Click` handlers using unsaved fixture-only access seams.

- Initial new coverage: **206 PASS / 15 FAIL**. Missing Add Selected,
  Add Disposition and Confirm Dispositions attempt/outcome records.
- Expanded rejection and protected-staging failure coverage: **210 PASS /
  35 FAIL**. No harness errors; all previous 190 checks remain GREEN.
- Independent guards prove two actual Add operations, three subsequently
  submitted exact event identities (including one direct-service staging call),
  all receipt or RETURN/DUMP event applications once, captured workbook binding,
  preservation of the unrelated workbook and extra staging header, and no
  user-control attribution for the direct staging service call.
- Invalid quantity and a protected staging sheet exercise the actual Add handler
  before runtime edits. Both preserve the entire staged row data and show the
  existing cause. Expected observation effects differ: pre-validation is
  Unchanged, while service failure remains Unknown without inferring rollback.

Ignored evidence: `reports/runtime/slice4be-receiving-activity/staging-red.json`
and `staging-outcomes-red.json`. Reports contain check names and booleans only;
raw activity, fixture values and screenshots remain outside committed evidence.

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File tests/tooling/Test-Slice4beConfigCommands.ps1 -DeployRoot deploy/validation-activity -Phase RED -CheckActivityEvidence -CheckActivityFoundation -CheckReceivingActivity -CheckReceivingStagingActivity -CaptureEvidence
```

GREEN, catalog compatibility, changed-package compile, layout, static,
live-role, full-chain and visible operator gates remain pending.
