# Codex Sub-Agent Definitions for invSys

## Purpose

This document defines the first-pass Codex sub-agent set for the invSys repo.
It is the fourth layer in the sub-agent stack:

1. Coordinator rules
2. Path ownership matrix
3. Agent prompt blocks
4. Codex sub-agent definitions

Use this document together with:

- [codex_sub-agent_policy.md](/c:/Users/Justin/repos/invSys_fork/0%20plan%20docs/codex_sub-agent_policy.md)
- [codex_sub-agent_prompts.md](/c:/Users/Justin/repos/invSys_fork/0%20plan%20docs/codex_sub-agent_prompts.md)
- [codex_sub-agent_usage.md](/c:/Users/Justin/repos/invSys_fork/0%20plan%20docs/codex_sub-agent_usage.md)

Do not define agents beyond this set until the first-pass definitions are stable in practice.

## Global Definition Rules

- All agents inherit the repo sub-agent policy.
- All agents must be launched with the matching prompt block from the prompt doc.
- Only one agent may run Excel/COM or operate in an active Excel session at a time.
- Only `runtime-packaging` may touch packaged XLAM outputs, add-in registration, or `tools/build-xlam.ps1`.
- `coordinator-reviewer` owns decomposition and review, not default feature editing.
- If a task overlaps multiple write scopes, the Coordinator must serialize it.

## Recommended Agent Types

Use these heuristics:

- `default`
  Best for planning, review, bounded reasoning, and coordinator work.

- `worker`
  Best for implementation inside a well-bounded write scope.

Do not use explorer-style delegation for broad repo work unless the question is narrowly scoped and read-only.

## Reasoning Level Note

`Preferred reasoning level` in this document is a human guidance label, not a hard Codex product setting.

- `high`
  Use the strongest available reasoning mode or model for complex, coupled, or high-risk work.

- `medium`
  Use the normal/default worker setting for bounded implementation tasks.

## First-Pass Agent Set

### 1. `coordinator-reviewer`

- Agent type: `default`
- Primary role: decomposition, sequencing, review, and integration
- Write scope: none by default
- Allowed to write only by exception after re-scoping the task explicitly
- Preferred reasoning level: high

Use when:

- a task spans multiple ownership lanes
- the next step is deciding who owns what
- a diff needs scope review
- packaging/runtime work must be sequenced separately from logic work

Do not use when:

- the task is a single-lane, bounded code change that can be done directly by one worker

Launch pattern:

```text
Use prompt block: coordinator-reviewer
Mode: review and assignment first
Write scope: none unless explicitly granted after review
```

### 2. `runtime-packaging`

- Agent type: `worker`
- Primary role: build, package, COM Excel runtime, deployed XLAM behavior, Admin/runtime surfaces, auth/config/bootstrap ownership
- Preferred reasoning level: high

Write scope:

- [tools/build-xlam.ps1](/c:/Users/Justin/repos/invSys_fork/tools/build-xlam.ps1)
- [deploy](/c:/Users/Justin/repos/invSys_fork/deploy)
- [src/Admin](/c:/Users/Justin/repos/invSys_fork/src/Admin)
- [modAuth.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modAuth.bas)
- [modConfig.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modConfig.bas)
- [modGlobals.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modGlobals.bas)
- [modRoleUiAccess.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modRoleUiAccess.bas)
- [modRoleWorkbookSurfaces.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modRoleWorkbookSurfaces.bas)
- [modRuntimeWorkbooks.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modRuntimeWorkbooks.bas)

Use when:

- packaged XLAMs must be built or repaired
- Excel startup/load errors occur
- COM validation is required
- add-in registration or deployment state is part of the task
- Admin/runtime/bootstrap wiring is the task

Do not use when:

- the task is pure business logic in Receiving, Shipping, Production, or InventoryDomain

Scope guard:

- if this lane begins to drift outside its intended ownership in practice, the Coordinator should re-scope it to named files only before the next delegation
- Runtime/Packaging may execute test-harness-owned `run_phase*` validation scripts for Excel/COM validation, but those scripts remain `test-harness` write scope unless explicitly reassigned

Launch pattern:

```text
Use prompt block: runtime-packaging
Excel/COM: allowed
Parallelism: no other Excel/COM agent active
Expected output: build/load/runtime result plus cleanup of opened Excel sessions
```

### 3. `core-event-inventory`

- Agent type: `worker`
- Primary role: processor, locks, event writing, inventory apply, runtime-neutral shared logic
- Preferred reasoning level: high

Write scope:

- approved subset of [src/Core/Modules](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules):
  - `modProcessor.bas`
  - `modRoleEventWriter.bas`
  - `modLockManager.bas`
  - `modItemSearch.bas`
  - `modUR_*`
  - closely related runtime-neutral helpers approved by the Coordinator
- [src/InventoryDomain](/c:/Users/Justin/repos/invSys_fork/src/InventoryDomain)

Boundary note:

- `runtime-neutral helpers` does not include `modConfig.bas`, `modGlobals.bas`, `modAuth.bas`, or `modDiagramCore.bas`
- those files require explicit reassignment, not a helper-scope exception

Use when:

- the task is about event flow, idempotency, locks, apply logic, or shared inventory semantics

Do not use when:

- auth/config/bootstrap/runtime state is the issue
- packaged XLAM load/build behavior is involved

Launch pattern:

```text
Use prompt block: core-event-inventory
Excel/COM: not allowed
Packaging/build: not allowed
Expected output: runtime-neutral code change plus tests where appropriate
```

### 4. `role-receiving`

- Agent type: `worker`
- Primary role: Receiving flow only
- Preferred reasoning level: medium

Write scope:

- [src/Receiving](/c:/Users/Justin/repos/invSys_fork/src/Receiving)

Use when:

- the task is receiving UI, receiving event creation, or receiving-local helpers

Do not use when:

- the task needs packaging/runtime/build work
- the task primarily changes shared event semantics
- the task touches `modDiagramCore.bas`

Launch pattern:

```text
Use prompt block: role-receiving
Excel/COM: not allowed
Cross-role changes: escalate
```

### 5. `role-shipping`

- Agent type: `worker`
- Primary role: Shipping flow only
- Preferred reasoning level: medium

Write scope:

- [src/Shipping](/c:/Users/Justin/repos/invSys_fork/src/Shipping)

Use when:

- the task is shipping UI, shipping event creation, or shipping-local helpers

Do not use when:

- the task needs packaging/runtime/build work
- the task primarily changes shared event semantics
- the task touches `modDiagramCore.bas`

Launch pattern:

```text
Use prompt block: role-shipping
Excel/COM: not allowed
Cross-role changes: escalate
```

### 6. `role-production`

- Agent type: `worker`
- Primary role: Production flow only
- Preferred reasoning level: medium

Write scope:

- [src/Production](/c:/Users/Justin/repos/invSys_fork/src/Production)
- [src/DesignsDomain](/c:/Users/Justin/repos/invSys_fork/src/DesignsDomain) only when explicitly tied to approved Production work

Use when:

- the task is production UI, production event creation, or production-local helpers
- the change is in DesignsDomain only because it is directly tied to production behavior

Do not use when:

- the task needs packaging/runtime/build work
- the task primarily changes shared event semantics
- the task touches `modDiagramCore.bas`

Launch pattern:

```text
Use prompt block: role-production
Excel/COM: not allowed
DesignsDomain access: only when explicitly scoped
```

### 7. `test-harness`

- Agent type: `worker`
- Primary role: fixtures, validation scripts, regression coverage
- Preferred reasoning level: medium

Write scope:

- [tests](/c:/Users/Justin/repos/invSys_fork/tests)
- [create_phase1_fixture_xlsx.ps1](/c:/Users/Justin/repos/invSys_fork/tools/create_phase1_fixture_xlsx.ps1)
- [create_phase2_fixture_xlsx.ps1](/c:/Users/Justin/repos/invSys_fork/tools/create_phase2_fixture_xlsx.ps1)
- [run_phase1_excel_validation.ps1](/c:/Users/Justin/repos/invSys_fork/tools/run_phase1_excel_validation.ps1)
- [run_phase2_excel_validation.ps1](/c:/Users/Justin/repos/invSys_fork/tools/run_phase2_excel_validation.ps1)
- [run_phase2_excel_validation_visible.ps1](/c:/Users/Justin/repos/invSys_fork/tools/run_phase2_excel_validation_visible.ps1)
- [run_phase3_excel_validation.ps1](/c:/Users/Justin/repos/invSys_fork/tools/run_phase3_excel_validation.ps1)
- [run_phase4_excel_validation.ps1](/c:/Users/Justin/repos/invSys_fork/tools/run_phase4_excel_validation.ps1)
- [run_phase5_excel_validation.ps1](/c:/Users/Justin/repos/invSys_fork/tools/run_phase5_excel_validation.ps1)
- [run_phase6_excel_validation.ps1](/c:/Users/Justin/repos/invSys_fork/tools/run_phase6_excel_validation.ps1)

Use when:

- a failure must be made reproducible
- regression coverage should be added in parallel to an already-scoped implementation task

Do not use when:

- the real problem is packaging/runtime/build behavior
- the task requires installed add-in mutation or deployment-state repair

Launch pattern:

```text
Use prompt block: test-harness
Excel/COM: not allowed. If test validation requires Excel execution, delegate that execution step to runtime-packaging, not test-harness.
Expected output: test/fixture/harness changes only
```

## Deferred Lane

### `diagram-core`

- Status: deferred for release 1
- Default owner: `coordinator-reviewer` by exception only
- Scope:
  - [modDiagramCore.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modDiagramCore.bas)

Promotion trigger:

- the work becomes release-scoped
- the Coordinator assigns an explicit owner
- the policy doc and prompt doc are updated first

Do not instantiate this as an active sub-agent before those conditions are met.

## Launch Guidance

### Recommended rollout

Start with only:

- `coordinator-reviewer`
- `runtime-packaging`
- `core-event-inventory`
- one role agent at a time
- `test-harness` as a sidecar

### Sequence for mixed logic + packaging work

If a task touches both role/Core logic and packaged XLAM/runtime behavior:

1. assign the logic lane first
2. complete the bounded logic change
3. return to `coordinator-reviewer`
4. review whether packaging/runtime work is required
5. if yes, assign `runtime-packaging` separately
6. do not mix the two in one delegated prompt

### Parallel use

Safe parallel pattern:

- one implementation agent
- one test-harness sidecar
- coordinator reviewing

Unsafe parallel pattern:

- more than one agent editing `Core`
- any second agent while `runtime-packaging` is running Excel/COM
- multiple role agents touching shared contracts simultaneously

## Practical Bottom Line

This definition set is intentionally conservative.
It is designed for this repo's actual failure modes:

- shared `Core` gravity
- packaged XLAM/runtime fragility
- Excel COM state
- deployment ambiguity

If those become stable, this set can expand later.
Until then, fewer agents with harder boundaries is the correct model.

