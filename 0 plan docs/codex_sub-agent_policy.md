# Codex Sub-Agent Policy for invSys

## Purpose

This document defines the first practical sub-agent operating model for this repo.
It is intentionally narrow.
The goal is not to maximize agent count.
The goal is to reduce collisions, protect packaged Excel/runtime work, and keep changes reviewable.

For practical launch and handoff use, also see:

- [codex_sub-agent_usage.md](/c:/Users/Justin/repos/invSys_fork/0%20plan%20docs/codex_sub-agent_usage.md)

This policy is organized into four implementation layers:

1. Coordinator rules
2. Path allow-lists and deny-lists
3. Agent-by-agent prompts
4. Codex sub-agent definitions

The dependency order is strict:

`Coordinator rules -> path ownership matrix -> agent prompts -> sub-agent definitions`

Do not create or use sub-agents for this repo until all four layers are defined.

## Core Principles

- Runtime/packaging is a hard single-owner lane.
- Only one agent may run Excel/COM automation in a session.
- `Core` is shared infrastructure and orchestration, not a dumping ground.
- Every writable path must have exactly one owner at a time.
- Every agent needs both an allow-list and a deny-list.
- The Coordinator is a reviewer and gatekeeper, not only a router.
- If write scopes overlap, work must be serialized.

## 1. Coordinator Rules

### Coordinator Role

The Coordinator owns:

- task decomposition
- write-scope assignment
- sequencing
- conflict detection
- diff review
- final integration decision

The Coordinator does **not** default to editing business logic directly.

### Hard Rules

- Only the Runtime/Packaging agent may touch packaged XLAM outputs, build scripts, add-in registration state, or Excel COM load validation.
- **Only one agent may run Excel/COM or operate in an active Excel session at a time.**
- No parallel Excel sessions across agents.
- No agent may edit outside its allow-list without explicit Coordinator approval.
- No task may be split across multiple agents if the immediate next step depends on one blocking result.
- Any change touching `deploy/`, XLAM packaging, or COM automation must be isolated from unrelated edits.

### Review Rules

The Coordinator must reject or re-scope a sub-agent diff if it:

- edits outside the assigned path scope
- adds new business logic into unrelated `Core` modules
- modifies build/deploy/runtime files without being the Runtime/Packaging agent
- mixes feature work and packaging work in one change
- changes tests in a way that hides runtime failures instead of exposing them

If Runtime/Packaging begins to drift outside its intended lane in practice, the Coordinator must tighten that lane from mixed path scope to named-files-only scope before delegating further Runtime/Packaging work.

### Delegation Rules

Delegate only when:

- the write set is bounded
- the subtask materially advances the main task
- the result does not require simultaneous editing in the same files by another agent

Do not delegate when:

- Excel/COM runtime state is involved
- packaged add-ins are loaded
- the issue is likely in build/import/export semantics
- the next step is blocked on one tightly coupled investigation

## 2. Path Ownership Matrix

This is the first-pass ownership model for this repo.

### A. Runtime/Packaging Agent

Allow-list:

- [tools/build-xlam.ps1](/c:/Users/Justin/repos/invSys_fork/tools/build-xlam.ps1)
- [deploy](/c:/Users/Justin/repos/invSys_fork/deploy)
- `tools/*xlam*`
- runtime bootstrap/load helpers in:
  - [modRuntimeWorkbooks.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modRuntimeWorkbooks.bas)
  - [modRoleWorkbookSurfaces.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modRoleWorkbookSurfaces.bas)
- auth/config/system-surface modules in:
  - [modAuth.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modAuth.bas)
  - [modConfig.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modConfig.bas)
  - [modGlobals.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modGlobals.bas)
  - [modRoleUiAccess.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modRoleUiAccess.bas)
- [src/Admin](/c:/Users/Justin/repos/invSys_fork/src/Admin)

Deny-list:

- role business logic
- inventory mutation logic
- event/inventory semantics outside bootstrap/auth/config/runtime behavior
- test assertions unrelated to packaging/runtime
- test-harness-owned fixture and validation scripts unless explicitly reassigned by the Coordinator

Special rule:

- single owner for Excel/COM, packaged XLAM rebuilds, add-in registration, and deployment validation
- release 1 owner for Admin surfaces, auth/config bootstrap, and shared runtime state
- if repeated scope drift appears in this lane, reduce Runtime/Packaging scope to named files only until the drift is resolved
- Runtime/Packaging may execute test-harness-owned `run_phase*` validation scripts for Excel/COM validation, but execution rights do not imply write ownership

### B. Core Event/Inventory Agent

Allow-list:

- [src/Core/Modules](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules) limited to:
  - `modProcessor.bas`
  - `modRoleEventWriter.bas`
  - `modLockManager.bas`
  - `modItemSearch.bas`
  - `modUR_*`
  - runtime-neutral shared helpers
- [src/InventoryDomain](/c:/Users/Justin/repos/invSys_fork/src/InventoryDomain)

Deny-list:

- [deploy](/c:/Users/Justin/repos/invSys_fork/deploy)
- [tools/build-xlam.ps1](/c:/Users/Justin/repos/invSys_fork/tools/build-xlam.ps1)
- any Excel-opening or add-in-mutating script under [tools](/c:/Users/Justin/repos/invSys_fork/tools)
- role UI modules
- Excel COM validation scripts
- [modAuth.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modAuth.bas)
- [modConfig.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modConfig.bas)
- [modGlobals.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modGlobals.bas)
- [src/Admin](/c:/Users/Justin/repos/invSys_fork/src/Admin)
- [modDiagramCore.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modDiagramCore.bas)

### C. Role Agent

Only one role agent should be active at a time unless write scopes are fully disjoint.

Receiving allow-list:

- [src/Receiving](/c:/Users/Justin/repos/invSys_fork/src/Receiving)

Shipping allow-list:

- [src/Shipping](/c:/Users/Justin/repos/invSys_fork/src/Shipping)

Production allow-list:

- [src/Production](/c:/Users/Justin/repos/invSys_fork/src/Production)
- [src/DesignsDomain](/c:/Users/Justin/repos/invSys_fork/src/DesignsDomain) when explicitly tied to Production work

Common deny-list for role agents:

- [deploy](/c:/Users/Justin/repos/invSys_fork/deploy)
- [tools/build-xlam.ps1](/c:/Users/Justin/repos/invSys_fork/tools/build-xlam.ps1)
- any Excel-opening or add-in-mutating script under [tools](/c:/Users/Justin/repos/invSys_fork/tools)
- unrelated role folders
- shared `Core` modules outside explicitly approved integration hooks
- [src/Admin](/c:/Users/Justin/repos/invSys_fork/src/Admin)
- [modAuth.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modAuth.bas)
- [modConfig.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modConfig.bas)
- [modGlobals.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modGlobals.bas)
- [modDiagramCore.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modDiagramCore.bas)

### D. Test Harness Agent

Allow-list:

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

Deny-list:

- [deploy](/c:/Users/Justin/repos/invSys_fork/deploy)
- [tools/build-xlam.ps1](/c:/Users/Justin/repos/invSys_fork/tools/build-xlam.ps1)
- any Excel-opening or add-in-mutating script under [tools](/c:/Users/Justin/repos/invSys_fork/tools)
- installed add-in state
- production business logic except minimal test hooks approved by the Coordinator
- [src/Admin](/c:/Users/Justin/repos/invSys_fork/src/Admin)
- [modAuth.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modAuth.bas)
- [modConfig.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modConfig.bas)
- [modGlobals.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modGlobals.bas)

## 3. Agent Prompts

These are the first-pass behavioral contracts.

### Coordinator / Reviewer

Mission:

- Break the task into bounded subtasks.
- Assign one write scope per agent.
- Review diffs for scope violations.
- Sequence risky changes.

Definition of done:

- each subtask has a bounded write set
- no path ownership conflicts remain
- runtime/build work is isolated
- final integrated diff is coherent

Escalate when:

- two agents need the same file
- Excel runtime state is required
- packaged add-ins are already loaded

### Runtime/Packaging Agent

Mission:

- Own build, package, XLAM import/export, COM Excel validation, and deployed runtime surfaces.

Definition of done:

- packaged files build cleanly
- `build-xlam.ps1 -Apply` completes without error
- add-ins load without compile/runtime startup failures
- packaged XLAM smoke load completes without startup compile error
- no stale deployment ambiguity remains

Must not do:

- redesign role workflows
- implement business rules unrelated to load/runtime

### Core Event/Inventory Agent

Mission:

- Own processor, event orchestration, locks, inventory apply, and shared event/search helpers.

Definition of done:

- logic is runtime-neutral
- schema and idempotency remain coherent
- tests cover the changed event/apply path

Must not do:

- packaging
- deployed XLAM manipulation
- role workbook UI redesign

### Role Agent

Mission:

- Own one role's entry flow, UI behavior, event creation, and role-local helpers.

Definition of done:

- the role flow works end to end against its current shared contracts
- no unrelated role is changed
- any shared contract change is kicked back to Coordinator first

Must not do:

- package XLAMs
- modify unrelated role modules
- change cross-role inventory semantics alone

### Test Harness Agent

Mission:

- Add regression coverage, fixtures, and isolated validation paths for the target scenario.

Definition of done:

- a failing behavior becomes reproducible
- a passing behavior remains documented by test output

Must not do:

- own packaging/runtime deployment
- mutate installed add-ins

## 4. Codex Sub-Agent Definitions

These are the initial sub-agents worth defining.
Do not define more until these are stable.

### `coordinator-reviewer`

- Type: `default`
- Responsibility: decomposition, sequencing, review, integration
- Write scope: none by default

### `runtime-packaging`

- Type: `worker`
- Responsibility: build, package, COM Excel runtime, deployed XLAM behavior, Admin/runtime surfaces, and auth/config/bootstrap ownership
- Write scope:
  - [tools/build-xlam.ps1](/c:/Users/Justin/repos/invSys_fork/tools/build-xlam.ps1)
  - [deploy](/c:/Users/Justin/repos/invSys_fork/deploy)
  - runtime/bootstrap packaging helpers
  - [src/Admin](/c:/Users/Justin/repos/invSys_fork/src/Admin)
  - [modAuth.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modAuth.bas)
  - [modConfig.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modConfig.bas)
  - [modGlobals.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modGlobals.bas)
  - [modRoleUiAccess.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modRoleUiAccess.bas)
  - [modRoleWorkbookSurfaces.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modRoleWorkbookSurfaces.bas)
  - [modRuntimeWorkbooks.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modRuntimeWorkbooks.bas)

### `core-event-inventory`

- Type: `worker`
- Responsibility: processor, locks, event writing, inventory apply, runtime-neutral shared logic
- Write scope:
  - [src/Core/Modules](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules) approved subset
  - [src/InventoryDomain](/c:/Users/Justin/repos/invSys_fork/src/InventoryDomain)

### `diagram-core`

- Status: deferred for release 1
- Owner for now: `coordinator-reviewer` by exception only
- Scope:
  - [modDiagramCore.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modDiagramCore.bas)
- Rule:
  - no delegated work unless diagramming becomes a release-scoped feature

### Deferred Item Promotion Trigger

A deferred lane may be promoted only when all of the following are true:

- the work becomes release-scoped rather than speculative
- the Coordinator assigns an explicit owner
- the policy doc and prompt doc are both updated before delegation starts

### `role-receiving`

- Type: `worker`
- Responsibility: Receiving flow only
- Write scope:
  - [src/Receiving](/c:/Users/Justin/repos/invSys_fork/src/Receiving)

### `role-shipping`

- Type: `worker`
- Responsibility: Shipping flow only
- Write scope:
  - [src/Shipping](/c:/Users/Justin/repos/invSys_fork/src/Shipping)

### `role-production`

- Type: `worker`
- Responsibility: Production flow only
- Write scope:
  - [src/Production](/c:/Users/Justin/repos/invSys_fork/src/Production)
  - [src/DesignsDomain](/c:/Users/Justin/repos/invSys_fork/src/DesignsDomain) when approved

### `test-harness`

- Type: `worker`
- Responsibility: fixtures, validation scripts, regression coverage
- Write scope:
  - [tests](/c:/Users/Justin/repos/invSys_fork/tests)
  - approved non-packaging validation scripts under [tools](/c:/Users/Justin/repos/invSys_fork/tools)

## Initial Rollout Recommendation

Start with only these active lanes:

- `coordinator-reviewer`
- `runtime-packaging`
- `core-event-inventory`
- one role agent at a time
- `test-harness` as a sidecar

Do not run multiple role agents plus runtime/packaging in the same active Excel session.

## Repo-Specific Non-Negotiables

- If Excel is open, assume runtime state is mutable and dangerous.
- If add-ins are installed, do not rebuild over them without explicit session control.
- If a task touches `deploy/` and role logic, split it into separate steps.
- If a task requires COM Excel inspection, serialize the entire workflow through the Runtime/Packaging agent.
- [modDiagramCore.bas](/c:/Users/Justin/repos/invSys_fork/src/Core/Modules/modDiagramCore.bas) is frozen for release 1 unless explicitly approved by the Coordinator.

## Practical Bottom Line

Sub-agents can help this repo, but only if ownership is enforced through:

- a Coordinator who reviews, not just routes
- path allow-lists
- explicit deny-lists
- hard single-ownership of runtime/packaging work

Without those controls, multi-agent work will amplify Excel/runtime failures instead of reducing them.
