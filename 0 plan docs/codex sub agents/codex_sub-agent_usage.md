P# Codex Sub-Agent Usage for invSys

## Purpose

This document explains how to use the sub-agent stack in practice.
It is the operational companion to:

- [codex_sub-agent_policy.md](/c:/Users/Justin/repos/invSys_fork/0%20plan%20docs/codex_sub-agent_policy.md)
- [codex_sub-agent_prompts.md](/c:/Users/Justin/repos/invSys_fork/0%20plan%20docs/codex_sub-agent_prompts.md)
- [codex_sub-agent_definitions.md](/c:/Users/Justin/repos/invSys_fork/0%20plan%20docs/codex_sub-agent_definitions.md)

Use this doc when you are about to delegate real work.

## How To Use The Docs In Practice

Use the docs in this order:

1. Read the policy doc first.
   Confirm the task is safe to delegate, identify the ownership lane, and check whether Excel/COM, `deploy/`, or packaged XLAM behavior is involved.

2. Check the definitions doc next.
   Pick the correct agent, confirm its write scope, and make sure the task does not belong to a different lane.

3. Copy the matching prompt block.
   Use the prompt doc to give the sub-agent its operating rules, boundaries, and reporting format.

4. Add a task-specific launch wrapper.
   State the exact task, allowed paths, forbidden paths, definition of done, and stop conditions.

5. Review the diff before any follow-on delegation.
   The Coordinator should check for scope drift, deny-list violations, and hidden packaging/runtime bleed before assigning the next lane.

6. If Runtime/Packaging is needed, serialize it.
   Do not let a role/Core/test agent and Runtime/Packaging work the same active Excel/runtime surface at the same time.

## Practical Sequence

For most work, use one of these sequences.

### Single-Lane Logic Change

1. Coordinator identifies the lane.
2. Launch one worker with the matching prompt and a bounded task wrapper.
3. Review the diff.
4. If needed, launch `test-harness` as a sidecar or follow-up.

### Logic Plus Packaging Change

1. Coordinator assigns the logic lane first.
2. Worker completes the bounded logic change only.
3. Coordinator reviews for scope drift.
4. If packaged build/load/runtime work is required, Coordinator launches `runtime-packaging` separately.
5. Runtime/Packaging runs build/load validation in isolation.

### Runtime Failure Investigation

1. Coordinator confirms the failure is really in build/load/runtime behavior.
2. Launch `runtime-packaging` directly.
3. Do not launch other workers into Excel/COM or `deploy/` while that work is active.
4. Return to Coordinator after runtime state is stabilized.

## Coordinator Checklist

Before delegation:

- identify the smallest ownership lane
- confirm the task does not cross multiple write scopes
- check whether Excel/COM is involved
- check whether `deploy/` or `tools/build-xlam.ps1` is involved
- decide whether the task is implementation, test, or runtime/package work

After delegation:

- verify changed files stayed inside scope
- verify deny-list files were not touched
- verify Runtime/Packaging did not expand beyond intended ownership
- tighten Runtime/Packaging to named-files-only scope if drift appears
- decide whether a second lane is actually needed

## Launch Template

Use this as the standard handoff format.

```text
Agent:

Task:

Allowed paths:

Forbidden paths:

Definition of done:

Stop and return to Coordinator if:

Required output:
1. What changed
2. Which files changed
3. What was intentionally not touched
4. Risks or follow-up for review
```

## Filled Templates

### Coordinator / Reviewer

```text
Agent: coordinator-reviewer

Task: classify the request, assign the correct ownership lane, and review resulting diffs for boundary violations

Allowed paths:
- none by default

Forbidden paths:
- no default write scope

Definition of done:
- correct owner selected
- write scope explicitly stated
- deny scope explicitly stated
- resulting diff reviewed for scope drift

Stop and return to Coordinator if:
- not applicable; this is the Coordinator lane

Required output:
1. Assigned owner
2. Allowed write scope
3. Forbidden scope
4. Review result and next lane if needed
```

### Runtime / Packaging

```text
Agent: runtime-packaging

Task: [fill in exact packaging/runtime task]

Allowed paths:
- tools/build-xlam.ps1
- deploy/*
- src/Admin/*
- src/Core/Modules/modAuth.bas
- src/Core/Modules/modConfig.bas
- src/Core/Modules/modGlobals.bas
- src/Core/Modules/modRoleUiAccess.bas
- src/Core/Modules/modRoleWorkbookSurfaces.bas
- src/Core/Modules/modRuntimeWorkbooks.bas

Forbidden paths:
- role business logic
- InventoryDomain business logic
- tests and test-harness-owned fixture/validation scripts unless explicitly reassigned by the Coordinator

Definition of done:
- build/load/runtime task completed
- no stale deployment ambiguity remains
- Excel/COM session cleaned up

Stop and return to Coordinator if:
- the real issue is feature logic rather than packaging/runtime
- requested changes would expand this lane beyond intended ownership
- the task requires editing a `tools/run_phase*` validation script rather than just executing it

Required output:
1. What changed
2. Which files changed
3. Whether Excel/COM was used
4. Remaining runtime risk or deployment ambiguity
```

### Core Event / Inventory

```text
Agent: core-event-inventory

Task: [fill in exact event/inventory task]

Allowed paths:
- src/Core/Modules/modProcessor.bas
- src/Core/Modules/modRoleEventWriter.bas
- src/Core/Modules/modLockManager.bas
- src/Core/Modules/modItemSearch.bas
- src/Core/Modules/modUR_*
- approved runtime-neutral helpers
- src/InventoryDomain/*

Forbidden paths:
- deploy/*
- tools/build-xlam.ps1
- Excel/COM scripts
- src/Core/Modules/modAuth.bas
- src/Core/Modules/modConfig.bas
- src/Core/Modules/modGlobals.bas
- src/Core/Modules/modDiagramCore.bas
- src/Admin/*

Definition of done:
- runtime-neutral logic change complete
- schema/idempotency/lock semantics remain coherent
- tests updated where appropriate

Stop and return to Coordinator if:
- auth/config/bootstrap/runtime files are needed
- packaged XLAM/load behavior is implicated
- the helper exception would require touching modAuth.bas, modConfig.bas, modGlobals.bas, or modDiagramCore.bas

Required output:
1. What changed
2. Which files changed
3. What was intentionally not touched
4. Risks or shared-contract implications
```

### Role Receiving

```text
Agent: role-receiving

Task: [fill in exact Receiving task]

Allowed paths:
- src/Receiving/*

Forbidden paths:
- src/Shipping/*
- src/Production/*
- src/Admin/*
- deploy/*
- tools/build-xlam.ps1
- Excel/COM scripts
- src/Core/Modules/modAuth.bas
- src/Core/Modules/modConfig.bas
- src/Core/Modules/modGlobals.bas
- src/Core/Modules/modDiagramCore.bas
- shared Core modules outside approved hooks

Definition of done:
- Receiving-only change complete
- no cross-role drift
- no packaging/runtime edits

Stop and return to Coordinator if:
- shared Core contract change is required
- packaged XLAM/runtime behavior is implicated

Required output:
1. What changed
2. Which files changed
3. What was intentionally not touched
4. Risks or follow-up
```

### Role Shipping

```text
Agent: role-shipping

Task: [fill in exact Shipping task]

Allowed paths:
- src/Shipping/*

Forbidden paths:
- src/Receiving/*
- src/Production/*
- src/Admin/*
- deploy/*
- tools/build-xlam.ps1
- Excel/COM scripts
- src/Core/Modules/modAuth.bas
- src/Core/Modules/modConfig.bas
- src/Core/Modules/modGlobals.bas
- src/Core/Modules/modDiagramCore.bas
- shared Core modules outside approved hooks

Definition of done:
- Shipping-only change complete
- no cross-role drift
- no packaging/runtime edits

Stop and return to Coordinator if:
- shared Core contract change is required
- packaged XLAM/runtime behavior is implicated

Required output:
1. What changed
2. Which files changed
3. What was intentionally not touched
4. Risks or follow-up
```

### Role Production

```text
Agent: role-production

Task: [fill in exact Production task]

Allowed paths:
- src/Production/*
- src/DesignsDomain/* only when explicitly tied to approved Production work

Forbidden paths:
- src/Receiving/*
- src/Shipping/*
- src/Admin/*
- deploy/*
- tools/build-xlam.ps1
- Excel/COM scripts
- src/Core/Modules/modAuth.bas
- src/Core/Modules/modConfig.bas
- src/Core/Modules/modGlobals.bas
- src/Core/Modules/modDiagramCore.bas
- shared Core modules outside approved hooks

Definition of done:
- Production-only change complete
- DesignsDomain touched only if explicitly scoped
- no packaging/runtime edits

Stop and return to Coordinator if:
- shared Core contract change is required
- packaged XLAM/runtime behavior is implicated
- work starts to drift toward diagram-core behavior

Required output:
1. What changed
2. Which files changed
3. What was intentionally not touched
4. Risks or follow-up
```

### Test Harness

```text
Agent: test-harness

Task: [fill in exact test/fixture task]

Allowed paths:
- tests/*
- tools/create_phase1_fixture_xlsx.ps1
- tools/create_phase2_fixture_xlsx.ps1
- tools/run_phase1_excel_validation.ps1
- tools/run_phase2_excel_validation.ps1
- tools/run_phase2_excel_validation_visible.ps1
- tools/run_phase3_excel_validation.ps1
- tools/run_phase4_excel_validation.ps1
- tools/run_phase5_excel_validation.ps1
- tools/run_phase6_excel_validation.ps1

Forbidden paths:
- deploy/*
- tools/build-xlam.ps1
- installed add-in state
- Excel/COM packaging/runtime workflows
- src/Admin/*
- src/Core/Modules/modAuth.bas
- src/Core/Modules/modConfig.bas
- src/Core/Modules/modGlobals.bas

Definition of done:
- failure reproduced or coverage added
- harness remains realistic
- no packaging/runtime repair is hidden in test code

Stop and return to Coordinator if:
- Excel execution is required
- the real issue is deployment/build/runtime state
- feature logic changes are needed outside a minimal approved test hook

Required output:
1. What changed
2. Which files changed
3. What was intentionally not touched
4. Remaining gaps or follow-up
```

## Bottom Line

The policy doc tells you the rules.
The prompt doc tells you how each agent should behave.
The definitions doc tells you which agent to use.
This doc tells you how to launch the work without improvising the handoff each time.
