# invSys Agent Instructions

## Scope and repository locations

These instructions apply to work in both invSys repositories:

- code repository: the repository containing this file, normally
  `/mnt/c/Users/justu/source/repos/invSys_fork`;
- documentation repository: the sibling repository `../invSys_docs`, normally
  `/mnt/c/Users/justu/source/repos/invSys_docs`.

Resolve the documentation repository relative to the code-repository root.
Do not assume the current working directory is either repository root. If the
expected sibling repository or a required pointer is missing, report that
condition instead of guessing from another clone or from GitHub.

## Authority and precedence

Repository architectural precedence is:

1. The current normative invSys design specification.
2. The explicitly designated current implementation plan.
3. The current handoff/baton.
4. Generated implementation and runtime evidence.
5. Historical guidance and handoffs.

This precedence applies to repository architecture and project records. It does
not override system, developer, or current user instructions.

Generated manifests and runtime reports describe observed reality. They never
override the normative specification.

If the current implementation plan or handoff contradicts the normative
specification:

1. stop before implementing the conflicting behavior;
2. cite the conflicting sections and describe the concrete conflict;
3. follow the normative specification for implementation;
4. propose the required plan or handoff correction; and
5. if the user intends to change the architecture, update and approve the
   normative specification before implementing the new architecture.

Do not silently reconcile a spec/plan contradiction by inventing a hybrid.

## Read at session start

From the code-repository root, resolve and read in this order:

1. this `AGENTS.md`;
2. `../invSys_docs/0 plan docs/xlam_invSys/CURRENT_SPEC.md`;
3. the normative specification named by that pointer;
4. `../invSys_docs/expert guidance docs/CURRENT.md`;
5. the implementation plan named by that pointer;
6. `../invSys_docs/0 plan docs/xlam_invSys/invSys-Controls-v1.md`;
7. `../invSys_docs/last handoff/CURRENT.md`;
8. the handoff named by that pointer, if one exists;
9. additional applicable `AGENTS.md` files; and
10. the current Git branch, status, and existing uncommitted changes in both
   repositories.

To determine additional applicable `AGENTS.md` files for a target file, start
at that file's repository root and walk each directory on the path to the
target file. Read every `AGENTS.md` encountered in root-to-leaf order. The
closest file applies most specifically within its directory tree, but no
nested instruction may override the normative specification.

Re-run this path check before changing files outside the directories initially
inspected.

The current plan applies to Receiving, Production, Boxing, Shipping, shared
Operations packaging, developer tooling, and related Domain/Core work according
to its declared scope.

`../invSys_docs/0 plan docs/xlam_invSys/invSys-Controls-v1.md` is the maintained
control catalog and user-acceptance record. Update it whenever a form or Ribbon
control, displayed wording, reachability, sizing behavior, workflow contract,
or visible acceptance status changes. It is subordinate to the normative
specification and current implementation plan and must not be used to redefine
architecture silently.

Prefer the local repositories. Use GitHub only when remote state, PR state, or
review information must be verified.

Do not infer the current specification, plan, or handoff from filesystem
timestamps or the highest filename alone.

## GitHub contribution workflow — this computer only

The user confirmed this workflow on 2026-10-16 for **this computer only** (the
current Windows laptop). It applies to both local repositories and persists
across chats on this computer. Although this file is versioned, this section does
not establish the GitHub workflow for another computer or clone. Other computers
must use their own user-authorized workflow; do not automatically copy these
remote, push or PR defaults to them.

- `soetrain/invSys_fork` and `soetrain/invSys_docs`, branch `main`, are the
  authoritative upstream repositories. Remote `origin` points to the matching
  soetrain repository. Pull the latest from these repositories when resuming work.
- The authenticated contributor is `justinwj`. Remote `fork` points to the
  matching `justinwj/invSys_fork` or `justinwj/invSys_docs` repository. Publishing
  feature branches there is authorized as part of submitting PRs to soetrain.
- Preserve the active feature branch and local work. Fetch `origin`, then merge
  `origin/main` into the feature branch when needed. Do not rebase published
  history, force-push, discard changes, or replace newer local work with an older
  upstream tree. Local `main` tracks `origin/main`; preserve divergent local-only
  history before realigning a stale local `main`.
- Push the feature branch to `fork`, then create or update its PR **in the
  corresponding soetrain repository**, targeting `main`. Keep incomplete or
  unqualified implementation as a draft PR with explicit remaining gates. Do not
  stop at pushing a branch without checking that the upstream PR exists.
- Set the GitHub CLI default repository to the matching soetrain repository.
  Keep `remote.pushDefault=fork`, the feature branch's `pushRemote=fork`, and
  `push.default=current`; feature branches may track their published fork branch.
  Use explicit upstream fetch/merge commands rather than assuming an unqualified
  `git pull` on a fork-tracking feature branch pulls from soetrain.
- Soetrain reviews and merges the PR, then pulls from his soetrain account.
  Do not merge PRs, enable auto-merge, or push directly to soetrain `main` without
  a separate explicit instruction.
- Before ending, verify both PR base repositories/branches, published head SHAs,
  and local Git status. Report the PR links and any actual remaining sync blocker.

## Before changing implementation

Before editing VBA, forms, RibbonX, build/deployment scripts, schemas, tests, or
contract documentation for Receiving, Production, Boxing, Shipping, Operations,
Admin, Core, or Domain, report:

- active slice number and name;
- role/package affected;
- contract being changed;
- focused test that protects the contract;
- expected behavioral RED; and
- files/packages expected to change.

Under D13, create and run the focused test before changing implementation when
no suitable test exists.

A compile failure, missing fixture, unavailable workbook, or broken harness is
not meaningful RED.

For form and Ribbon work, the protecting test must exercise the same public
callback or form-action handler used by the operator whenever practical.
Direct service tests supplement packaged action tests; they do not replace
them.

## Handling new user requests

Do not silently force a new request into the active slice.

When a request arrives:

1. determine whether it is part of the active slice, a newly discovered blocker,
   non-contract work, or a deliberate priority change;
2. state any resulting slice change;
3. flag architectural conflict with D12 or D13 before implementation; and
4. follow the user's direction after explaining the concrete conflict.

## Architectural invariants

- The exact managed inventory identity header is `System_Key`.
- `System_Key` is generated once at the owning creation/service boundary,
  system-wide unique, immutable, and preserved through sorting, refresh,
  save/reopen, movement, condition changes, snapshots, and projection rebuild.
- `ITEM_CODE`/SKU identifies what an item is. It does not identify one durable
  inventory entity.
- Location, quantity, `Condition`, and custom fields are attributes, not
  identity.
- `ROW` is prohibited as a managed runtime header, migration key, display key,
  compatibility field, or authority path.
- This is a greenfield reset. Do not import, translate, reconcile, repair, or
  map old business inventory into `System_Key`.
- Supported test inventory begins with Admin Generate Warehouse/Create
  Warehouse and optional bootstrap or Admin `Seed Demo Inventory`.
- Managed tables define a required header subset and tolerate additional
  end-user columns. Resolve managed columns by normalized header name, never
  ordinal position, and preserve unknown columns through refresh/resize/rebuild.
- Shared custom values persist by `System_Key`; workbook-only display columns
  remain local.
- `Condition` is a managed inventory header and seeded demo inventory defaults
  it to `GOOD`.
- When `DesignsEnabled=True`, Production reads released Designs Domain recipes
  and must not silently fall back to legacy recipe storage.
- Within one VBA project, use direct typed procedure calls.
- Use `Application.Run` only at declared cross-XLAM compatibility or bridge
  boundaries.
- Cross-XLAM contracts use declared primitive or serialized result envelopes.
- Operator workbooks and forms are projections/staging surfaces, not canonical
  Domain authority.
- Core and Domain add-ins remain headless.
- Receiving, Production, Boxing, and Shipping are packaged in
  `invSys.Operations.xlam` under D12.

## Editing, generated evidence, and security

- Preserve unrelated user changes in a dirty working tree.
- Do not overwrite, delete, or regenerate operational workbooks without explicit
  authorization.
- Do not deploy or rebuild XLAMs while Excel has the relevant workbooks/add-ins
  open.
- Never place passwords, tokens, PINs, credential hashes, service credentials,
  or Windows credential material in source, documentation, generated reports,
  tests, fixtures, logs, or handoffs.
- Do not commit generated runtime reports containing machine, session, user,
  inventory, customer, shipment, recipe, or credential-sensitive data.
- Runtime-state extraction is read-only by default. It must not mutate, save,
  refresh, repair, process, recalculate, or close an operational workbook or
  add-in.
- Runtime-state extraction must redact secrets and sensitive values before
  writing JSON, Markdown, logs, fixtures, or console output.
- Row-level operational values require an explicit diagnostic opt-in and a
  documented redaction policy. Default reports contain schemas, counts,
  identifiers, statuses, versions, and hashes only.
- Machine/runtime reports are ignored by Git by default. Review and sanitize a
  report before attaching it to an issue, PR, handoff, or test record.
- Tool B changes require redaction tests and a before/after proof that inspected
  operational workbooks were not changed.
- Do not delete scanner-reported dead code automatically. Require reviewed
  reachability evidence, compile success, and protecting regression tests.

## Completion evidence

For contract-affecting implementation work, a slice is not complete until:

- focused RED and GREEN are recorded;
- applicable packaged form-action tests pass;
- relevant regression tests pass;
- static maintenance evidence is regenerated;
- code-bloat and dynamic-call metrics do not regress without an explicit
  exception;
- documentation affected by the contract is updated; and
- Git status is reviewed for unintended generated or duplicate files.

Manual success alone does not satisfy D13.

For non-contract work, such as spelling corrections, pointer maintenance,
comment-only changes, repository housekeeping, or documentation formatting,
state why D13 does not apply and validate proportionally. At minimum, inspect
the resulting diff, run relevant formatting or link checks when available, and
review Git status. Do not manufacture a RED test for work that changes no
runtime or architectural contract.

## Handoff policy

Create or update a handoff only when the chat/work session is ending or the user
explicitly requests one.

Store handoffs in:

```text
../invSys_docs/last handoff/
```

Use the next available zero-padded numeric prefix and a descriptive Markdown
filename. After creating a handoff, update:

```text
../invSys_docs/last handoff/CURRENT.md
```

The pointer must name the handoff explicitly. Do not select a handoff by
filesystem modification time.

A handoff is a decision-centered continuation record, not a chronological chat
summary. Preserve durable facts, decisions, constraints, unresolved questions,
and exact next actions. Do not infer missing details or turn tentative ideas
into confirmed facts.

Use these compact sections:

1. **Goal and release outcome**
   - State the current goal in one sentence.
   - Identify the specification requirement and acceptance evidence it advances.
2. **Current verified state**
   - Record the code and documentation branches and latest commits.
   - Record the active slice, completed gates, modified/uncommitted files, and
     whether Excel must be closed before continuing.
   - Add `Last verified` dates to facts likely to become stale.
3. **Decisions and constraints**
   - Record confirmed decisions, relevant specification references, and user
     preferences that must survive the session boundary.
   - Label tentative ideas and proposals explicitly.
4. **Evidence and traceability**
   - Record exact focused/regression results and generated evidence used.
   - For an unresolved defect, map symptom -> known or suspected root cause ->
     governing requirement -> protecting or required test.
5. **Do Not Repeat**
   - Record failed approaches, rejected designs, and dead ends only when doing
     so prevents likely repeated work.
6. **Assumptions to Re-verify**
   - List facts that may be stale, environment-dependent, or not yet proven.
7. **Open questions and blockers**
   - Mark unresolved items as unresolved; never fill gaps with a best guess.
8. **Immediate next action**
   - State one concrete test-first next action in one sentence.
9. **Critical references**
   - List only the exact files, procedures, forms, sheets, tables, controls,
     event IDs, or other identifiers needed to resume safely.

Keep the handoff concise enough to load at the beginning of a fresh session.
Prefer references to authoritative files and generated evidence over copied
narrative. Prioritize the release outcome over the most recent visible symptom;
include a local defect only when it blocks or materially advances an applicable
acceptance criterion.

Do not include secrets or unredacted runtime data in a handoff.

Do not use handoffs as a substitute for updating the specification, current
plan, tests, or generated evidence.

Commit and Sync/Push to Github before exiting prompt.
