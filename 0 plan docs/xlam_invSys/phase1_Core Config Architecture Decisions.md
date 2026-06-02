# Core.Config Architecture Decisions — invSys R1

## Purpose

This document locks the ten architectural decisions required before coding `Core.Config` for Release 1 of the invSys multi-warehouse inventory system. Each decision is grounded in the authoritative design document (`invSys-Design-v4.2.md`) and the consolidated design spec, with the recommended path justified against R1 constraints (VBA-only, Excel + SharePoint, offline-first).

***

## Decision 1 — Authoritative Config Source

**Question:** Is `WHx.invSys.Config.xlsb` the sole authoritative source, or should workbook-local overrides also exist?

**Recommendation: `WHx.invSys.Config.xlsb` is the single authoritative source for R1. No workbook-local overrides.**

The design doc defines exactly one config workbook per warehouse (`WHx.invSys.Config.xlsb`) with two tables: `tblWarehouseConfig` and `tblStationConfig`. The architecture's "clear ownership boundary" principle (D3) assigns config loading exclusively to Core. Introducing a second source (e.g., per-workbook named ranges or hidden sheets) would:

- Create ambiguity about which value is authoritative when they diverge
- Complicate troubleshooting across stations
- Violate the single-source-of-truth principle that underpins event sourcing

**R2 hook:** If workbook-local overrides become necessary (e.g., station-specific debug flags), they can be added later as a new `tblLocalOverrides` table with explicit documentation that it overrides the warehouse config. The `Core.Config` API below is designed to accommodate this without breaking callers.

***

## Decision 2 — Precedence Rules

**Question:** Should precedence rules be defined now?

**Recommendation: Yes — define precedence now, even if R1 uses only one layer. This ensures the API contract is forward-compatible.**

### R1 Precedence Stack (top wins)

| Priority | Layer | Source | R1 Status |
|----------|-------|--------|-----------|
| 1 | Station config | `tblStationConfig` row matching current StationId | Active |
| 2 | Warehouse config | `tblWarehouseConfig` row matching current WarehouseId | Active |
| 3 | Hardcoded defaults | `Const` values in `modConfigDefaults` | Active |
| 4 | Workbook-local overrides | *(reserved — not implemented in R1)* | Deferred |

Station-specific values (e.g., `RoleDefault`, station-specific paths) override warehouse-level values. Warehouse-level values override hardcoded defaults. This three-tier model mirrors how the design already separates warehouse identity from station identity in the config tables.

**Implementation:** `Core.Config.Get(key)` walks the stack top-down and returns the first non-empty value. `GetRequired(key)` raises an error if all layers return empty.

***

## Decision 3 — MVP Config Keys

**Question:** What exact config keys are required for R1?

The following keys are derived directly from the design doc's processor workflow, lock model, backup cadence, and topology sections:

### Warehouse-Scope Keys

| Key | Type | Default | Source in Design |
|-----|------|---------|------------------|
| `WarehouseId` | String | *(required, no default)* | `tblWarehouseConfig.WarehouseId` |
| `WarehouseName` | String | *(required)* | `tblWarehouseConfig.WarehouseName` |
| `Timezone` | String | `"UTC"` | `tblWarehouseConfig.Timezone` |
| `DefaultLocation` | String | `""` | `tblWarehouseConfig.DefaultLocation` |
| `BatchSize` | Long | `500` | Processor workflow: "batchSize=500" |
| `LockTimeoutMinutes` | Long | `3` | "expires in 3 min" |
| `HeartbeatIntervalSeconds` | Long | `30` | "every 30 seconds" |
| `MaxLockHoldMinutes` | Long | `2` | "If a batch exceeds 2 minutes, log a warning" |
| `SnapshotCadence` | String | `"PER_BATCH"` | "per batch (not per event)" |
| `BackupCadence` | String | `"DAILY"` | "Daily (or per shift)" |
| `PathDataRoot` | String | `"C:\invSys\{WarehouseId}\"` | Implied by restore playbook paths |
| `PathBackupRoot` | String | `"C:\invSys\Backups\{WarehouseId}\"` | Backup section |
| `PathSharePointRoot` | String | `""` | SharePoint sync folder |
| `DesignsEnabled` | Boolean | `False` | Designs are "optional per warehouse" |
| `PoisonRetryMax` | Long | `3` | Poison handling: RetryCount++ |
| `AuthCacheTTLSeconds` | Long | `28800` | "Cached locally for a normal work session; explicit Sign Out ends the session sooner." |

### Station-Scope Keys

| Key | Type | Default | Source in Design |
|-----|------|---------|------------------|
| `StationId` | String | *(required, no default)* | `tblStationConfig.StationId` |
| `StationName` | String | `""` | `tblStationConfig.StationName` |
| `RoleDefault` | String | `"RECEIVE"` | `tblStationConfig.RoleDefault` |

### Feature Flags (R1)

| Key | Type | Default | Purpose |
|-----|------|---------|---------|
| `FF_DesignsEnabled` | Boolean | `False` | Toggle Designs domain per warehouse |
| `FF_OutlookAlerts` | Boolean | `False` | Optional VBA email alerts |
| `FF_AutoSnapshot` | Boolean | `True` | Auto-generate snapshot after processor batch |

***

## Decision 4 — Strongly Typed with Schema Validation

**Question:** Should config be strongly typed with schema validation/defaulting, or permissive string lookup first?

**Recommendation: Strongly typed with schema validation and automatic defaulting from the start.**

The design doc already mandates schema self-heal for all workbooks — "Missing tables/columns are recreated with defaults". Config should follow the same pattern. A permissive string-only lookup creates a class of silent bugs where a typo in a key name returns empty instead of raising an error.

### Implementation Pattern

```text
' modConfigSchema.bas — defines the manifest
Type ConfigKeyDef
    Key         As String
    DataType    As String    ' STRING | LONG | BOOLEAN | DATETIME
    DefaultVal  As String    ' string representation of default
    Required    As Boolean   ' True = no default, must exist
    Scope       As String    ' WAREHOUSE | STATION
End Type
```

On load, `Core.Config.Load()` reads the raw values from the config workbook tables, then validates each against the schema manifest. Missing keys get their defaults applied. Type mismatches log a warning and fall back to defaults. Required keys that are missing cause a validation failure (see Decision 7 for failure policy).

This approach costs ~50 lines of VBA over the permissive alternative and catches config errors at startup instead of mid-batch.

***

## Decision 5 — Warehouse/Station Context Resolution

**Question:** How should warehouse/station context be resolved — passed explicitly, inferred from workbook, or both?

**Recommendation: Both — infer on startup, allow explicit override for testing and Admin scenarios.**

### Resolution Strategy

1. **Primary (startup inference):** When `Core.Config.Load()` runs, it scans `Application.Workbooks` for an open workbook matching the pattern `WH*.invSys.Config.xlsb`. The `WarehouseId` is extracted from `tblWarehouseConfig`. The `StationId` is resolved by matching `Environ("COMPUTERNAME")` against `tblStationConfig.StationName` or by explicit `StationId` column.

2. **Override (explicit pass):** `Core.Config.Load(Optional warehouseId As String, Optional stationId As String)` accepts explicit values. This enables the test harness and the Admin XLAM (which may manage multiple warehouses) to load config for a specific context.

3. **Cached context:** After resolution, `WarehouseId` and `StationId` are cached in module-level variables and accessible via `Core.Config.GetWarehouseId()` and `Core.Config.GetStationId()`.

This matches the design's principle that "each warehouse operates autonomously" while allowing the Admin XLAM and test harness to explicitly target a warehouse/station context.

***

## Decision 6 — Cache Behavior

**Question:** Load once on startup, or support reload/invalidation while Excel is open?

**Recommendation: Load once on startup + explicit `Reload()` method. No automatic invalidation in R1.**

The design doc specifies cache with TTL for auth capabilities, but config changes are infrequent (admin-initiated, not per-transaction). Auto-reload adds complexity for a case that occurs perhaps once a month.

### Behavior

- `Core.Config.Load()` — called during add-in initialization (`Workbook_Open` or `Auto_Open`). Reads all config values into a `Scripting.Dictionary` cache.
- `Core.Config.Reload()` — re-reads the config workbook and rebuilds the cache. Called explicitly by the Admin XLAM after config changes, or by the processor at the start of each batch.
- `Core.Config.IsLoaded() As Boolean` — returns `True` if cache is populated, `False` if never loaded or explicitly cleared.
- Config values are read from the in-memory `Dictionary` on every `Get()` call — no disk I/O after initial load.

**R2 hook:** If config-change-while-running becomes a real need, a version stamp in `tblWarehouseConfig` can trigger automatic reload when the stamp changes. The `Reload()` method already exists to support this.

***

## Decision 7 — Failure Policy

**Question:** If the config workbook/table is missing or corrupt, fail closed or run with defaults + warnings?

**Recommendation: Fail closed for required keys; run with defaults + warnings for optional keys.**

This aligns with the design's existing offline behavior rule: "If config cannot be refreshed and cache expired, fail closed for write operations".

### Policy Matrix

| Scenario | Behavior |
|----------|----------|
| Config workbook not found | Fail closed — `Load()` returns `False`, all `GetRequired()` calls raise error. Log `CONFIG_MISSING` error. |
| Config workbook found but `tblWarehouseConfig` missing | Attempt self-heal (create table with defaults per schema validation pattern). If self-heal fails, fail closed. |
| Required key missing after self-heal | Fail closed — `Load()` returns `False`. Log `CONFIG_KEY_MISSING` with the key name. |
| Optional key missing or malformed | Apply hardcoded default, log `CONFIG_KEY_DEFAULT` warning. Continue loading. |
| Config workbook is locked by another user | Retry once after 2 seconds. If still locked, fail closed. Log `CONFIG_LOCKED`. |

**Practical effect:** If an operator accidentally deletes the config workbook, the system will refuse to process events (preventing data corruption) and the Admin XLAM will surface the error. Recovery follows the existing restore playbook: copy backup, reopen, schema self-heals.

***

## Decision 8 — API Contract

**Question:** What API surface should `Core.Config` expose?

**Recommendation: Five public methods + two context accessors.**

### Public API

| Method | Signature | Purpose |
|--------|-----------|---------|
| `Load` | `Function Load(Optional whId As String, Optional stId As String) As Boolean` | Read config workbook, validate schema, populate cache. Returns `True` if all required keys resolved. |
| `Get` | `Function Get(key As String) As Variant` | Returns cached value for key. Returns `Empty` if key not found (does not error). |
| `GetRequired` | `Function GetRequired(key As String) As Variant` | Returns cached value. Raises `ERR_CONFIG_KEY_MISSING` if key is absent or empty. |
| `TryGet` | `Function TryGet(key As String, ByRef outVal As Variant) As Boolean` | Returns `True` and populates `outVal` if key exists; `False` otherwise. No error raised. |
| `Reload` | `Function Reload() As Boolean` | Re-reads config workbook using previously resolved context. Returns `True` on success. |
| `Validate` | `Function Validate() As String` | Returns empty string if valid; otherwise returns semicolon-delimited list of issues (e.g., `"BatchSize: expected LONG, got TEXT; WarehouseId: MISSING"`). |
| `GetWarehouseId` | `Function GetWarehouseId() As String` | Returns resolved warehouse ID from cache. |
| `GetStationId` | `Function GetStationId() As String` | Returns resolved station ID from cache. |

### Typed Convenience Wrappers (in same module)

```text
Function GetLong(key As String, defaultVal As Long) As Long
Function GetBool(key As String, defaultVal As Boolean) As Boolean
Function GetString(key As String, defaultVal As String) As String
```

These prevent every caller from needing `CLng(Config.Get("BatchSize"))` with error handling. The wrappers use `TryGet` internally and fall back to `defaultVal` on failure.

***

## Decision 9 — Read-Only in R1

**Question:** Should `Core.Config` be read-only in R1, or include controlled write/update methods?

**Recommendation: Read-only in R1. No write methods in the `Core.Config` module.**

Config changes in R1 are an admin task performed by directly editing the `WHx.invSys.Config.xlsb` workbook (opening it in Excel, changing cells, saving). This is consistent with how Auth data is also managed — via Admin XLAM forms that write to the Auth workbook, not via Core.Auth write methods.

Adding write methods to `Core.Config` in R1 would:

- Require locking logic for the config workbook (currently not designed)
- Create a second code path for config mutation alongside direct Excel editing
- Add surface area for bugs in a Phase 1 deliverable

**R2 hook:** If programmatic config writes become necessary (e.g., Admin XLAM "Settings" panel), add `Set(key, value)` and `Save()` methods. The internal `Dictionary` cache already supports this — only the persistence layer needs to be added.

***

## Decision 10 — Test Harness

**Question:** Should a VBA test harness be included in this same pass?

**Recommendation: Yes — include a `TestCoreConfig.bas` module in the `tests/unit/` directory.**

The design doc already defines a test harness pattern in `TestRunner.bas` and the repo has a `tests/unit/` directory ready for it. Config is a pure-logic module with no UI — ideal for the first unit test target.

### Test Cases

| Test | Setup | Expected |
|------|-------|----------|
| `TestLoad_ValidConfig` | Sample Config.xlsb with all required keys | `Load() = True`, all `Get()` calls return expected values |
| `TestLoad_MissingWorkbook` | No config workbook open | `Load() = False` |
| `TestLoad_MissingRequiredKey` | Config workbook open but `WarehouseId` row deleted | `Load() = False`, `Validate()` returns `"WarehouseId: MISSING"` |
| `TestGet_ExistingKey` | Loaded config with `BatchSize = 250` | `Get("BatchSize") = 250` |
| `TestGet_MissingKey` | Loaded config without `FutureKey` | `Get("FutureKey") = Empty` |
| `TestGetRequired_MissingKey` | Loaded config without `WarehouseId` | Raises `ERR_CONFIG_KEY_MISSING` |
| `TestTryGet_HitAndMiss` | Loaded config | `TryGet("BatchSize", v) = True` and `v = 250`; `TryGet("NoSuchKey", v) = False` |
| `TestReload_UpdatedValue` | Change `BatchSize` in workbook after initial load | After `Reload()`, `Get("BatchSize")` returns new value |
| `TestPrecedence_StationOverridesWarehouse` | Station row has `RoleDefault = "SHIP"`, warehouse default is `"RECEIVE"` | `Get("RoleDefault") = "SHIP"` |
| `TestGetLong_TypeConversion` | `BatchSize = "500"` (text in cell) | `GetLong("BatchSize", 100) = 500` |
| `TestGetBool_TypeConversion` | `DesignsEnabled = "TRUE"` | `GetBool("DesignsEnabled", False) = True` |

### Harness Pattern (consistent with design doc)

```text
' TestCoreConfig.bas
Sub RunConfigTests()
    Dim passed As Long, failed As Long
    passed = passed + TestLoad_ValidConfig()
    passed = passed + TestLoad_MissingWorkbook()
    passed = passed + TestGet_ExistingKey()
    ' ... etc
    Debug.Print "Core.Config tests — Passed: " & passed & " Failed: " & failed
End Sub
```

***

## Module Layout Summary

The following files should be created as part of the Core.Config implementation:

| File | Location | Purpose |
|------|----------|---------|
| `modConfig.bas` | `src/Core/Modules/` | Main `Core.Config` API (Load, Get, GetRequired, TryGet, Reload, Validate, typed wrappers) |
| `modConfigDefaults.bas` | `src/Core/Modules/` | Schema manifest array and hardcoded default values |
| `TestCoreConfig.bas` | `tests/unit/` | Unit test harness for all config operations |
| `WHx.invSys.Config.xlsb` (sample) | `tests/fixtures/` | Sample config workbook for testing |

The existing `modGlobals.bas` status constants remain untouched — they are domain-level constants, not configuration.

***

## Implementation Sequence

1. **Create `modConfigDefaults.bas`** — Define the `ConfigKeyDef` UDT and the schema manifest array listing all MVP keys, types, defaults, and required flags.
2. **Create `modConfig.bas`** — Implement `Load()` with workbook discovery, table reading, schema validation, and `Scripting.Dictionary` cache population. Then implement `Get`, `GetRequired`, `TryGet`, `Reload`, `Validate`, and typed wrappers.
3. **Create sample `WHx.invSys.Config.xlsb`** — Populate `tblWarehouseConfig` and `tblStationConfig` with test data for WH1/S1.
4. **Create `TestCoreConfig.bas`** — Implement all 11 test cases above.
5. **Verify** — Run `RunConfigTests()` in the test harness workbook and confirm all pass.

This sequence delivers a fully tested, schema-validated, read-only config module that Auth, LockManager, and Processor can depend on — exactly what the Phase 1 roadmap calls for.
