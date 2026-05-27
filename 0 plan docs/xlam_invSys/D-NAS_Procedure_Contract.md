# D-NAS Public API Contract
**Document:** invSys Architecture v4.9 — D-NAS Addendum  
**Module:** `Core.NasConnection` (and related `Core.Auth` sequencing)  
**Status:** Ratified working contract — promoted from d5 on May 21, 2026  
**Date:** May 21, 2026

---

## Overview

D-NAS defines a three-layer warehouse connection model owned by Core:

1. **NAS / Windows credential layer** — establishes the Excel session's SMB access to a warehouse root path
2. **Warehouse target layer** — selects the active warehouse runtime (config, auth, inboxes, processor identity, HQ context)
3. **invSys user layer** — signs the operator in against the selected warehouse's auth workbook

All role XLAMs (Receiving, Shipping, Production) and Admin XLAM call into `Core.NasConnection` and `Core.Auth` using these signatures. No role XLAM implements its own resolver logic.

---

## Constants and Enumerations

All status codes and source type values are `Public Enum`. Free-string comparisons are forbidden throughout the codebase.

```vb
' Core.NasConnection

Public Enum NasStatusCode
    NAS_OK = 0
    NAS_ROOT_UNREACHABLE = 1
    NAS_ROOT_NO_CONFIG = 2
    NAS_CREDENTIAL_REJECTED = 3         ' WinAPI returned ERROR_ACCESS_DENIED (5) or
                                        ' ERROR_LOGON_FAILURE (1326) explicitly.
                                        ' NOT used for generic filesystem or workbook errors.
    NAS_ROOT_NOT_IN_SESSION = 4
    WH_RUNTIME_NOT_FOUND = 5
    WH_CONFIG_INVALID = 6
    WH_AUTH_NOT_FOUND = 7
    WH_TARGET_INCOMPLETE = 8
    NAS_TARGET_UNREACHABLE = 9          ' path probe failed; cause unknown or filesystem error
    WH_NO_TARGET = 10
End Enum

Public Enum WH_SourceType
    WH_SOURCE_NAS = 0        ' resolved from live NAS root this session
    WH_SOURCE_REMEMBERED = 1 ' restored from profile store, revalidated without new credential entry
    WH_SOURCE_LOCAL = 2      ' explicitly operator-selected local path (no remembered NAS)
    WH_SOURCE_FALLBACK = 3   ' default dev root; no remembered NAS target existed
End Enum
```

```vb
' Core.Auth

Public Enum AuthStatusCode
    AUTH_OK = 0
    AUTH_CANCELLED = 1
    AUTH_WAREHOUSE_MISMATCH = 2
    AUTH_USER_NOT_FOUND = 3
    AUTH_CREDENTIAL_REJECTED = 4
    AUTH_WORKBOOK_UNREADABLE = 5
    AUTH_NO_CAPABILITIES = 6
    AUTH_NOT_SIGNED_IN = 7
    AUTH_REAUTH_REQUIRED = 8    ' capability cache TTL expired; re-authentication needed
    AUTH_CACHE_EXPIRED = 9      ' auth workbook read is stale beyond acceptable window;
                                ' distinction: REAUTH_REQUIRED = silent re-check possible;
                                ' CACHE_EXPIRED = operator must sign in again explicitly
End Enum
```

`NasStatusCode` is used exclusively by `Core.NasConnection` procedures.  
`AuthStatusCode` is used exclusively by `Core.Auth` procedures.  
No procedure returns a mixed type. No `AUTH_*` code appears in `NasStatusCode` and no `NAS_*` code appears in `AuthStatusCode`.

**TTL semantics:**
- `AUTH_REAUTH_REQUIRED` — `IsSignedIn()` returns `False`; capability cache is stale but the user record in the auth workbook may be re-read without a new PIN entry if the user's session key is still valid. Used for silent background re-check paths.
- `AUTH_CACHE_EXPIRED` — `IsSignedIn()` returns `False`; operator must sign in again explicitly through `ShowSignInPrompt`. Used when the auth workbook has not been readable within the defined cache window.

---

## `WarehouseTarget` Class

`WarehouseTarget` is a plain mutable DTO. Fields are `Public` by design for VBA simplicity. **By convention, only `Core.NasConnection` procedures may write to a `WarehouseTarget` instance.** Role XLAMs, Admin XLAM, and `Core.Auth` treat received `WarehouseTarget` objects as read-only. This convention is enforced by code review.

If stricter enforcement is needed in a future revision, `Public` fields can be replaced with `Private` backing variables and `Property Get` accessors as a non-breaking refactor.

**`GetCurrentTarget()` returns a deep copy, not a reference.** VBA object variables are references by default. `GetCurrentTarget()` must explicitly construct a new `WarehouseTarget` instance and copy each field value before returning. Callers that mutate the returned object do not affect resolver state.

```vb
' Core.WarehouseTarget (Class Module)
Public WarehouseId     As String
Public WarehouseName   As String        ' display name from config workbook
Public StationId       As String        ' "" = roaming/unscoped; specific value = station-scoped
Public HubRoot         As String        ' NAS share root, e.g. "\\192.168.1.5\invSysWH1"
Public RuntimeRoot     As String        ' runtime folder within hub (may equal HubRoot)
Public ConfigPath      As String        ' cached hint — revalidated from HubRoot+RuntimeRoot on restore
Public AuthPath        As String        ' cached hint — revalidated on restore
Public InboxRoot       As String        ' cached hint — revalidated on restore
Public SourceType      As WH_SourceType
Public LastResolvedUTC As Date
```

**Cached hints rule:** `ConfigPath`, `AuthPath`, and `InboxRoot` stored in the profile store are hints only. On any restore from the profile store (priorities 2–3 in the resolver), Core recomputes and validates these paths from `HubRoot` + `RuntimeRoot` + config workbook contents before marking the target live. A restored target whose hints cannot be revalidated is treated as `NAS_TARGET_UNREACHABLE` for that session.

**`WarehouseId` rule:** Always read from the config workbook. Never inferred from folder or file names.

**Station semantics:**

| `StationId` value | Meaning |
|---|---|
| `""` (empty string) | Roaming / unscoped. Warehouse-level capabilities only. Inbox path resolves to a warehouse-global inbox. The warehouse processor must be configured to accept events from a non-station-specific inbox path. |
| `"S2"` (specific value) | Station-scoped. Station-scoped capabilities are loaded in addition to warehouse-level capabilities. Inbox path resolves to the station-specific workbook (e.g., `invSys.Inbox.Receiving.S2.xlsb`). |
| `"*"` | Config query wildcard only. Must never appear in a live `WarehouseTarget`. |

Station inbox policy is enforced by `SelectWarehouseTarget` via the `requireStationInbox` parameter. See below.

---

## Startup Modal Policy — Non-Negotiable

**`Workbook_Open` must not show any modal form, message box, or prompt.**

`Workbook_Open` calls only `ResolveWarehouseTarget`, which is non-modal and safe in any context. It stores the result (or the `NAS_TARGET_UNREACHABLE` state) in module-level session state and updates ribbon label/enabled state via `IRibbonUI.Invalidate`. The operator sees the ribbon status change (e.g., "NAS unreachable" label, Connect button enabled) and acts explicitly.

The modal connection flow (`EnsureWarehouseTargetInteractive`) is triggered only by:
1. Admin/setup storage-management UI or Runtime Context troubleshooting — explicit operator action outside the normal role workflow
2. A role write ribbon button (`onAction` callback) only if the implementation intentionally offers a recovery prompt; the preferred role behavior is fail-closed with clear status/message

`EnsureWarehouseTargetInteractive` must never be called from `Workbook_Open`, `Workbook_Activate`, `Application.OnTime` callbacks, ribbon `getEnabled`/`getLabel`/`getImage` callbacks, or background refresh paths.

Storage connection and invSys authentication are separate workflows. The warehouse connection prompt may collect Windows/NAS credentials and select a target, but it must not claim to sign in the invSys operator. Receiving, Shipping, and Production **Connect Server** buttons are non-modal: they re-run target resolution, refresh server status, and fail closed if no acceptable NAS target is available. The Sign In prompt is the invSys user/PIN or user/password action against the selected warehouse auth workbook and must not show NAS credential fields.

### Fallback Policy by Context

A `WH_SOURCE_FALLBACK` target (priority 5, local dev root) is handled differently depending on `requireNasTarget`:

- **`requireNasTarget = False` (Admin XLAM default):** Fallback is an acceptable resolved state. Post/write controls enabled normally.
- **`requireNasTarget = True` (role write/sign-in policy):** Fallback is not an acceptable resolved state. Role XLAMs normally enforce this non-modally through cached target checks, server status labels, and disabled post/write controls. Admin/setup may still use `EnsureWarehouseTargetInteractive` to repair or select storage.

The role XLAM fallback restriction is also enforced independently in the ribbon `getEnabled` callback by checking `GetCurrentTarget().SourceType <> WH_SOURCE_FALLBACK` so that the button stays disabled even if `EnsureWarehouseTargetInteractive` is not called first.

---

## Module: `Core.NasConnection`

### NAS Root Layer

---

#### `ConnectNasRootWithCredentials`

```vb
Public Function ConnectNasRootWithCredentials( _
    ByVal rootPath As String, _
    ByVal userName As String, _
    ByVal windowsPassword As String _
) As NasStatusCode
```

**Purpose:** Establish or verify the Windows/Excel session's SMB access to a warehouse root path using explicit Windows/NAS credentials. Stores the root in the session scan list on success.

**Parameters:**
- `rootPath` — UNC path or mapped drive root (e.g., `\\192.168.1.5\invSysWH1`).
- `userName` — Windows account in `DOMAIN\user` format. This is a Windows/NAS credential, not an invSys user identity.
- `windowsPassword` — Windows/NAS account password. This is distinct from the invSys user PIN/password used in `Core.Auth`.

**Language note:** `windowsPassword` is the Windows SMB account password. The term "PIN" appears nowhere in this procedure or its documentation. PIN is reserved for invSys user authentication in `Core.Auth`.

**Returns:** `NAS_OK` on success. Appropriate `NasStatusCode` on failure.

**Credential handling — WinAPI only:**
- SMB authentication is performed exclusively through the `WNetAddConnection2` Windows API. `Shell`, `CreateObject("WScript.Shell")`, and `net use` command strings are forbidden for any call that passes credentials.
- `windowsPassword` is passed as the `lpPassword` parameter in the `NETRESOURCE` struct call. It is not written to any module-level variable, registry key, `SaveSetting` store, workbook cell, or log entry.
- On success, `userName` is stored for display and reconnect prompts. `windowsPassword` is not retained.
- On failure, logs: timestamp, `rootPath`, `userName`, `NasStatusCode`. `windowsPassword` is never logged.

---

#### `TryRevalidateRememberedRoot`

```vb
Public Function TryRevalidateRememberedRoot( _
    ByVal rootPath As String _
) As NasStatusCode
```

**Purpose:** Probe a remembered NAS root path using the existing Windows SMB session (no new credential entry). If the root is reachable under the current Windows session credentials, add it to the session scan list. This is the resolver's path for restoring remembered targets after Excel restart when the Windows SMB session is still active from a previous connection.

**Parameters:**
- `rootPath` — Remembered hub root path to probe.

**Returns:** `NAS_OK`, `NAS_TARGET_UNREACHABLE`, or `NAS_CREDENTIAL_REJECTED`.

**Error mapping — precise WinAPI codes required:**
Implementation must map WinAPI return codes to `NasStatusCode` as follows. Do not infer credential failure from filesystem behavior alone:

| WinAPI / VBA error | Mapped `NasStatusCode` |
|---|---|
| `WNetAddConnection2` returns `ERROR_ACCESS_DENIED` (5) | `NAS_CREDENTIAL_REJECTED` |
| `WNetAddConnection2` returns `ERROR_LOGON_FAILURE` (1326) | `NAS_CREDENTIAL_REJECTED` |
| `Dir$` or workbook open raises error 76 (path not found) | `NAS_TARGET_UNREACHABLE` |
| `Dir$` or workbook open raises error 52 / 53 (bad file name / not found) | `NAS_TARGET_UNREACHABLE` |
| Network timeout, error 53 (network path not found) | `NAS_TARGET_UNREACHABLE` |
| Any other unhandled VBA error | `NAS_TARGET_UNREACHABLE` — do not overclaim credential rejection |

`NAS_CREDENTIAL_REJECTED` is returned only when the Windows API explicitly reports an authentication failure. All ambiguous errors map to `NAS_TARGET_UNREACHABLE`.

**Rules:**
- Does not prompt for credentials. Returns failure code and the caller (resolver priority 2–3) surfaces the appropriate ribbon status.
- On `NAS_OK`, adds `rootPath` to the session scan list. Subsequent `SelectWarehouseTarget` calls using this root proceed normally.
- Called internally by `ResolveWarehouseTarget` for priority 2–3 candidates. Not a public entry point for ribbon callbacks.

---

#### `ShowWarehouseConnectionPrompt`

```vb
Public Sub ShowWarehouseConnectionPrompt( _
    Optional ByVal reason As String = "" _
)
```

**Purpose:** Surface the Core-owned modal form for NAS root connection, warehouse target selection, and Windows credential entry. Called only from `EnsureWarehouseTargetInteractive`. This form establishes storage access only; it does not sign in an invSys user.

**Parameters:**
- `reason` — Optional display string explaining why the prompt appeared.

**Rules:**
- The form owns credential input. It calls `ConnectNasRootWithCredentials` internally. `windowsPassword` never leaves the form.
- UI copy must use storage/network language such as "Connect to Server", "Select Warehouse Storage", and "Network user/password". It must not present the Windows/NAS credential as an invSys sign-in.
- On successful connect and target selection, the form calls `SelectWarehouseTarget` and `RememberTarget` before closing.
- On cancel, resolved target is unchanged.
- **Must not be called directly from `Workbook_Open`, ribbon `getEnabled`/`getLabel`/`getImage`, `Application.OnTime` callbacks, or background refresh paths.** Only `EnsureWarehouseTargetInteractive` may call this procedure.

---

#### `DisconnectNasRoot`

```vb
Public Sub DisconnectNasRoot( _
    ByVal rootPath As String, _
    Optional ByVal disconnectWindowsSession As Boolean = False _
)
```

**Purpose:** Remove a root from the session scan list. Optionally tear down the Windows SMB session via `WNetCancelConnection2`.

**Parameters:**
- `rootPath` — Root to remove.
- `disconnectWindowsSession` — When `True`, Core calls `WNetCancelConnection2`. Core owns all WinAPI credential calls; callers do not issue their own.

**Rules:**
- If the disconnected root is the current warehouse target's root, `ClearWarehouseTarget` is called automatically before the root is removed.

---

#### `ForgetRoot`

```vb
Public Sub ForgetRoot(ByVal rootPath As String)
```

**Purpose:** Remove a NAS root from the remembered root list in the profile store. Does not affect the current session.

**Rules:**
- Any remembered `WarehouseTarget` whose `HubRoot` matches `rootPath` is also removed.
- Does not call `DisconnectNasRoot`. Session and profile store are managed independently.

---

#### `ScanNasRoot`

```vb
Public Function ScanNasRoot( _
    ByVal rootPath As String _
) As Collection
```

**Purpose:** Enumerate candidate warehouse runtime folder paths under a connected root by filesystem pattern only.

**Returns:** `Collection` of candidate folder path strings. Empty `Collection` if none found or root unreachable.

**Rules:**
- Filesystem enumeration only. Does not open workbooks.
- Returns path strings only — no `WarehouseId` inference from folder or file names.
- Results are passed to `SelectWarehouseTarget` for validation. The calling form calls `SelectWarehouseTarget` on each candidate to obtain `WarehouseId` and display name.

---

### Warehouse Target Layer

---

#### `SelectWarehouseTarget`

```vb
Public Function SelectWarehouseTarget( _
    ByVal hubRoot As String, _
    ByVal runtimeRoot As String, _
    ByRef outTarget As WarehouseTarget, _
    Optional ByVal stationId As String = "", _
    Optional ByVal requireStationInbox As Boolean = False _
) As NasStatusCode
```

**Purpose:** Validate a candidate warehouse runtime and populate a fully resolved `WarehouseTarget`. Reads the config workbook to obtain `WarehouseId` and `WarehouseName`. Validates auth workbook reachability. Recomputes `ConfigPath`, `AuthPath`, and `InboxRoot` from source.

**Parameters:**
- `hubRoot` — Must be in the session scan list (added by `ConnectNasRootWithCredentials` or `TryRevalidateRememberedRoot`).
- `runtimeRoot` — Specific warehouse runtime folder within hub. May equal `hubRoot`.
- `outTarget` — Output. Populated on `NAS_OK`; null target on failure.
- `stationId` — See station semantics. Empty string accepted for roaming unless `requireStationInbox = True`.
- `requireStationInbox` — When `True`, a `stationId` of `""` or `"*"` causes the function to return `WH_TARGET_INCOMPLETE`. Role XLAMs that require a station-scoped inbox path pass `True`. Admin XLAM and roaming roles pass `False` (default).

**Returns:** `NAS_OK` on success. Appropriate `NasStatusCode` on failure.

**Rules:**
- Opens and reads config workbook to populate `WarehouseId` and `WarehouseName`. Filename inference is forbidden.
- Verifies auth workbook via `Workbooks.Open`, not `FileSystemObject.FileExists`.
- Config workbook is closed after reading.
- On `NAS_OK`, sets `outTarget.SourceType = WH_SOURCE_NAS`, stores as in-session override, calls `RememberTarget` automatically.
- Does not sign in the invSys user.

---

#### `ResolveWarehouseTarget`

```vb
Public Function ResolveWarehouseTarget( _
    ByRef outTarget As WarehouseTarget, _
    ByRef statusCode As NasStatusCode _
) As Boolean
```

**Purpose:** Apply the five-priority resolver rule. Returns status only. **Never modal. Never shows any UI.**

**Parameters:**
- `outTarget` — Output. On `NAS_TARGET_UNREACHABLE`, populated with stale remembered values for display only — not usable as live runtime paths.
- `statusCode` — Output. `NasStatusCode` describing the outcome.

**Returns:** `True` if a live, validated target was resolved. `False` otherwise.

**Resolver priority (strict order):**

```
1. Current in-session WarehouseTarget override (set by SelectWarehouseTarget this session)

2. Remembered WarehouseTarget from profile store
   → Core calls TryRevalidateRememberedRoot(rememberedHubRoot)
   → NAS_OK: root added to session list; SelectWarehouseTarget called to revalidate;
     hints recomputed from HubRoot + RuntimeRoot + config workbook
   → NAS_TARGET_UNREACHABLE or NAS_CREDENTIAL_REJECTED: returns False, NAS_TARGET_UNREACHABLE
     (fail closed — does not fall through to priority 5)
   → No credential prompt

3. Other remembered HubRoots in profile store
   → Same TryRevalidateRememberedRoot probe; same fail-closed behavior

4. Open workbook-local or active runtime config, only when explicitly selected or unambiguous

5. Default local dev root — ONLY when no remembered NAS target exists at priorities 2–3,
   or when the operator has explicitly chosen local fallback via the connection form
```

**Fail-closed rule:**
- If a remembered NAS target exists at priority 2 or 3 and is unreachable for any reason (path, permissions, or network): returns `False`, `NAS_TARGET_UNREACHABLE`. Priority 5 is not attempted.
- Priority 5 is reached only when no remembered NAS entry exists at priorities 2–3.
- When priority 5 is used: returns `True`, `statusCode = NAS_OK`, `outTarget.SourceType = WH_SOURCE_FALLBACK`.

**Safe contexts:** `Workbook_Open`, ribbon `getEnabled`/`getLabel`/`getImage`, background refresh, `Application.OnTime` callbacks, any non-interactive path.

---

#### `EnsureWarehouseTargetInteractive`

```vb
Public Function EnsureWarehouseTargetInteractive( _
    Optional ByVal reason As String = "", _
    Optional ByVal requireNasTarget As Boolean = False _
) As Boolean
```

**Purpose:** Interactive coordinator. Calls `ResolveWarehouseTarget`; if unresolved, or if `requireNasTarget = True` and the resolved target is `WH_SOURCE_FALLBACK`, shows `ShowWarehouseConnectionPrompt` and re-resolves. Returns `True` only when a live, appropriate target is confirmed.

**Parameters:**
- `reason` — Optional display string forwarded to `ShowWarehouseConnectionPrompt`.
- `requireNasTarget` — When `True`, a `WH_SOURCE_FALLBACK` result is not accepted. The function prompts again until the operator connects to a NAS target or cancels. Admin/setup may pass `True` for repair workflows. The normal role ribbon path does not call this form.

**Returns:** `True` if a live target meeting the `requireNasTarget` constraint is resolved. `False` if the operator cancelled or connection failed.

**Recommended call pattern:**

```vb
' Role XLAM Connect Server button onAction
Dim target As WarehouseTarget
Dim statusCode As NasStatusCode
If Core.NasConnection.ResolveWarehouseTarget(target, statusCode) Then
    ' server status label is refreshed; no storage credential form opens
End If

' Admin/setup Connect / Select Warehouse Storage button onAction
If Core.NasConnection.EnsureWarehouseTargetInteractive() Then
    ' storage target selected; fallback accepted for Admin
End If
```

**Permitted callers:**
- Admin/setup **Connect / Select Warehouse Storage** button `onAction` callback
- Runtime Context troubleshooting action

**Forbidden callers:** `Workbook_Open`, `Workbook_Activate`, `Application.OnTime`, ribbon `getEnabled`/`getLabel`/`getImage`, any background path.

---

#### `ClearWarehouseTarget`

```vb
Public Sub ClearWarehouseTarget()
```

**Purpose:** Clear the current in-session override. Does not affect the remembered target in the profile store. Calls `Core.Auth.SignOut` automatically before clearing.

---

#### `IsTargetResolved`

```vb
Public Function IsTargetResolved() As Boolean
```

**Returns:** `True` if a warehouse target is resolved and live. A `WH_SOURCE_FALLBACK` target returns `True`; callers that need to distinguish fallback check `GetCurrentTarget().SourceType` directly.

**Safe contexts:** Ribbon `getEnabled`, any non-modal path. Does not trigger network access or UI.

---

#### `GetCurrentTarget`

```vb
Public Function GetCurrentTarget() As WarehouseTarget
```

**Returns:** A deep copy of the resolved `WarehouseTarget`. Returns a null target if none is resolved.

**Implementation requirement:** The function must construct a new `WarehouseTarget` instance and assign each field by value before returning. It must not return the module-level reference directly. Example:

```vb
Public Function GetCurrentTarget() As WarehouseTarget
    Dim copy As New WarehouseTarget
    copy.WarehouseId     = m_CurrentTarget.WarehouseId
    copy.WarehouseName   = m_CurrentTarget.WarehouseName
    copy.StationId       = m_CurrentTarget.StationId
    copy.HubRoot         = m_CurrentTarget.HubRoot
    copy.RuntimeRoot     = m_CurrentTarget.RuntimeRoot
    copy.ConfigPath      = m_CurrentTarget.ConfigPath
    copy.AuthPath        = m_CurrentTarget.AuthPath
    copy.InboxRoot       = m_CurrentTarget.InboxRoot
    copy.SourceType      = m_CurrentTarget.SourceType
    copy.LastResolvedUTC = m_CurrentTarget.LastResolvedUTC
    Set GetCurrentTarget = copy
End Function
```

---

#### `RememberTarget`

```vb
Public Sub RememberTarget(ByVal target As WarehouseTarget)
```

**Purpose:** Persist a resolved `WarehouseTarget` to the Windows user profile via `SaveSetting`, keyed by `Environ$("USERDOMAIN") & "\" & Environ$("USERNAME")` plus Office account identity when available.

**Rules:**
- Stores: `WarehouseId`, `WarehouseName`, `HubRoot`, `RuntimeRoot`, `StationId`, `LastResolvedUTC`.
- Stores `ConfigPath`, `AuthPath`, `InboxRoot` as cached hints. On restore, Core recomputes from `HubRoot` + `RuntimeRoot` + config workbook.
- Does not store Windows passwords or invSys `secretText`.
- Maximum 10 remembered targets per user. Oldest by `LastResolvedUTC` is dropped when exceeded.

---

#### `ForgetTarget`

```vb
Public Sub ForgetTarget(ByVal warehouseId As String)
```

**Purpose:** Remove a remembered `WarehouseTarget` from the profile store. If it matches the in-session target, `ClearWarehouseTarget` is called first.

---

### Session State

---

#### `IsConnected`

```vb
Public Function IsConnected() As Boolean
```

**Returns:** `True` if at least one NAS root is in the session scan list and was reachable at last check. Cached state only — no network probe. Safe for ribbon `getEnabled`.

---

#### `GetConnectionStatus`

```vb
Public Function GetConnectionStatus() As String
```

**Purpose:** Human-readable status for ribbon labels and Admin console. Display binding only — never for logic branching.

**Returns:** One of:
- `"Connected — WH1 (Seattle) at \\192.168.1.5\invSysWH1"`
- `"No warehouse target selected"`
- `"NAS unreachable — last known: WH1 at \\192.168.1.5\invSysWH1"`
- `"Local fallback active — writes disabled"` (role XLAMs)
- `"Local fallback active — WH1 at C:\invSys\WH1"` (Admin XLAM)

---

## Module: `Core.Auth`

---

#### `ShowSignInPrompt`

```vb
Public Function ShowSignInPrompt( _
    ByVal target As WarehouseTarget, _
    Optional ByVal requiredCapability As String = "" _
) As AuthStatusCode
```

**Purpose:** Surface the Core-owned modal sign-in form. Returns an `AuthStatusCode` so callers can distinguish cancel, failed credential, and success without re-checking `IsSignedIn`.

**Parameters:**
- `target` — Resolved `WarehouseTarget` to sign in against.

**Returns:** `AUTH_OK` on success. `AUTH_CANCELLED` if the operator dismissed the form. `AUTH_CREDENTIAL_REJECTED` or other code if validation failed after submission.

**Rules:**
- Must not be called before `IsTargetResolved() = True`.
- The form calls `ValidateUserCredentialForTarget` internally. `secretText` (invSys PIN/password) never leaves the form.
- The sign-in account field accepts **exact `tblUsers.UserId` only**. `DisplayName` is display-only for ribbon/status labels and is not accepted as a login alias. If future releases use email addresses, the email address belongs in `UserId`.
- **Cancel:** Form closes. Currently signed-in user (if any) unchanged. Returns `AUTH_CANCELLED`.
- **Failed credential during attempted user switch:** Currently signed-in user unchanged. Failed attempt logged. Form remains open for retry or cancel. Returns failure code only after form is closed by operator.
- **Success:** In-session user replaced with newly validated user. Returns `AUTH_OK`.

---

#### `ValidateUserCredentialForTarget`

```vb
Public Function ValidateUserCredentialForTarget( _
    ByVal userId As String, _
    ByVal secretText As String, _
    ByVal target As WarehouseTarget, _
    Optional ByVal requiredCapability As String = "" _
) As AuthStatusCode
```

**Purpose:** Internal validation call used by `ShowSignInPrompt` and Phase 6 automation. Role and Admin UI sign-in always goes through `ShowSignInPrompt`.

**Parameters:**
- `userId` — exact invSys user identifier from `tblUsers.UserId`. This is the login identity; `DisplayName` is never used for credential lookup.
- `secretText` — invSys user PIN or password. This is the invSys credential, distinct from the Windows/NAS password used in `Core.NasConnection`. Never logged or persisted.
- `target` — Must be the currently resolved `WarehouseTarget`.
- `requiredCapability` — Optional capability required for this sign-in context, e.g. `RECEIVE_POST`.

**Returns:** `AuthStatusCode`.

**Credential language note:** `secretText` is the invSys user PIN or password. It has no relationship to the `windowsPassword` parameter in `ConnectNasRootWithCredentials`. These are two separate credential domains.

**User-switch semantics:**
- On `AUTH_OK`: replaces in-session user; populates capability cache.
- On any failure code: does not modify existing in-session user or capability cache.

**Error states:**

| Code | Condition |
|---|---|
| `AUTH_WAREHOUSE_MISMATCH` | `target.WarehouseId` does not match the current resolver target or the loaded config identity |
| `AUTH_USER_NOT_FOUND` | `userId` not in `tblUsers` for this warehouse |
| `AUTH_CREDENTIAL_REJECTED` | User found but `secretText` does not validate |
| `AUTH_WORKBOOK_UNREADABLE` | Auth workbook cannot be opened |
| `AUTH_NO_CAPABILITIES` | User found and validated but has no active capabilities |

On failure, logs: `userId`, warehouse, station, timestamp, `AuthStatusCode`. `secretText` is never logged.

---

#### `SignOut`

```vb
Public Sub SignOut()
```

**Purpose:** Clear the current invSys user session and capability cache. Called automatically by `Core.NasConnection.ClearWarehouseTarget`.

---

#### `GetCurrentUserId`

```vb
Public Function GetCurrentUserId() As String
```

**Returns:** `UserId` of the signed-in invSys user. Empty string if none.

---

#### `GetCurrentUserDisplayName`

```vb
Public Function GetCurrentUserDisplayName() As String
```

**Returns:** Display name for the signed-in invSys user. Falls back to `UserId` if the display name is blank or unavailable. Display name is presentation only and is never a credential lookup key.

---

#### `IsSignedIn`

```vb
Public Function IsSignedIn() As Boolean
```

**Returns:** `True` if a user is signed in and the capability cache TTL has not expired. Returns `False` when TTL has lapsed (`AUTH_REAUTH_REQUIRED`) or the auth workbook cache has expired (`AUTH_CACHE_EXPIRED`). Safe for ribbon `getEnabled`.

When this function returns `False` due to TTL expiry rather than explicit sign-out, the ribbon status label should reflect the expiry state (e.g., "Session expired — please sign in again") rather than a blank user display. The specific `AuthStatusCode` is available via `GetAuthStatus` (see below).

---

#### `GetAuthStatus`

```vb
Public Function GetAuthStatus() As AuthStatusCode
```

**Purpose:** Return the current `AuthStatusCode` for display and diagnostic purposes. Used by ribbon `getLabel` callbacks to surface `AUTH_REAUTH_REQUIRED` versus `AUTH_NOT_SIGNED_IN` versus `AUTH_CACHE_EXPIRED` in the status label. Never used for logic branching in capability checks — use `CanPerform` for that.

---

#### `CanPerform`

```vb
Public Function CanPerform( _
    ByVal capability As String, _
    ByVal userId As String, _
    Optional ByVal warehouseId As String = "", _
    Optional ByVal stationId As String = "", _
    Optional ByVal source As String = "UI", _
    Optional ByVal requestId As String = "" _
) As Boolean
```

**Purpose:** Check whether the signed-in user holds a named capability against the current warehouse and station. This is the single signed-in operator capability gate used by all ribbon `getEnabled` callbacks and current-user write action guards. It is defined in full in the `Core.Auth` capability contract (separate document); this reference is included here so that the D-NAS startup sequence and ribbon integration table use a consistent entry point and no role XLAM invents an alternative capability path.

**Returns:** `True` if the user is signed in, the TTL is valid, and the capability is active for the given warehouse/station scope. `False` in all other cases (signed out, TTL expired, capability not held, station mismatch).

**Rules:**
- `CanPerform` is the only capability check entry point for signed-in role UI/operator writes. Role XLAMs do not read capability flags directly from `WarehouseTarget` or the auth workbook.
- If `IsSignedIn()` returns `False` for any reason, `CanPerform` returns `False` without opening the auth workbook.
- Role ribbon `getEnabled` callbacks use `Core.RoleUiAccess.CanCurrentUserPerformCapabilityCached`, which is a cached-state facade over the current target/auth state and must remain non-modal and network-free.
- See `Core.Auth` capability contract for capability name constants, scope rules, and TTL refresh behavior.

---

#### `HasProvisionedCapabilityForSystem`

```vb
Public Function HasProvisionedCapabilityForSystem( _
    ByVal capability As String, _
    ByVal userId As String, _
    Optional ByVal warehouseId As String = "", _
    Optional ByVal stationId As String = "" _
) As Boolean
```

**Purpose:** System-side provisioning check for bootstrap, processor validation, and legacy explicit-user queue paths. This does not establish or depend on an operator sign-in session and must not be used by role UI `getEnabled` callbacks or current-user write buttons.

---

## Required Ribbon Integration

**Callback type rules:**

| Callback type | Permitted calls |
|---|---|
| `getEnabled`, `getLabel`, `getImage` | `IsTargetResolved`, `IsConnected`, `IsSignedIn`, `GetCurrentUserId`, `GetCurrentTarget`, `GetConnectionStatus`, `GetAuthStatus` — cached state, non-modal, no network probe |
| `onAction` (button click) | May additionally call `EnsureWarehouseTargetInteractive`, `ShowSignInPrompt`, `SignOut`, role event creators |

**Ribbon surface (all role and Admin XLAMs):**

| Ribbon element | `onAction` / binding | Enabled condition |
|---|---|---|
| **Connect Server** button (role XLAMs) | Non-modal `ResolveWarehouseTarget`; refresh ribbon; no storage credential form | Always enabled |
| **Connect / Select Warehouse Storage** button (Admin/setup) | `EnsureWarehouseTargetInteractive(requireNasTarget:=False)` | Always enabled |
| **Server status** label (role XLAMs) | cached target status label | Always visible; no network probe |
| **Warehouse status** label | `GetConnectionStatus` | Always visible |
| **Sign In** button | If target is acceptable, `ShowSignInPrompt(GetCurrentTarget(), requiredCapability)`; otherwise message/status only | `IsSignedIn() = False`; writes remain disabled until target and auth pass |
| **Sign Out** button | `SignOut` | `IsSignedIn() = True` |
| **Current user / auth status** label | `IsSignedIn` + `GetCurrentUserDisplayName` + `GetAuthStatus` | Always visible; signed-out display is `Sign In` or `<not signed in>`, never Windows/NAS identity; runtime diagnostics may separately show `UserId` |
| **Post / Confirm Writes** button (role XLAMs) | Role event creator | `IsSignedIn() = True` AND `CanPerform(cap, ...)` AND `GetCurrentTarget().SourceType <> WH_SOURCE_FALLBACK` |
| **Post / Confirm Writes** button (Admin XLAM) | Admin action | `IsSignedIn() = True` AND `CanPerform(cap, ...)` |

---

## Required Call Sequence (Startup and Session Resume)

```vb
' Workbook_Open (all role and Admin XLAMs)
'
' Non-modal only. No prompts. Ribbon state updated; operator acts.

Dim tgt As WarehouseTarget
Dim sc  As NasStatusCode

Core.NasConnection.ResolveWarehouseTarget tgt, sc
' ResolveWarehouseTarget internally calls TryRevalidateRememberedRoot for
' priorities 2–3. No credential prompt is issued. If the Windows SMB session
' is still active from a prior connect, the remembered root is revalidated
' silently. If the Windows session has expired, sc = NAS_TARGET_UNREACHABLE
' and the ribbon shows "NAS unreachable" with Connect button enabled.

ribbonUI.Invalidate   ' refresh all ribbon getEnabled / getLabel callbacks


' Role Connect Server ribbon button onAction
'
' Explicit operator action. No modal storage form.

Dim connectTarget As WarehouseTarget
Dim connectStatus As NasStatusCode
Call Core.NasConnection.ResolveWarehouseTarget(connectTarget, connectStatus)
ribbonUI.Invalidate


' Sign In ribbon button onAction
'
' Explicit operator action. InvSys auth is separate from storage access.

If Not Core.NasConnection.IsTargetResolved() Then
    MsgBox "Warehouse storage is not connected. Use Connect Server or Runtime Context before signing in.", vbExclamation
    Exit Sub
End If
Dim result As AuthStatusCode
result = Core.Auth.ShowSignInPrompt(Core.NasConnection.GetCurrentTarget(), requiredCapability)
' AUTH_OK: ribbon write controls enabled by IsSignedIn() getEnabled
' AUTH_CANCELLED or failure: ribbon remains in signed-out state
ribbonUI.Invalidate


' Role write button onAction (e.g., Confirm Writes in Receiving)
'
' Explicit operator action. Fail closed if storage is unavailable.

If Not Core.NasConnection.IsTargetResolved() Then
    MsgBox "Warehouse storage is not connected. Use Connect Server before posting.", vbExclamation
    Exit Sub
End If
If Core.NasConnection.GetCurrentTarget().SourceType = WH_SOURCE_FALLBACK Then
    MsgBox "Connect to a NAS warehouse to enable writes.", vbExclamation
    Exit Sub
End If
If Not Core.Auth.IsSignedIn() Then
    If Core.Auth.ShowSignInPrompt(Core.NasConnection.GetCurrentTarget(), requiredCapability) <> AUTH_OK Then Exit Sub
End If
If Not Core.Auth.CanPerform(requiredCapability, Core.Auth.GetCurrentUserId(), Core.NasConnection.GetCurrentTarget().WarehouseId, Core.NasConnection.GetCurrentTarget().StationId) Then
    MsgBox "You do not have permission to perform this action.", vbExclamation
    Exit Sub
End If
' proceed with role event creation
```

**Config load:** `Core.Config.Load(outTarget)` is called after `ResolveWarehouseTarget` returns a live target, using `outTarget.ConfigPath` (revalidated by resolver, not a stale hint). Config load is not called before target resolution.

---

## Phase 6 Tests Required by This Contract

| Test | Entry point | Pass condition |
|---|---|---|
| `Workbook_Open` completes without modal UI regardless of NAS state | `Workbook_Open` with NAS offline | No prompt shown; ribbon label shows "NAS unreachable"; Connect button enabled |
| `Workbook_Open` with NAS online and Windows session active silently resolves remembered target | `ResolveWarehouseTarget` priority 2 via `TryRevalidateRememberedRoot` | `outTarget.SourceType = WH_SOURCE_NAS`; no credential prompt |
| `Workbook_Open` with NAS online but Windows session expired: no prompt, ribbon shows unreachable | `TryRevalidateRememberedRoot` → `NAS_CREDENTIAL_REJECTED` | Returns `False`; `NAS_TARGET_UNREACHABLE`; no modal |
| Stale remembered root / active Windows session: resolves to NAS, not `C:\invSys\WH1` | Priority 2 restore via `TryRevalidateRememberedRoot` (Windows session still active) | `outTarget.HubRoot` = NAS path; `SourceType ≠ WH_SOURCE_FALLBACK` |
| `ConnectNasRootWithCredentials` uses `WNetAddConnection2`; no shell process spawned | API call audit | No `Shell` or `WScript.Shell` invocation; `windowsPassword` not in any log |
| `TryRevalidateRememberedRoot` maps WinAPI 5 / 1326 to `NAS_CREDENTIAL_REJECTED`; all other errors map to `NAS_TARGET_UNREACHABLE` | Error injection per WinAPI code | Code table applied exactly; no overclaim of credential rejection |
| `ScanNasRoot` returns `Collection` of path strings; no `WarehouseId` inference | `ScanNasRoot` | Each item is a path string; none contain inferred warehouse IDs |
| `SelectWarehouseTarget` on folder `invsys_Zenbook_WH` returns `WarehouseId` from config workbook | `SelectWarehouseTarget` | `outTarget.WarehouseId` = config value, not folder name |
| Select `invsys_Zenbook_WH`, restart Excel, Windows session active: resolves to NAS, not `C:\invSys\WH1` | Priority 2 restore | `outTarget.HubRoot` = NAS path; `SourceType ≠ WH_SOURCE_FALLBACK` |
| Select `invsys_Zenbook_WH`, restart Excel, Windows session expired: ribbon shows unreachable, no fallback | Priority 2 probe → `NAS_CREDENTIAL_REJECTED` | Returns `False`; `NAS_TARGET_UNREACHABLE`; local root not loaded |
| Role XLAM Connect Server does not open storage credential form | Ribbon `onAction` | Calls non-modal resolver, refreshes server label, and shows no warehouse connection form |
| Role XLAM server status label reflects target state | Ribbon `getLabel` | Shows `Server: Connected ...` for acceptable NAS target; `Server: Not connected` otherwise |
| Admin XLAM `EnsureWarehouseTargetInteractive()` accepts fallback without re-prompt | `EnsureWarehouseTargetInteractive` with `WH_SOURCE_FALLBACK` active, `requireNasTarget` default `False` | Returns `True`; no re-prompt |
| Role XLAM `WH_SOURCE_FALLBACK` target: Post button disabled via `getEnabled` | Ribbon `getEnabled` | `CanPerform = True` but Post button disabled; `SourceType = WH_SOURCE_FALLBACK` |
| Admin XLAM `WH_SOURCE_FALLBACK` target: Post button enabled | Ribbon `getEnabled` | Post button enabled normally |
| Signed-out role ribbon user label does not display Windows/NAS identity | Ribbon `getLabel` / runtime status label | Shows `Sign In` or `<not signed in>` until `IsSignedIn() = True` |
| Role current write rejects signed-out operator | Role event creator using current user | Write blocked before queueing; no fallback user is substituted |
| Role current write rejects signed-in user without required capability | Role event creator using current user | Write blocked before queueing; capability error surfaced |
| Role current write rejects fallback target even when auth/capability otherwise pass | Role event creator using current user | Write blocked before queueing; operator is told to connect to NAS warehouse |
| Role current write allows signed-in user with required capability on NAS target | Role event creator using current user | Event queued under current invSys user |
| Sign-in uses exact `UserId`, not `DisplayName` | `ValidateUserCredentialForTarget("DisplayName", ...)` where `DisplayName <> UserId` | Returns `AUTH_USER_NOT_FOUND`; signed-in user unchanged |
| Password/PIN reset for existing `UserId` works | Admin updates `tblUsers.PinHash`, then `ValidateUserCredentialForTarget(UserId, newSecret, ...)` | Returns `AUTH_OK`; display label may show `DisplayName` |
| Packaged ribbon capability buttons include `getEnabled` callback mapping | Packaged RibbonX validation | Required-capability buttons use centralized enabled callback; missing callback fails validation |
| `SelectWarehouseTarget` with `requireStationInbox:=True` and `stationId = ""` returns `WH_TARGET_INCOMPLETE` | `SelectWarehouseTarget` | `WH_TARGET_INCOMPLETE`; `outTarget` null |
| `SelectWarehouseTarget` with `requireStationInbox:=False` and `stationId = ""` returns `NAS_OK` | `SelectWarehouseTarget` | `NAS_OK`; `InboxRoot` = warehouse-global path |
| `GetCurrentTarget()` returns copy; mutation by caller does not affect resolver state | `GetCurrentTarget` + field mutation | Second `GetCurrentTarget()` call returns original values |
| Dilbert signed in; Calvin submits wrong PIN; Dilbert remains signed in | `ShowSignInPrompt` / `ValidateUserCredential` | `GetCurrentUserId() = "Dilbert"` after failed Calvin attempt |
| `ShowSignInPrompt` returns `AUTH_CANCELLED` on operator dismiss | `ShowSignInPrompt` | Return value = `AUTH_CANCELLED`; `IsSignedIn()` unchanged |
| Failed sign-in log row contains no `secretText` | `ValidateUserCredential` failure | Log fields: `userId`, warehouse, station, timestamp, `AuthStatusCode` only |
| `ValidateUserCredentialForTarget` rejects mismatched `target.WarehouseId` | Mismatched target | Returns `AUTH_WAREHOUSE_MISMATCH`; cache unchanged |
| `IsSignedIn()` returns `False` after TTL expiry; `GetAuthStatus()` returns `AUTH_REAUTH_REQUIRED` | TTL expiry simulation | `IsSignedIn() = False`; ribbon label reflects expiry; `CanPerform` returns `False` |
| `CanPerform` returns `False` when `IsSignedIn() = False` without opening auth workbook | `CanPerform` with signed-out state | Returns `False`; no workbook open event fired |
| Two LAN stations: same `WarehouseId`, distinct `StationId`, independent `InboxRoot` | `GetCurrentTarget()` per station | Each station's `StationId` and `InboxRoot` independent |
| Restored remembered target recomputes `ConfigPath` from `HubRoot` + `RuntimeRoot`; stale hint discarded | Priority 2 restore | `outTarget.ConfigPath` matches live path; stale stored hint not used |
| Roaming user (`StationId = ""`, `requireStationInbox = False`) accepted; warehouse-global inbox assigned | `SelectWarehouseTarget` | `NAS_OK`; `InboxRoot` = warehouse-global path |
| Station-required role (`requireStationInbox = True`) rejects `StationId = ""` | `SelectWarehouseTarget` | Returns `WH_TARGET_INCOMPLETE` |
| Stale capability TTL fails closed at write; refreshes before next processor run | `IsSignedIn` TTL expiry + `CanPerform` | Write blocked after TTL; refreshed before next batch |
