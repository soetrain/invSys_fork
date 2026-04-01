## Slice 1 — Retire/Migrate spec and validation

## Slice 2 — Password / re-auth gate

## Slice 3 — Archive package writer

## Slice 4 — Migration / import routine

## Slice 5 — Retirement marker + tombstone

## Slice 6 — Admin form

## Slice 7 — Ribbon button

## Slice 8 — End-to-end evidence

## Key invariants Codex must never violate

```
INVARIANTS — do not violate:
- Archive must complete before any retirement or deletion step runs
- DeleteLocalRuntime is only callable when OperationMode = MODE_ARCHIVE_RETIRE_DELETE
  AND spec.ConfirmedByUser = True AND tombstone already exists
- Auth passwords are never written plaintext to any export
- Target warehouse config identity (WarehouseId, WarehouseName, StationId) is never overwritten by migration
- Inbox files are never copied during migration
- SharePoint is never made authoritative at any step
- modProcessor remains the single writer — MigrateInventoryToTarget posts events, not direct table writes
```