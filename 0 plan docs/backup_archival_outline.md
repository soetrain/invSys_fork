# Backup + Archival System (Draft Outline)

Purpose: provide external, long‑term history independent of Excel/OneDrive while keeping the workbook small and recoverable.

## Goals
- Recover data after user mistakes even if tables are corrupted.
- Avoid hitting Excel row limits (1,048,576).
- Keep daily workflow fast; archival should be lightweight and optional.

## Core approach
**External snapshots** of critical tables to a user‑chosen folder, plus **append‑only in‑workbook logs** for audit.

### A) External snapshots (primary recovery)
- Export key tables to timestamped files (CSV or XLSX) in `BackupRoot`.
- One folder per workbook + date partitioning.
- Optionally zip older days to reduce clutter.

**Suggested export sets**
- invSys (core inventory)
- Production tables (Recipe chooser, Inventory chooser, Batch codes, Check_invSys)
- Shipping tables (ShipmentsTally, NotShipped, Aggregate tables)
- Builder tables (RecipeList, AcceptableItems, ShippingBOM)
- _SystemMeta (schema definitions)

### B) In‑workbook logs (secondary audit)
- Continue writing InventoryLog / ActionLog rows.
- Used for “what happened” but not relied on for full restore.

## Backup triggers
- Manual button: **“Create Backup Snapshot”**
- Optional auto‑trigger:
  - On workbook close
  - On “major actions” (Confirm Made, Shipments Sent, Save Recipe)

## File format
Preferred: CSV for speed + size.
Optional: XLSX for readability (slower, larger).

Example folder structure:
```
BackupRoot/
  invSys_2025-12-29_235900.csv
  Production_RecipeChooser_2025-12-29_235900.csv
  Shipping_ShipmentsTally_2025-12-29_235900.csv
  _SystemMeta_2025-12-29_235900.csv
```

## Minimal metadata for backup system
Stored in `_SystemMeta`:
- BackupRoot path
- LastBackupTimestamp
- BackupFormat (CSV/XLSX)
- AutoBackupEnabled (true/false)

## Restore flow (future)
1. Regenerate schema from `_SystemMeta`.
2. Choose backup timestamp.
3. Import CSVs to recreate tables.

## Risks / mitigations
- **Large file sizes**: use CSV + optional zip rotation.
- **User forgets to set folder**: prompt on first backup.
- **Schema mismatch**: check `SCHEMA_VERSION` before restore.
