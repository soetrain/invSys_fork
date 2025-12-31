# Production System - Minimal Metadata Outline

Purpose: store just enough metadata inside the workbook to support undo/redo and regeneration later.

## Location
- Hidden worksheet: `_SystemMeta`
- All metadata stored as ListObjects for easy rebuild and export.

## Tables

### tblSystemVersion
| Field | Type | Notes |
| --- | --- | --- |
| SCHEMA_VERSION | text | Increment when schema changes |
| BUILD_DATE | datetime | For traceability |
| OWNER | text | Optional (workbook author) |

### tblSchemaTables
| Field | Type | Notes |
| --- | --- | --- |
| TABLE_NAME | text | ListObject name (unique) |
| SHEET_NAME | text | Sheet where the table lives |
| ANCHOR_CELL | text | Top-left cell (e.g., `B6`) |
| HAS_HEADERS | bool | Always true for ListObjects |
| ROW_TEMPLATE | text | Optional: named range for row formatting |
| VERSION | text | Optional per-table version |

### tblSchemaColumns
| Field | Type | Notes |
| --- | --- | --- |
| TABLE_NAME | text | FK to tblSchemaTables |
| COL_NAME | text | Column header |
| COL_ORDER | number | 1-based order |
| DATA_TYPE | text | text / number / date / bool |
| REQUIRED | bool | Validation rule |
| DEFAULT_FORMULA | text | Excel formula to apply |
| READONLY | bool | For system-maintained columns |

### tblUILayout
| Field | Type | Notes |
| --- | --- | --- |
| TABLE_NAME | text | FK to tblSchemaTables |
| COLUMN_WIDTHS | text | Comma list or JSON (optional) |
| HIDDEN_COLUMNS | text | Comma list (optional) |
| STYLE_NAME | text | Optional table style |
| BUTTON_ANCHORS | text | JSON or key:value pairs |

### tblRegenNotes
| Field | Type | Notes |
| --- | --- | --- |
| TABLE_NAME | text | FK to tblSchemaTables |
| REGEN_PRIORITY | number | Order to rebuild |
| CLEAR_BEFORE_REGEN | bool | Whether to purge data |
| NOTES | text | Extra guidance |

## Minimal usage (Phase 1)
- Populate tblSchemaTables and tblSchemaColumns only.
- Regeneration recreates missing tables and columns.

## Undo/Redo readiness (Phase 2)
- Add GUID columns to transactional tables.
- Snapshot tables in `_SystemMeta` or `UndoLog` (separate system).

## Example tables to register (initial)
- RecipeBuilder_Header
- RecipeBuilder_Ingredients
- AcceptableItems_Header
- AcceptableItems_Ingredients
- AcceptableItems_Items
- Production_RecipeChooser
- Production_InventoryChooser
- Production_BatchCodes
- Production_CheckInvSys
