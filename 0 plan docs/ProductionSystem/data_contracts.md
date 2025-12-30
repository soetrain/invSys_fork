# Production System - Data Contracts (Draft)

## Builder sheets
| Table | Columns | Notes |
| --- | --- | --- |
| Recipe list builder | RECIPE, INGREDIENT, UOM, QTY_PER_BATCH | Source of recipe rows |
| Acceptable items builder | RECIPE, INGREDIENT, ITEM, ROW | Controls item search filtering |

## Production sheet
| Table | Columns | Notes |
| --- | --- | --- |
| Recipe chooser | INGREDIENT, UOM, QUANTITY | Quantities are batch-adjusted |
| Inventory item chooser | ITEM, UOM, QUANTITY, LOCATION, ROW | ITEM must be acceptable per ingredient |
| Batch / recall codes | BATCH, RECALL CODE | Captured before Confirm MADE |
| Check_invSys | RECEIVING, USED, MADE, TOTAL INV, ROW | Filtered to rows in use |

## invSys + logging
| Table | Columns | Notes |
| --- | --- | --- |
| invSys | ROW, ITEM, ITEM_CODE, UOM, LOCATION, USED, MADE, TOTAL INV | Source of truth |
| InventoryLog | LOG_ID, USER, ACTION, ROW, ITEM_CODE, ITEM_NAME, QTY_CHANGE, NEW_QTY, TIMESTAMP | All deltas logged via InvMan |
