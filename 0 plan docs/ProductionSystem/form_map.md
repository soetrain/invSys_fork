# Production System - Form Map (Draft)

## UI controls
- Buttons: Select Recipe, Select # batches, To USED, Confirm MADE, Send to TOTAL INV
- Builders: Recipe list builder, Acceptable item list builder

## Runtime tables (Production sheet)
| Table | Columns |
| --- | --- |
| Recipe chooser | INGREDIENT, UOM, QUANTITY |
| Inventory item chooser | ITEM, UOM, QUANTITY, LOCATION, ROW |
| Batch tracking / recall codes | BATCH, RECALL CODE |
| Check_invSys | RECEIVING, USED, MADE, TOTAL INV, ROW |

## Flow (high level)
Select Recipe -> Recipe chooser -> Inventory item chooser -> Batch/recall -> Check_invSys

## Item search behavior
- Opens from Inventory item chooser
- Filtered to acceptable items per recipe ingredient
- Optional filter: shippable only
