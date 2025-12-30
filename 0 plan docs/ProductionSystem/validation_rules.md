# Production System - Validation Rules (Draft)

## Recipe and batch
- Recipe is required.
- # batches is required.

## Quantities
- QUANTITY must be > 0.
- QTY_PER_BATCH must be present in recipe list.

## Item selection
- ITEM must be from Acceptable list for the chosen ingredient.
- ROW must resolve in invSys.

## Inventory checks
- Enough TOTAL INV for To USED.
- No negative results after staging.

## Batch tracking
- BATCH required when recall tracking is enabled.
- RECALL CODE format must validate (format TBD).

## Logging
- All invSys deltas must be logged via InvMan.
