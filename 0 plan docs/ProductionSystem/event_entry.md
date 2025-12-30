# Production System - Event Entry Map (Draft)

## Sequence (user + system)
1. User selects Recipe (button or dropdown).
2. System loads recipe rows into Recipe chooser.
3. User selects # batches.
4. System multiplies recipe quantities by batch count.
5. User selects an INGREDIENT row.
6. System opens item search filtered to acceptable items.
7. User picks ITEM and enters QUANTITY.
8. System writes selection to Inventory item chooser and rebuilds Check_invSys.
9. User clicks To USED.
10. System stages components to invSys.USED and logs via InvMan.
11. User clicks Confirm MADE.
12. System moves finished qty to invSys.MADE and logs via InvMan.
13. User clicks Send to TOTAL INV.
14. System moves MADE to TOTAL INV, clears staging tables, logs via InvMan.
