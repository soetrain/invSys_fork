# Shipping / Production Feature Completion Slice List

Date: 2026-05-27

Baseline before this slice:
- Phase 6 split suite: 126/126 passing
- Packaged ribbon: 136/136 passing
- Live role workflow: 34/34 passing

## Slice 1 - Existing Inventory Shipping Path

Status: completed

Goal: validate the practical operator flow for shipping inventory that already exists in the warehouse.

Flow:
1. Aggregate package demand exists in `AggregatePackages`.
2. `BtnToShipments` stages package demand into operator `invSys.SHIPMENTS`.
3. `BtnShipmentsSent` queues a `SHIP` event.
4. Processor applies the event to canonical inventory and `tblInventoryLog`.
5. Local shipping staging is cleared so the operator cannot accidentally post the same shipment twice.

Current implementation notes:
- `BtnToShipments` now applies staging directly to the operator workbook `invSys` table instead of the legacy `modInvMan` active-workbook resolver.
- Live workflow now includes `Shipping.BtnToShipments.Local`.

Validation:
- Covered by `tools/validate_phase6_live_role_workflows.ps1`; validated at 34/34 before Slice 4 hold checks were added.

## Slice 2 - Shipping Box Build Path

Status: in progress

Goal: bring the old box-build flow into the D-NAS model.

Candidate flow:
1. Package builder saves package and component BOM.
2. `BtnConfirmInventory` validates component availability without writing canonical state.
3. `BtnBoxesMade` converts components and package outputs into a canonical production-style event.
4. Runtime processor applies component consumption and package completion.
5. Local box staging clears or refreshes predictably.

Open decision:
- Current event semantics keep intermediate made-output records audit/local-only. Canonical on-hand increases at the completion/`To Total Inv` step.

Current implementation notes:
- `BtnBoxesMade` now applies component usage and package `MADE` staging directly to the operator workbook `invSys` table instead of using the legacy `modInvMan` active-workbook resolver.
- Shipping `BtnToTotalInv` now applies `MADE -> TOTAL INV` directly to the operator workbook `invSys` table.
- Live workflow now includes `Shipping.BtnBoxesMade.Local` and `Shipping.BtnToTotalInv.Local`.

Validation:
- Covered by `tools/validate_phase6_live_role_workflows.ps1`; validated at 34/34 before Slice 4 hold checks were added.

## Slice 3 - Production Send To MADE

Status: completed

Goal: validate production's first write step, not only the final `BtnToTotalInv` completion step.

Flow:
1. Process/output tables generate `ProductionOutput`.
2. `BtnToMade` queues `PROD_CONSUME`.
3. Runtime processor applies component consumption and staged made output.
4. Operator workbook refresh leaves `USED`/`MADE` state coherent for the next step.

Current implementation notes:
- Existing code already queues `PROD_CONSUME`.
- Existing code already applies local `USED` and `MADE` changes against the operator workbook.
- Live workflow now includes `Production.BtnToMade.Local`, `.Queue`, `.Process`, and `.InventoryLog`.
- `PROD_CONSUME` MADE payload lines are audit/local staging only; `PROD_COMPLETE` changes canonical finished-goods on-hand.

## Slice 4 - Hold / Unship

Status: completed

Goal: validate row movement between `ShipmentsTally` and `NotShipped`.

Scope:
- `BtnUnship`
- `BtnSendHold`
- `BtnReturnHold`
- Aggregate rebuild correctness after moving rows.

Current implementation notes:
- Added `MoveShipmentHoldForAutomation` so the live harness can exercise hold movement without relying on Excel selection or modal quantity prompts.
- Live workflow now includes `Shipping.Hold.ToggleNotShipped`, `Shipping.Hold.Send`, and `Shipping.Hold.Return`.

Validation:
- Run `tools/validate_phase6_live_role_workflows.ps1`; expected total is 37 checks while live workflow preflight diagnostics remain enabled.

## Slice 5 - Operator UI Tune-Up

Status: in progress

Goal: polish role sheets for user-side testing.

Scope:
- Button text and status messages.
- Reduce modal popups for normal success paths where ribbon/status labels can carry the state.
- Make stale default-path inbox warnings actionable.
- Confirm user display name, selected warehouse, and connected state are visible where needed.

Current implementation notes:
- Shipping clean success paths now write concise completion text to Excel's status bar instead of opening modal dialogs.
- Production clean success paths now write concise completion text to Excel's status bar instead of opening modal dialogs.
- Actionable warnings and errors still use modal dialogs.
- Receiving pending-inbox warnings now explain that the row was written to the station inbox and identify the path to inspect if processing remains stuck.
