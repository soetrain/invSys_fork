# Handoff Baton - BoxBuilder UserForm Migration

Date: 2026-06-12

## Current State

The Shipping BoxBuilder/BoxMaker work hit a structural limit with the ribbon-buttons-editing-worksheet-tables approach. The user likes the visual sheet/table style, but performance and stability are not acceptable:

- Excel blinks and opens/closes runtime workbooks during cell selection.
- `Worksheet_SelectionChange` and `Worksheet_Change` recursively trigger picker logic, table writes, version autofill, and refresh logic.
- BoxBOM picker/source bugs repeatedly returned because the sheet event path mixes real inventory, ShippingBOM recipes, and version projections.
- First-row/table edge cases remain fragile.
- Version rows have drifted visually between boxes during sheet/table refreshes.

`tests/latest problem help.md` gives the architectural diagnosis: move BoxBuilder interaction into a UserForm, keep the sheet tables as a projection/display layer, and do one quiet batch commit on Save.

## Latest Implemented Slice

This session added the first UserForm-based BoxBuilder path beside the old sheet workflow.

Files changed:

- `src/Shipping/Forms/frmShippingBoxBuilder.frm`
  - New runtime-built UserForm.
  - Loads current BoxBuilder metadata into textboxes.
  - Loads current BoxBOM rows into a form ListBox.
  - Loads managed inventory once into memory using `modTS_Shipments.LoadShippingComponentPickerItems`.
  - Provides in-form inventory search, Add, Remove, Save, Cancel.
  - Save sends the form state to the bridge commit routine.

- `src/Shipping/Modules/modTS_Shipments.bas`
  - Added `BtnOpenBoxBuilder`.
  - Added public bridge functions:
    - `BoxBuilderFormCurrentMeta`
    - `BoxBuilderFormCurrentComponents`
    - `CommitBoxBuilderFormState`
  - Added helper:
    - `EnsureListObjectHasRowsShipping`
  - `BtnOpenBoxBuilder` now only runs full `InitializeShipmentsUiForWorkbook` if the Shipping sheet or required tables are missing; normal form opening is lightweight.
  - `CommitBoxBuilderFormState` does a quiet batch write:
    - disables events
    - suppresses the generated identity edit guard
    - writes BoxBuilder and BoxBOM table rows once
    - restores UI/event state
    - calls existing `BtnSaveBox` persistence path

- `tools/build-xlam.ps1`
  - Added new Shipping ribbon button:
    - Id: `btnShippingBoxBuilderForm`
    - Label: `Box Builder`
    - Macro: `modTS_Shipments.BtnOpenBoxBuilder`
    - Required capability: `SHIP_POST`

## Validation Status

Latest validation after the final lightweight-open fix:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./tools/build-xlam.ps1 -Apply
```

Result: `Build complete.`

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./tools/validate_phase6_packaged_ribbon.ps1
```

Result:

```text
PHASE6_PACKAGED_RIBBON_VALIDATION_OK
PASSED=227 FAILED=0 TOTAL=227
```

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./tools/run_phase6_excel_validation.ps1 -StartAt 109
```

Result:

```text
PHASE6_VALIDATION_OK
PASSED=20 FAILED=0 TOTAL=20
RANGE=109-128 AVAILABLE=128
```

The user tested the new form path and reported: "its a lot better". This confirms the UserForm direction is viable, but the form is not feature-complete yet.

## Immediate User Feedback / Next Requirements

The form now needs dropdowns and full box/version management:

1. Add dropdown/list for existing boxes.
   - User should be able to find/select saved boxes from warehouse ShippingBOM.
   - Selecting a box should load its metadata, version list, and the selected version's components into the form.

2. Add version dropdown/list.
   - Versions must be strongly tied to the selected box.
   - Selecting a version should load that version's components.
   - If a box has only `v1`, it must still show `v1`.
   - No cross-box version drift.

3. Add explicit edit/save/version commands.
   - Update selected version.
   - Save as New Version.
   - Delete selected version.
   - Delete selected box.
   - Status dropdown: `Active` / `Retired`.
   - Admin-only actions should still require `ADMIN_MAINT`.

4. Improve form layout.
   - Current form is a functional first slice, not final UX.
   - It needs clear sections:
     - box selector
     - version selector/status
     - box metadata
     - managed inventory search/list
     - BOM component list
     - action buttons

5. Keep old sheet path available temporarily.
   - Do not delete worksheet/table picker code yet.
   - Once the form supports box/version workflows reliably, make the form the canonical BoxBuilder path.
   - Then disable the `SelectionChange -> picker` trigger for `BoxBOM`.

## Design Direction

Preferred end architecture:

```text
Shipping ribbon button
  -> frmShippingBoxBuilder
      -> loads boxes, versions, and managed inventory once
      -> user edits entirely in form memory
      -> Save does one quiet batch commit
      -> existing runtime ShippingBOM persistence remains authoritative
      -> sheet tables become projection/display, not primary interaction UI
```

The form should not write to sheet cells on every selection. It should keep all user edits in memory until the user clicks Save/Update/New Version/Delete.

## Important Existing Behavior To Preserve

- Warehouse runtime ShippingBOM remains authoritative.
- Operator workbook tables are projections.
- `BtnSaveBox` currently contains the proven persistence path to runtime ShippingBOM.
- BoxBOM rows represent components from managed inventory, not ShippingBOM recipes.
- Shippable boxes are rows from `*.invSys.Data.ShippingBOM.xlsb`.
- BoxMaker uses saved BoxBOM recipes to make/unmake real inventory.
- `ROW`, `ITEM_CODE`, and `Version` are generated/managed identity fields, not normal user-edit fields.

## Known Risks / Watch Points

- `modTS_Shipments.bas` is still very large and fragile.
- The new form currently commits by projecting back to sheet tables and then calling `BtnSaveBox`; this is safer than live cell editing but still not the final clean architecture.
- Eventually extract form commit into a smaller module, e.g. `modShipping_BoxUI.bas`, and reduce direct dependence on the 117 KB monolith.
- Full `git status` was slow/hung in this worktree; use targeted `git diff -- <files>` or `git diff --name-only -- <files>` when possible.
- The new form file is new/untracked until added to git:
  - `src/Shipping/Forms/frmShippingBoxBuilder.frm`

## Suggested Next Implementation Steps

1. Add data-provider functions in `modTS_Shipments` for the form:
   - `BoxBuilderFormLoadSavedBoxes() As Variant`
   - `BoxBuilderFormLoadVersions(packageRow) As Variant`
   - `BoxBuilderFormLoadVersionComponents(packageRow, versionLabel) As Variant`
   - These should source from `ShippingBOMView` or runtime ShippingBOM in one refresh/load, not from worksheet selection events.

2. Add controls to `frmShippingBoxBuilder`:
   - `ComboBox` or `ListBox` for boxes.
   - `ComboBox` or `ListBox` for versions.
   - Status `ComboBox` with `Active`, `Retired`.
   - Buttons: `Update Version`, `New Version`, `Delete Version`, `Delete Box`.

3. Update save semantics:
   - `Update Version` should pass selected version into the existing replace-version path.
   - `New Version` should force a new version.
   - Do not show ambiguous Yes/No prompts for version creation.
   - The user explicitly asked for clear buttons, not vague prompt choices.

4. Keep performance discipline:
   - No `Workbooks.Open` or runtime refresh inside list selection events unless absolutely required.
   - Load/caches at form open or explicit Refresh.
   - Use `modUiQuiet.BeginQuietUi` and `Application.EnableEvents = False` only around final batch commit.

5. After user validates the form path:
   - Disable `HandleShippingSelectionChange` picker behavior for `BoxBOM`.
   - Keep table display/projection.
   - Remove dead picker/autofill patches gradually.

## Commands For Next Session

Build:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./tools/build-xlam.ps1 -Apply
```

Ribbon validation:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./tools/validate_phase6_packaged_ribbon.ps1
```

Focused Phase 6 validation:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File ./tools/run_phase6_excel_validation.ps1 -StartAt 109
```

Useful targeted diffs:

```bash
git diff -- tools/build-xlam.ps1 src/Shipping/Modules/modTS_Shipments.bas src/Shipping/Forms/frmShippingBoxBuilder.frm
git ls-files --others --exclude-standard src/Shipping/Forms/frmShippingBoxBuilder.frm
```

