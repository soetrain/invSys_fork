[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

$spec = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw
$form = Get-Content -LiteralPath (Join-Path $repo "src\Admin\Forms\frmAddInventoryItem.frm") -Raw
$admin = Get-Content -LiteralPath (Join-Path $repo "src\Admin\Modules\modAdmin.bas") -Raw
$catalogCompletion = Get-Content -LiteralPath (Join-Path $repo "src\Admin\Modules\modAdminInventoryCreate.bas") -Raw
$createWire = Get-Content -LiteralPath (Join-Path $repo "src\Core\Modules\modInventoryCreateWire.bas") -Raw
$apply = Get-Content -LiteralPath (Join-Path $repo "src\InventoryDomain\Modules\modInventoryApply.bas") -Raw
$queries = Get-Content -LiteralPath (Join-Path $repo "src\InventoryDomain\Modules\modInventoryQueries.bas") -Raw
$productionRun = Get-Content -LiteralPath (Join-Path $repo "src\Production\Modules\modProductionReusableRun.bas") -Raw
$viewer = Get-Content -LiteralPath (Join-Path $repo "src\Core\Modules\modInventoryViewerData.bas") -Raw
$validator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_phase6_packaged_xlams.ps1") -Raw
$addFunction = [regex]::Match(
    $admin,
    'Public Function AddInventoryItemForWarehouse(?s:.*?)End Function').Value
$setFunction = [regex]::Match(
    $admin,
    'Private Function SetInventoryQuantityForWarehouse(?s:.*?)End Function').Value
$createPayload = [regex]::Match(
    $admin,
    'Private Function BuildInventoryCreatePayloadAdmin(?s:.*?)End Function').Value
$wirePayload = [regex]::Match($createWire, 'Public Function BuildPayload(?s:.*?)End Function').Value
$wireQueue = [regex]::Match($createWire, 'Public Function QueueCreate(?s:.*?)End Function').Value
$owningCreateWire =
    $wirePayload.Contains('modRoleEventWriter.CreateInventoryEntityPayloadItem(') -and
    $wirePayload.Contains('modRoleEventWriter.BuildPayloadJsonFromCollection(items)') -and
    $wireQueue.Contains('modRoleEventWriter.QueueInventoryCreateEvent(warehouse, station,') -and
    $wireQueue.Contains('user, payload, note, 0, Nothing, eventId, report, "", attempted)')

$checks = [ordered]@{}
$checks["Docs.Slice4akContract"] =
    $spec.Contains("single-item **Add Item** action") -and
    $plan.Contains("Slice 4ak -- Admin-created inventory visibility and catalog dropdowns") -and
    $controls.Contains("Slice 4ak Admin-created inventory visibility and dropdowns")
$checks["Admin.AddCreatesExactEntity"] =
    $addFunction.Contains('CreateSystemKey') -and
    $addFunction.Contains('BuildInventoryCreatePayloadAdmin') -and
    $createPayload.Contains('modInventoryCreateWire.BuildPayload(') -and
    $addFunction.Contains('modInventoryCreateWire.QueueCreate(') -and
    $owningCreateWire -and
    -not $addFunction.Contains('QueueMigrationSeedEvent')
$checks["Admin.CatalogOnlyExplicitCompletion"] =
    $admin.Contains("InventorySkuHasManagedEntityAdmin") -and
    $setFunction.Contains("CreateFirstInventoryEntityForCatalogItemAdmin") -and
    $admin.Contains('modAdminInventoryCreate.CompleteCatalogItem(') -and
    $catalogCompletion.Contains('modRoleEventWriter.CreateSystemKey()') -and
    $catalogCompletion.Contains('modInventoryCreateWire.BuildPayload(') -and
    $catalogCompletion.Contains('modInventoryCreateWire.QueueCreate(') -and
    $owningCreateWire -and
    $catalogCompletion.Contains('modInventoryCommands.ConfirmCreate(context, eventId, report)')
$checks["Form.LocationAndCategoryDropdowns"] =
    $form.Contains("Private WithEvents mCmbLocation As MSForms.ComboBox") -and
    $form.Contains("Private WithEvents mCmbCategory As MSForms.ComboBox") -and
    $form.Contains("Private Sub mCmbLocation_Click()") -and
    $form.Contains("Private Sub mCmbCategory_Click()") -and
    $form.Contains('modInventoryUiText.AddCombo(Me.Controls, "cmbLocation"') -and
    $form.Contains('modInventoryUiText.AddCombo(Me.Controls, "cmbCategory"') -and
    $form.Contains("LoadInventoryDimensionOptions")
$checks["Form.RealSubmitHandlerEvidence"] =
    $form.Contains("Public Function TestAddItemVisibilityDropdownContract") -and
    $form -match 'TestAddItemVisibilityDropdownContract(?s).*?mBtnOK_Click'
$checks["ViewerAndProductionUseManagedProjection"] =
    $queries.Contains('Set loEntities = FindInventoryQueryTable(wb, "tblInventoryEntities")') -and
    $queries.Contains('If systemKey = "" Or sku = "" Then GoTo NextEntity') -and
    $queries.Contains('InventoryQueryCatalogIsNonCounted') -and
    $apply.Contains('eventType = EVENT_TYPE_INVENTORY_CREATE And qty = 0 Then GoTo QtyAccepted') -and
    $productionRun.Contains('ExactEntityIsNonCounted') -and
    $viewer.Contains('Set snapshotTable = ViewerFindTable(snapshotWb, SNAPSHOT_TABLE)') -and
    $viewer.Contains('If inventoryState = "RETIRED" Then GoTo ContinueRow') -and
    $viewer.Contains('If CDbl(quantities(key)) >= 0 Then visibleCount = visibleCount + 1')
$checks["Packaged.AdminAddEvidence"] =
    $admin.Contains("InventoryAddVisibilityDropdownContractForAutomation") -and
    $validator.Contains("Admin.InventoryAddVisibilityDropdowns") -and
    $validator.Contains("ExactEntityCreate=True") -and
    $validator.Contains("LocationDropdown=True") -and
    $validator.Contains("CategoryDropdown=True")

$passed = 0
$red = 0
foreach ($entry in $checks.GetEnumerator()) {
    if ($entry.Value) {
        Write-Host ("PASS " + $entry.Key)
        $passed++
    }
    else {
        Write-Host ("RED  " + $entry.Key)
        $red++
    }
}

Write-Host ("PLAN022_SLICE4AK_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
