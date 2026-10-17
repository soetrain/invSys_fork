[CmdletBinding()]
param(
    [string]$RepoRoot = (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
)

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

function Read-OptionalText([string]$Path) {
    if (Test-Path -LiteralPath $Path) { return Get-Content -LiteralPath $Path -Raw }
    return ""
}

$spec = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md") -Raw
$plan = Get-Content -LiteralPath (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md") -Raw
$controls = Get-Content -LiteralPath (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md") -Raw
$build = Get-Content -LiteralPath (Join-Path $repo "tools\build-xlam.ps1") -Raw
$ribbonValidator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_phase6_packaged_ribbon.ps1") -Raw
$packagedValidator = Get-Content -LiteralPath (Join-Path $repo "tools\validate_phase6_packaged_xlams.ps1") -Raw
$form = Get-Content -LiteralPath (Join-Path $repo "src\Admin\Forms\frmAddInventoryItem.frm") -Raw
$admin = Get-Content -LiteralPath (Join-Path $repo "src\Admin\Modules\modAdmin.bas") -Raw
$worksheet = Read-OptionalText (Join-Path $repo "src\Admin\Modules\modAdminInventoryWorksheet.bas")

$checks = [ordered]@{}
$checks["Docs.Slice4aiNormativeContract"] =
    $spec.Contains("Admin inventory worksheet workbench") -and
    $plan.Contains("Slice 4ai -- Admin bulk inventory worksheet staging") -and
    $controls.Contains("Slice 4ai Admin inventory worksheet workbench")
$checks["Ribbon.AddEditInventoryItems"] =
    $build.Contains('Label = "Add/Edit Inventory Items"; Macro = "modAdmin.Add_InventoryItem"') -and
    $ribbonValidator.Contains('Label = "Add/Edit Inventory Items"; Macro = "modAdmin.Add_InventoryItem"')
$checks["Form.VisibleWorksheetActions"] =
    $form.Contains("mBtnCreateInventoryTable") -and
    $form.Contains("mBtnUploadInventoryTable") -and
    $form.Contains('"Create Inventory Table"') -and
    $form.Contains('"Upload Selected Inventory Table"')
$checks["Form.CapturedWorkbookHandlers"] =
    $form.Contains("Public Sub SetOperatorWorkbook") -and
    $form -match 'mBtnCreateInventoryTable_Click\(\)(?s).*?CreateInventoryWorksheetTable' -and
    $form -match 'mBtnUploadInventoryTable_Click\(\)(?s).*?UploadSelectedInventoryWorksheetTable'
$checks["Worksheet.ManagedShapeAndValidation"] =
    $worksheet.Contains('EDITOR_SHEET As String = "invSys Inventory Editor"') -and
    $worksheet.Contains('TABLE_PREFIX As String = "invSys_Inventory_"') -and
    $worksheet.Contains('"Action", "Item Code", "Item Name", "UOM", "Qty Mode", "Quantity"') -and
    $worksheet.Contains("ApplyInventoryWorksheetValidation") -and
    $worksheet.Contains('"COUNTED,UTILITY,SERVICE,NOT COUNTED"')
$checks["Worksheet.IdentityAndCustomHeaders"] =
    $worksheet.Contains("ValidateInventoryWorksheetHeaders") -and
    $worksheet.Contains('Chr$(82) & Chr$(79) & Chr$(87)') -and
    $worksheet.Contains('normalized = "SYSTEMKEY"') -and
    $worksheet.Contains("AppendInventoryWorksheetCustomFields") -and
    $worksheet.Contains("GenerateInventoryItemCodeForWorksheet")
$checks["Worksheet.WholeTablePreflightAndStatus"] =
    $worksheet.Contains("PreflightInventoryWorksheetTable") -and
    $worksheet.Contains("Validate every pending row before") -and
    $worksheet.Contains('"Upload Status"') -and
    $worksheet.Contains('"Upload Result"') -and
    $worksheet.Contains("InventoryItemCatalogContainsForWorksheet")
$checks["Packaged.RealHandlerContract"] =
    $form.Contains("Public Function TestInventoryWorksheetActionContract") -and
    $form -match 'TestInventoryWorksheetActionContract(?s).*?mBtnCreateInventoryTable_Click(?s).*?mBtnUploadInventoryTable_Click' -and
    $admin.Contains("Public Function InventoryWorksheetContractForAutomation") -and
    $packagedValidator.Contains("Admin.InventoryWorksheetActions") -and
    $packagedValidator.Contains("TableCreated=True") -and
    $packagedValidator.Contains("Preflight=True") -and
    $packagedValidator.Contains("Utility=True")

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

Write-Host ("PLAN022_SLICE4AI_SOURCE passed={0} red={1} total={2}" -f $passed, $red, $checks.Count)
if ($red -gt 0) { exit 1 }
