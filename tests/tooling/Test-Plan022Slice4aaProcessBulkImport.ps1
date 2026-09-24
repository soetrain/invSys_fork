[CmdletBinding()]
param([string]$RepoRoot = ".")

$ErrorActionPreference = "Stop"
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$docs = (Resolve-Path -LiteralPath (Join-Path $repo "..\invSys_docs")).Path

function Read-Text([string]$Path) { Get-Content -Raw -LiteralPath $Path }

$spec = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Design-v4.11.md")
$plan = Read-Text (Join-Path $docs "expert guidance docs\022 Deployed Operations Launcher and NAS Runtime Stabilization Plan.md")
$controls = Read-Text (Join-Path $docs "0 plan docs\xlam_invSys\invSys-Controls-v1.md")
$form = Read-Text (Join-Path $repo "src\Production\Forms\frmProduction.frm")
$worksheet = Read-Text (Join-Path $repo "src\Production\Modules\modProductionProcessWorksheet.bas")
$events = Read-Text (Join-Path $repo "src\Production\ClassModules\cProductionAppEvents.cls")
$picker = Read-Text (Join-Path $repo "src\Core\ClassModules\cDynItemSearch.cls")
$designs = Read-Text (Join-Path $repo "src\DesignsDomain\Modules\modDesignsApply.bas")
$eventWriter = Read-Text (Join-Path $repo "src\Core\Modules\modRoleEventWriter.bas")
$processor = Read-Text (Join-Path $repo "src\Core\Modules\modProcessor.bas")
$validator = Read-Text (Join-Path $repo "tools\validate_plan022_packaged_launchers.ps1")

$checks = @(
    [pscustomobject]@{ Name = "Docs.Slice4aaContract"; Pass =
        $spec -match 'Ctrl\+click multi-area selection' -and
        $plan -match 'Slice 4aa -- text-safe Process identities' -and
        $controls -match 'btnProcessWorksheetAddAlternative' },
    [pscustomobject]@{ Name = "Worksheet.TextSafeGeneratedIds"; Pass =
        $worksheet -match 'ApplyProcessWorksheetTextIdentityFormats' -and
        $worksheet -match 'NumberFormat\s*=\s*"@"' -and
        $worksheet -match 'COL_REQUIREMENT_ID\)\.Formula\s*=\s*"=\[@ID\]"' -and
        $designs -match 'IsTextDesignIdentityColumn' -and
        $designs -match 'targetCell\.NumberFormat\s*=\s*"@"' -and
        $eventWriter -match 'StrComp\(columnName, "DesignId"' -and
        $eventWriter -match 'targetCell\.NumberFormat\s*=\s*"@"' -and
        $processor -match 'CanonicalReusableDesignEventIdProcessor' },
    [pscustomobject]@{ Name = "Worksheet.CatalogUomValidation"; Pass =
        $worksheet -match 'GetConfiguredUomsPackedText' -and
        $worksheet -match 'ApplyProcessWorksheetUomValidation' -and
        $worksheet -match 'UOM is not in the Recipe UOM Catalog' },
    [pscustomobject]@{ Name = "Worksheet.NumberedAlternatives"; Pass =
        $worksheet -match 'Acceptable Managed Item 1' -and
        $worksheet -match 'AddAcceptableItemPairToSelectedTable' -and
        $worksheet -match 'Accepted SKU " & CStr' },
    [pscustomobject]@{ Name = "Picker.NumberedCommit"; Pass =
        $picker -match 'ProcessAlternativePairNumber' -and
        $picker -match 'Acceptable Managed Item " & CStr' -and
        $events -match 'ShowProductionProcessItemSearch' },
    [pscustomobject]@{ Name = "Form.AddAlternativeAction"; Pass =
        $form -match 'btnProcessWorksheetAddAlternative' -and
        $form -match 'Private Sub mBtnProcessWorksheetAddAlternative_Click\(\)' },
    [pscustomobject]@{ Name = "Worksheet.MultiAreaSelection"; Pass =
        $worksheet -match 'FindSelectedProcessWorksheetTables' -and
        $worksheet -match 'selectedRange\.Areas' -and
        $form -match 'SubmitDesignerAction\(True, "PROCESS_SAVE"' },
    [pscustomobject]@{ Name = "Packaged.PublicOperatorBoundary"; Pass =
        $validator -match 'RunProcessWorksheetBulkImportContractTest' -and
        $validator -match 'TextSafeIds=True' -and
        $validator -match 'PickerOpened=True' -and
        $validator -match 'MultiTableDrafts=True' }
)

$failed = @($checks | Where-Object { -not $_.Pass })
foreach ($check in $checks) {
    "{0} {1}" -f $(if ($check.Pass) { "PASS" } else { "RED" }), $check.Name
}
"PLAN022_SLICE4AA_SOURCE passed=$($checks.Count - $failed.Count) red=$($failed.Count) total=$($checks.Count)"
if ($failed.Count -gt 0) { exit 1 }
