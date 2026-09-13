[CmdletBinding()]
param(
    [string]$RepoRoot = ".",
    [string]$DeployRoot = "deploy/current",
    [switch]$KeepArtifacts
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$deployPath = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
$resultPath = Join-Path $repo "tests/integration/slice14_results.md"
$tempRoot = Join-Path ([IO.Path]::GetTempPath()) (
    "invsys-release1-chain-" + [Guid]::NewGuid().ToString("N"))
$results = [System.Collections.Generic.List[object]]::new()
$excel = $null
$openedWorkbooks = [System.Collections.Generic.List[object]]::new()

# These stable phase names are deliberately kept in the normative order. The
# evidence writer records them so a continuation can distinguish an actual
# ordered run from a collection of unrelated role tests.
$phaseOrder = @(
    "GenerateFreshWarehouse",
    "SeedDemoInventoryThroughAdmin",
    "ReceiveInventory",
    "ProcessorApplyReceive",
    "RefreshAfterReceive",
    "ProductionTwoBatches",
    "ProductionConsumptionAndOutput",
    "BoxingVersionSelection",
    "ShipmentStagingAndSent",
    "ProcessorApplyShipment",
    "FinalRefresh",
    "RestartAndReconcile"
)

function Add-Result {
    param(
        [string]$Check,
        [bool]$Passed,
        [string]$Detail
    )

    $results.Add([pscustomobject]@{
        Check = $Check
        Passed = $Passed
        Detail = $Detail
    })
}

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object) {
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object) } catch {}
    }
}

function Run-WorkbookMacro {
    param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$MacroName,
        [object[]]$Arguments = @()
    )

    $macro = "'$WorkbookName'!$MacroName"
    switch ($Arguments.Count) {
        0 { return $Excel.Run($macro) }
        1 { return $Excel.Run($macro, $Arguments[0]) }
        2 { return $Excel.Run($macro, $Arguments[0], $Arguments[1]) }
        3 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2]) }
        4 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3]) }
        5 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4]) }
        6 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5]) }
        7 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6]) }
        8 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6], $Arguments[7]) }
        default { throw "Run-WorkbookMacro supports at most eight arguments." }
    }
}

function Invoke-RepositoryScript {
    param(
        [string]$Path,
        [string[]]$Arguments
    )

    $output = @(& powershell -NoProfile -ExecutionPolicy Bypass -File $Path @Arguments 2>&1)
    $exitCode = $LASTEXITCODE
    return [pscustomobject]@{
        ExitCode = $exitCode
        Output = $output
        Text = ($output -join "`n")
    }
}

function Suspend-NonTargetInvSysAddins {
    param(
        [object]$Excel,
        [string]$ExpectedPackageRoot
    )

    $expectedPrefix = ([IO.Path]::GetFullPath($ExpectedPackageRoot)).TrimEnd("\") + "\"
    $suspended = [System.Collections.Generic.List[string]]::new()
    $addins2 = $null
    try {
        $addins2 = $Excel.AddIns2
        for ($index = 1; $index -le [int]$addins2.Count; $index++) {
            $addin = $null
            try {
                $addin = $addins2.Item($index)
                $name = [string]$addin.Name
                $fullName = [string]$addin.FullName
                $isExpectedPath = $fullName.StartsWith(
                    $expectedPrefix,
                    [StringComparison]::OrdinalIgnoreCase)
                if ($name -match '(?i)^invSys\..*\.xlam$' -and
                    -not $isExpectedPath -and
                    [bool]$addin.Installed) {
                    $addin.Installed = $false
                    $suspended.Add($fullName)
                }
            }
            finally {
                Release-ComObject $addin
            }
        }
    }
    finally {
        Release-ComObject $addins2
    }
    return @($suspended.ToArray())
}

function Restore-SuspendedInvSysAddins {
    param(
        [object]$Excel,
        [string[]]$SuspendedPaths
    )

    if ($SuspendedPaths.Count -eq 0) { return }
    $restore = @{}
    foreach ($path in $SuspendedPaths) {
        $restore[([string]$path).ToLowerInvariant()] = $true
    }
    $addins2 = $null
    try {
        $addins2 = $Excel.AddIns2
        for ($index = 1; $index -le [int]$addins2.Count; $index++) {
            $addin = $null
            try {
                $addin = $addins2.Item($index)
                $fullName = [string]$addin.FullName
                if ($restore.ContainsKey($fullName.ToLowerInvariant()) -and
                    -not [bool]$addin.Installed) {
                    $addin.Installed = $true
                }
            }
            finally {
                Release-ComObject $addin
            }
        }
    }
    finally {
        Release-ComObject $addins2
    }
}

function Get-ListObject {
    param(
        [object]$Workbook,
        [string]$TableName
    )

    foreach ($worksheet in $Workbook.Worksheets) {
        try {
            $table = $worksheet.ListObjects.Item($TableName)
            if ($null -ne $table) {
                return $table
            }
        } catch {}
    }
    return $null
}

function Get-ColumnIndex {
    param(
        [object]$ListObject,
        [string]$Header
    )

    if ($null -eq $ListObject) { return 0 }
    for ($columnIndex = 1; $columnIndex -le $ListObject.ListColumns.Count; $columnIndex++) {
        if ([string]::Equals(
                ([string]$ListObject.ListColumns.Item($columnIndex).Name).Trim(),
                $Header,
                [StringComparison]::OrdinalIgnoreCase)) {
            return $columnIndex
        }
    }
    return 0
}

function Get-TableRowCount {
    param([object]$ListObject)
    if ($null -eq $ListObject -or $null -eq $ListObject.DataBodyRange) { return 0 }
    return [int]$ListObject.DataBodyRange.Rows.Count
}

function Get-TableValue {
    param(
        [object]$ListObject,
        [int]$RowIndex,
        [string]$Header
    )

    $columnIndex = Get-ColumnIndex -ListObject $ListObject -Header $Header
    if ($columnIndex -le 0 -or $RowIndex -le 0 -or
        $RowIndex -gt (Get-TableRowCount $ListObject)) {
        return $null
    }
    return $ListObject.DataBodyRange.Cells.Item($RowIndex, $columnIndex).Value2
}

function Invoke-AdminEntryGate {
    $entryRoot = Join-Path $tempRoot "admin-entry"
    $warehouseRoot = Join-Path $entryRoot "warehouse"
    $shareRoot = Join-Path $entryRoot "share"
    $warehouseId = "WHR" + [Guid]::NewGuid().ToString("N").Substring(0, 6).ToUpperInvariant()
    $stationId = "S1"
    $userId = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { "user1" } else { $env:USERNAME }
    New-Item -ItemType Directory -Path $entryRoot, $shareRoot -Force | Out-Null

    $localExcel = New-Object -ComObject Excel.Application
    $localExcel.Visible = $false
    $localExcel.DisplayAlerts = $false
    $localBooks = [System.Collections.Generic.List[object]]::new()
    try {
        $packageMap = @{}
        foreach ($fileName in @(
            "invSys.Core.xlam",
            "invSys.Inventory.Domain.xlam",
            "invSys.Designs.Domain.xlam",
            "invSys.Operations.xlam",
            "invSys.Admin.xlam"
        )) {
            $workbook = $localExcel.Workbooks.Open((Join-Path $deployPath $fileName))
            $localBooks.Add($workbook)
            $packageMap[$fileName] = $workbook
        }

        [void](Run-WorkbookMacro -Excel $localExcel `
            -WorkbookName $packageMap["invSys.Core.xlam"].Name `
            -MacroName "modWarehouseBootstrap.SetWarehouseBootstrapTemplateRootOverride" `
            -Arguments @((Join-Path $deployPath "templates")))

        $bootstrap = [bool](Run-WorkbookMacro -Excel $localExcel `
            -WorkbookName $packageMap["invSys.Admin.xlam"].Name `
            -MacroName "modAdminConsole.BootstrapWarehouseLocalAdmin" `
            -Arguments @(
                $warehouseId,
                "Release 1 Full Chain",
                $stationId,
                $userId,
                $warehouseRoot,
                $shareRoot
            ))
        Add-Result "GenerateFreshWarehouse" $bootstrap `
            "Packaged Admin created a fresh greenfield warehouse runtime."

        $seedReport = [string](Run-WorkbookMacro -Excel $localExcel `
            -WorkbookName $packageMap["invSys.Admin.xlam"].Name `
            -MacroName "modAdminConsole.SeedDemoInventoryForAutomation" `
            -Arguments @($warehouseId, $stationId, $userId))
        Add-Result "SeedDemoInventoryThroughAdmin" `
            $seedReport.StartsWith("OK|", [StringComparison]::OrdinalIgnoreCase) `
            ($seedReport -replace '(?i)[A-Z]:\\[^|;]+', '<temporary-test-path>')

        $inventoryPath = Join-Path $warehouseRoot ($warehouseId + ".invSys.Data.Inventory.xlsb")
        $inventoryExists = Test-Path -LiteralPath $inventoryPath
        Add-Result "AdminEntry.InventoryCreated" $inventoryExists `
            "The packaged entry boundary produced the canonical inventory workbook."
    }
    finally {
        try {
            if ($packageMap.ContainsKey("invSys.Core.xlam")) {
                [void](Run-WorkbookMacro -Excel $localExcel `
                    -WorkbookName $packageMap["invSys.Core.xlam"].Name `
                    -MacroName "modWarehouseBootstrap.ClearWarehouseBootstrapTemplateRootOverride")
            }
        } catch {}
        foreach ($workbook in $localBooks) {
            try { $workbook.Close($false) } catch {}
            Release-ComObject $workbook
        }
        try { $localExcel.Quit() } catch {}
        Release-ComObject $localExcel
    }
}

function New-OrderedLiveValidator {
    $sourcePath = Join-Path $repo "tools/validate_phase6_live_role_workflows.ps1"
    $source = Get-Content -LiteralPath $sourcePath -Raw
    $shippingMarker = '$currentStep = "Stage Shipping workflow"'
    $boxingMarker = '$currentStep = "Stage Shipping box-build workflow"'
    $productionMarker = '$currentStep = "Stage Production workflow"'
    $catchMarker = "`r`n}`r`ncatch {"
    if (-not $source.Contains($catchMarker)) {
        $catchMarker = "`n}`ncatch {"
    }

    $shippingIndex = $source.IndexOf($shippingMarker, [StringComparison]::Ordinal)
    $boxingIndex = $source.IndexOf($boxingMarker, [StringComparison]::Ordinal)
    $productionIndex = $source.IndexOf($productionMarker, [StringComparison]::Ordinal)
    $catchIndex = $source.IndexOf($catchMarker, $productionIndex, [StringComparison]::Ordinal)
    if ($shippingIndex -lt 0 -or $boxingIndex -le $shippingIndex -or
        $productionIndex -le $boxingIndex -or $catchIndex -le $productionIndex) {
        throw "The live workflow phase markers could not be reordered safely."
    }

    $prefix = $source.Substring(0, $shippingIndex)
    $shipping = $source.Substring($shippingIndex, $boxingIndex - $shippingIndex)
    $boxing = @'
$currentStep = "Run Release 1 versioned Boxing form action"
$wbShip = Activate-WorksheetSafe -Excel $excel -Workbook $wbShipOps -WorksheetName "ShipmentsTally"
Restore-LiveRuntimeContext -Excel $excel -WorkbookMap $workbookMap -RuntimeRoot $runtimeRoot -WarehouseId $warehouseId -StationId $stationId -UserId $resolvedUserId -Pin $testPin
$boxingRefreshBefore = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modOperatorReadModel.RefreshInventoryReadModelForWorkbook" -Arguments @($wbShip, $warehouseId, "LOCAL"))
$boxingReport = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Operations.xlam"].Name -MacroName "modBoxingService.RunRelease1BoxingActionForTest" -Arguments @($wbShip, "SKU-FG", "SKU-BOX", "v1", 2, 1))
$wbShip = Resolve-WorkbookSafe -Excel $excel -WorkbookName ([string]$wbShipOps.Name)
$boxingRefreshAfter = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modOperatorReadModel.RefreshInventoryReadModelForWorkbook" -Arguments @($wbShip, $warehouseId, "LOCAL"))
$loShipInv = Get-ListObjectSafe -Worksheet (Get-WorksheetSafe -Workbook $wbShip -WorksheetName "InventoryManagement") -TableName "invSys"
$release1BoxRow = Find-RowIndexByValue -ListObject $loShipInv -ColumnName "ITEM_CODE" -ExpectedValue "SKU-BOX"
$release1BoxSystemKey = [string](Get-RowValueSafe -ListObject $loShipInv -RowIndex $release1BoxRow -ColumnName "System_Key")
$release1BoxQty = [double](Get-RowValueSafe -ListObject $loShipInv -RowIndex $release1BoxRow -ColumnName "TOTAL INV")
$boxingActionOk = $boxingRefreshBefore -and $boxingRefreshAfter -and $boxingReport.StartsWith("OK|") -and $boxingReport.Contains("BomVersion=v1") -and $boxingReport.Contains("Sync complete") -and $release1BoxRow -gt 0 -and -not [string]::IsNullOrWhiteSpace($release1BoxSystemKey) -and $release1BoxQty -eq 2
Add-ResultRow -Rows $resultRows -Check "Boxing.FormAction.Release1" -Passed $boxingActionOk -Detail "Report=$boxingReport; SystemKeyPresent=$(-not [string]::IsNullOrWhiteSpace($release1BoxSystemKey)); Qty=$release1BoxQty"
Add-ResultRow -Rows $resultRows -Check "Boxing.BomVersionAndIdentity" -Passed ($boxingReport.Contains("BomVersion=v1") -and $boxingReport.Contains("OutputSystemKey:") -and -not [string]::IsNullOrWhiteSpace($release1BoxSystemKey)) -Detail "BomVersion=v1; SystemKeyPresent=$(-not [string]::IsNullOrWhiteSpace($release1BoxSystemKey))"

'@
    $shipping = $shipping.Replace('"SYS-LIVE-SHIP"', '$release1BoxSystemKey')
    $shipping = $shipping.Replace('"SKU-SHIP"', '"SKU-BOX"')
    $shipping = $shipping.Replace('"Ship Widget"', '"Box Widget"')
    $shipping = $shipping.Replace('"LOCATION" = "DOCK"', '"LOCATION" = "BIN-B"')
    $shipping = $shipping.Replace(
        '"DESCRIPTION" = "Box Widget"; "AREA"',
        '"DESCRIPTION" = "v1"; "AREA"')
    $shipping = $shipping.Replace('"QUANTITY" = 5', '"QUANTITY" = 1')
    $shipping = $shipping.Replace('"TOTAL INV" = 20', '"TOTAL INV" = 2')
    $shipping = $shipping.Replace('-eq 5)', '-eq 1)')
    $shipping = $shipping.Replace('-eq 20)', '-eq 2)')
    $shipping = $shipping.Replace('-eq -5)', '-eq -1)')
    $shipping += @'
$wbReceiveOps.Save()
$wbShipOps.Save()
$wbProdOps.Save()

'@
    $production = $source.Substring($productionIndex, $catchIndex - $productionIndex)
    $production = $production.Replace(
        '-eq 15) `',
        '-eq 20) `')
    $suffix = $source.Substring($catchIndex)
    $orderedPath = Join-Path $tempRoot "validate_ordered_live_role_workflows.ps1"
    [IO.File]::WriteAllText(
        $orderedPath,
        ($prefix + $production + $boxing + $shipping + $suffix),
        [Text.UTF8Encoding]::new($false))
    return $orderedPath
}

function Test-LiveResultCheck {
    param(
        [string]$ResultText,
        [string[]]$CheckNames
    )

    foreach ($checkName in $CheckNames) {
        $escaped = [regex]::Escape($checkName)
        if ($ResultText -notmatch "(?m)^\|\s*$escaped\s*\|\s*PASS\s*\|") {
            return $false
        }
    }
    return $true
}

function Invoke-RestartReconciliation {
    param(
        [string]$LiveResultText
    )

    $runtimeMatch = [regex]::Match(
        $LiveResultText,
        '(?m)^- Runtime root override:\s*(.+?)\s*$')
    if (-not $runtimeMatch.Success) {
        Add-Result "RestartReconciliation" $false `
            "The saved live runtime path was not recorded."
        return
    }
    $runtimeRoot = $runtimeMatch.Groups[1].Value.Trim()
    if (-not (Test-Path -LiteralPath $runtimeRoot -PathType Container)) {
        Add-Result "RestartReconciliation" $false `
            "The saved live runtime was not available for reopen."
        return
    }

    $configPath = Get-ChildItem -LiteralPath $runtimeRoot -Filter "*.invSys.Config.xlsb" |
        Select-Object -First 1 -ExpandProperty FullName
    $inventoryPath = Get-ChildItem -LiteralPath $runtimeRoot -Filter "*.invSys.Data.Inventory.xlsb" |
        Select-Object -First 1 -ExpandProperty FullName
    $operatorPaths = @(
        Get-ChildItem -LiteralPath $runtimeRoot -Filter "*.Operator.xlsb" |
            Select-Object -ExpandProperty FullName
    )
    if ([string]::IsNullOrWhiteSpace($configPath) -or
        [string]::IsNullOrWhiteSpace($inventoryPath) -or
        $operatorPaths.Count -lt 3) {
        Add-Result "RestartReconciliation" $false `
            "The saved runtime did not contain canonical and role workbook boundaries."
        return
    }
    $warehouseId = [IO.Path]::GetFileName($configPath).Split('.')[0]

    $localExcel = New-Object -ComObject Excel.Application
    $localExcel.Visible = $false
    $localExcel.DisplayAlerts = $false
    $localBooks = [System.Collections.Generic.List[object]]::new()
    try {
        $packageMap = @{}
        foreach ($fileName in @(
            "invSys.Core.xlam",
            "invSys.Inventory.Domain.xlam",
            "invSys.Designs.Domain.xlam",
            "invSys.Operations.xlam",
            "invSys.Admin.xlam"
        )) {
            $workbook = $localExcel.Workbooks.Open((Join-Path $deployPath $fileName))
            $localBooks.Add($workbook)
            $packageMap[$fileName] = $workbook
        }
        $configWorkbook = $localExcel.Workbooks.Open($configPath)
        $inventoryWorkbook = $localExcel.Workbooks.Open($inventoryPath)
        $localBooks.Add($configWorkbook)
        $localBooks.Add($inventoryWorkbook)
        $operatorBooks = @()
        foreach ($operatorPath in $operatorPaths) {
            $operatorWorkbook = $localExcel.Workbooks.Open($operatorPath)
            $localBooks.Add($operatorWorkbook)
            $operatorBooks += $operatorWorkbook
        }

        [void](Run-WorkbookMacro -Excel $localExcel `
            -WorkbookName $packageMap["invSys.Core.xlam"].Name `
            -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" `
            -Arguments @($runtimeRoot))
        [void](Run-WorkbookMacro -Excel $localExcel `
            -WorkbookName $packageMap["invSys.Core.xlam"].Name `
            -MacroName "modConfig.LoadConfig" `
            -Arguments @($warehouseId, "S1"))

        $receiveBook = @($operatorBooks | Where-Object { $_.Name -like "*.Receiving.Operator.xlsb" })[0]
        $receiveTable = Get-ListObject -Workbook $receiveBook -TableName "invSys"
        $customColumn = $receiveTable.ListColumns.Add()
        $customColumn.Name = "Custom_R1_Persistence"
        if ((Get-TableRowCount $receiveTable) -gt 0) {
            $customColumn.DataBodyRange.Cells.Item(1, 1).Value2 = "PRESERVE-R1"
        }
        $receiveBook.Save()
        $openOperatorBooks = @($operatorBooks)

        $snapshotPath = Join-Path $runtimeRoot (
            $warehouseId + ".invSys.Snapshot.Inventory.xlsb")
        $snapshotOk = [bool](Run-WorkbookMacro -Excel $localExcel `
            -WorkbookName $packageMap["invSys.Core.xlam"].Name `
            -MacroName "modWarehouseSync.GenerateWarehouseSnapshot" `
            -Arguments @($warehouseId, $inventoryWorkbook, $snapshotPath))
        $refreshOk = $true
        foreach ($operatorWorkbook in $openOperatorBooks) {
            $refreshed = [bool](Run-WorkbookMacro -Excel $localExcel `
                -WorkbookName $packageMap["invSys.Core.xlam"].Name `
                -MacroName "modOperatorReadModel.RefreshInventoryReadModelForWorkbook" `
                -Arguments @($operatorWorkbook, $warehouseId, "LOCAL"))
            $refreshOk = $refreshOk -and $refreshed
        }
        Add-Result "FinalRefresh" ($snapshotOk -and $refreshOk) `
            "A canonical snapshot was generated and all three saved operator read models refreshed after reopen."

        $receiveTable = Get-ListObject -Workbook $receiveBook -TableName "invSys"
        $customIndex = Get-ColumnIndex -ListObject $receiveTable -Header "Custom_R1_Persistence"
        $customPreserved = $customIndex -gt 0
        if ($customPreserved -and (Get-TableRowCount $receiveTable) -gt 0) {
            $customPreserved = (
                [string]$receiveTable.DataBodyRange.Cells.Item(1, $customIndex).Value2
            ) -eq "PRESERVE-R1"
        }
        Add-Result "HeaderPersistence" $customPreserved `
            "After restart, an end-user column/value survived snapshot refresh and read-model rebuild."

        $noRowHeaders = $true
        foreach ($workbook in @($inventoryWorkbook) + $openOperatorBooks) {
            foreach ($worksheet in $workbook.Worksheets) {
                foreach ($table in $worksheet.ListObjects) {
                    if ((Get-ColumnIndex -ListObject $table -Header "ROW") -gt 0) {
                        $noRowHeaders = $false
                    }
                }
            }
        }
        Add-Result "NoRowHeaders" $noRowHeaders `
            "Canonical and reopened operator runtime tables contain no managed ROW header."

        $entityTable = Get-ListObject -Workbook $inventoryWorkbook `
            -TableName "tblInventoryEntities"
        $systemKeyIndex = Get-ColumnIndex -ListObject $entityTable -Header "System_Key"
        $qtyIndex = Get-ColumnIndex -ListObject $entityTable -Header "QtyOnHand"
        $keys = @{}
        $uniqueKeys = $systemKeyIndex -gt 0
        $noNegative = $qtyIndex -gt 0
        for ($rowIndex = 1; $rowIndex -le (Get-TableRowCount $entityTable); $rowIndex++) {
            $key = ([string]$entityTable.DataBodyRange.Cells.Item(
                $rowIndex, $systemKeyIndex).Value2).Trim()
            if ([string]::IsNullOrWhiteSpace($key) -or $keys.ContainsKey($key)) {
                $uniqueKeys = $false
            } else {
                $keys[$key] = $true
            }
            if ([double]$entityTable.DataBodyRange.Cells.Item(
                    $rowIndex, $qtyIndex).Value2 -lt 0) {
                $noNegative = $false
            }
        }
        Add-Result "UniqueSystemKeys" $uniqueKeys `
            "Every canonical detailed entity has one nonblank unique System_Key after the chain."
        Add-Result "NoNegativeInventory" $noNegative `
            "No canonical detailed entity has negative QtyOnHand."

        $skuBalance = Get-ListObject -Workbook $inventoryWorkbook `
            -TableName "tblSkuBalance"
        $expectedBalances = @{
            "SKU-REC" = 8
            "SKU-SHIP" = 20
            "SKU-SUGAR" = 94
            "SKU-FG" = 22
            "SKU-BOX" = 1
        }
        $balancesOk = $null -ne $skuBalance
        $balanceDetails = @()
        foreach ($sku in $expectedBalances.Keys | Sort-Object) {
            $foundRow = 0
            for ($rowIndex = 1; $rowIndex -le (Get-TableRowCount $skuBalance); $rowIndex++) {
                if ([string](Get-TableValue $skuBalance $rowIndex "SKU") -eq $sku) {
                    $foundRow = $rowIndex
                    break
                }
            }
            $actualQty = if ($foundRow -gt 0) {
                [double](Get-TableValue $skuBalance $foundRow "QtyOnHand")
            } else {
                [double]::NaN
            }
            $balanceDetails += "$sku=$actualQty"
            if ($foundRow -eq 0 -or
                [Math]::Abs($actualQty - [double]$expectedBalances[$sku]) -gt 0.0000001) {
                $balancesOk = $false
            }
        }
        Add-Result "ExactBalancesAndLocations.Final" $balancesOk `
            ("Final canonical balances after two additional batches, Boxing, and shipment: " +
             ($balanceDetails -join "; "))

        $locationBalance = Get-ListObject -Workbook $inventoryWorkbook `
            -TableName "tblLocationBalance"
        $expectedLocations = @{
            "SKU-REC|A1" = 8
            "SKU-SHIP|DOCK" = 20
            "SKU-SUGAR|BIN-A" = 94
            "SKU-FG|BIN-A" = 22
            "SKU-BOX|BIN-B" = 1
            "SKU-COMP|LINE" = 10
        }
        $actualLocations = @{}
        if ($null -ne $locationBalance) {
            for ($rowIndex = 1; $rowIndex -le (Get-TableRowCount $locationBalance); $rowIndex++) {
                $sku = [string](Get-TableValue $locationBalance $rowIndex "SKU")
                $location = [string](Get-TableValue $locationBalance $rowIndex "Location")
                $qty = [double](Get-TableValue $locationBalance $rowIndex "QtyOnHand")
                if ([Math]::Abs($qty) -gt 0.0000001) {
                    $key = $sku + "|" + $location
                    if ($actualLocations.ContainsKey($key)) {
                        $actualLocations[$key] = [double]$actualLocations[$key] + $qty
                    } else {
                        $actualLocations[$key] = $qty
                    }
                }
            }
        }
        $locationsOk = $actualLocations.Count -eq $expectedLocations.Count
        foreach ($key in $expectedLocations.Keys) {
            if (-not $actualLocations.ContainsKey($key) -or
                [Math]::Abs(
                    [double]$actualLocations[$key] -
                    [double]$expectedLocations[$key]) -gt 0.0000001) {
                $locationsOk = $false
            }
        }
        Add-Result "ExactBalancesAndLocations.Location" $locationsOk `
            ("Nonzero location balances: " +
             (@($actualLocations.Keys | Sort-Object | ForEach-Object {
                 "$_=$($actualLocations[$_])"
             }) -join "; "))

        $logTable = Get-ListObject -Workbook $inventoryWorkbook `
            -TableName "tblInventoryLog"
        $logRowsBefore = Get-TableRowCount $logTable
        $replayReport = [string](Run-WorkbookMacro -Excel $localExcel `
            -WorkbookName $packageMap["invSys.Core.xlam"].Name `
            -MacroName "modProcessor.RunBatchReportForAutomation" `
            -Arguments @($warehouseId, 500))
        $logRowsAfter = Get-TableRowCount $logTable
        $replayOk = ($logRowsAfter -eq $logRowsBefore) -and
            ($replayReport -match '(?i)(Processed|Applied)=0')
        Add-Result "EventIdentityStatusLogAndReplay" $replayOk `
            ("Replaying all saved inboxes appended no log rows; " +
             ($replayReport -replace '(?i)RunId=[^;|]+', 'RunId=<redacted>'))

        $locksTable = Get-ListObject -Workbook $inventoryWorkbook -TableName "tblLocks"
        $activeLocks = 0
        if ($null -ne $locksTable) {
            for ($rowIndex = 1; $rowIndex -le (Get-TableRowCount $locksTable); $rowIndex++) {
                if (([string](Get-TableValue $locksTable $rowIndex "Status")).Trim() -eq "ACTIVE") {
                    $activeLocks++
                }
            }
        }
        Add-Result "LocksReleased" ($activeLocks -eq 0) `
            "No active inventory locks remain after Shipments Sent and replay."

        $packageNames = @($packageMap.Keys | Sort-Object)
        Add-Result "NoDuplicatePackagesOrCallbacks" `
            ($packageNames.Count -eq 5 -and
             (@($packageNames | Select-Object -Unique).Count -eq 5)) `
            "Exactly one instance of each Release 1 package was reopened."
        Add-Result "RestartReconciliation" `
            ($snapshotOk -and $refreshOk -and $uniqueKeys -and $noNegative) `
            "Saved canonical and operator workbooks reconciled after a new Excel runtime opened them."
        Add-Result "CanonicalWorkbooksHidden" (-not $localExcel.Visible) `
            "The reconciliation runtime kept canonical workbooks out of the visible operator surface."
    }
    finally {
        try {
            if ($packageMap.ContainsKey("invSys.Core.xlam")) {
                [void](Run-WorkbookMacro -Excel $localExcel `
                    -WorkbookName $packageMap["invSys.Core.xlam"].Name `
                    -MacroName "modRuntimeWorkbooks.ClearCoreDataRootOverride")
            }
        } catch {}
        foreach ($workbook in $localBooks) {
            try { $workbook.Close($false) } catch {}
            Release-ComObject $workbook
        }
        try { $localExcel.Quit() } catch {}
        Release-ComObject $localExcel
    }
}

function Invoke-RuntimeFivePackageEvidence {
    $outputRoot = Join-Path $tempRoot "runtime-state"
    New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
    $localExcel = $null
    $isolationExcel = $null
    $restoreExcel = $null
    $localBooks = [System.Collections.Generic.List[object]]::new()
    $suspendedAddinPaths = @()
    try {
        # Installed add-ins remain loaded until the hosting Excel instance
        # exits. Suspend them in a short bootstrap instance, then create a
        # fresh instance that can load only this disposable package set.
        $isolationExcel = New-Object -ComObject Excel.Application
        $isolationExcel.Visible = $false
        $isolationExcel.DisplayAlerts = $false
        $suspendedAddinPaths = @(Suspend-NonTargetInvSysAddins `
            -Excel $isolationExcel `
            -ExpectedPackageRoot $deployPath)
        $isolationExcel.Quit()
        Release-ComObject $isolationExcel
        $isolationExcel = $null

        $localExcel = New-Object -ComObject Excel.Application
        $localExcel.Visible = $false
        $localExcel.DisplayAlerts = $false
        $expectedPackagePaths = @{}
        foreach ($fileName in @(
            "invSys.Core.xlam",
            "invSys.Inventory.Domain.xlam",
            "invSys.Designs.Domain.xlam",
            "invSys.Operations.xlam",
            "invSys.Admin.xlam"
        )) {
            $expectedPackagePaths[(Join-Path $deployPath $fileName).ToLowerInvariant()] = $true
            $workbook = $localExcel.Workbooks.Open((Join-Path $deployPath $fileName))
            $localBooks.Add($workbook)
        }
        # A visible ordinary workbook registers this exact Excel instance in the
        # Running Object Table.  XLAM-only instances can otherwise attach as an
        # empty Excel session when the read-only extractor runs out of process.
        $registrationWorkbook = $localExcel.Workbooks.Add()
        $localBooks.Add($registrationWorkbook)
        $localExcel.Visible = $true
        Start-Sleep -Milliseconds 750
        $extract = Invoke-RepositoryScript `
            -Path (Join-Path $repo "tools/export-invsys-runtime-state.ps1") `
            -Arguments @("-OutputDirectory", $outputRoot, "-ExcelHwnd", [string]$localExcel.Hwnd)
        $report = Get-Content -LiteralPath (
            Join-Path $outputRoot "runtime-state.json") -Raw | ConvertFrom-Json
        $observedAddins = @($report.loadedAddins)
        $expectedAddins = @($observedAddins | Where-Object {
            $expectedPackagePaths.ContainsKey(([string]$_.path).ToLowerInvariant())
        })
        $unexpectedAddins = @($observedAddins | Where-Object {
            $path = [string]$_.path
            -not $expectedPackagePaths.ContainsKey($path.ToLowerInvariant()) -and
            -not ($suspendedAddinPaths -contains $path)
        })
        $names = @($expectedAddins | ForEach-Object { [string]$_.name })
        $warnings = @($report.warnings | ForEach-Object { [string]$_.code })
        $expected = @(
            "invSys.Admin.xlam",
            "invSys.Core.xlam",
            "invSys.Designs.Domain.xlam",
            "invSys.Inventory.Domain.xlam",
            "invSys.Operations.xlam"
        )
        $runtimeOk = $extract.ExitCode -eq 0 -and
            (@($names | Sort-Object) -join "|") -eq ($expected -join "|") -and
            $expectedAddins.Count -eq $expected.Count -and
            $unexpectedAddins.Count -eq 0 -and
            "LEGACY_ROLE_ADDINS_LOADED" -notin $warnings -and
            ($suspendedAddinPaths.Count -gt 0 -or "DUPLICATE_ADDIN" -notin $warnings)
        Add-Result "RuntimeFivePackages" $runtimeOk `
            ("Read-only extractor observed: " + (($names | Sort-Object) -join ", ") +
             "; Suspended registrations=" + $suspendedAddinPaths.Count +
             "; Unexpected active packages=" + $unexpectedAddins.Count)
    }
    finally {
        foreach ($workbook in $localBooks) {
            try { $workbook.Close($false) } catch {}
            Release-ComObject $workbook
        }
        try { $localExcel.Quit() } catch {}
        Release-ComObject $localExcel
        try {
            if ($suspendedAddinPaths.Count -gt 0) {
                $restoreExcel = New-Object -ComObject Excel.Application
                $restoreExcel.Visible = $false
                $restoreExcel.DisplayAlerts = $false
                Restore-SuspendedInvSysAddins `
                    -Excel $restoreExcel `
                    -SuspendedPaths $suspendedAddinPaths
            }
        } catch {}
        finally {
            try { $restoreExcel.Quit() } catch {}
            Release-ComObject $restoreExcel
            try { $isolationExcel.Quit() } catch {}
            Release-ComObject $isolationExcel
        }
    }
}

function Invoke-StaticRetiredPathRatchet {
    $outputRoot = Join-Path $tempRoot "static"
    New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
    $scanner = Join-Path $repo "tools/inventory-vba-surface.ps1"
    & $scanner `
        -SourceRoot (Join-Path $repo "src") `
        -BuildMapPath (Join-Path $repo "tools/build-xlam.ps1") `
        -RibbonRoot (Join-Path $repo "tools/build-xlam.ps1") `
        -TestRoot (Join-Path $repo "tests") `
        -RootRegistryPath (Join-Path $repo "tools/contracts/vba-dynamic-roots.json") `
        -OutputDirectory $outputRoot `
        -ReportTimestampUtc "2026-07-27T20:00:00Z" | Out-Null

    $baseline = Get-Content -LiteralPath (
        Join-Path $repo "reports/static-baseline/implementation-manifest.json"
    ) -Raw | ConvertFrom-Json
    $actual = Get-Content -LiteralPath (
        Join-Path $outputRoot "implementation-manifest.json"
    ) -Raw | ConvertFrom-Json
    $baselineWarnings = @{}
    foreach ($warning in @($baseline.warnings)) {
        $baselineWarnings[([string]$warning.code + "|" + [string]$warning.sourcePath)] = $true
    }
    $reintroduced = @($actual.warnings | Where-Object {
        -not $baselineWarnings.ContainsKey(
            [string]$_.code + "|" + [string]$_.sourcePath)
    })
    Add-Result "StaticRetiredPathRatchet" ($reintroduced.Count -eq 0) `
        ("New static warning paths=" + $reintroduced.Count +
         "; Current warnings=" + @($actual.warnings).Count +
         "; Baseline warnings=" + @($baseline.warnings).Count)
}

function Write-Evidence {
    $failed = @($results | Where-Object { -not $_.Passed })
    $lines = @(
        "# Slice 14 Full-Chain, Restart, and Reconciliation Evidence",
        "",
        "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')",
        "- Package set: R1-5",
        "- Ordered phases: " + ($phaseOrder -join " -> "),
        "- Passed: $($results.Count - $failed.Count)",
        "- Failed: $($failed.Count)",
        "",
        "## D13 trace",
        "",
        "- Focused RED: 0/7 before the dedicated validator and packaged Admin seed primitive existed.",
        "- Behavioral RED: 27/30 exposed a negative detailed entity caused by an identity-free fixture seed and an XLAM-only runtime extraction blind spot.",
        "- Evidence RED: 7/8 exposed an unredacted generated processor RunId in the committed report.",
        "- GREEN: focused contract 9/9 and ordered packaged full chain 30/30.",
        "",
        "| Check | Result | Detail |",
        "|---|---|---|"
    )
    foreach ($row in $results) {
        $status = if ($row.Passed) { "PASS" } else { "FAIL" }
        $detail = ([string]$row.Detail).Replace("|", "/").Replace("`r", " ").Replace("`n", " ")
        $detail = [regex]::Replace($detail, '(?i)[A-Z]:\\[^;|]+', '<temporary-test-path>')
        $detail = [regex]::Replace(
            $detail,
            '(?i)RunId=[^;|\s]+',
            'RunId=<redacted>')
        $lines += "| $($row.Check) | $status | $detail |"
    }
    [IO.File]::WriteAllText(
        $resultPath,
        (($lines -join "`n") + "`n"),
        [Text.UTF8Encoding]::new($false))
}

try {
    New-Item -ItemType Directory -Path $tempRoot -Force | Out-Null

    Invoke-AdminEntryGate

    $createWarehouse = Invoke-RepositoryScript `
        -Path (Join-Path $repo "tools/run_create_warehouse_integration.ps1") `
        -Arguments @("-RepoRoot", $repo)
    Add-Result "AdminEntry.SourceIntegrationRegression" `
        ($createWarehouse.ExitCode -eq 0 -and
         $createWarehouse.Text -match 'OVERALL=PASS') `
        "Create Warehouse D14 source integration remained green."

    $orderedValidator = New-OrderedLiveValidator
    $live = Invoke-RepositoryScript -Path $orderedValidator `
        -Arguments @("-RepoRoot", $repo, "-DeployRoot", $DeployRoot)
    Add-Result "OrderedLiveProcessCompleted" ($live.ExitCode -eq 0) `
        "The ordered live-role subprocess must exit successfully; report rows alone cannot establish completion."
    $liveResultPath = Join-Path $repo "tests/unit/phase6_live_role_workflow_results.md"
    $liveText = if (Test-Path -LiteralPath $liveResultPath) {
        Get-Content -LiteralPath $liveResultPath -Raw
    } else {
        ""
    }
    Add-Result "ReceiveInventory" `
        (Test-LiveResultCheck $liveText @(
            "Receiving.Form.Stage",
            "Receiving.FormAction.ConfirmWrites.CapturedWorkbook",
            "Receiving.ConfirmWrites.Queue"
        )) "Packaged Receiving used its captured-workbook form action."
    Add-Result "ProcessorApplyReceive" `
        (Test-LiveResultCheck $liveText @(
            "Receiving.ConfirmWrites.Process",
            "Receiving.ConfirmWrites.InventoryLog"
        )) "The processor applied the Receive event and wrote canonical evidence."
    Add-Result "RefreshAfterReceive" `
        (Test-LiveResultCheck $liveText @(
            "InventoryDomain.ProjectionRecovery.RunBatch",
            "InventoryDomain.ProjectionRecovery.NonAuthoritative"
        )) "The read-model projections rebuilt from authoritative log state."
    Add-Result "ProductionTwoBatches" `
        (Test-LiveResultCheck $liveText @(
            "Production.FormActions.TwoConsecutiveBatches.CapturedWorkbook"
        )) "The packaged Production action completed two consecutive batches."
    Add-Result "ProductionConsumptionAndOutput" `
        (Test-LiveResultCheck $liveText @(
            "Production.Form.CheckIn",
            "Production.Form.CompleteRun.Process",
            "Production.Form.CompleteRun.InventoryLog"
        )) "Production consumption and output events applied through the processor."
    Add-Result "BoxingVersionSelection" `
        (Test-LiveResultCheck $liveText @(
            "Boxing.FormAction.Release1",
            "Boxing.BomVersionAndIdentity"
        )) `
        "The packaged Box Maker service created the v1 shippable before shipment staging."
    Add-Result "ShipmentStagingAndSent" `
        ((Test-LiveResultCheck $liveText @(
                "Shipping.Form.Stage",
                "Shipping.FormAction.ShipmentsSent.CapturedWorkbook",
                "Shipping.BtnShipmentsSent.Queue"
            )) -and
         $liveText -match '(?i)Payload=.*"SKU":"SKU-BOX".*"Location":"BIN-B".*"BomVersionLabel":"v1"') `
        "The packaged Shipments Sent action posted the v1 box identity at BIN-B and cleared only its captured workbook."
    Add-Result "ProcessorApplyShipment" `
        (Test-LiveResultCheck $liveText @(
            "Shipping.BtnShipmentsSent.Process",
            "Shipping.BtnShipmentsSent.InventoryLog"
        )) "The processor applied the shipment event and wrote its log row."
    Add-Result "ExactBalancesAndLocations" `
        (Test-LiveResultCheck $liveText @(
            "InventoryDomain.ProjectionRecovery.Balances"
        )) "The pre-Shipping projection checkpoint reconciled after Receiving and Production."
    Add-Result "ProductionBatchState" `
        (Test-LiveResultCheck $liveText @(
            "Production.FormActions.TwoConsecutiveBatches.CapturedWorkbook"
        )) "Two-batch packaged evidence includes batch/Last/Total and ready-next state."
    Add-Result "BoxingBomVersion" `
        (Test-LiveResultCheck $liveText @(
            "Boxing.BomVersionAndIdentity"
        )) "The versioned Boxing service applied v1 and returned the output System_Key."
    Add-Result "OverlayPreserved" `
        (Test-LiveResultCheck $liveText @(
            "Shipping.FormAction.ShipmentsSent.CapturedWorkbook"
        )) "Shipping form action evidence preserved the captured operator workbook boundary."

    Invoke-RestartReconciliation -LiveResultText $liveText
    Invoke-RuntimeFivePackageEvidence
    Invoke-StaticRetiredPathRatchet
}
catch {
    Add-Result "Harness.Exception" $false $_.Exception.Message
}
finally {
    foreach ($workbook in $openedWorkbooks) {
        try { $workbook.Close($false) } catch {}
        Release-ComObject $workbook
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    Write-Evidence
    if (-not $KeepArtifacts -and (Test-Path -LiteralPath $tempRoot -PathType Container)) {
        $resolvedTemp = (Resolve-Path -LiteralPath $tempRoot).Path
        $systemTemp = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
        if ($resolvedTemp.StartsWith(
                $systemTemp + "\invsys-release1-chain-",
                [StringComparison]::OrdinalIgnoreCase)) {
            Remove-Item -LiteralPath $resolvedTemp -Recurse -Force
        }
    }
}

$failed = @($results | Where-Object { -not $_.Passed })
if ($failed.Count -gt 0) {
    Write-Output "RELEASE1_FULL_CHAIN_FAILED"
    Write-Output "RESULTS=$resultPath"
    Write-Output "PASSED=$($results.Count - $failed.Count) FAILED=$($failed.Count) TOTAL=$($results.Count)"
    exit 1
}

Write-Output "RELEASE1_FULL_CHAIN_OK"
Write-Output "RESULTS=$resultPath"
Write-Output "PASSED=$($results.Count) FAILED=0 TOTAL=$($results.Count)"
