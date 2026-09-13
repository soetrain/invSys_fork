[CmdletBinding()]
param(
    [string]$RepoRoot = ".",
    [string]$DeployRoot = "deploy/current",
    [string]$OutputDirectory = "reports/runtime/plan022-slice0",
    [ValidateSet("", "Receiving", "Production", "Shipping")]
    [string]$CallbackFilter = "",
    [ValidateSet("NoEligible", "ConfigActive", "SavedEligible", "UnrelatedActive", "CapturedClosed", "ReceivingDurability", "ReceivingFormClosed", "ShippingLayout", "ProductionReusable")]
    [string]$WorkbookState = "NoEligible",
    [switch]$ProductionRunOnly,
    [switch]$ProductionEditExportOnly
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Import-LiveValidationHelpers {
    param([string]$ScriptPath)

    $tokens = $null
    $errors = $null
    $ast = [System.Management.Automation.Language.Parser]::ParseFile(
        $ScriptPath,
        [ref]$tokens,
        [ref]$errors
    )
    if ($errors.Count -gt 0) {
        throw "Unable to parse validation helper source: $($errors[0].Message)"
    }

    $definitions = $ast.FindAll({
        param($node)
        $node -is [System.Management.Automation.Language.FunctionDefinitionAst]
    }, $true)
    foreach ($definition in $definitions) {
        $scriptScopedDefinition = $definition.Extent.Text -replace (
            '^(?i)function\s+' + [regex]::Escape($definition.Name)
        ), ('function script:' + $definition.Name)
        . ([scriptblock]::Create($scriptScopedDefinition))
    }

    $required = @(
        "Release-ComObject",
        "Get-InvSysCredentialHash",
        "Run-WorkbookMacro",
        "New-ConfigWorkbook",
        "New-AuthWorkbook",
        "New-InventoryWorkbook",
        "New-OperationalWorkbook",
        "Get-WorksheetSafe",
        "Get-ListObjectSafe",
        "Get-ColumnIndexSafe",
        "Get-RowCountSafe"
    )
    $missing = @($required | Where-Object { -not (Get-Command $_ -CommandType Function -ErrorAction SilentlyContinue) })
    if ($missing.Count -gt 0) {
        throw "Missing imported validation helpers: $($missing -join ', ')"
    }
}

function Get-ExcelProcessId {
    param([object]$Excel)

    [uint32]$processId = 0
    [void][Plan022LauncherWindow]::GetWindowThreadProcessId(
        [intptr]$Excel.Hwnd,
        [ref]$processId
    )
    return [int]$processId
}

function Get-OpenWorkbookNames {
    param([object]$Excel)

    $names = New-Object System.Collections.Generic.List[string]
    $count = [int]$Excel.Workbooks.Count
    for ($i = 1; $i -le $count; $i++) {
        try {
            $name = [string]$Excel.Workbooks.Item($i).Name
            if (-not [string]::IsNullOrWhiteSpace($name)) {
                $names.Add($name) | Out-Null
            }
        }
        catch {}
    }
    return @($names)
}

function Start-DialogCaptureAndDismiss {
    param(
        [int]$ExcelProcessId,
        [int]$TimeoutSeconds = 20,
        [string]$StopPath = ""
    )
    $observerPath = Join-Path $repo 'tools/plan022-dialog-observer.ps1'
    return Start-Job -ScriptBlock {
        param($processId, $timeoutSeconds, $stopPath, $observerPath)
        . $observerPath
        Invoke-Plan022NativeDialogObservation -ProcessId $processId -TimeoutSeconds $timeoutSeconds -StopPath $stopPath
    } -ArgumentList $ExcelProcessId, $TimeoutSeconds, $StopPath, $observerPath
}

function Invoke-PackagedCallback {
    param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$MacroName
    )

    $stopPath = Join-Path $runtimeRoot ('dialog-stop-'+[guid]::NewGuid().ToString('N'))
    $job = Start-DialogCaptureAndDismiss -ExcelProcessId (Get-ExcelProcessId $Excel) -StopPath $stopPath
    $errorText = ""
    try {
        [void](Run-WorkbookMacro -Excel $Excel -WorkbookName $WorkbookName -MacroName $MacroName)
    }
    catch {
        $errorText = $_.Exception.Message
    }
    finally {
        Start-Sleep -Milliseconds 600
        [IO.File]::WriteAllText($stopPath, '')
        Wait-Job -Job $job -Timeout 25 -ErrorAction SilentlyContinue | Out-Null
        if ($job.State -notin @('Completed','Failed','Stopped')) {
            throw 'Dialog observer did not finish; callback evidence is incomplete.'
        }
        Remove-Item -LiteralPath $stopPath -ErrorAction SilentlyContinue
    }

    $observerCompleted = $job.State -eq 'Completed'
    if (-not $observerCompleted -and $errorText -eq '') { $errorText = 'Dialog observer did not complete normally.' }
    $captured = @(
        Receive-Job -Job $job -ErrorAction SilentlyContinue |
            ForEach-Object { [string]$_ } |
            Sort-Object -Unique
    )
    Remove-Job -Job $job -ErrorAction SilentlyContinue
    return [pscustomobject]@{
        Macro = $MacroName
        Error = $errorText
        WindowText = $captured
        ObserverCompleted = $observerCompleted
    }
}

function Add-Evidence {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Callback,
        [string]$Expected,
        [bool]$Passed,
        [string]$Observed
    )

    $Rows.Add([pscustomobject]@{
        Callback = $Callback
        Expected = $Expected
        Passed = $Passed
        Observed = $Observed
    }) | Out-Null
}

function Get-WorkbookSurfaceCounts {
    param([object]$Workbook)

    if ($null -eq $Workbook) {
        return [pscustomobject]@{ Worksheets = 0; Tables = 0 }
    }
    $tableCount = 0
    foreach ($worksheet in $Workbook.Worksheets) {
        $tableCount += [int]$worksheet.ListObjects.Count
    }
    return [pscustomobject]@{
        Worksheets = [int]$Workbook.Worksheets.Count
        Tables = $tableCount
    }
}

function Get-FileHashMap {
    param([string[]]$Paths)

    $result = @{}
    foreach ($path in $Paths) {
        if (-not (Test-Path -LiteralPath $path -PathType Leaf)) {
            $result[$path] = "<missing>"
            continue
        }
        $result[$path] = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    }
    return $result
}

function Compare-FileHashMaps {
    param(
        [hashtable]$Before,
        [hashtable]$After
    )

    foreach ($path in $Before.Keys) {
        if (-not $After.ContainsKey($path) -or $Before[$path] -ne $After[$path]) {
            return $false
        }
    }
    return $true
}

function Get-ChangedFileNames {
    param(
        [hashtable]$Before,
        [hashtable]$After
    )

    $changed = @()
    foreach ($path in $Before.Keys) {
        if (-not $After.ContainsKey($path) -or $Before[$path] -ne $After[$path]) {
            $changed += [IO.Path]::GetFileName($path)
        }
    }
    return @($changed)
}

function Get-RoleFormWindowCount {
    param([int]$ExcelProcessId)

    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $processCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
        $ExcelProcessId
    )
    $classCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ClassNameProperty,
        "ThunderDFrame"
    )
    $condition = New-Object System.Windows.Automation.AndCondition(
        $processCondition,
        $classCondition
    )
    $windows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        $condition
    )
    return [int]$windows.Count
}

function Close-RoleFormWindow {
    param(
        [int]$ExcelProcessId,
        [string]$Caption
    )

    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $processCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ProcessIdProperty,
        $ExcelProcessId
    )
    $classCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::ClassNameProperty,
        "ThunderDFrame"
    )
    $nameCondition = New-Object System.Windows.Automation.PropertyCondition(
        [System.Windows.Automation.AutomationElement]::NameProperty,
        $Caption
    )
    $condition = New-Object System.Windows.Automation.AndCondition(
        $processCondition,
        $classCondition,
        $nameCondition
    )
    $windows = [System.Windows.Automation.AutomationElement]::RootElement.FindAll(
        [System.Windows.Automation.TreeScope]::Descendants,
        $condition
    )
    if ($windows.Count -ne 1) {
        return $false
    }
    $nativeHandle = [IntPtr]$windows.Item(0).Current.NativeWindowHandle
    if ($nativeHandle -eq [IntPtr]::Zero) {
        return $false
    }
    [void][Plan022LauncherWindow]::SendMessage(
        $nativeHandle,
        0x0010,
        [IntPtr]::Zero,
        [IntPtr]::Zero
    )
    Start-Sleep -Milliseconds 750
    return $true
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
public static class Plan022LauncherWindow
{
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);
    [DllImport("user32.dll")]
    public static extern IntPtr SendMessage(
        IntPtr hWnd,
        uint message,
        IntPtr wParam,
        IntPtr lParam
    );
}
"@

$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$deployPath = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
$helperSource = Join-Path $repo "tools/validate_phase6_live_role_workflows.ps1"
Import-LiveValidationHelpers -ScriptPath $helperSource

$outputPath = Join-Path $repo $OutputDirectory
New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
$progressPath = Join-Path $outputPath "progress.txt"
$runtimeRoot = Join-Path ([IO.Path]::GetTempPath()) (
    "invsys-plan022-launcher-red-" + [guid]::NewGuid().ToString("N")
)
$warehouseId = "WHL" + [guid]::NewGuid().ToString("N").Substring(0, 6).ToUpperInvariant()
$stationId = "S1"
$testUser = if ([string]::IsNullOrWhiteSpace($env:USERNAME)) { "user1" } else { $env:USERNAME }
$testPin = [guid]::NewGuid().ToString("N")
$testPinHash = Get-InvSysCredentialHash -Credential $testPin
$configPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Config.xlsb")
$authPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Auth.xlsb")
$inventoryPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Data.Inventory.xlsb")
$snapshotPath = Join-Path $runtimeRoot ($warehouseId + ".invSys.Snapshot.Inventory.xlsb")
$operatorRoot = Join-Path $runtimeRoot "operator-workbooks"
$callbacks = @(
    @{
        Name = "Receiving"
        Macro = "modTS_Received.ShowReceivingForm"
        Expected = "Creates or opens one station-local saved Receiving operator workbook and opens the modeless form."
    },
    @{
        Name = "Production"
        Macro = "mProduction.BtnOpenProductionForm"
        Expected = "Creates or opens one station-local saved Production operator workbook and opens one modeless form."
    },
    @{
        Name = "Shipping"
        Macro = "modTS_Shipments.BtnOpenShipmentsForm"
        Expected = "Creates or opens one station-local saved Shipping operator workbook and opens one modeless form."
    }
)
if (-not [string]::IsNullOrWhiteSpace($CallbackFilter)) {
    $callbacks = @($callbacks | Where-Object { $_.Name -eq $CallbackFilter })
}
if ($WorkbookState -eq "CapturedClosed" -and [string]::IsNullOrWhiteSpace($CallbackFilter)) {
    throw "CapturedClosed requires -CallbackFilter so each role starts from its own captured workbook."
}
if ($WorkbookState -eq "ReceivingDurability") {
    if ([string]::IsNullOrWhiteSpace($CallbackFilter)) {
        $callbacks = @($callbacks | Where-Object { $_.Name -eq "Receiving" })
    }
    elseif ($CallbackFilter -ne "Receiving") {
        throw "ReceivingDurability supports only -CallbackFilter Receiving."
    }
}
if ($WorkbookState -eq "ReceivingFormClosed") {
    if ([string]::IsNullOrWhiteSpace($CallbackFilter)) {
        $callbacks = @($callbacks | Where-Object { $_.Name -eq "Receiving" })
    }
    elseif ($CallbackFilter -ne "Receiving") {
        throw "ReceivingFormClosed supports only -CallbackFilter Receiving."
    }
}

function Add-ProductionPickerProjectionFixture {
    param([object]$InventoryWorkbook)

    $wsSku = Get-WorksheetSafe -Workbook $InventoryWorkbook -WorksheetName "SkuCatalog"
    $loSku = Get-ListObjectSafe -Worksheet $wsSku -TableName "tblSkuCatalog"
    $wsLog = Get-WorksheetSafe -Workbook $InventoryWorkbook -WorksheetName "InventoryLog"
    $loLog = Get-ListObjectSafe -Worksheet $wsLog -TableName "tblInventoryLog"
    if ($null -eq $loSku) { throw "Production picker fixture requires tblSkuCatalog." }
    if ($null -eq $loLog) { throw "Production picker fixture requires tblInventoryLog." }
    foreach ($header in @("ITEM_CODE", "ITEM", "UOM", "LOCATION", "DESCRIPTION", "CATEGORY")) {
        if ((Get-ColumnIndexSafe -ListObject $loSku -ColumnName $header) -eq 0) {
            $newColumn = $loSku.ListColumns.Add()
            $newColumn.Name = $header
        }
    }

    $fixture = @{
        "SKU-RUN-RAW" = @("SYS-LIVE-PRODUCTION-RUN-RAW-A", "Production Raw Material", 3.0)
        "SKU-RUN-STALE" = @("SYS-LIVE-PRODUCTION-RUN-STALE", "Production Stale Material", 6.0)
    }
    $entityRows = @()
    foreach ($sku in @("SKU-RUN-RAW", "SKU-RUN-STALE")) {
        $rowIndex = 0
        if ($null -ne $loSku.DataBodyRange) {
            for ($candidateRow = 1; $candidateRow -le [int]$loSku.ListRows.Count; $candidateRow++) {
                $candidateSku = [string]$loSku.DataBodyRange.Cells($candidateRow, (Get-ColumnIndexSafe -ListObject $loSku -ColumnName "SKU")).Value2
                if ($candidateSku -eq $sku) { $rowIndex = $candidateRow; break }
            }
        }
        if ($rowIndex -eq 0) {
            [void]$loSku.ListRows.Add()
            $rowIndex = [int]$loSku.ListRows.Count
            $loSku.DataBodyRange.Cells($rowIndex, (Get-ColumnIndexSafe -ListObject $loSku -ColumnName "SKU")).Value2 = $sku
        }
        $loSku.DataBodyRange.Cells($rowIndex, (Get-ColumnIndexSafe -ListObject $loSku -ColumnName "ITEM_CODE")).Value2 = $sku
        $loSku.DataBodyRange.Cells($rowIndex, (Get-ColumnIndexSafe -ListObject $loSku -ColumnName "ITEM")).Value2 = [string]$fixture[$sku][1]
        $loSku.DataBodyRange.Cells($rowIndex, (Get-ColumnIndexSafe -ListObject $loSku -ColumnName "UOM")).Value2 = "LB"
        $loSku.DataBodyRange.Cells($rowIndex, (Get-ColumnIndexSafe -ListObject $loSku -ColumnName "LOCATION")).Value2 = "LINE"
        $loSku.DataBodyRange.Cells($rowIndex, (Get-ColumnIndexSafe -ListObject $loSku -ColumnName "DESCRIPTION")).Value2 = "isolated packaged picker fixture"
        $loSku.DataBodyRange.Cells($rowIndex, (Get-ColumnIndexSafe -ListObject $loSku -ColumnName "CATEGORY")).Value2 = "INGREDIENT"
        $entityRows += ,@([string]$fixture[$sku][0], $sku, [double]$fixture[$sku][2], "LINE", "GOOD", "ACTIVE", "{}", [datetime]::UtcNow)
    }
    foreach ($rawPart in @(
        @("B", 3.0), @("C", 3.0), @("D", 3.0), @("E", 3.0),
        @("F", 3.0), @("G", 2.0)
    )) {
        $entityRows += ,@(
            "SYS-LIVE-PRODUCTION-RUN-RAW-$($rawPart[0])", "SKU-RUN-RAW",
            [double]$rawPart[1], "LINE", "GOOD", "ACTIVE", "{}", [datetime]::UtcNow
        )
    }

    $rawSeedRow = 0
    for ($candidateRow = 1; $candidateRow -le [int]$loLog.ListRows.Count; $candidateRow++) {
        $candidateSku = [string]$loLog.DataBodyRange.Cells(
            $candidateRow,
            (Get-ColumnIndexSafe -ListObject $loLog -ColumnName "SKU")
        ).Value2
        if ($candidateSku -eq "SKU-RUN-RAW") { $rawSeedRow = $candidateRow; break }
    }
    if ($rawSeedRow -eq 0) { throw "Production picker fixture requires the raw seed event." }
    $fixtureWarehouseId = [string]$loLog.DataBodyRange.Cells(
        $rawSeedRow,
        (Get-ColumnIndexSafe -ListObject $loLog -ColumnName "WarehouseId")
    ).Value2
    $loLog.DataBodyRange.Cells(
        $rawSeedRow,
        (Get-ColumnIndexSafe -ListObject $loLog -ColumnName "System_Key")
    ).Value2 = "SYS-LIVE-PRODUCTION-RUN-RAW-A"
    $loLog.DataBodyRange.Cells(
        $rawSeedRow,
        (Get-ColumnIndexSafe -ListObject $loLog -ColumnName "QtyDelta")
    ).Value2 = 3.0
    foreach ($rawPart in @(
        @("B", 3.0), @("C", 3.0), @("D", 3.0), @("E", 3.0),
        @("F", 3.0), @("G", 2.0)
    )) {
        Add-ListObjectRow -ListObject $loLog -Values @{
            "EventID" = "EVT-LIVE-SEED-SKU-RUN-RAW"
            "UndoOfEventId" = ""
            "AppliedSeq" = 1
            "EventType" = "INVENTORY_CREATE"
            "OccurredAtUTC" = [datetime]::UtcNow
            "AppliedAtUTC" = [datetime]::UtcNow
            "WarehouseId" = $fixtureWarehouseId
            "StationId" = "S1"
            "UserId" = "svc_processor"
            "System_Key" = "SYS-LIVE-PRODUCTION-RUN-RAW-$($rawPart[0])"
            "SKU" = "SKU-RUN-RAW"
            "QtyDelta" = [double]$rawPart[1]
            "Location" = "LINE"
            "Condition" = "GOOD"
            "AttributesJson" = "{}"
            "Note" = "isolated packaged split stock fixture"
        }
    }

    $wsEntities = $InventoryWorkbook.Worksheets.Add()
    $wsEntities.Name = "InventoryEntities"
    Add-Table -Worksheet $wsEntities -TableName "tblInventoryEntities" -Headers @(
        "System_Key", "SKU", "QtyOnHand", "Location", "Condition", "InventoryState",
        "AttributesJson", "LastAppliedUTC"
    ) -Rows $entityRows | Out-Null
    $InventoryWorkbook.Save()
}
if ($WorkbookState -eq "ShippingLayout") {
    if ([string]::IsNullOrWhiteSpace($CallbackFilter)) {
        $callbacks = @($callbacks | Where-Object { $_.Name -eq "Shipping" })
    }
    elseif ($CallbackFilter -ne "Shipping") {
        throw "ShippingLayout supports only -CallbackFilter Shipping."
    }
}
if ($WorkbookState -eq "ProductionReusable") {
    if ([string]::IsNullOrWhiteSpace($CallbackFilter)) {
        $callbacks = @($callbacks | Where-Object { $_.Name -eq "Production" })
    }
    elseif ($CallbackFilter -ne "Production") {
        throw "ProductionReusable supports only -CallbackFilter Production."
    }
}
if ($ProductionRunOnly -and
    ($WorkbookState -ne "ProductionReusable" -or $CallbackFilter -ne "Production")) {
    throw "ProductionRunOnly requires -WorkbookState ProductionReusable -CallbackFilter Production."
}
if ($ProductionEditExportOnly -and
    ($WorkbookState -ne "ProductionReusable" -or $CallbackFilter -ne "Production")) {
    throw "ProductionEditExportOnly requires -WorkbookState ProductionReusable -CallbackFilter Production."
}
if ($ProductionRunOnly -and $ProductionEditExportOnly) {
    throw "ProductionRunOnly and ProductionEditExportOnly are mutually exclusive."
}
$packageNames = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Operations.xlam"
)

$excel = $null
$excelProcessId = 0
$opened = New-Object System.Collections.Generic.List[object]
$packages = @{}
$evidence = New-Object System.Collections.Generic.List[object]
$setupEvidence = New-Object System.Collections.Generic.List[string]
$currentStep = "startup"
$stateWb = $null
$inventoryWb = $null
$canonicalPaths = @()
$canonicalHashesBefore = @{}
$reusableRecipeId = ""
$reusableRecipeVersion = "2"
$productionOperatorPath = ""

try {
    [IO.File]::WriteAllText($progressPath, "create runtime root")
    New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null
    [IO.File]::WriteAllText($progressPath, "start Excel")
    $excel = New-Object -ComObject Excel.Application
    $excelProcessId = Get-ExcelProcessId $excel
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1

    $currentStep = "create isolated config and auth"
    [IO.File]::WriteAllText($progressPath, $currentStep)
    $configWb = New-ConfigWorkbook -Excel $excel -Path $configPath `
        -WarehouseId $warehouseId -StationId $stationId -RuntimeRoot $runtimeRoot
    $authWb = New-AuthWorkbook -Excel $excel -Path $authPath `
        -WarehouseId $warehouseId -StationId $stationId `
        -CurrentUserIds @($testUser) -CredentialHash $testPinHash
    $opened.Add($configWb) | Out-Null
    $opened.Add($authWb) | Out-Null
    if ($WorkbookState -eq "ReceivingDurability") {
        $inventoryWb = New-InventoryWorkbook -Excel $excel -Path $inventoryPath `
            -WarehouseId $warehouseId -SkuRows @("SKU-R1-DURABILITY")
        $opened.Add($inventoryWb) | Out-Null
    }
    elseif ($WorkbookState -eq "ProductionReusable") {
        $inventoryWb = New-InventoryWorkbook -Excel $excel -Path $inventoryPath `
            -WarehouseId $warehouseId -SkuRows @("SKU-RUN-RAW", "SKU-RUN-STALE")
        Add-ProductionPickerProjectionFixture -InventoryWorkbook $inventoryWb
        $opened.Add($inventoryWb) | Out-Null
    }

    $currentStep = "open packaged add-ins"
    [IO.File]::WriteAllText($progressPath, $currentStep)
    foreach ($packageName in $packageNames) {
        $packagePath = Join-Path $deployPath $packageName
        $packageWb = $excel.Workbooks.Open($packagePath)
        $opened.Add($packageWb) | Out-Null
        $packages[$packageName] = $packageWb
    }

    $coreName = [string]$packages["invSys.Core.xlam"].Name
    $operationsName = [string]$packages["invSys.Operations.xlam"].Name
    [IO.File]::WriteAllText($progressPath, "configure and sign in")
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($runtimeRoot))
    $configLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modConfig.LoadConfig" -Arguments @($warehouseId, $stationId))
    $authLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modAuth.LoadAuth" -Arguments @($warehouseId))
    $targetResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modNasConnection.SelectWarehouseTargetForAutomation" `
        -Arguments @($runtimeRoot, $runtimeRoot, $stationId, $true))
    $testHubRoot = "\\plan022-test\warehouse"
    $targetPathsSet = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modNasConnection.SetCurrentTargetPathsForTest" `
        -Arguments @($testHubRoot, $runtimeRoot))
    $operatorRootOverrideSet = $false
    try {
        $operatorRootOverrideSet = [bool](Run-WorkbookMacro -Excel $excel `
            -WorkbookName $coreName `
            -MacroName "modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation" `
            -Arguments @($operatorRoot))
    }
    catch {}
    $signInResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
        -MacroName "modAuth.SignInCurrentTargetForAutomation" `
        -Arguments @($testUser, $testPin, "RECEIVE_POST"))
    $setupEvidence.Add("ConfigLoaded=$configLoaded") | Out-Null
    $setupEvidence.Add("AuthLoaded=$authLoaded") | Out-Null
    $setupEvidence.Add("TargetSelected=$($targetResult.StartsWith('OK|'))") | Out-Null
    $setupEvidence.Add("TestNasTargetApplied=$targetPathsSet") | Out-Null
    $setupEvidence.Add("OperatorRootOverrideSet=$operatorRootOverrideSet") | Out-Null
    $setupEvidence.Add("SignedIn=$($signInResult.StartsWith('OK|'))") | Out-Null

    $currentStep = "validate signed-in fixture before callbacks"
    if (-not $signInResult.StartsWith('OK|')) {
        throw "Fixture sign-in failed before any packaged workflow callback."
    }

    if ($WorkbookState -eq "ReceivingDurability") {
        $currentStep = "generate isolated canonical snapshot"
        [IO.File]::WriteAllText($progressPath, $currentStep)
        $snapshotCreated = [bool](Run-WorkbookMacro -Excel $excel `
            -WorkbookName $coreName `
            -MacroName "modWarehouseSync.GenerateWarehouseSnapshot" `
            -Arguments @($warehouseId, $inventoryWb, $snapshotPath))
        if (-not $snapshotCreated) {
            throw "Unable to generate the isolated canonical snapshot."
        }
        $setupEvidence.Add("SnapshotCreated=True") | Out-Null
    }

    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
        -MacroName "modOperationsInit.Auto_Open")

    $currentStep = "prepare workbook state"
    [IO.File]::WriteAllText($progressPath, $currentStep)
    if ($WorkbookState -in @("NoEligible", "ReceivingDurability", "ReceivingFormClosed", "ShippingLayout", "ProductionReusable")) {
        if ($null -ne $inventoryWb) {
            $inventoryWb.Close($true)
        }
        $authWb.Close($false)
        $configWb.Close($false)
    }
    elseif ($WorkbookState -eq "ConfigActive") {
        $authWb.Close($false)
        $configWb.Activate()
        $stateWb = $configWb
    }
    else {
        $authWb.Close($false)
        $configWb.Close($false)
        if ($WorkbookState -in @("SavedEligible", "CapturedClosed")) {
            $stateFileName = "Plan022.Role.Operator.xlsb"
        }
        else {
            $stateFileName = "Plan022.Unrelated.xlsb"
        }
        $statePath = Join-Path $runtimeRoot $stateFileName
        $stateWb = New-OperationalWorkbook -Excel $excel `
            -NameHint ([IO.Path]::GetFileName($statePath)) -Path $statePath
        $opened.Add($stateWb) | Out-Null
        if ($WorkbookState -in @("SavedEligible", "CapturedClosed")) {
            foreach ($surfaceMacro in @(
                "modRoleWorkbookSurfaces.EnsureReceivingWorkbookSurface",
                "modRoleWorkbookSurfaces.EnsureProductionWorkbookSurface",
                "modRoleWorkbookSurfaces.EnsureShippingWorkbookSurface"
            )) {
                $surfaceOk = [bool](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $coreName -MacroName $surfaceMacro `
                    -Arguments @($stateWb))
                if (-not $surfaceOk) {
                    throw "Unable to create saved eligible role surface: $surfaceMacro"
                }
            }
            $stateWb.Save()
        }
        $stateWb.Activate()
    }
    if ($WorkbookState -eq "ReceivingDurability") {
        $canonicalPaths = @($configPath, $authPath, $inventoryPath, $snapshotPath)
        $canonicalHashesBefore = Get-FileHashMap -Paths $canonicalPaths
    }

    $currentStep = "invoke packaged callbacks for state $WorkbookState"
    [IO.File]::WriteAllText($progressPath, $currentStep)
    foreach ($callback in $callbacks) {
        [IO.File]::WriteAllText($progressPath, "invoke " + $callback.Macro)
        $beforeNames = @(Get-OpenWorkbookNames -Excel $excel)
        $beforeSurface = Get-WorkbookSurfaceCounts -Workbook $stateWb
        $capture = Invoke-PackagedCallback -Excel $excel `
            -WorkbookName $operationsName -MacroName $callback.Macro
        $afterNames = @(Get-OpenWorkbookNames -Excel $excel)
        $afterSurface = Get-WorkbookSurfaceCounts -Workbook $stateWb
        $observedText = (@($capture.WindowText) -join " || ")
        if (-not [string]::IsNullOrWhiteSpace($capture.Error)) {
            $observedText += " || COM_ERROR=" + $capture.Error
        }
        $newWorkbooks = @($afterNames | Where-Object { $_ -notin $beforeNames })
        $newWorkbookPaths = @()
        foreach ($newWorkbookName in $newWorkbooks) {
            try {
                $newWorkbookPaths += [string]$excel.Workbooks.Item($newWorkbookName).FullName
            }
            catch {}
        }
        if ($newWorkbooks.Count -gt 0) {
            $observedText += " || NEW_WORKBOOKS=" + ($newWorkbooks -join ",")
            $observedText += " || NEW_WORKBOOKS_STATION_LOCAL=" +
                [bool](@($newWorkbookPaths | Where-Object {
                    [IO.Path]::GetFullPath($_).StartsWith(
                        [IO.Path]::GetFullPath($operatorRoot),
                        [StringComparison]::OrdinalIgnoreCase
                    )
                }).Count -eq $newWorkbookPaths.Count)
        }
        $surfaceChanged = $beforeSurface.Worksheets -ne $afterSurface.Worksheets -or
            $beforeSurface.Tables -ne $afterSurface.Tables
        if ($null -ne $stateWb) {
            $observedText += " || SURFACE_BEFORE=$($beforeSurface.Worksheets)/$($beforeSurface.Tables)"
            $observedText += " || SURFACE_AFTER=$($afterSurface.Worksheets)/$($afterSurface.Tables)"
            $observedText += " || SURFACE_CHANGED=$surfaceChanged"
        }

        $passed = $false
        if ($WorkbookState -eq "ShippingLayout" -and $callback.Name -eq "Shipping") {
            $statusLayoutReport = [string](Run-WorkbookMacro -Excel $excel `
                -WorkbookName $operationsName `
                -MacroName "modTS_Shipments.RunShippingStatusAnchorTest")
            $boxingLayoutReport = [string](Run-WorkbookMacro -Excel $excel `
                -WorkbookName $operationsName `
                -MacroName "modTS_Shipments.RunShippingBoxingLayoutTest")
            $shippingIdentityReport = [string](Run-WorkbookMacro -Excel $excel `
                -WorkbookName $operationsName `
                -MacroName "modTS_Shipments.RunShippingSystemKeyIdentityTest")
            $observedText += " || STATUS_LAYOUT=" + $statusLayoutReport
            $observedText += " || BOXING_LAYOUT=" + $boxingLayoutReport
            $observedText += " || SHIPPING_IDENTITY=" + $shippingIdentityReport
            $passed = [string]::IsNullOrWhiteSpace($capture.Error) -and
                $statusLayoutReport -match '^OK\|' -and
                $statusLayoutReport -match '(?:^|\|)TopStable=True(?:\||$)' -and
                $statusLayoutReport -match '(?:^|\|)HeightStable=True(?:\||$)' -and
                $statusLayoutReport -match '(?:^|\|)AboveSearch=True(?:\||$)' -and
                $boxingLayoutReport -match '^OK\|' -and
                $boxingLayoutReport -match '(?:^|\|)BuilderInventoryGrew=True(?:\||$)' -and
                $boxingLayoutReport -match '(?:^|\|)MakerDesignsGrew=True(?:\||$)' -and
                $boxingLayoutReport -match '(?:^|\|)BoxingHeaderWidthsMatchLists=True(?:\||$)' -and
                $boxingLayoutReport -match '(?:^|\|)HeadersMatch=True(?:\||$)' -and
                $boxingLayoutReport -match '(?:^|\|)SearchFiltered=True(?:\||$)' -and
                $boxingLayoutReport -match '(?:^|\|)NonVersionIsNA=True(?:\||$)' -and
                $shippingIdentityReport -match '^OK\|' -and
                $shippingIdentityReport -match '(?:^|\|)ControlReady=True(?:\||$)' -and
                $shippingIdentityReport -match '(?:^|\|)ValuePreserved=True(?:\||$)' -and
                $shippingIdentityReport -match '(?:^|\|)ReservationUsesKey=True(?:\||$)'
        }
        elseif ($WorkbookState -eq "ReceivingDurability" -and $callback.Name -eq "Receiving") {
            $operatorPath = @($newWorkbookPaths | Where-Object {
                [IO.Path]::GetFullPath($_).StartsWith(
                    [IO.Path]::GetFullPath($operatorRoot),
                    [StringComparison]::OrdinalIgnoreCase
                )
            } | Select-Object -First 1)
            $operatorWb = if ($operatorPath.Count -eq 1) {
                $excel.Workbooks.Item([IO.Path]::GetFileName($operatorPath[0]))
            }
            else {
                $null
            }
            $customHeader = "Custom_R1_Launcher_Persistence"
            $customValue = "PRESERVE"
            $customAdded = $false
            $historyRowsAdded = 0
            $receivingHistoryReport = ""
            $receivingControlReport = ""
            if ($null -ne $operatorWb) {
                $inventorySheet = Get-WorksheetSafe -Workbook $operatorWb `
                    -WorksheetName "InventoryManagement"
                $inventoryTable = Get-ListObjectSafe -Worksheet $inventorySheet `
                    -TableName "invSys"
                if ($null -ne $inventoryTable) {
                    $customColumn = $inventoryTable.ListColumns.Add()
                    $customColumn.Name = $customHeader
                    if ((Get-RowCountSafe $inventoryTable) -gt 0) {
                        $customColumn.DataBodyRange.Cells.Item(1, 1).Value2 = $customValue
                    }
                    $operatorWb.Save()
                    $customAdded = (Get-ColumnIndexSafe -ListObject $inventoryTable -ColumnName $customHeader) -gt 0
                }
                $historySheet = Get-WorksheetSafe -Workbook $operatorWb `
                    -WorksheetName "ReceivedLog"
                $historyTable = Get-ListObjectSafe -Worksheet $historySheet `
                    -TableName "ReceivedLog"
                if ($null -ne $historyTable) {
                    foreach ($historyOrdinal in 1..2) {
                        $historyRow = $historyTable.ListRows.Add()
                        $historyValues = @{
                            "ENTRY_DATE" = [DateTime]::UtcNow.AddMinutes(-$historyOrdinal).ToOADate()
                            "USER" = $testUser
                            "REF_NUMBER" = "R1-HISTORY-$historyOrdinal"
                            "ITEMS" = "R1 durability item"
                            "QUANTITY" = $historyOrdinal
                            "UOM" = "EA"
                            "VENDOR" = "R1 test vendor"
                            "LOCATION" = "TEST"
                            "ITEM_CODE" = "SKU-R1-DURABILITY"
                            "System_Key" = "SYS-R1-HISTORY-$historyOrdinal"
                            "EventId" = "EVT-R1-HISTORY-$historyOrdinal"
                        }
                        foreach ($historyColumn in $historyValues.Keys) {
                            $historyColumnIndex = Get-ColumnIndexSafe `
                                -ListObject $historyTable -ColumnName $historyColumn
                            if ($historyColumnIndex -gt 0) {
                                $historyRow.Range.Cells.Item(1, $historyColumnIndex).Value2 = `
                                    $historyValues[$historyColumn]
                            }
                        }
                        $historyRowsAdded++
                    }
                    $operatorWb.Save()
                }
                $receivingHistoryReport = [string](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $operationsName `
                    -MacroName "modTS_Received.RunReceivingRefreshFormActionForTest" `
                    -Arguments @([string]$operatorWb.Name, ""))
                $receivingControlReport = [string](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $operationsName `
                    -MacroName "modTS_Received.RunReceivingSearchAndHeaderContractTest")
            }
            $formsBeforeClose = Get-RoleFormWindowCount -ExcelProcessId $excelProcessId
            if ($null -ne $operatorWb) {
                $operatorWb.Close($true)
            }

            $unrelatedPath = Join-Path $runtimeRoot "Plan022.Durability.Unrelated.xlsb"
            $unrelatedWb = New-OperationalWorkbook -Excel $excel `
                -NameHint ([IO.Path]::GetFileName($unrelatedPath)) -Path $unrelatedPath
            $opened.Add($unrelatedWb) | Out-Null
            $unrelatedBefore = Get-WorkbookSurfaceCounts -Workbook $unrelatedWb
            $unrelatedWb.Activate()

            $reopenNamesBefore = @(Get-OpenWorkbookNames -Excel $excel)
            $reopenCapture = Invoke-PackagedCallback -Excel $excel `
                -WorkbookName $operationsName -MacroName $callback.Macro
            $reopenNamesAfter = @(Get-OpenWorkbookNames -Excel $excel)
            $reopenNewNames = @($reopenNamesAfter | Where-Object { $_ -notin $reopenNamesBefore })
            $operatorWb = if ($operatorPath.Count -eq 1) {
                try { $excel.Workbooks.Item([IO.Path]::GetFileName($operatorPath[0])) } catch { $null }
            }
            else {
                $null
            }
            $refreshOk = $false
            $customPreserved = $false
            if ($null -ne $operatorWb) {
                $refreshOk = [bool](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $coreName `
                    -MacroName "modOperationsPrimitiveBridge.RefreshInventoryReadModel" `
                    -Arguments @([string]$operatorWb.Name, $warehouseId, "LOCAL"))
                $inventorySheet = Get-WorksheetSafe -Workbook $operatorWb `
                    -WorksheetName "InventoryManagement"
                $inventoryTable = Get-ListObjectSafe -Worksheet $inventorySheet `
                    -TableName "invSys"
                $customIndex = Get-ColumnIndexSafe -ListObject $inventoryTable -ColumnName $customHeader
                $customPreserved = $customIndex -gt 0
                if ($customPreserved -and (Get-RowCountSafe $inventoryTable) -gt 0) {
                    $customPreserved = (
                        [string]$inventoryTable.DataBodyRange.Cells.Item(1, $customIndex).Value2
                    ) -eq $customValue
                }
            }

            $formCountAfterReopen = Get-RoleFormWindowCount -ExcelProcessId $excelProcessId
            $unrelatedWb.Activate()
            Start-Sleep -Milliseconds 300
            $formCountAfterUnrelatedActivation = Get-RoleFormWindowCount `
                -ExcelProcessId $excelProcessId
            $unrelatedAfter = Get-WorkbookSurfaceCounts -Workbook $unrelatedWb
            $unrelatedChanged = $unrelatedBefore.Worksheets -ne $unrelatedAfter.Worksheets -or
                $unrelatedBefore.Tables -ne $unrelatedAfter.Tables
            $canonicalHashesAfter = Get-FileHashMap -Paths $canonicalPaths
            $canonicalHashesUnchanged = Compare-FileHashMaps `
                -Before $canonicalHashesBefore -After $canonicalHashesAfter
            $changedCanonicalNames = @(
                Get-ChangedFileNames `
                    -Before $canonicalHashesBefore -After $canonicalHashesAfter
            )
            $operatorFileCount = @(
                Get-ChildItem -LiteralPath $operatorRoot -Filter "*.Receiving.Operator.xlsm" `
                    -File -Recurse -ErrorAction SilentlyContinue
            ).Count
            $reopenText = @($reopenCapture.WindowText) -join " // "
            $observedText += " || CUSTOM_ADDED=$customAdded"
            $observedText += " || RECEIVING_HISTORY_ROWS_ADDED=$historyRowsAdded"
            $observedText += " || RECEIVING_HISTORY=" + $receivingHistoryReport
            $observedText += " || RECEIVING_CONTROLS=" + $receivingControlReport
            $observedText += " || REFRESH_OK=$refreshOk"
            $observedText += " || CUSTOM_PRESERVED=$customPreserved"
            $observedText += " || CANONICAL_HASHES_UNCHANGED=$canonicalHashesUnchanged"
            if ($changedCanonicalNames.Count -gt 0) {
                $observedText += " || CHANGED_CANONICAL_FILES=" +
                    ($changedCanonicalNames -join ",")
            }
            $observedText += " || FORMS_BEFORE_CLOSE=$formsBeforeClose"
            $observedText += " || FORMS_AFTER_REOPEN=$formCountAfterReopen"
            $observedText += " || FORMS_AFTER_UNRELATED_ACTIVATION=$formCountAfterUnrelatedActivation"
            $observedText += " || UNRELATED_SURFACE_CHANGED=$unrelatedChanged"
            $observedText += " || OPERATOR_FILES=$operatorFileCount"
            $observedText += " || REOPEN_NEW_WORKBOOKS=" + ($reopenNewNames -join ",")
            if (-not [string]::IsNullOrWhiteSpace($reopenText)) {
                $observedText += " || REOPEN_DIALOGS=" + $reopenText
            }
            $passed = $operatorRootOverrideSet -and
                $newWorkbooks.Count -eq 1 -and
                $customAdded -and
                $historyRowsAdded -eq 2 -and
                $receivingHistoryReport -match '^OK\|' -and
                $receivingHistoryReport -match '(?:^|\|)HistoryRows=2(?:\||$)' -and
                $receivingHistoryReport -match '(?:^|\|)LoaderHistoryRowsBefore=2(?:\||$)' -and
                $receivingHistoryReport -match '(?:^|\|)LoaderHistoryRowsAfter=2(?:\||$)' -and
                $receivingHistoryReport -match '(?:^|\|)DirectHistoryRows=2(?:\||$)' -and
                $receivingControlReport -match '^OK\|' -and
                $receivingControlReport -match '(?:^|\|)DedicatedItemResults=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)Location=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)OptionalLot=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)ReceivingHeaderColumnsAligned=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)CapacityStub=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)SearchRowsLoaded=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)HiddenSystemKeyMap=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)TenColumnItemResults=True(?:\||$)' -and
                $receivingControlReport -match '(?:^|\|)HeadersSingleLine=True(?:\||$)' -and
                $refreshOk -and
                $customPreserved -and
                $canonicalHashesUnchanged -and
                $formsBeforeClose -eq 1 -and
                $formCountAfterReopen -eq 1 -and
                $formCountAfterUnrelatedActivation -eq 1 -and
                -not $unrelatedChanged -and
                $operatorFileCount -eq 1 -and
                $reopenNewNames.Count -eq 1 -and
                $reopenText -notmatch "(?i)(failed|Type mismatch|operator workbook)"
        }
        elseif ($WorkbookState -eq "ReceivingFormClosed") {
            $formCountBeforeClose = Get-RoleFormWindowCount `
                -ExcelProcessId $excelProcessId
            $formClosed = Close-RoleFormWindow -ExcelProcessId $excelProcessId `
                -Caption "Receiving"
            $formCountAfterClose = Get-RoleFormWindowCount `
                -ExcelProcessId $excelProcessId
            $secondBeforeNames = @(Get-OpenWorkbookNames -Excel $excel)
            $secondCapture = Invoke-PackagedCallback -Excel $excel `
                -WorkbookName $operationsName -MacroName $callback.Macro
            $automationDisconnectedAfterRelaunch = $false
            try {
                $secondAfterNames = @(Get-OpenWorkbookNames -Excel $excel)
            }
            catch {
                $automationDisconnectedAfterRelaunch = $true
                $secondAfterNames = @($secondBeforeNames)
            }
            $secondNewWorkbooks = @(
                $secondAfterNames | Where-Object { $_ -notin $secondBeforeNames }
            )
            $formCountAfterRelaunch = Get-RoleFormWindowCount `
                -ExcelProcessId $excelProcessId
            $ownedExcelProcess = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
            $excelAliveAfterRelaunch = $null -ne $ownedExcelProcess -and
                $ownedExcelProcess.ProcessName -eq "EXCEL" -and
                $ownedExcelProcess.Responding
            $operatorFileCount = @(
                Get-ChildItem -LiteralPath $operatorRoot `
                    -Filter "*.Receiving.Operator.xlsm" -File -Recurse `
                    -ErrorAction SilentlyContinue
            ).Count
            $secondText = @($secondCapture.WindowText) -join " // "
            $observedText += " || FORM_CLOSED=$formClosed"
            $observedText += " || FORMS_BEFORE_CLOSE=$formCountBeforeClose"
            $observedText += " || FORMS_AFTER_CLOSE=$formCountAfterClose"
            $observedText += " || FORMS_AFTER_RELAUNCH=$formCountAfterRelaunch"
            $observedText += " || RELAUNCH_NEW_WORKBOOKS=" + $secondNewWorkbooks.Count
            $observedText += " || OPERATOR_FILES=$operatorFileCount"
            $observedText += " || EXCEL_RESPONDING_AFTER_RELAUNCH=$excelAliveAfterRelaunch"
            $observedText += " || AUTOMATION_CLIENT_DISCONNECTED_AFTER_RELAUNCH=" +
                $automationDisconnectedAfterRelaunch
            if (-not [string]::IsNullOrWhiteSpace($secondCapture.Error)) {
                $observedText += " || RELAUNCH_COM_ERROR=" + $secondCapture.Error
            }
            if (-not [string]::IsNullOrWhiteSpace($secondText)) {
                $observedText += " || RELAUNCH_DIALOGS=" + $secondText
            }
            $passed = $newWorkbooks.Count -eq 1 -and
                $formCountBeforeClose -eq 1 -and
                $formClosed -and
                $formCountAfterClose -eq 0 -and
                $secondNewWorkbooks.Count -eq 0 -and
                $formCountAfterRelaunch -eq 1 -and
                $operatorFileCount -eq 1 -and
                $excelAliveAfterRelaunch -and
                [string]::IsNullOrWhiteSpace($capture.Error) -and
                [string]::IsNullOrWhiteSpace($secondCapture.Error) -and
                $secondText -notmatch "(?i)(failed|automation error)"
        }
        elseif ($WorkbookState -in @("NoEligible", "ProductionReusable")) {
            $workflowControlReport = ""
            $workflowControlPassed = $true
            $secondBeforeNames = @(Get-OpenWorkbookNames -Excel $excel)
            $secondCapture = Invoke-PackagedCallback -Excel $excel `
                -WorkbookName $operationsName -MacroName $callback.Macro
            $secondAfterNames = @(Get-OpenWorkbookNames -Excel $excel)
            $secondNewWorkbooks = @($secondAfterNames | Where-Object { $_ -notin $secondBeforeNames })
            $observedText += " || SECOND_LAUNCH_NEW_WORKBOOKS=" + $secondNewWorkbooks.Count
            if ($callback.Name -eq "Production") {
                [IO.File]::WriteAllText($progressPath, "invoke Production batch-scale contract")
                $workflowControlReport = [string](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $operationsName `
                    -MacroName "mProduction.RunProductionBatchScaleContractTest")
                $workflowControlPassed = $workflowControlReport -match '^OK\|' -and
                    $workflowControlReport -match '(?:^|\|)Min=\.001%(?:\||$)' -and
                    $workflowControlReport -match '(?:^|\|)Default=100%(?:\||$)' -and
                    $workflowControlReport -match '(?:^|\|)Max=1000%(?:\||$)' -and
                    $workflowControlReport -match '(?:^|\|)BoundsRejected=True(?:\||$)'
                $observedText += " || PRODUCTION_BATCH_SCALE=" + $workflowControlReport
                [IO.File]::WriteAllText($progressPath, "invoke Production Run List responsive-layout handler")
                $productionRunListLayoutReport = [string](Run-WorkbookMacro -Excel $excel `
                    -WorkbookName $operationsName `
                    -MacroName "mProduction.RunProductionRunListResponsiveLayoutTest")
                $productionRunListLayoutPassed = $productionRunListLayoutReport -match '^OK\|' -and
                    $productionRunListLayoutReport -match '(?:^|\|)CheckEightRows=True(?:\||$)' -and
                    $productionRunListLayoutReport -match '(?:^|\|)InstructionsFourRows=True(?:\||$)' -and
                    $productionRunListLayoutReport -match '(?:^|\|)AllListsGrew=True(?:\||$)' -and
                    $productionRunListLayoutReport -match '(?:^|\|)InstructionLeftStable=True(?:\||$)' -and
                    $productionRunListLayoutReport -match '(?:^|\|)HeadersAligned=True(?:\||$)' -and
                    $productionRunListLayoutReport -match '(?:^|\|)RunListVerticalScrollbar=True(?:\||$)' -and
                    $productionRunListLayoutReport -match '(?:^|\|)RunPlanQtyHeaderAligned=True(?:\||$)' -and
                    $productionRunListLayoutReport -match '(?:^|\|)GeometryHealthy=True(?:\||$)'
                $workflowControlPassed = $workflowControlPassed -and $productionRunListLayoutPassed
                $observedText += " || PRODUCTION_RUN_LIST_LAYOUT=" + $productionRunListLayoutReport
                if ($WorkbookState -eq "ProductionReusable") {
                    [IO.File]::WriteAllText($progressPath, "invoke Production EA whole-unit handler")
                    $productionEaWholeUnitReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunEaWholeUnitActionContractTest")
                    $productionEaWholeUnitPassed = $productionEaWholeUnitReport -match '^OK\|' -and
                        $productionEaWholeUnitReport -match '(?:^|\|)RequirementHandlerRejected=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionEaWholeUnitPassed
                    $observedText += " || PRODUCTION_EA_WHOLE_UNIT=" + $productionEaWholeUnitReport
                    [IO.File]::WriteAllText($progressPath, "invoke Production output regulation settings handler")
                    $productionOutputRegulationReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunProductionOutputRegulationActionContractTest")
                    $productionOutputRegulationPassed = $productionOutputRegulationReport -match '^OK\|' -and
                        $productionOutputRegulationReport -match '(?:^|\|)SettingsPage=True(?:\||$)' -and
                        $productionOutputRegulationReport -match '(?:^|\|)VersionedDefault=True(?:\||$)' -and
                        $productionOutputRegulationReport -match '(?:^|\|)FractionalEaRejected=True(?:\||$)' -and
                        $productionOutputRegulationReport -match '(?:^|\|)CompletionHandler=ValidateReusableActualOutput(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionOutputRegulationPassed
                    $observedText += " || PRODUCTION_OUTPUT_REGULATION=" + $productionOutputRegulationReport
                    [IO.File]::WriteAllText($progressPath, "invoke Production variable quantity mode handler")
                    $productionVariableQuantityReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunProductionVariableQuantityModeActionContractTest")
                    $productionVariableQuantityPassed = $productionVariableQuantityReport -match '^OK\|' -and
                        $productionVariableQuantityReport -match '(?:^|\|)RequirementActual=True(?:\||$)' -and
                        $productionVariableQuantityReport -match '(?:^|\|)OutputActual=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionVariableQuantityPassed
                    $observedText += " || PRODUCTION_VARIABLE_QUANTITY=" + $productionVariableQuantityReport
                    if ($ProductionEditExportOnly) {
                    [IO.File]::WriteAllText($progressPath, "invoke focused Production released Process edit and export")
                    $productionEditExportReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunProcessEditExportContractTest")
                    $productionEditExportPassed = $productionEditExportReport -match '^OK\|' -and
                        $productionEditExportReport -match '(?:^|\|)ReleasedProcessEditable=True(?:\||$)' -and
                        $productionEditExportReport -match '(?:^|\|)ExistingProcessExported=True(?:\||$)' -and
                        $productionEditExportReport -match '(?:^|\|)ExportRoundTrip=True(?:\||$)' -and
                        $productionEditExportReport -match '(?:^|\|)OutputDesignVersionRebased=True(?:\||$)' -and
                        $productionEditExportReport -match '(?:^|\|)OutputYieldRebased=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionEditExportPassed
                    $observedText += " || PRODUCTION_PROCESS_EDIT_EXPORT=" + $productionEditExportReport
                    }
                    else {
                    if (-not $ProductionRunOnly) {
                    [IO.File]::WriteAllText($progressPath, "invoke Production reusable-surface contract")
                    $productionDesignReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunReusableProductionSurfaceContractTest")
                    $productionDesignPassed = $productionDesignReport -match '^OK\|' -and
                        $productionDesignReport -match '(?:^|\|)Pages=6(?:\||$)' -and
                        $productionDesignReport -match '(?:^|\|)ProcessDesigner=True(?:\||$)' -and
                        $productionDesignReport -match '(?:^|\|)RecipeDesigner=True(?:\||$)' -and
                        $productionDesignReport -match '(?:^|\|)ProductionSettings=True(?:\||$)' -and
                        $productionDesignReport -match '(?:^|\|)LegacyRecipeBuilder=False(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionDesignPassed
                    $observedText += " || PRODUCTION_REUSABLE_DESIGN=" + $productionDesignReport

                    [IO.File]::WriteAllText($progressPath, "invoke Production Process worksheet workbench")
                    $productionWorkbenchReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunProcessWorksheetWorkbenchContractTest")
                    $productionWorkbenchPassed = $productionWorkbenchReport -match '^OK\|' -and
                        $productionWorkbenchReport -match '(?:^|\|)SeparateActions=True(?:\||$)' -and
                        $productionWorkbenchReport -match '(?:^|\|)MultipleTables=True(?:\||$)' -and
                        $productionWorkbenchReport -match '(?:^|\|)SelectedOnly=True(?:\||$)' -and
                        $productionWorkbenchReport -match '(?:^|\|)RecordTypeDropdown=True(?:\||$)' -and
                        $productionWorkbenchReport -match '(?:^|\|)CalculatedPercent=True(?:\||$)' -and
                        $productionWorkbenchReport -match '(?:^|\|)GeneratedDesign=True(?:\||$)' -and
                        $productionWorkbenchReport -match '(?:^|\|)ItemCodeRemoved=True(?:\||$)' -and
                        $productionWorkbenchReport -match '(?:^|\|)Assignments=True(?:\||$)' -and
                        $productionWorkbenchReport -match '(?:^|\|)ItemSearch=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionWorkbenchPassed
                    $observedText += " || PRODUCTION_PROCESS_WORKBENCH=" + $productionWorkbenchReport

                    [IO.File]::WriteAllText($progressPath, "invoke Production released Process edit and export")
                    $productionEditExportReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunProcessEditExportContractTest")
                    $productionEditExportPassed = $productionEditExportReport -match '^OK\|' -and
                        $productionEditExportReport -match '(?:^|\|)ReleasedProcessEditable=True(?:\||$)' -and
                        $productionEditExportReport -match '(?:^|\|)ExistingProcessExported=True(?:\||$)' -and
                        $productionEditExportReport -match '(?:^|\|)ExportRoundTrip=True(?:\||$)' -and
                        $productionEditExportReport -match '(?:^|\|)OutputDesignVersionRebased=True(?:\||$)' -and
                        $productionEditExportReport -match '(?:^|\|)OutputYieldRebased=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionEditExportPassed
                    $observedText += " || PRODUCTION_PROCESS_EDIT_EXPORT=" + $productionEditExportReport

                    [IO.File]::WriteAllText($progressPath, "invoke Production Process worksheet bulk import")
                    $productionBulkImportReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunProcessWorksheetBulkImportContractTest")
                    $productionBulkImportPassed = $productionBulkImportReport -match '^OK\|' -and
                        $productionBulkImportReport -match '(?:^|\|)TextSafeIds=True(?:\||$)' -and
                        $productionBulkImportReport -match '(?:^|\|)RequirementIds=True(?:\||$)' -and
                        $productionBulkImportReport -match '(?:^|\|)UomCatalog=True(?:\||$)' -and
                        $productionBulkImportReport -match '(?:^|\|)NumberedAlternatives=True(?:\||$)' -and
                        $productionBulkImportReport -match '(?:^|\|)AddedAlternative=True(?:\||$)' -and
                        $productionBulkImportReport -match '(?:^|\|)PickerOpened=True(?:\||$)' -and
                        $productionBulkImportReport -match '(?:^|\|)PickerInventoryRows=True(?:\||$)' -and
                        $productionBulkImportReport -match '(?:^|\|)MultiAreaSelection=True(?:\||$)' -and
                        $productionBulkImportReport -match '(?:^|\|)MultiTableDrafts=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionBulkImportPassed
                    $observedText += " || PRODUCTION_PROCESS_BULK_IMPORT=" + $productionBulkImportReport

                    [IO.File]::WriteAllText($progressPath, "invoke Production Process worksheet output picker")
                    $productionOutputPickerReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunProcessWorksheetOutputPickerContractTest")
                    $productionOutputPickerPassed = $productionOutputPickerReport -match '^OK\|' -and
                        $productionOutputPickerReport -match '(?:^|\|)OutputPickerOpened=True(?:\||$)' -and
                        $productionOutputPickerReport -match '(?:^|\|)OutputPickerCommitted=True(?:\||$)' -and
                        $productionOutputPickerReport -match '(?:^|\|)OutputSkuHidden=True(?:\||$)' -and
                        $productionOutputPickerReport -match '(?:^|\|)OutputSkuRoundTrip=True(?:\||$)' -and
                        $productionOutputPickerReport -match '(?:^|\|)OutputNameRetained=True(?:\||$)' -and
                        $productionOutputPickerReport -match '(?:^|\|)OutputNamePickerSuppressed=True(?:\||$)' -and
                        $productionOutputPickerReport -match '(?:^|\|)UniqueRowIds=True(?:\||$)' -and
                        $productionOutputPickerReport -match '(?:^|\|)FirstAssignedIdRetained=True(?:\||$)' -and
                        $productionOutputPickerReport -match '(?:^|\|)NoPhysicalKey=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionOutputPickerPassed
                    $observedText += " || PRODUCTION_PROCESS_OUTPUT_PICKER=" + $productionOutputPickerReport

                    [IO.File]::WriteAllText($progressPath, "invoke Production reusable lifecycle actions")
                    $productionActionReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunReusableProductionFormActionContractTest")
                    $productionActionPassed = $productionActionReport -match '^OK\|' -and
                        $productionActionReport -match '(?:^|\|)ProcessSaved=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)ProcessReleased=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)ProcessObsoleted=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)ProcessReused=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeConnected=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeOrdered=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeSaved=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeReleased=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeObsoleted=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeIdGenerated=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeVersionGenerated=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeIdLocked=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeVersionEditable=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)EditedRecipeVersionRetained=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeOutputNameVisible=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeOutputIdPreserved=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeUomCatalog=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeConnectionUpdated=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeNodeNamesVisible=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeRequirementNameVisible=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeConnectionNamesVisible=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeConnectionHeaders=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeConnectionsFullWidth=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeFinishedOutputGuidance=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeConnectionSelected=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeDisconnected=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeSelfReferenceRejected=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeOutputFirstRouting=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeCompatibleTargetsOnly=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeRequirementInternallyBound=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeNoIngredientDropdown=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeForkConvergenceVisible=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeTerminalOutputVisible=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)RecipeStagesDerived=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)AlternativesSaved=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)OutputYieldDefaults=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)OutputFlowUsesProcessYield=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)ProcessAssignmentHeaders=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)AssignmentSystemKeyReadable=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)AcceptableItemsNamed=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)ProcessOutputEditorCompact=True(?:\||$)' -and
                        $productionActionReport -match '(?:^|\|)ProcessOutputUomCatalog=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionActionPassed
                    $observedText += " || PRODUCTION_REUSABLE_ACTIONS=" + $productionActionReport
                    }

                    [IO.File]::WriteAllText($progressPath, "invoke Production reusable run actions")
                    $productionRunReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunReusableProductionRunActionContractTest")
                    $productionRunPassed = $productionRunReport -match '^OK\|' -and
                        $productionRunReport -match '(?:^|\|)Batches=2(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)ScaleMin=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)ScaleDefault=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)ScaleMax=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)ExactInputKeys=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)InsufficiencyRejected=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)StaleRejected=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)DistinctOutputKeys=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)IntermediateConsumed=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)CoProductRemaining=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)PercentageYieldBasis=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)ActualOutputAccepted=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)LastActualDisplayed=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)ActualInventoryQty=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)SystemKeyHeadersReadable=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)BatchHistoryRows=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)ProcessTotal=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)UtilityDisplay=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)MultiProcessRunPlan=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)TargetOutputScaleStub=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)LocationStockBuckets=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)LocationStockExactExpansion=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)SelectedProcessOnly=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)RunInstructionsVisible=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)WholeRecipeStatus=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)EightPaletteRows=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)UpstreamWaitThenReady=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)RoutedInputVisible=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)RoutedInputDetails=True(?:\||$)' -and
                        $productionRunReport -match '(?:^|\|)GroupedUsedGoods=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $productionRunPassed
                    $observedText += " || PRODUCTION_REUSABLE_RUN=" + $productionRunReport
                    if ($productionRunReport -match '(?:^|\|)ReusableRecipe=([^|]+)') {
                        $reusableRecipeId = [string]$Matches[1]
                    }

                    [IO.File]::WriteAllText($progressPath, "invoke Production Chai fork/convergence actions")
                    $chaiRunReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunChaiForkConvergenceRunActionContractTest")
                    $chaiRunPassed = $chaiRunReport -match '^OK\|' -and
                        $chaiRunReport -match '(?:^|\|)ChaiInitialWaitingUpstream=True(?:\||$)' -and
                        $chaiRunReport -match '(?:^|\|)ChaiBatchNoteFrozen=True(?:\||$)' -and
                        $chaiRunReport -match '(?:^|\|)ChaiBatchNoteRetained=True(?:\||$)' -and
                        $chaiRunReport -match '(?:^|\|)ChaiRoutedConvergenceInputs=True(?:\||$)' -and
                        $chaiRunReport -match '(?:^|\|)ChaiUpstreamExactKeysConsumed=True(?:\||$)' -and
                        $chaiRunReport -match '(?:^|\|)ChaiFinalBottlingRoutedInput=True(?:\||$)' -and
                        $chaiRunReport -match '(?:^|\|)ChaiFinalOutputNewKey=True(?:\||$)' -and
                        $chaiRunReport -match '(?:^|\|)ChaiFourProcessesCompleted=True(?:\||$)' -and
                        $chaiRunReport -match '(?:^|\|)ChaiFinalBottlingCompleted=True(?:\||$)' -and
                        $chaiRunReport -match '(?:^|\|)ChaiRunNotRestarted=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $chaiRunPassed
                    $observedText += " || PRODUCTION_CHAI_FORK_CONVERGENCE=" + $chaiRunReport

                    if (-not $ProductionRunOnly) {
                    [IO.File]::WriteAllText($progressPath, "invoke Production Process worksheet round-trip")
                    $processWorksheetReport = [string](Run-WorkbookMacro -Excel $excel `
                        -WorkbookName $operationsName `
                        -MacroName "mProduction.RunProcessWorksheetRoundTripContractTest")
                    $processWorksheetPassed = $processWorksheetReport -match '^OK\|' -and
                        $processWorksheetReport -match '(?:^|\|)RecipeIdentityInitialized=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)ProcessIdGenerated=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)RecipeIdGenerated=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)RecipeVersionGenerated=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)RecipeIdLocked=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)RecipeVersionEditable=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)RequirementIdGenerated=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)OutputIdGenerated=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)IdentityControlsLocked=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)WorksheetHandler=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)MixedUomAccepted=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)MixedUomRowsPreserved=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)TableRemoved=True(?:\||$)' -and
                        $processWorksheetReport -match '(?:^|\|)RepeatRoundTrip=True(?:\||$)'
                    $workflowControlPassed = $workflowControlPassed -and $processWorksheetPassed
                    $observedText += " || PRODUCTION_PROCESS_WORKSHEET=" + $processWorksheetReport
                    }
                    }
                }
            }
            if (@($secondCapture.WindowText).Count -gt 0) {
                $observedText += " || SECOND_LAUNCH=" + (@($secondCapture.WindowText) -join " // ")
            }
            $stationLocalCount = @($newWorkbookPaths | Where-Object {
                [IO.Path]::GetFullPath($_).StartsWith(
                    [IO.Path]::GetFullPath($operatorRoot),
                    [StringComparison]::OrdinalIgnoreCase
                )
            }).Count
            if ($WorkbookState -eq "ProductionReusable" -and $stationLocalCount -eq 1) {
                $productionOperatorPath = [string](@($newWorkbookPaths | Where-Object {
                    [IO.Path]::GetFullPath($_).StartsWith(
                        [IO.Path]::GetFullPath($operatorRoot),
                        [StringComparison]::OrdinalIgnoreCase
                    )
                })[0])
            }
            $passed = $operatorRootOverrideSet -and
                $stationLocalCount -eq 1 -and
                $secondNewWorkbooks.Count -eq 0 -and
                $workflowControlPassed -and
                [string]::IsNullOrWhiteSpace($capture.Error) -and
                [string]::IsNullOrWhiteSpace($secondCapture.Error) -and
                $observedText -notmatch "(?i)(failed|Type mismatch|operator workbook)"
        }
        elseif ($WorkbookState -in @("ConfigActive", "UnrelatedActive")) {
            $stationLocalCount = @($newWorkbookPaths | Where-Object {
                [IO.Path]::GetFullPath($_).StartsWith(
                    [IO.Path]::GetFullPath($operatorRoot),
                    [StringComparison]::OrdinalIgnoreCase
                )
            }).Count
            $passed = $operatorRootOverrideSet -and
                -not $surfaceChanged -and
                $stationLocalCount -eq 1 -and
                $observedText -notmatch "(?i)(failed|Type mismatch|operator workbook)"
        }
        elseif ($WorkbookState -in @("SavedEligible", "CapturedClosed")) {
            $passed = [string]::IsNullOrWhiteSpace($capture.Error) -and
                $observedText -notmatch "(?i)(failed|Type mismatch|operator workbook)"
        }

        if ($WorkbookState -eq "CapturedClosed") {
            $initialLaunchPassed = $passed
            $stateWb.Close($false)
            $recoveryPath = Join-Path $runtimeRoot "Plan022.AfterCapturedClose.Unrelated.xlsb"
            $recoveryWb = New-OperationalWorkbook -Excel $excel `
                -NameHint ([IO.Path]::GetFileName($recoveryPath)) -Path $recoveryPath
            $opened.Add($recoveryWb) | Out-Null
            $recoveryWb.Activate()
            $recoveryBefore = Get-WorkbookSurfaceCounts -Workbook $recoveryWb
            $recoveryNamesBefore = @(Get-OpenWorkbookNames -Excel $excel)
            $recoveryCapture = Invoke-PackagedCallback -Excel $excel `
                -WorkbookName $operationsName -MacroName $callback.Macro
            $recoveryNamesAfter = @(Get-OpenWorkbookNames -Excel $excel)
            $recoveryAfter = Get-WorkbookSurfaceCounts -Workbook $recoveryWb
            $recoveryNewNames = @($recoveryNamesAfter | Where-Object { $_ -notin $recoveryNamesBefore })
            $recoveryText = (@($recoveryCapture.WindowText) -join " // ")
            $recoverySurfaceChanged = $recoveryBefore.Worksheets -ne $recoveryAfter.Worksheets -or
                $recoveryBefore.Tables -ne $recoveryAfter.Tables
            $observedText += " || AFTER_CAPTURED_CLOSE=" + $recoveryText
            $observedText += " || RECOVERY_SURFACE_CHANGED=" + $recoverySurfaceChanged
            $observedText += " || RECOVERY_NEW_WORKBOOKS=" + ($recoveryNewNames -join ",")

            $recoveryLocalCount = 0
            foreach ($recoveryName in $recoveryNewNames) {
                try {
                    $recoveryFullName = [string]$excel.Workbooks.Item($recoveryName).FullName
                    If ([IO.Path]::GetFullPath($recoveryFullName).StartsWith(
                            [IO.Path]::GetFullPath($operatorRoot),
                            [StringComparison]::OrdinalIgnoreCase)) {
                        $recoveryLocalCount += 1
                    }
                }
                catch {}
            }
            $passed = $initialLaunchPassed -and
                -not $recoverySurfaceChanged -and
                $recoveryLocalCount -eq 1 -and
                $recoveryText -notmatch "(?i)(failed|Type mismatch|operator workbook)"
        }
        Add-Evidence -Rows $evidence -Callback $callback.Macro `
            -Expected $callback.Expected -Passed $passed -Observed $observedText
        [IO.File]::WriteAllText($progressPath, "completed " + $callback.Macro)
    }

    if ($WorkbookState -eq "ProductionReusable" -and -not $ProductionRunOnly -and
        -not $ProductionEditExportOnly) {
        $currentStep = "restart reusable Production in a clean Excel process"
        [IO.File]::WriteAllText($progressPath, $currentStep)
        if ([string]::IsNullOrWhiteSpace($reusableRecipeId)) {
            throw "The first Production session did not report its released reusable Recipe identity."
        }
        if ([string]::IsNullOrWhiteSpace($productionOperatorPath) -or
            -not (Test-Path -LiteralPath $productionOperatorPath -PathType Leaf)) {
            throw "The first Production session did not create its saved station-local operator workbook."
        }

        for ($workbookIndex = [int]$excel.Workbooks.Count; $workbookIndex -ge 1; $workbookIndex--) {
            $restartWb = $null
            try {
                $restartWb = $excel.Workbooks.Item($workbookIndex)
                if (-not [string]::IsNullOrWhiteSpace([string]$restartWb.Path) -and
                    -not [bool]$restartWb.IsAddin) {
                    $restartWb.Save()
                }
                $restartWb.Close($false)
            }
            catch {}
            Release-ComObject $restartWb
        }
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
        $excel = $null
        if ($excelProcessId -gt 0) {
            Start-Sleep -Milliseconds 750
            $firstProcess = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
            if ($null -ne $firstProcess -and $firstProcess.ProcessName -eq "EXCEL") {
                Stop-Process -Id $excelProcessId -Force
            }
        }

        $opened = New-Object 'System.Collections.Generic.List[object]'
        $packages = @{}
        $excel = New-Object -ComObject Excel.Application
        $excelProcessId = Get-ExcelProcessId $excel
        $excel.Visible = $true
        $excel.DisplayAlerts = $false
        $excel.EnableEvents = $true
        $excel.AutomationSecurity = 1
        foreach ($packageName in $packageNames) {
            $packagePath = Join-Path $deployPath $packageName
            $packageWb = $excel.Workbooks.Open($packagePath)
            $opened.Add($packageWb) | Out-Null
            $packages[$packageName] = $packageWb
        }

        $coreName = [string]$packages["invSys.Core.xlam"].Name
        $operationsName = [string]$packages["invSys.Operations.xlam"].Name
        [void](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
            -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($runtimeRoot))
        $restartConfigLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
            -MacroName "modConfig.LoadConfig" -Arguments @($warehouseId, $stationId))
        $restartAuthLoaded = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
            -MacroName "modAuth.LoadAuth" -Arguments @($warehouseId))
        $restartTargetResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
            -MacroName "modNasConnection.SelectWarehouseTargetForAutomation" `
            -Arguments @($runtimeRoot, $runtimeRoot, $stationId, $true))
        $restartTargetPathsSet = [bool](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
            -MacroName "modNasConnection.SetCurrentTargetPathsForTest" `
            -Arguments @($testHubRoot, $runtimeRoot))
        $restartOperatorOverrideSet = [bool](Run-WorkbookMacro -Excel $excel `
            -WorkbookName $coreName `
            -MacroName "modWarehouseBootstrap.SetLocalOperatorRootOverrideForAutomation" `
            -Arguments @($operatorRoot))
        $restartSignInResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreName `
            -MacroName "modAuth.SignInCurrentTargetForAutomation" `
            -Arguments @($testUser, $testPin, "PROD_POST"))
        [void](Run-WorkbookMacro -Excel $excel -WorkbookName $operationsName `
            -MacroName "modOperationsInit.Auto_Open")

        $restartBeforeNames = @(Get-OpenWorkbookNames -Excel $excel)
        $restartCapture = Invoke-PackagedCallback -Excel $excel `
            -WorkbookName $operationsName -MacroName "mProduction.BtnOpenProductionForm"
        $restartAfterNames = @(Get-OpenWorkbookNames -Excel $excel)
        $restartNewWorkbooks = @(
            $restartAfterNames | Where-Object { $_ -notin $restartBeforeNames }
        )
        $restartNewOperatorWorkbooks = New-Object 'System.Collections.Generic.List[string]'
        $restartOperatorFullName = ""
        foreach ($restartName in $restartNewWorkbooks) {
            try {
                $candidateFullName = [string]$excel.Workbooks.Item($restartName).FullName
                if ([IO.Path]::GetFullPath($candidateFullName).StartsWith(
                        [IO.Path]::GetFullPath($operatorRoot),
                        [StringComparison]::OrdinalIgnoreCase)) {
                    $restartNewOperatorWorkbooks.Add($candidateFullName) | Out-Null
                    $restartOperatorFullName = $candidateFullName
                }
            }
            catch {}
        }
        $restartSameOperatorWorkbook = -not [string]::IsNullOrWhiteSpace($restartOperatorFullName) -and
            [string]::Equals(
                [IO.Path]::GetFullPath($restartOperatorFullName),
                [IO.Path]::GetFullPath($productionOperatorPath),
                [StringComparison]::OrdinalIgnoreCase
            )
        $probeBeforeNames = @(Get-OpenWorkbookNames -Excel $excel)
        $productionRestartReport = [string](Run-WorkbookMacro -Excel $excel `
            -WorkbookName $operationsName `
            -MacroName "mProduction.RunReusableProductionRestartActionContractTest" `
            -Arguments @($reusableRecipeId, $reusableRecipeVersion, $productionOperatorPath))
        $probeAfterNames = @(Get-OpenWorkbookNames -Excel $excel)
        $restartProbeNewWorkbooks = @(
            $probeAfterNames | Where-Object { $_ -notin $probeBeforeNames }
        ).Count
        $restartOperatorFileCount = @(
            Get-ChildItem -LiteralPath $operatorRoot `
                -Filter "*.Production.Operator.xlsm" -File -Recurse `
                -ErrorAction SilentlyContinue
        ).Count
        $restartPassed = $restartConfigLoaded -and $restartAuthLoaded -and
            $restartTargetResult.StartsWith('OK|') -and $restartTargetPathsSet -and
            $restartOperatorOverrideSet -and $restartSignInResult.StartsWith('OK|') -and
            [string]::IsNullOrWhiteSpace($restartCapture.Error) -and
            $restartNewOperatorWorkbooks.Count -eq 1 -and $restartSameOperatorWorkbook -and
            $restartOperatorFileCount -eq 1 -and $restartProbeNewWorkbooks -eq 0 -and
            $productionRestartReport -match '^OK\|' -and
            $productionRestartReport -match '(?:^|\|)RecipeFound=True(?:\||$)' -and
            $productionRestartReport -match '(?:^|\|)Loaded=True(?:\||$)' -and
            $productionRestartReport -match '(?:^|\|)SameWorkbook=True(?:\||$)' -and
            $productionRestartReport -match '(?:^|\|)WorksheetRediscovered=True(?:\||$)' -and
            $productionRestartReport -match '(?:^|\|)WorksheetRetrieved=True(?:\||$)' -and
            $productionRestartReport -match '(?:^|\|)MultipleTablesRediscovered=True(?:\||$)' -and
            $productionRestartReport -match '(?:^|\|)SelectedOnly=True(?:\||$)' -and
            $productionRestartReport -match '(?:^|\|)AllRetrieved=True(?:\||$)'
        $restartObserved = "PRODUCTION_RESTART=" + $productionRestartReport +
            " || RestartSameOperatorWorkbook=" + $restartSameOperatorWorkbook +
            " || RestartNewWorkbooks=" + $restartNewWorkbooks.Count +
            " || RestartNewOperatorWorkbooks=" + $restartNewOperatorWorkbooks.Count +
            " || ProbeNewWorkbooks=" + $restartProbeNewWorkbooks +
            " || OperatorFiles=" + $restartOperatorFileCount
        if (-not [string]::IsNullOrWhiteSpace($restartCapture.Error)) {
            $restartObserved += " || COM_ERROR=" + $restartCapture.Error
        }
        Add-Evidence -Rows $evidence `
            -Callback "mProduction.BtnOpenProductionForm [clean restart]" `
            -Expected "A new Excel process reuses the saved Production workbook and loads the persisted exact released Recipe through the Run List handler." `
            -Passed $restartPassed -Observed $restartObserved
    }
}
catch {
    Add-Evidence -Rows $evidence -Callback "HARNESS" `
        -Expected "The isolated packaged harness reaches the launcher callbacks." `
        -Passed $false -Observed (
            "Step=" + $currentStep + "; " + $_.Exception.Message +
            "; Location=" + ($_.InvocationInfo.PositionMessage -replace '\r?\n', ' ')
        )
}
finally {
    $reportLines = New-Object System.Collections.Generic.List[string]
    $reportTitle = if ($WorkbookState -eq "ReceivingDurability") {
        "# Plan 022 Slice 1 Receiving Durability Evidence"
    }
    elseif ($WorkbookState -eq "ShippingLayout") {
        "# Plan 022 Slices 4g-4h Shipping Layout Evidence"
    }
    elseif ($WorkbookState -eq "ProductionReusable" -and $ProductionRunOnly) {
        "# Plan 022 Slice 4av Focused Packaged Production Evidence"
    }
    elseif ($WorkbookState -eq "ProductionReusable" -and $ProductionEditExportOnly) {
        "# Plan 022 Slice 4aw Focused Packaged Production Evidence"
    }
    elseif ($WorkbookState -eq "ProductionReusable") {
        "# Plan 022 Slice 4x Packaged Reusable Production Evidence"
    }
    else {
        "# Plan 022 Slice 0 Packaged Launcher Evidence"
    }
    $reportLines.Add($reportTitle) | Out-Null
    $reportLines.Add("") | Out-Null
    $reportLines.Add("- Captured: $([DateTime]::UtcNow.ToString('yyyy-MM-ddTHH:mm:ssZ'))") | Out-Null
    $reportLines.Add("- Runtime: isolated temporary test runtime (not NAS acceptance)") | Out-Null
    $reportLines.Add("- Package source: $DeployRoot") | Out-Null
    $reportLines.Add("- Workbook state: connected and signed in; $WorkbookState") | Out-Null
    foreach ($line in $setupEvidence) {
        $reportLines.Add("- $line") | Out-Null
    }
    $reportLines.Add("") | Out-Null
    $reportLines.Add("| Callback | Result | Expected | Observed |") | Out-Null
    $reportLines.Add("|---|---|---|---|") | Out-Null
    foreach ($row in $evidence) {
        $result = if ($row.Passed) { "PASS" } else { "RED" }
        $expected = ([string]$row.Expected).Replace("|", "/")
        $observed = ([string]$row.Observed).Replace("|", "/")
        $reportLines.Add("| $($row.Callback) | $result | $expected | $observed |") | Out-Null
    }
    $reportFile = if ($WorkbookState -eq "NoEligible") {
        "packaged-launcher-noeligible.md"
    }
    elseif ($WorkbookState -eq "ShippingLayout") {
        "shipping-layout.md"
    }
    elseif ($WorkbookState -eq "CapturedClosed") {
        "packaged-launcher-capturedclosed-$($CallbackFilter.ToLowerInvariant()).md"
    }
    elseif ($WorkbookState -eq "ReceivingDurability") {
        "receiving-launcher-durability.md"
    }
    elseif ($WorkbookState -eq "ReceivingFormClosed") {
        "receiving-launcher-formclosed.md"
    }
    elseif ($WorkbookState -eq "ProductionReusable") {
        "production-reusable-production.md"
    }
    else {
        "packaged-launcher-$($WorkbookState.ToLowerInvariant()).md"
    }
    $reportPath = Join-Path $outputPath $reportFile
    [IO.File]::WriteAllLines($reportPath, $reportLines)

    foreach ($wb in $opened) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
    if ($excelProcessId -gt 0) {
        Start-Sleep -Milliseconds 500
        $ownedExcelProcess = Get-Process -Id $excelProcessId -ErrorAction SilentlyContinue
        if ($null -ne $ownedExcelProcess -and $ownedExcelProcess.ProcessName -eq "EXCEL") {
            Stop-Process -Id $excelProcessId -Force
        }
    }

    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
    $resolvedRuntime = [IO.Path]::GetFullPath($runtimeRoot)
    if ($resolvedRuntime.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase) -and
        (Split-Path -Leaf $resolvedRuntime) -like "invsys-plan022-launcher-red-*") {
        Remove-Item -LiteralPath $resolvedRuntime -Recurse -Force -ErrorAction SilentlyContinue
    }
}

$redCount = @($evidence | Where-Object { -not $_.Passed }).Count
if ($redCount -eq 0) {
    Write-Output "PLAN022_PACKAGED_LAUNCHER_GREEN"
}
else {
    Write-Output "PLAN022_PACKAGED_LAUNCHER_RED"
}
Write-Output "REPORT=$reportPath"
Write-Output "PASS=$($evidence.Count - $redCount) RED=$redCount TOTAL=$($evidence.Count)"
if ($redCount -eq 0) {
    exit 0
}
exit 1
