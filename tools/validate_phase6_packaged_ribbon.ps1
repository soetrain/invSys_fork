[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$DeployRoot = "deploy/current"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
Add-Type -AssemblyName System.IO.Compression.FileSystem

function Release-ComObject {
    param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
    }
}

function Add-ResultRow {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [string]$Check,
        [bool]$Passed,
        [string]$Detail = ""
    )

    $Rows.Add([pscustomobject]@{
        Check  = $Check
        Passed = $Passed
        Detail = $Detail
    }) | Out-Null
}

function Get-CustomUiXml {
    param([string]$WorkbookPath)

    $tempCopy = Join-Path ([System.IO.Path]::GetTempPath()) ([System.IO.Path]::GetFileNameWithoutExtension($WorkbookPath) + "-" + [guid]::NewGuid().ToString("N") + ".xlam")
    Copy-Item -LiteralPath $WorkbookPath -Destination $tempCopy -Force

    $zip = $null
    try {
        $zip = [System.IO.Compression.ZipFile]::OpenRead($tempCopy)
        $entry = $zip.Entries | Where-Object { $_.FullName -ieq "customUI/customUI.xml" } | Select-Object -First 1
        if ($null -eq $entry) { return $null }

        $reader = New-Object System.IO.StreamReader($entry.Open())
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        if ($null -ne $zip) { $zip.Dispose() }
        Remove-Item -LiteralPath $tempCopy -Force -ErrorAction SilentlyContinue
    }
}

function Get-RibbonButtons {
    param([string]$CustomUiXml)

    if ([string]::IsNullOrWhiteSpace($CustomUiXml)) { return @{} }

    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($CustomUiXml)

    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("cu", "http://schemas.microsoft.com/office/2006/01/customui")

    $result = @{}
    $buttons = $doc.SelectNodes("//cu:button", $ns)
    foreach ($button in $buttons) {
        $screentip = ""
        if ($null -ne $button.Attributes["screentip"]) {
            $screentip = [string]$button.Attributes["screentip"].Value
        }
        $result[$button.id] = [pscustomobject]@{
            Id        = [string]$button.id
            Label     = [string]$button.label
            OnAction  = [string]$button.onAction
            Screentip = $screentip
        }
    }

    return $result
}

function Get-ModuleText {
    param(
        [object]$Workbook,
        [string]$ComponentName
    )

    try {
        $component = $Workbook.VBProject.VBComponents.Item($ComponentName)
        $module = $component.CodeModule
        if ($module.CountOfLines -le 0) { return "" }
        return $module.Lines(1, $module.CountOfLines)
    }
    catch {
        return ""
    }
}

function Test-ProcedureExists {
    param(
        [object]$Workbook,
        [string]$QualifiedMacroName
    )

    $parts = $QualifiedMacroName.Split(".")
    if ($parts.Count -ne 2) { return $false }

    $componentName = $parts[0]
    $procedureName = $parts[1]
    $moduleText = Get-ModuleText -Workbook $Workbook -ComponentName $componentName
    if ([string]::IsNullOrWhiteSpace($moduleText)) { return $false }

    $pattern = "(?im)^\s*(Public|Private)?\s*(Sub|Function)\s+" + [regex]::Escape($procedureName) + "\b"
    return [regex]::IsMatch($moduleText, $pattern)
}

function Run-WorkbookMacro {
    param(
        [object]$Excel,
        [string]$WorkbookName,
        [string]$MacroName,
        [object[]]$Arguments = @()
    )

    $fullMacro = "'$WorkbookName'!$MacroName"
    switch ($Arguments.Count) {
        0 { return $Excel.Run($fullMacro) }
        1 { return $Excel.Run($fullMacro, $Arguments[0]) }
        2 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1]) }
        3 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2]) }
        4 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3]) }
        5 { return $Excel.Run($fullMacro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4]) }
        default { throw "Run-WorkbookMacro supports at most 5 arguments." }
    }
}

$repo = (Resolve-Path $RepoRoot).Path
$deployPath = Join-Path $repo $DeployRoot
$resultPath = Join-Path $repo "tests/unit/phase6_packaged_ribbon_results.md"
$runtimeRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("invsys-phase6-ribbon-" + [guid]::NewGuid().ToString("N"))

$openOrder = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Receiving.xlam",
    "invSys.Shipping.xlam",
    "invSys.Production.xlam",
    "invSys.Admin.xlam"
)

$ribbonSpecs = @(
    @{
        Name = "Receiving"
        File = "invSys.Receiving.xlam"
        Callback = "RibbonOnActionReceiving"
        Buttons = @(
            @{ Id = "btnReceivingSetup"; Label = "Setup UI"; Macro = "modTS_Received.EnsureGeneratedButtons"; Execute = $true },
            @{ Id = "btnReceivingConfirm"; Label = "Confirm Writes"; Macro = "modTS_Received.ConfirmWrites"; Execute = $false },
            @{ Id = "btnReceivingUndo"; Label = "Undo"; Macro = "modTS_Received.MacroUndo"; Execute = $false },
            @{ Id = "btnReceivingRedo"; Label = "Redo"; Macro = "modTS_Received.MacroRedo"; Execute = $false }
        )
    }
    @{
        Name = "Shipping"
        File = "invSys.Shipping.xlam"
        Callback = "RibbonOnActionShipping"
        Buttons = @(
            @{ Id = "btnShippingSetup"; Label = "Setup UI"; Macro = "modTS_Shipments.InitializeShipmentsUI"; Execute = $true },
            @{ Id = "btnShippingConfirm"; Label = "Confirm Inventory"; Macro = "modTS_Shipments.BtnConfirmInventory"; Execute = $false },
            @{ Id = "btnShippingStage"; Label = "To Shipments"; Macro = "modTS_Shipments.BtnToShipments"; Execute = $false },
            @{ Id = "btnShippingSend"; Label = "Shipments Sent"; Macro = "modTS_Shipments.BtnShipmentsSent"; Execute = $false }
        )
    }
    @{
        Name = "Production"
        File = "invSys.Production.xlam"
        Callback = "RibbonOnActionProduction"
        Buttons = @(
            @{ Id = "btnProductionSetup"; Label = "Setup UI"; Macro = "mProduction.InitializeProductionUI"; Execute = $true },
            @{ Id = "btnProductionLoad"; Label = "Load Recipe"; Macro = "mProduction.BtnLoadRecipe"; Execute = $false },
            @{ Id = "btnProductionUsed"; Label = "To Used"; Macro = "mProduction.BtnToUsed"; Execute = $false },
            @{ Id = "btnProductionMade"; Label = "To Made"; Macro = "mProduction.BtnToMade"; Execute = $false },
            @{ Id = "btnProductionTotal"; Label = "To Total Inv"; Macro = "mProduction.BtnToTotalInv"; Execute = $false },
            @{ Id = "btnProductionPrintCodes"; Label = "Print Recall Codes"; Macro = "mProduction.BtnPrintRecallCodes"; Execute = $false }
        )
    }
    @{
        Name = "Admin"
        File = "invSys.Admin.xlam"
        Callback = "RibbonOnActionAdmin"
        Buttons = @(
            @{ Id = "btnAdminOpen"; Label = "Admin Console"; Macro = "modAdmin.Admin_Click"; Execute = $true },
            @{ Id = "btnAdminUsers"; Label = "Users and Roles"; Macro = "modAdmin.Open_CreateDeleteUser"; Execute = $true },
            @{ Id = "btnAdminCreateWarehouse"; Label = "Create New Warehouse"; Macro = "modAdmin.Open_CreateWarehouse"; Execute = $false },
            @{ Id = "btnAdminRetireMigrateWarehouse"; Label = "Retire / Migrate Warehouse"; Macro = "modAdmin.Admin_RetireMigrateWarehouse_Click"; Execute = $false; Screentip = "Archive, migrate, retire, or delete a warehouse runtime" }
        )
    }
)

$resultRows = New-Object 'System.Collections.Generic.List[object]'
$excel = $null
$openedWorkbooks = New-Object 'System.Collections.Generic.List[object]'
$workbookMap = @{}

try {
    New-Item -ItemType Directory -Path $runtimeRoot -Force | Out-Null

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $true
    $excel.AutomationSecurity = 1

    foreach ($fileName in $openOrder) {
        $path = Join-Path $deployPath $fileName
        if (-not (Test-Path -LiteralPath $path)) {
            Add-ResultRow -Rows $resultRows -Check "$fileName.Open" -Passed $false -Detail "Missing packaged XLAM: $path"
            continue
        }

        try {
            $wb = $excel.Workbooks.Open($path)
            $openedWorkbooks.Add($wb) | Out-Null
            $workbookMap[$fileName] = $wb
            Add-ResultRow -Rows $resultRows -Check "$fileName.Open" -Passed $true -Detail "Opened from $path"
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "$fileName.Open" -Passed $false -Detail $_.Exception.Message
        }
    }

    if ($workbookMap.ContainsKey("invSys.Core.xlam")) {
        try {
            [void](Run-WorkbookMacro -Excel $excel -WorkbookName $workbookMap["invSys.Core.xlam"].Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($runtimeRoot))
            Add-ResultRow -Rows $resultRows -Check "Core.RuntimeRootOverride" -Passed $true -Detail $runtimeRoot
        }
        catch {
            Add-ResultRow -Rows $resultRows -Check "Core.RuntimeRootOverride" -Passed $false -Detail $_.Exception.Message
        }
    }

    foreach ($spec in $ribbonSpecs) {
        $fileName = $spec.File
        if (-not $workbookMap.ContainsKey($fileName)) {
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonXml" -Passed $false -Detail "Workbook not open."
            continue
        }

        $wb = $workbookMap[$fileName]
        $path = Join-Path $deployPath $fileName

        $customUiXml = Get-CustomUiXml -WorkbookPath $path
        if ([string]::IsNullOrWhiteSpace($customUiXml)) {
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonXml" -Passed $false -Detail "customUI/customUI.xml missing."
            continue
        }

        Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonXml" -Passed $true -Detail "customUI/customUI.xml present."
        $buttons = Get-RibbonButtons -CustomUiXml $customUiXml
        $callbackModuleText = Get-ModuleText -Workbook $wb -ComponentName "modRibbonGenerated"
        Add-ResultRow -Rows $resultRows -Check "$($spec.Name).CallbackModule" -Passed (-not [string]::IsNullOrWhiteSpace($callbackModuleText)) -Detail "modRibbonGenerated"

        foreach ($button in $spec.Buttons) {
            $buttonId = [string]$button.Id
            if ($buttons.ContainsKey($buttonId)) {
                $buttonInfo = $buttons[$buttonId]
                $detail = "Label=$($buttonInfo.Label); OnAction=$($buttonInfo.OnAction); Screentip=$($buttonInfo.Screentip)"
                $passed = ($buttonInfo.Label -eq $button.Label -and $buttonInfo.OnAction -eq $spec.Callback)
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonButton.$buttonId" -Passed $passed -Detail $detail
                if ($button.ContainsKey("Screentip")) {
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonButtonScreentip.$buttonId" -Passed ($buttonInfo.Screentip -eq $button.Screentip) -Detail $buttonInfo.Screentip
                }
            }
            else {
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonButton.$buttonId" -Passed $false -Detail "Button missing from Ribbon XML."
                if ($button.ContainsKey("Screentip")) {
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonButtonScreentip.$buttonId" -Passed $false -Detail "Button missing from Ribbon XML."
                }
            }

            $macroName = [string]$button.Macro
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).MacroExists.$buttonId" -Passed (Test-ProcedureExists -Workbook $wb -QualifiedMacroName $macroName) -Detail $macroName

            $callbackHasButton = (-not [string]::IsNullOrWhiteSpace($callbackModuleText)) -and $callbackModuleText.Contains($buttonId)
            $callbackHasMacro = (-not [string]::IsNullOrWhiteSpace($callbackModuleText)) -and $callbackModuleText.Contains($macroName)
            Add-ResultRow -Rows $resultRows -Check "$($spec.Name).CallbackMap.$buttonId" -Passed ($callbackHasButton -and $callbackHasMacro) -Detail "$buttonId -> $macroName"

            if ($button.Execute) {
                try {
                    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $wb.Name -MacroName $macroName)
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).SafeExec.$buttonId" -Passed $true -Detail $macroName
                }
                catch {
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).SafeExec.$buttonId" -Passed $false -Detail $_.Exception.Message
                }
            }
        }
    }
}
finally {
    $failedCount = @($resultRows | Where-Object { -not $_.Passed }).Count
    $passedCount = $resultRows.Count - $failedCount

    $lines = @()
    $lines += "# Phase 6 Packaged Ribbon Validation Results"
    $lines += ""
    $lines += "- Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $lines += "- Deploy root: $deployPath"
    $lines += "- Runtime root override: $runtimeRoot"
    $lines += "- Passed: $passedCount"
    $lines += "- Failed: $failedCount"
    $lines += ""
    $lines += "| Check | Result | Detail |"
    $lines += "|---|---|---|"
    foreach ($row in $resultRows) {
        $result = if ($row.Passed) { "PASS" } else { "FAIL" }
        $detail = [string]$row.Detail
        $detail = $detail.Replace("|", "/")
        $lines += "| $($row.Check) | $result | $detail |"
    }
    [System.IO.File]::WriteAllLines($resultPath, $lines)

    foreach ($wb in $openedWorkbooks) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }

    Remove-Item -LiteralPath $runtimeRoot -Recurse -Force -ErrorAction SilentlyContinue
}

$failed = @($resultRows | Where-Object { -not $_.Passed }).Count
if ($failed -gt 0) {
    Write-Output "PHASE6_PACKAGED_RIBBON_VALIDATION_FAILED"
    Write-Output "RESULTS=$resultPath"
    Write-Output "PASSED=$($resultRows.Count - $failed) FAILED=$failed TOTAL=$($resultRows.Count)"
    exit 1
}

Write-Output "PHASE6_PACKAGED_RIBBON_VALIDATION_OK"
Write-Output "RESULTS=$resultPath"
Write-Output "PASSED=$($resultRows.Count) FAILED=0 TOTAL=$($resultRows.Count)"
