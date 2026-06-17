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
        $label = ""
        if ($null -ne $button.Attributes["label"]) {
            $label = [string]$button.Attributes["label"].Value
        }
        $getLabel = ""
        if ($null -ne $button.Attributes["getLabel"]) {
            $getLabel = [string]$button.Attributes["getLabel"].Value
        }
        $onAction = ""
        if ($null -ne $button.Attributes["onAction"]) {
            $onAction = [string]$button.Attributes["onAction"].Value
        }
        $screentip = ""
        if ($null -ne $button.Attributes["screentip"]) {
            $screentip = [string]$button.Attributes["screentip"].Value
        }
        $getEnabled = ""
        if ($null -ne $button.Attributes["getEnabled"]) {
            $getEnabled = [string]$button.Attributes["getEnabled"].Value
        }
        $result[$button.id] = [pscustomobject]@{
            Id         = [string]$button.id
            Label      = $label
            GetLabel   = $getLabel
            OnAction   = $onAction
            Screentip  = $screentip
            GetEnabled = $getEnabled
        }
    }

    return $result
}

function Get-RibbonLabelControls {
    param([string]$CustomUiXml)

    if ([string]::IsNullOrWhiteSpace($CustomUiXml)) { return @{} }

    $doc = New-Object System.Xml.XmlDocument
    $doc.LoadXml($CustomUiXml)

    $ns = New-Object System.Xml.XmlNamespaceManager($doc.NameTable)
    $ns.AddNamespace("cu", "http://schemas.microsoft.com/office/2006/01/customui")

    $result = @{}
    $labels = $doc.SelectNodes("//cu:labelControl", $ns)
    foreach ($label in $labels) {
        $getLabel = ""
        if ($null -ne $label.Attributes["getLabel"]) {
            $getLabel = [string]$label.Attributes["getLabel"].Value
        }
        $result[$label.id] = [pscustomobject]@{
            Id       = [string]$label.id
            GetLabel = $getLabel
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
        EnabledCallback = "RibbonRequiredCapabilityGetEnabledReceiving"
        StatusLabels = @(
            @{ Id = "lblReceivingServerStatus"; GetLabel = "RibbonServerStatusGetLabel" },
            @{ Id = "lblReceivingAccessStatus"; GetLabel = "RibbonAccessStatusGetLabel" }
        )
        Buttons = @(
            @{ Id = "btnReceivingConnectServer"; Label = "Connect Server"; DirectAction = 'modRoleEventWriter.ConnectWarehouseStorageForCapability "RECEIVE_POST"'; Execute = $false },
            @{ Id = "btnReceivingCurrentUser"; GetLabel = "RibbonCurrentUserGetLabel"; DirectAction = 'modRoleEventWriter.PromptSetCurrentUserForCapability "RECEIVE_POST"'; Execute = $false; Screentip = "Sign in as an invSys user" },
            @{ Id = "btnReceivingSignOut"; Label = "Sign Out"; DirectAction = "modRoleEventWriter.SignOutCurrentUser"; Execute = $false },
            @{ Id = "btnReceivingSetup"; Label = "Setup UI"; Macro = "modTS_Received.EnsureGeneratedButtons"; Execute = $false; RequiredCapability = "RECEIVE_POST" },
            @{ Id = "btnReceivingConfirm"; Label = "Confirm Writes"; Macro = "modTS_Received.ConfirmWrites"; Execute = $false; RequiredCapability = "RECEIVE_POST" },
            @{ Id = "btnReceivingUndo"; Label = "Undo"; Macro = "modTS_Received.MacroUndo"; Execute = $false; RequiredCapability = "RECEIVE_POST" },
            @{ Id = "btnReceivingRedo"; Label = "Redo"; Macro = "modTS_Received.MacroRedo"; Execute = $false; RequiredCapability = "RECEIVE_POST" }
        )
    }
    @{
        Name = "Shipping"
        File = "invSys.Shipping.xlam"
        Callback = "RibbonOnActionShipping"
        EnabledCallback = "RibbonRequiredCapabilityGetEnabledShipping"
        StatusLabels = @(
            @{ Id = "lblShippingServerStatus"; GetLabel = "RibbonServerStatusGetLabel" },
            @{ Id = "lblShippingAccessStatus"; GetLabel = "RibbonAccessStatusGetLabel" }
        )
        Buttons = @(
            @{ Id = "btnShippingConnectServer"; Label = "Connect Server"; DirectAction = 'modRoleEventWriter.ConnectWarehouseStorageForCapability "SHIP_POST"'; Execute = $false },
            @{ Id = "btnShippingCurrentUser"; GetLabel = "RibbonCurrentUserGetLabel"; DirectAction = 'modRoleEventWriter.PromptSetCurrentUserForCapability "SHIP_POST"'; Execute = $false; Screentip = "Sign in as an invSys user" },
            @{ Id = "btnShippingSignOut"; Label = "Sign Out"; DirectAction = "modRoleEventWriter.SignOutCurrentUser"; Execute = $false },
            @{ Id = "btnShippingSetup"; Label = "Setup UI"; Macro = "modTS_Shipments.InitializeShipmentsUI"; Execute = $false; RequiredCapability = "SHIP_POST" },
            @{ Id = "btnShippingBoxMode"; GetLabel = "RibbonBoxMakerModeGetLabel"; Macro = "modTS_Shipments.BtnSwitchToBoxMaker"; Execute = $false; RequiredCapability = "SHIP_POST" },
            @{ Id = "btnShippingSaveBox"; Label = "Save Box"; Macro = "modTS_Shipments.BtnSaveBox"; Execute = $false; RequiredCapability = "SHIP_POST" },
            @{ Id = "btnShippingDeleteBoxVersion"; Label = "Delete Version"; Macro = "modTS_Shipments.BtnDeleteBoxVersion"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnShippingDeleteBox"; Label = "Delete Box"; Macro = "modTS_Shipments.BtnDeleteBox"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnShippingConfirm"; Label = "Confirm Inventory"; Macro = "modTS_Shipments.BtnConfirmInventory"; Execute = $false; RequiredCapability = "SHIP_POST" },
            @{ Id = "btnShippingBoxCreated"; Label = "Box Created"; Macro = "modTS_Shipments.BtnBoxCreated"; Execute = $false; RequiredCapability = "SHIP_POST" },
            @{ Id = "btnShippingBoxUnboxed"; Label = "Box Unboxed"; Macro = "modTS_Shipments.BtnBoxUnboxed"; Execute = $false; RequiredCapability = "SHIP_POST" },
            @{ Id = "btnShippingStage"; Label = "To Shipments"; Macro = "modTS_Shipments.BtnToShipments"; Execute = $false; RequiredCapability = "SHIP_POST" },
            @{ Id = "btnShippingSend"; Label = "Shipments Sent"; Macro = "modTS_Shipments.BtnShipmentsSent"; Execute = $false; RequiredCapability = "SHIP_POST" }
        )
    }
    @{
        Name = "Production"
        File = "invSys.Production.xlam"
        Callback = "RibbonOnActionProduction"
        EnabledCallback = "RibbonRequiredCapabilityGetEnabledProduction"
        StatusLabels = @(
            @{ Id = "lblProductionServerStatus"; GetLabel = "RibbonServerStatusGetLabel" },
            @{ Id = "lblProductionAccessStatus"; GetLabel = "RibbonAccessStatusGetLabel" }
        )
        Buttons = @(
            @{ Id = "btnProductionConnectServer"; Label = "Connect Server"; DirectAction = 'modRoleEventWriter.ConnectWarehouseStorageForCapability "PROD_POST"'; Execute = $false },
            @{ Id = "btnProductionCurrentUser"; GetLabel = "RibbonCurrentUserGetLabel"; DirectAction = 'modRoleEventWriter.PromptSetCurrentUserForCapability "PROD_POST"'; Execute = $false; Screentip = "Sign in as an invSys user" },
            @{ Id = "btnProductionSignOut"; Label = "Sign Out"; DirectAction = "modRoleEventWriter.SignOutCurrentUser"; Execute = $false },
            @{ Id = "btnProductionSetup"; Label = "Setup UI"; Macro = "mProduction.InitializeProductionUI"; Execute = $false; RequiredCapability = "PROD_POST" },
            @{ Id = "btnProductionLoad"; Label = "Load Recipe"; Macro = "mProduction.BtnLoadRecipe"; Execute = $false; RequiredCapability = "PROD_POST" },
            @{ Id = "btnProductionUsed"; Label = "To Used"; Macro = "mProduction.BtnToUsed"; Execute = $false; RequiredCapability = "PROD_POST" },
            @{ Id = "btnProductionMade"; Label = "To Made"; Macro = "mProduction.BtnToMade"; Execute = $false; RequiredCapability = "PROD_POST" },
            @{ Id = "btnProductionTotal"; Label = "To Total Inv"; Macro = "mProduction.BtnToTotalInv"; Execute = $false; RequiredCapability = "PROD_POST" },
            @{ Id = "btnProductionPrintCodes"; Label = "Print Recall Codes"; Macro = "mProduction.BtnPrintRecallCodes"; Execute = $false; RequiredCapability = "PROD_POST" }
        )
    }
    @{
        Name = "Admin"
        File = "invSys.Admin.xlam"
        Callback = "RibbonOnActionAdmin"
        EnabledCallback = "RibbonRequiredCapabilityGetEnabledAdmin"
        StatusLabels = @(
            @{ Id = "lblAdminServerStatus"; GetLabel = "RibbonServerStatusGetLabel" },
            @{ Id = "lblAdminAccessStatus"; GetLabel = "RibbonAccessStatusGetLabel" }
        )
        Buttons = @(
            @{ Id = "btnAdminOpen"; Label = "Admin Console"; Macro = "modAdmin.Admin_Click"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnAdminConnectServer"; Label = "Connect Server"; DirectAction = 'modRoleEventWriter.ConnectWarehouseStorageForCapability "ADMIN_MAINT"'; Execute = $false },
            @{ Id = "btnAdminCurrentUser"; GetLabel = "RibbonCurrentUserGetLabel"; DirectAction = 'modRoleEventWriter.PromptSetCurrentUserForCapability "ADMIN_MAINT"'; Execute = $false; Screentip = "Sign in as an invSys user" },
            @{ Id = "btnAdminSignOut"; Label = "Sign Out"; DirectAction = "modRoleEventWriter.SignOutCurrentUser"; Execute = $false },
            @{ Id = "btnAdminUsers"; Label = "Users and Roles"; Macro = "modAdmin.Open_CreateDeleteUser"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnAdminSettings"; Label = "Settings"; Macro = "modAdmin.Open_Settings"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnAdminWarehouses"; Label = "View Warehouses"; Macro = "modAdmin.Open_WarehouseDirectory"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnAdminWarehouseRoot"; Label = "Add Warehouse Root"; Macro = "modAdmin.Add_WarehouseDirectoryRoot"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnAdminCreateWarehouse"; Label = "Create New Warehouse"; Macro = "modAdmin.Open_CreateWarehouse"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnAdminSetupTesterStation"; Label = "Setup Tester Station"; Macro = "modAdmin.Admin_SetupTesterStation_Click"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnAdminSeedInventory"; Label = "Seed Demo Inventory"; Macro = "modAdmin.Seed_DemoInventory"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnAdminVerifyAddinsPublished"; Label = "Verify Add-ins Published"; Macro = "modAdmin.Verify_AddinsPublished"; Execute = $false; RequiredCapability = "ADMIN_MAINT" },
            @{ Id = "btnAdminRetireMigrateWarehouse"; Label = "Retire / Migrate Warehouse"; Macro = "modAdmin.Admin_RetireMigrateWarehouse_Click"; Execute = $false; Screentip = "Archive, migrate, retire, or delete a warehouse runtime"; RequiredCapability = "ADMIN_MAINT" }
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
        $labels = Get-RibbonLabelControls -CustomUiXml $customUiXml
        $callbackModuleText = Get-ModuleText -Workbook $wb -ComponentName "modRibbonGenerated"
        Add-ResultRow -Rows $resultRows -Check "$($spec.Name).CallbackModule" -Passed (-not [string]::IsNullOrWhiteSpace($callbackModuleText)) -Detail "modRibbonGenerated"

        if ($spec.ContainsKey("StatusLabels")) {
            foreach ($statusLabel in $spec.StatusLabels) {
                $labelId = [string]$statusLabel.Id
                if ($labels.ContainsKey($labelId)) {
                    $labelInfo = $labels[$labelId]
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).StatusLabel.$labelId" -Passed ($labelInfo.GetLabel -eq $statusLabel.GetLabel) -Detail "GetLabel=$($labelInfo.GetLabel)"
                }
                else {
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).StatusLabel.$labelId" -Passed $false -Detail "Label control missing from Ribbon XML."
                }
            }
        }

        foreach ($button in $spec.Buttons) {
            $buttonId = [string]$button.Id
            if ($buttons.ContainsKey($buttonId)) {
                $buttonInfo = $buttons[$buttonId]
                $detail = "Label=$($buttonInfo.Label); OnAction=$($buttonInfo.OnAction); GetEnabled=$($buttonInfo.GetEnabled); Screentip=$($buttonInfo.Screentip)"
                if ($button.ContainsKey("GetLabel")) {
                    $passed = ($buttonInfo.GetLabel -eq $button.GetLabel -and $buttonInfo.OnAction -eq $spec.Callback)
                }
                else {
                    $passed = ($buttonInfo.Label -eq $button.Label -and $buttonInfo.OnAction -eq $spec.Callback)
                }
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonButton.$buttonId" -Passed $passed -Detail $detail
                if ($button.ContainsKey("Screentip")) {
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonButtonScreentip.$buttonId" -Passed ($buttonInfo.Screentip -eq $button.Screentip) -Detail $buttonInfo.Screentip
                }
                if ($button.ContainsKey("RequiredCapability")) {
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonButtonGetEnabled.$buttonId" -Passed ($buttonInfo.GetEnabled -eq $spec.EnabledCallback) -Detail $buttonInfo.GetEnabled
                }
            }
            else {
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonButton.$buttonId" -Passed $false -Detail "Button missing from Ribbon XML."
                if ($button.ContainsKey("Screentip")) {
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonButtonScreentip.$buttonId" -Passed $false -Detail "Button missing from Ribbon XML."
                }
                if ($button.ContainsKey("RequiredCapability")) {
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).RibbonButtonGetEnabled.$buttonId" -Passed $false -Detail "Button missing from Ribbon XML."
                }
            }

            $callbackHasButton = (-not [string]::IsNullOrWhiteSpace($callbackModuleText)) -and $callbackModuleText.Contains($buttonId)
            if ($button.ContainsKey("Macro")) {
                $macroName = [string]$button.Macro
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).MacroExists.$buttonId" -Passed (Test-ProcedureExists -Workbook $wb -QualifiedMacroName $macroName) -Detail $macroName

                $callbackHasMacro = (-not [string]::IsNullOrWhiteSpace($callbackModuleText)) -and $callbackModuleText.Contains($macroName)
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).CallbackMap.$buttonId" -Passed ($callbackHasButton -and $callbackHasMacro) -Detail "$buttonId -> $macroName"
            }
            elseif ($button.ContainsKey("DirectAction")) {
                $directAction = [string]$button.DirectAction
                $callbackHasDirectAction = (-not [string]::IsNullOrWhiteSpace($callbackModuleText)) -and $callbackModuleText.Contains($directAction)
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).CallbackMap.$buttonId" -Passed ($callbackHasButton -and $callbackHasDirectAction) -Detail "$buttonId -> $directAction"
            }
            if ($button.ContainsKey("RequiredCapability")) {
                $callbackHasEnabled = (-not [string]::IsNullOrWhiteSpace($callbackModuleText)) -and $callbackModuleText.Contains([string]$spec.EnabledCallback) -and $callbackModuleText.Contains("RibbonRequiredCapabilityIsEnabledById") -and $callbackModuleText.Contains([string]$button.RequiredCapability)
                Add-ResultRow -Rows $resultRows -Check "$($spec.Name).CallbackGetEnabled.$buttonId" -Passed $callbackHasEnabled -Detail "$buttonId -> $($button.RequiredCapability)"
                try {
                    $enabledOffline = Run-WorkbookMacro -Excel $excel -WorkbookName $wb.Name -MacroName "modRibbonGenerated.RibbonRequiredCapabilityIsEnabledById" -Arguments @($buttonId)
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).DisabledOffline.$buttonId" -Passed ([bool]$enabledOffline -eq $false) -Detail "$buttonId enabled=$enabledOffline"
                }
                catch {
                    Add-ResultRow -Rows $resultRows -Check "$($spec.Name).DisabledOffline.$buttonId" -Passed $false -Detail $_.Exception.Message
                }
            }

            if ($button.ContainsKey("Macro") -and $button.Execute) {
                $macroName = [string]$button.Macro
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
    if ($resultRows.Count -gt 0) {
        [System.IO.File]::WriteAllLines($resultPath, $lines)
    }

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
