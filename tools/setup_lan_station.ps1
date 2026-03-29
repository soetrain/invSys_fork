[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$DeployRoot = "deploy/current",

    [Parameter(Mandatory = $true)]
    [string]$WarehouseId,

    [Parameter(Mandatory = $true)]
    [string]$StationId,

    [Parameter(Mandatory = $true)]
    [string]$SharedRuntimeRoot,

    [Parameter(Mandatory = $true)]
    [string]$StationInboxRoot,

    [Parameter(Mandatory = $false)]
    [string]$StationInboxShareName = "",

    [Parameter(Mandatory = $false)]
    [string]$StationInboxShareHost = "",

    [switch]$PublishStationInboxShare,

    [Parameter(Mandatory = $false)]
    [string]$RoleDefault = "RECEIVE",

    [Parameter(Mandatory = $false)]
    [string]$StationName = $env:COMPUTERNAME,

    [Parameter(Mandatory = $false)]
    [string]$StationUserId = $env:USERNAME,

    [Parameter(Mandatory = $false)]
    [string]$LocalConfigRoot = "",

    [switch]$CreateOperatorWorkbook,

    [Parameter(Mandatory = $false)]
    [string]$OperatorWorkbookPath = "",

    [switch]$SkipSharedBootstrap,

    [switch]$VisibleExcel
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$Obj)
    if ($null -ne $Obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Obj) } catch {}
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
        9 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6], $Arguments[7], $Arguments[8]) }
        10 { return $Excel.Run($macro, $Arguments[0], $Arguments[1], $Arguments[2], $Arguments[3], $Arguments[4], $Arguments[5], $Arguments[6], $Arguments[7], $Arguments[8], $Arguments[9]) }
        default { throw "Too many macro arguments for $MacroName" }
    }
}

function Convert-RoleToEventType {
    param([string]$RoleName)

    switch ($RoleName.Trim().ToUpperInvariant()) {
        "RECEIVE" { return "RECEIVE" }
        "RECEIVING" { return "RECEIVE" }
        "SHIP" { return "SHIP" }
        "SHIPPING" { return "SHIP" }
        "PROD" { return "PROD_CONSUME" }
        "PRODUCTION" { return "PROD_CONSUME" }
        "PROD_CONSUME" { return "PROD_CONSUME" }
        "PROD_COMPLETE" { return "PROD_CONSUME" }
        default { return $RoleName.Trim().ToUpperInvariant() }
    }
}

function Convert-RoleToCapability {
    param([string]$RoleName)

    switch ($RoleName.Trim().ToUpperInvariant()) {
        "RECEIVE" { return "RECEIVE_POST" }
        "RECEIVING" { return "RECEIVE_POST" }
        "SHIP" { return "SHIP_POST" }
        "SHIPPING" { return "SHIP_POST" }
        "PROD" { return "PROD_POST" }
        "PRODUCTION" { return "PROD_POST" }
        "PROD_CONSUME" { return "PROD_POST" }
        "PROD_COMPLETE" { return "PROD_POST" }
        "ADMIN" { return "ADMIN_MAINT" }
        default { return "" }
    }
}

function Get-RoleSetup {
    param(
        [string]$RoleName,
        [string]$CoreWorkbookName
    )

    $roleKey = $RoleName.Trim().ToUpperInvariant()
    switch ($roleKey) {
        "RECEIVE" {
            return @{
                RoleLabel = "Receiving"
                Addins = @("invSys.Inventory.Domain.xlam", "invSys.Receiving.xlam")
                InitSteps = @(
                    @{ Workbook = "invSys.Receiving.xlam"; Macro = "modReceivingInit.InitReceivingAddin" }
                )
                EnsureSteps = @(
                    @{ Workbook = "invSys.Receiving.xlam"; Macro = "modReceivingInit.EnsureReceivingSurfaceForWorkbook" }
                )
            }
        }
        "RECEIVING" {
            return Get-RoleSetup -RoleName "RECEIVE" -CoreWorkbookName $CoreWorkbookName
        }
        "SHIP" {
            return @{
                RoleLabel = "Shipping"
                Addins = @("invSys.Inventory.Domain.xlam", "invSys.Shipping.xlam")
                InitSteps = @(
                    @{ Workbook = "invSys.Shipping.xlam"; Macro = "modShippingInit.InitShippingAddin" }
                )
                EnsureSteps = @(
                    @{ Workbook = "invSys.Shipping.xlam"; Macro = "modShippingInit.EnsureShippingSurfaceForWorkbook" }
                )
            }
        }
        "SHIPPING" {
            return Get-RoleSetup -RoleName "SHIP" -CoreWorkbookName $CoreWorkbookName
        }
        "PROD" {
            return @{
                RoleLabel = "Production"
                Addins = @("invSys.Inventory.Domain.xlam", "invSys.Designs.Domain.xlam", "invSys.Production.xlam")
                InitSteps = @(
                    @{ Workbook = "invSys.Production.xlam"; Macro = "modProductionInit.InitProductionAddin" }
                )
                EnsureSteps = @(
                    @{ Workbook = "invSys.Production.xlam"; Macro = "modProductionInit.EnsureProductionSurfaceForWorkbook" }
                )
            }
        }
        "PRODUCTION" {
            return Get-RoleSetup -RoleName "PROD" -CoreWorkbookName $CoreWorkbookName
        }
        "ADMIN" {
            return @{
                RoleLabel = "Admin"
                Addins = @("invSys.Admin.xlam")
                InitSteps = @(
                    @{ Workbook = "invSys.Admin.xlam"; Macro = "modAdminInit.InitAdminAddin" }
                )
                EnsureSteps = @(
                    @{ Workbook = $CoreWorkbookName; Macro = "modRoleWorkbookSurfaces.EnsureAdminLegacyWorkbookSurface" }
                    @{ Workbook = "invSys.Admin.xlam"; Macro = "modAdminConsole.EnsureAdminSchema" }
                )
            }
        }
        default {
            throw "Unsupported role for operator workbook bootstrap: $RoleName"
        }
    }
}

function Ensure-Directory {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Test-IsUncPath {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $false }
    return $Path.StartsWith("\\")
}

function Normalize-ShareHost {
    param([string]$HostName)

    $normalized = $HostName.Trim()
    if ([string]::IsNullOrWhiteSpace($normalized)) {
        $normalized = $env:COMPUTERNAME
    }
    return $normalized
}

function Normalize-ShareName {
    param(
        [string]$ProvidedName,
        [string]$FallbackPath,
        [string]$StationId
    )

    $name = $ProvidedName.Trim()
    if ([string]::IsNullOrWhiteSpace($name) -and -not [string]::IsNullOrWhiteSpace($FallbackPath)) {
        $leaf = Split-Path -Path $FallbackPath -Leaf
        if (-not [string]::IsNullOrWhiteSpace($leaf)) {
            $name = $leaf
        }
    }
    if ([string]::IsNullOrWhiteSpace($name)) {
        $name = "invSysStation" + $StationId
    }
    return $name
}

function Ensure-StationInboxShare {
    param(
        [string]$LocalInboxRoot,
        [string]$ShareName
    )

    Ensure-Directory -Path $LocalInboxRoot

    $existing = Get-SmbShare -Name $ShareName -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        New-SmbShare -Name $ShareName -Path $LocalInboxRoot -FullAccess "Everyone" | Out-Null
        return "CREATED"
    }

    $currentPath = [string]$existing.Path
    if (-not [string]::Equals($currentPath, $LocalInboxRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Existing SMB share '$ShareName' points to '$currentPath' instead of '$LocalInboxRoot'."
    }

    try {
        Grant-SmbShareAccess -Name $ShareName -AccountName "Everyone" -AccessRight Full -Force -ErrorAction Stop | Out-Null
    }
    catch {
        if ($_.Exception.Message -notmatch "already has") { throw }
    }

    return "EXISTS"
}

function Join-PackedArgs {
    param([string[]]$Values)
    return ($Values -join "|")
}

function Resolve-OperatorWorkbookPath {
    param(
        [string]$RequestedPath,
        [string]$WarehouseId,
        [string]$StationId,
        [string]$RoleLabel
    )

    if (-not [string]::IsNullOrWhiteSpace($RequestedPath)) {
        return [System.IO.Path]::GetFullPath($RequestedPath)
    }

    $documentsRoot = [Environment]::GetFolderPath("MyDocuments")
    if ([string]::IsNullOrWhiteSpace($documentsRoot)) {
        $documentsRoot = $env:USERPROFILE
    }
    return (Join-Path $documentsRoot ($WarehouseId + "_" + $StationId + "_" + $RoleLabel + "_Operator.xlsb"))
}

function Open-WorkbookOnce {
    param(
        [object]$Excel,
        [string]$FullPath,
        [hashtable]$WorkbookMap
    )

    foreach ($wb in $Excel.Workbooks) {
        if ([string]::Equals([string]$wb.FullName, $FullPath, [System.StringComparison]::OrdinalIgnoreCase)) {
            $WorkbookMap[$wb.Name] = $wb
            return $wb
        }
    }

    $opened = $Excel.Workbooks.Open($FullPath)
    $WorkbookMap[$opened.Name] = $opened
    return $opened
}

function Get-DiagnosticValue {
    param(
        [string]$DiagnosticText,
        [string]$Key
    )

    if ([string]::IsNullOrWhiteSpace($DiagnosticText) -or [string]::IsNullOrWhiteSpace($Key)) {
        return ""
    }

    foreach ($line in ($DiagnosticText -split "`r?`n")) {
        if ($line.StartsWith($Key + "=", [System.StringComparison]::OrdinalIgnoreCase)) {
            return $line.Substring($Key.Length + 1)
        }
    }
    return ""
}

function Test-ExcelWorkbookOpenable {
    param(
        [object]$Excel,
        [string]$FullPath
    )

    if ([string]::IsNullOrWhiteSpace($FullPath)) { return $false }
    if (-not (Test-Path -LiteralPath $FullPath)) { return $false }

    $wb = $null
    try {
        $wb = $Excel.Workbooks.Open($FullPath, 0, $true)
        return ($null -ne $wb)
    }
    catch {
        return $false
    }
    finally {
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch {}
            Release-ComObject $wb
        }
    }
}

$repoPath = (Resolve-Path $RepoRoot).Path
$deployPath = Join-Path $repoPath $DeployRoot
$coreAddinPath = Join-Path $deployPath "invSys.Core.xlam"

if (-not (Test-Path -LiteralPath $coreAddinPath)) {
    throw "Core add-in not found: $coreAddinPath"
}

$resolvedSharedRoot = [System.IO.Path]::GetFullPath($SharedRuntimeRoot)
if (-not (Test-Path -LiteralPath $resolvedSharedRoot)) {
    throw "Shared runtime root not found: $resolvedSharedRoot"
}

$sharedConfigPath = Join-Path $resolvedSharedRoot ($WarehouseId + ".invSys.Config.xlsb")
$sharedAuthPath = Join-Path $resolvedSharedRoot ($WarehouseId + ".invSys.Auth.xlsb")
$sharedSnapshotPath = Join-Path $resolvedSharedRoot ($WarehouseId + ".invSys.Snapshot.Inventory.xlsb")

if ([string]::IsNullOrWhiteSpace($LocalConfigRoot)) {
    $LocalConfigRoot = Join-Path "C:\invSys" $WarehouseId
}

$resolvedLocalConfigRoot = [System.IO.Path]::GetFullPath($LocalConfigRoot)
$resolvedInboxRoot = ""
$configuredInboxRoot = ""
$stationInboxShareStatus = "SKIPPED"
$stationInboxSharePath = ""
$stationInboxShellAccessible = $false
$stationInboxExcelOpenable = $false

if ($PublishStationInboxShare) {
    if (Test-IsUncPath -Path $StationInboxRoot) {
        throw "PublishStationInboxShare requires a local StationInboxRoot path, not a UNC path: $StationInboxRoot"
    }

    $resolvedInboxRoot = [System.IO.Path]::GetFullPath($StationInboxRoot)
    $shareHost = Normalize-ShareHost -HostName $StationInboxShareHost
    $shareName = Normalize-ShareName -ProvidedName $StationInboxShareName -FallbackPath $resolvedInboxRoot -StationId $StationId
    $stationInboxShareStatus = Ensure-StationInboxShare -LocalInboxRoot $resolvedInboxRoot -ShareName $shareName
    $stationInboxSharePath = "\\" + $shareHost + "\" + $shareName
    $configuredInboxRoot = $stationInboxSharePath
}
else {
    $resolvedInboxRoot = [System.IO.Path]::GetFullPath($StationInboxRoot)
    $configuredInboxRoot = $resolvedInboxRoot
}

$localConfigPath = Join-Path $resolvedLocalConfigRoot ($WarehouseId + ".invSys.Config.xlsb")
$eventType = Convert-RoleToEventType -RoleName $RoleDefault
$roleCapability = Convert-RoleToCapability -RoleName $RoleDefault
$operatorWorkbookOut = ""
$operatorRefreshStatus = "SKIPPED"
$operatorRefreshReport = ""
$authProvisionStatus = ""
$authValidationStatus = ""
$snapshotShellAccessible = $false
$snapshotExcelOpenable = $false
$roleReady = $false

$excel = $null
$coreWb = $null
$openedWorkbooks = @()
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$VisibleExcel
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 1

    $coreWb = $excel.Workbooks.Open($coreAddinPath)
    $openedWorkbooks += $coreWb
    $workbookMap = @{}
    $workbookMap[$coreWb.Name] = $coreWb
    [void](Run-WorkbookMacro -Excel $excel -WorkbookName $coreWb.Name -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @(
        $resolvedSharedRoot
    ))

    if (-not $SkipSharedBootstrap) {
        $sharedPacked = Join-PackedArgs @(
            $WarehouseId,
            $StationId,
            $StationName,
            ($configuredInboxRoot + "\"),
            $RoleDefault,
            $sharedConfigPath,
            $resolvedSharedRoot
        )
        $sharedResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreWb.Name -MacroName "modConfig.EnsureStationConfigEntryPackedForAutomation" -Arguments @(
            $sharedPacked
        ))
        if (-not $sharedResult.StartsWith("OK", [System.StringComparison]::OrdinalIgnoreCase)) {
            throw "Shared config bootstrap failed. Result=$sharedResult"
        }
    }

    if (-not (Test-Path -LiteralPath $sharedConfigPath)) {
        throw "Shared config workbook not found after bootstrap step: $sharedConfigPath"
    }

    $authPacked = Join-PackedArgs @(
        $WarehouseId,
        $StationId,
        $StationUserId,
        $StationUserId,
        $RoleDefault,
        $sharedAuthPath,
        "svc_processor"
    )
    $authProvisionStatus = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreWb.Name -MacroName "modAuth.EnsureStationRoleAuthPackedForAutomation" -Arguments @(
        $authPacked
    ))
    if (-not $authProvisionStatus.StartsWith("OK|", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Shared auth provisioning failed. Result=$authProvisionStatus"
    }
    $authValidationStatus = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreWb.Name -MacroName "modAuth.ValidateStationRoleAuthPackedForAutomation" -Arguments @(
        $authPacked
    ))
    if (-not $authValidationStatus.StartsWith("OK|", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Shared auth validation failed. Result=$authValidationStatus"
    }

    Ensure-Directory -Path $resolvedLocalConfigRoot
    Copy-Item -LiteralPath $sharedConfigPath -Destination $localConfigPath -Force

    $localPacked = Join-PackedArgs @(
        $WarehouseId,
        $StationId,
        $StationName,
        ($configuredInboxRoot + "\"),
        $RoleDefault,
        $localConfigPath,
        $resolvedSharedRoot
    )
    $localResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreWb.Name -MacroName "modConfig.EnsureStationConfigEntryPackedForAutomation" -Arguments @(
        $localPacked
    ))
    if (-not $localResult.StartsWith("OK", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Local config bootstrap failed. Result=$localResult"
    }

    $inboxPacked = Join-PackedArgs @(
        $WarehouseId,
        $StationId,
        $RoleDefault,
        $localConfigPath
    )
    $inboxResult = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreWb.Name -MacroName "modConfig.EnsureStationInboxPackedForAutomation" -Arguments @(
        $inboxPacked
    ))
    if (-not $inboxResult.StartsWith("OK|", [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Station inbox bootstrap failed. Result=$inboxResult"
    }
    $inboxPath = $inboxResult.Substring(3)
    $stationInboxShellAccessible = Test-Path -LiteralPath $inboxPath
    $stationInboxExcelOpenable = Test-ExcelWorkbookOpenable -Excel $excel -FullPath $inboxPath

    if ($CreateOperatorWorkbook) {
        $roleSetup = Get-RoleSetup -RoleName $RoleDefault -CoreWorkbookName $coreWb.Name
        foreach ($addinFile in $roleSetup.Addins) {
            $addinPath = Join-Path $deployPath $addinFile
            if (-not (Test-Path -LiteralPath $addinPath)) {
                throw "Required role add-in not found: $addinPath"
            }
            $roleWb = Open-WorkbookOnce -Excel $excel -FullPath $addinPath -WorkbookMap $workbookMap
            if ($openedWorkbooks -notcontains $roleWb) {
                $openedWorkbooks += $roleWb
            }
        }

        foreach ($step in $roleSetup.InitSteps) {
            [void](Run-WorkbookMacro -Excel $excel -WorkbookName $step.Workbook -MacroName $step.Macro)
        }

        $operatorWorkbookOut = Resolve-OperatorWorkbookPath -RequestedPath $OperatorWorkbookPath -WarehouseId $WarehouseId -StationId $StationId -RoleLabel $roleSetup.RoleLabel
        Ensure-Directory -Path ([System.IO.Path]::GetDirectoryName($operatorWorkbookOut))

        $operatorWb = $excel.Workbooks.Add()
        $openedWorkbooks += $operatorWb
        try { [void]$operatorWb.Activate() } catch {}

        foreach ($step in $roleSetup.EnsureSteps) {
            [void](Run-WorkbookMacro -Excel $excel -WorkbookName $step.Workbook -MacroName $step.Macro -Arguments @($operatorWb))
        }

        try {
            $operatorDiagnostic = [string](Run-WorkbookMacro -Excel $excel -WorkbookName $coreWb.Name -MacroName "modOperatorReadModel.DiagnoseInventoryReadModelRefresh" -Arguments @(
                $operatorWb,
                $WarehouseId,
                "LOCAL"
            ))
            $operatorRefreshReport = Get-DiagnosticValue -DiagnosticText $operatorDiagnostic -Key "RefreshReport"
            if ([string]::IsNullOrWhiteSpace($operatorRefreshReport)) {
                $operatorRefreshReport = "UNKNOWN"
            }
            $operatorRefreshStatus = if ([string]::Equals($operatorRefreshReport, "OK", [System.StringComparison]::OrdinalIgnoreCase)) { "OK" } else { "STALE_OR_FAILED" }
        }
        catch {
            $operatorRefreshStatus = "STALE_OR_FAILED"
            $operatorRefreshReport = $_.Exception.Message
        }

        if (Test-Path -LiteralPath $operatorWorkbookOut) {
            Remove-Item -LiteralPath $operatorWorkbookOut -Force
        }
        $operatorWb.SaveAs($operatorWorkbookOut, 50)
        $operatorWb.Close($false)
    }

    $snapshotShellAccessible = Test-Path -LiteralPath $sharedSnapshotPath
    $snapshotExcelOpenable = Test-ExcelWorkbookOpenable -Excel $excel -FullPath $sharedSnapshotPath
    $roleReady = $authValidationStatus.StartsWith("OK|", [System.StringComparison]::OrdinalIgnoreCase) `
        -and $stationInboxShellAccessible `
        -and $stationInboxExcelOpenable `
        -and $snapshotShellAccessible `
        -and $snapshotExcelOpenable `
        -and ((-not $CreateOperatorWorkbook) -or $operatorRefreshStatus -eq "OK")

    Write-Output "LAN_STATION_SETUP_OK"
    Write-Output ("WarehouseId=" + $WarehouseId)
    Write-Output ("StationId=" + $StationId)
    Write-Output ("StationUserId=" + $StationUserId)
    Write-Output ("SharedRuntimeRoot=" + $resolvedSharedRoot)
    Write-Output ("SharedConfigPath=" + $sharedConfigPath)
    Write-Output ("SharedAuthPath=" + $sharedAuthPath)
    Write-Output ("LocalConfigPath=" + $localConfigPath)
    Write-Output ("StationInboxRoot=" + $resolvedInboxRoot)
    Write-Output ("ConfiguredPathInboxRoot=" + $configuredInboxRoot)
    Write-Output ("StationInboxShareStatus=" + $stationInboxShareStatus)
    if (-not [string]::IsNullOrWhiteSpace($stationInboxSharePath)) {
        Write-Output ("StationInboxSharePath=" + $stationInboxSharePath)
    }
    Write-Output ("InboxPath=" + $inboxPath)
    Write-Output ("InboxShellAccessible=" + $stationInboxShellAccessible)
    Write-Output ("InboxExcelOpenable=" + $stationInboxExcelOpenable)
    Write-Output ("RoleDefault=" + $RoleDefault.ToUpperInvariant())
    Write-Output ("RoleCapability=" + $roleCapability)
    Write-Output ("AuthProvision=" + $authProvisionStatus)
    Write-Output ("AuthValidation=" + $authValidationStatus)
    Write-Output ("SharedSnapshotPath=" + $sharedSnapshotPath)
    Write-Output ("SnapshotShellAccessible=" + $snapshotShellAccessible)
    Write-Output ("SnapshotExcelOpenable=" + $snapshotExcelOpenable)
    if ($CreateOperatorWorkbook) {
        Write-Output ("OperatorWorkbookPath=" + $operatorWorkbookOut)
        Write-Output ("OperatorReadModelRefresh=" + $operatorRefreshStatus)
        Write-Output ("OperatorReadModelRefreshReport=" + ($operatorRefreshReport -replace "`r?`n", " ; "))
    }
    Write-Output ("RoleReady=" + $roleReady)
}
finally {
    foreach ($wb in ($openedWorkbooks | Select-Object -Unique)) {
        if ($null -ne $wb) {
            try { $wb.Close($false) } catch {}
            Release-ComObject $wb
        }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
}
