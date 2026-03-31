[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$DeployRoot = "deploy/current",

    [Parameter(Mandatory = $true)]
    [ValidateSet("WarehouseBatch", "WarehousePublish", "HqAggregation")]
    [string]$JobType,

    [Parameter(Mandatory = $false)]
    [string]$WarehouseId = "",

    [Parameter(Mandatory = $false)]
    [int]$BatchSize = 0,

    [Parameter(Mandatory = $false)]
    [string]$SharePointRoot = "",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = "",

    [Parameter(Mandatory = $false)]
    [string]$RuntimeRootOverride = "",

    [Parameter(Mandatory = $false)]
    [string]$LogPath = "",

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

function Ensure-Directory {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-JobLog {
    param(
        [string]$Path,
        [string]$Message
    )

    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    Ensure-Directory -Path (Split-Path -Parent $Path)
    Add-Content -LiteralPath $Path -Value ("[{0}] {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Message)
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
        default { throw "Too many macro arguments for $MacroName" }
    }
}

$repoPath = (Resolve-Path $RepoRoot).Path
$deployPath = if ([System.IO.Path]::IsPathRooted($DeployRoot)) { $DeployRoot } else { Join-Path $repoPath $DeployRoot }
$deployPath = (Resolve-Path $deployPath).Path

$addinFiles = @(
    "invSys.Core.xlam",
    "invSys.Inventory.Domain.xlam",
    "invSys.Designs.Domain.xlam",
    "invSys.Admin.xlam"
)

$excel = $null
$openedWorkbooks = New-Object System.Collections.Generic.List[object]
$resultText = ""

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = [bool]$VisibleExcel
    $excel.DisplayAlerts = $false

    foreach ($fileName in $addinFiles) {
        $fullPath = Join-Path $deployPath $fileName
        if (-not (Test-Path -LiteralPath $fullPath)) {
            throw "Missing add-in: $fullPath"
        }
        $wb = $excel.Workbooks.Open($fullPath, 0, $true)
        $openedWorkbooks.Add($wb) | Out-Null
    }

    if (-not [string]::IsNullOrWhiteSpace($RuntimeRootOverride)) {
        [void](Run-WorkbookMacro -Excel $excel -WorkbookName "invSys.Core.xlam" -MacroName "modRuntimeWorkbooks.SetCoreDataRootOverride" -Arguments @($RuntimeRootOverride))
    }

    switch ($JobType) {
        "WarehouseBatch" {
            $resultText = [string](Run-WorkbookMacro -Excel $excel -WorkbookName "invSys.Admin.xlam" -MacroName "modAdminConsole.RunScheduledWarehouseBatchForAutomation" -Arguments @($WarehouseId, [int]$BatchSize))
        }
        "WarehousePublish" {
            $resultText = [string](Run-WorkbookMacro -Excel $excel -WorkbookName "invSys.Admin.xlam" -MacroName "modAdminConsole.RunScheduledWarehousePublishForAutomation" -Arguments @($WarehouseId, $SharePointRoot))
        }
        "HqAggregation" {
            $resultText = [string](Run-WorkbookMacro -Excel $excel -WorkbookName "invSys.Admin.xlam" -MacroName "modAdminConsole.RunScheduledHQAggregationForAutomation" -Arguments @($SharePointRoot, $OutputPath))
        }
    }

    if ([string]::IsNullOrWhiteSpace($resultText)) {
        $resultText = "FAIL|Report=Automation macro returned no result."
    }
}
catch {
    $resultText = "FAIL|Report=$($_.Exception.Message.Replace('|', '/'))"
}
finally {
    if ($null -ne $excel -and -not [string]::IsNullOrWhiteSpace($RuntimeRootOverride)) {
        try { [void](Run-WorkbookMacro -Excel $excel -WorkbookName "invSys.Core.xlam" -MacroName "modRuntimeWorkbooks.ClearCoreDataRootOverride") } catch {}
    }
    foreach ($wb in $openedWorkbooks) {
        try { $wb.Close($false) } catch {}
        Release-ComObject $wb
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        Release-ComObject $excel
    }
}

Write-JobLog -Path $LogPath -Message ("JobType={0} {1}" -f $JobType, $resultText)
Write-Output $resultText

if ($resultText.StartsWith("FAIL|", [System.StringComparison]::OrdinalIgnoreCase)) {
    exit 1
}

exit 0
