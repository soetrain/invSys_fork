[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$RepoRoot = ".",

    [Parameter(Mandatory = $false)]
    [string]$RunnerScriptPath = "tools/run_wan_scheduler_job.ps1",

    [Parameter(Mandatory = $false)]
    [string]$TaskPrefix = "invSys",

    [Parameter(Mandatory = $false)]
    [ValidateSet("WarehouseBatch", "WarehousePublish")]
    [string]$WarehouseJobType = "WarehousePublish",

    [Parameter(Mandatory = $false)]
    [string]$WarehouseId = "",

    [Parameter(Mandatory = $false)]
    [string]$WarehouseTaskTime = "06:00",

    [Parameter(Mandatory = $false)]
    [string]$HqTaskTime = "06:15",

    [Parameter(Mandatory = $false)]
    [string]$SharePointRoot = "",

    [Parameter(Mandatory = $false)]
    [string]$HqOutputPath = "",

    [Parameter(Mandatory = $false)]
    [string]$RuntimeRootOverride = "",

    [Parameter(Mandatory = $false)]
    [string]$LogRoot = "logs/scheduler",

    [switch]$Apply
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Quote-TaskArg {
    param([string]$Value)
    if ($null -eq $Value) { return '""' }
    return '"' + $Value.Replace('"', '\"') + '"'
}

function Build-TaskCommand {
    param([string[]]$Arguments)

    return ($Arguments | ForEach-Object { Quote-TaskArg $_ }) -join " "
}

$repoPath = (Resolve-Path $RepoRoot).Path
$runnerFullPath = if ([System.IO.Path]::IsPathRooted($RunnerScriptPath)) { $RunnerScriptPath } else { Join-Path $repoPath $RunnerScriptPath }
$runnerFullPath = (Resolve-Path $runnerFullPath).Path
$logRootPath = if ([System.IO.Path]::IsPathRooted($LogRoot)) { $LogRoot } else { Join-Path $repoPath $LogRoot }

$taskSpecs = New-Object System.Collections.Generic.List[object]

if (-not [string]::IsNullOrWhiteSpace($WarehouseId)) {
    $warehouseLog = Join-Path $logRootPath ("warehouse_{0}.log" -f $WarehouseId)
    $warehouseArgs = @(
        "powershell.exe",
        "-ExecutionPolicy", "Bypass",
        "-File", $runnerFullPath,
        "-RepoRoot", $repoPath,
        "-JobType", $WarehouseJobType,
        "-WarehouseId", $WarehouseId,
        "-LogPath", $warehouseLog
    )
    if (-not [string]::IsNullOrWhiteSpace($SharePointRoot)) {
        $warehouseArgs += @("-SharePointRoot", $SharePointRoot)
    }
    if (-not [string]::IsNullOrWhiteSpace($RuntimeRootOverride)) {
        $warehouseArgs += @("-RuntimeRootOverride", $RuntimeRootOverride)
    }

    $taskSpecs.Add([pscustomobject]@{
        Name    = ("{0}.Warehouse.{1}.{2}" -f $TaskPrefix, $WarehouseJobType, $WarehouseId)
        Time    = $WarehouseTaskTime
        Command = (Build-TaskCommand -Arguments $warehouseArgs)
    }) | Out-Null
}

$hqLog = Join-Path $logRootPath "hq_aggregation.log"
$hqArgs = @(
    "powershell.exe",
    "-ExecutionPolicy", "Bypass",
    "-File", $runnerFullPath,
    "-RepoRoot", $repoPath,
    "-JobType", "HqAggregation",
    "-LogPath", $hqLog
)
if (-not [string]::IsNullOrWhiteSpace($SharePointRoot)) {
    $hqArgs += @("-SharePointRoot", $SharePointRoot)
}
if (-not [string]::IsNullOrWhiteSpace($HqOutputPath)) {
    $hqArgs += @("-OutputPath", $HqOutputPath)
}
if (-not [string]::IsNullOrWhiteSpace($RuntimeRootOverride)) {
    $hqArgs += @("-RuntimeRootOverride", $RuntimeRootOverride)
}

$taskSpecs.Add([pscustomobject]@{
    Name    = ("{0}.HQ.Aggregation" -f $TaskPrefix)
    Time    = $HqTaskTime
    Command = (Build-TaskCommand -Arguments $hqArgs)
}) | Out-Null

foreach ($task in $taskSpecs) {
    $schtasksArgs = @("/Create", "/F", "/SC", "DAILY", "/TN", $task.Name, "/ST", $task.Time, "/TR", $task.Command)
    if ($Apply) {
        & schtasks.exe @schtasksArgs | Out-Host
    }
    else {
        Write-Output ("schtasks.exe {0}" -f (($schtasksArgs | ForEach-Object { Quote-TaskArg $_ }) -join " "))
    }
}
