[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CoreXlamPath,

    [Parameter(Mandatory = $true)]
    [string]$AdminXlamPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch {}
    }
}

$core = (Resolve-Path -LiteralPath $CoreXlamPath).Path
$admin = (Resolve-Path -LiteralPath $AdminXlamPath).Path
$excel = $null
$wbCore = $null
$wbAdmin = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $wbCore = $excel.Workbooks.Open($core, 0, $true)
    $wbAdmin = $excel.Workbooks.Open($admin, 0, $true)
    $result = [string]$excel.Run("'$($wbAdmin.Name)'!modAdmin.Admin_AggregationSourcesFormSmokeForAutomation")
    $passed = $result.StartsWith("OK|", [StringComparison]::Ordinal)
    [pscustomobject]@{
        Check = "Slice4bd.AdminFormSmoke"
        Passed = $passed
        Detail = $result
    } | Format-Table -AutoSize
    if (-not $passed) { throw "Admin source-set form smoke failed: $result" }
}
finally {
    if ($null -ne $wbAdmin) { try { $wbAdmin.Close($false) } catch {}; Release-ComObject $wbAdmin }
    if ($null -ne $wbCore) { try { $wbCore.Close($false) } catch {}; Release-ComObject $wbCore }
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
