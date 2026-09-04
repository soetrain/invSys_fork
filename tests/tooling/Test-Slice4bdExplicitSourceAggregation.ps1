[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$CoreXlamPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Release-ComObject {
    param([object]$Object)
    if ($null -ne $Object) {
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Object) } catch {}
    }
}

function New-Snapshot {
    param([object]$Excel, [string]$Path, [string]$WarehouseId, [string]$SystemKey, [double]$Quantity)

    $wb = $Excel.Workbooks.Add()
    $ws = $wb.Worksheets.Item(1)
    $ws.Name = "InventorySnapshot"
    $headers = @("WarehouseId", "System_Key", "SKU", "QtyOnHand", "LastAppliedAtUTC")
    for ($i = 0; $i -lt $headers.Count; $i++) { $ws.Cells.Item(1, $i + 1).Value2 = $headers[$i] }
    $ws.Cells.Item(2, 1).Value2 = $WarehouseId
    $ws.Cells.Item(2, 2).Value2 = $SystemKey
    $ws.Cells.Item(2, 3).Value2 = "SKU-" + $WarehouseId
    $ws.Cells.Item(2, 4).Value2 = $Quantity
    $ws.Cells.Item(2, 5).Value2 = "2026-09-03 00:00:00"
    $range = $ws.Range("A1", "E2")
    $table = $ws.ListObjects.Add(1, $range, $null, 1)
    $table.Name = "tblInventorySnapshot"
    $wb.SaveAs($Path, 50)
    $wb.Close($false)
    Release-ComObject $table
    Release-ComObject $ws
    Release-ComObject $wb
}

$core = (Resolve-Path -LiteralPath $CoreXlamPath).Path
$root = Join-Path ([IO.Path]::GetTempPath()) ("invsys-slice4bd-explicit-" + [guid]::NewGuid().ToString("N"))
$excel = $null
$wbCore = $null
$wbHarness = $null
$wbOutput = $null

try {
    New-Item -ItemType Directory -Path $root -Force | Out-Null
    $sourceA = Join-Path $root "WHA.invSys.Snapshot.Inventory.xlsb"
    $sourceB = Join-Path $root "WHB.invSys.Snapshot.Inventory.xlsb"
    $output = Join-Path $root "Global\invSys.Global.InventorySnapshot.xlsb"

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    New-Snapshot -Excel $excel -Path $sourceA -WarehouseId "WHA" -SystemKey "KEY-WHA-1" -Quantity 5
    New-Snapshot -Excel $excel -Path $sourceB -WarehouseId "WHB" -SystemKey "KEY-WHB-1" -Quantity 8

    $wbCore = $excel.Workbooks.Open($core, 0, $true)
    $wbHarness = $excel.Workbooks.Add()
    $harnessName = [string]$wbHarness.Name
    [void]$wbHarness.VBProject.References.AddFromFile($core)
    $component = $wbHarness.VBProject.VBComponents.Add(1)
    $component.Name = "modExplicitSourceHarness"
    $component.CodeModule.AddFromString(@'
Public Function AggregateExplicitSources(ByVal sourceList As String, ByVal outputPath As String) As String
    Dim report As String
    If modHqAggregator.GenerateGlobalSnapshotFromFiles(sourceList, outputPath, report) Then
        AggregateExplicitSources = "OK|" & report
    Else
        AggregateExplicitSources = "FAIL|" & report
    End If
End Function
'@)
    $result = [string]$excel.Run("'$harnessName'!modExplicitSourceHarness.AggregateExplicitSources", ($sourceA + "`n" + $sourceB), $output)
    if (-not $result.StartsWith("OK|")) { throw "Explicit aggregation failed: $result" }

    $wbOutput = $excel.Workbooks.Open($output, 0, $true)
    $table = $wbOutput.Worksheets.Item("GlobalInventorySnapshot").ListObjects.Item("tblGlobalInventorySnapshot")
    $sourceColumn = $table.ListColumns.Item("SourceSnapshot").DataBodyRange
    $sourceValues = @($sourceColumn.Value2 | ForEach-Object { [string]$_ })
    $passed = ($table.ListRows.Count -eq 2) -and ($sourceValues -contains $sourceA) -and ($sourceValues -contains $sourceB)
    [pscustomobject]@{
        Check = "Slice4bd.ExplicitSourceAggregation"
        Passed = $passed
        Detail = "Rows=$($table.ListRows.Count); ExplicitSources=$($sourceValues.Count); Result=$result"
    } | Format-Table -AutoSize
    if (-not $passed) { throw "Explicit aggregation did not preserve both selected source identities." }
}
finally {
    if ($null -ne $wbOutput) { try { $wbOutput.Close($false) } catch {}; Release-ComObject $wbOutput }
    if ($null -ne $wbHarness) { try { $wbHarness.Close($false) } catch {}; Release-ComObject $wbHarness }
    if ($null -ne $wbCore) { try { $wbCore.Close($false) } catch {}; Release-ComObject $wbCore }
    if ($null -ne $excel) { try { $excel.Quit() } catch {}; Release-ComObject $excel }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
