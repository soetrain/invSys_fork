[CmdletBinding()]
param()
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before the isolated cleanup calibration.'}
. (Join-Path $PSScriptRoot 'ProductionReusableCleanup.ps1')
function Release-ComObject($Value){if($null -ne $Value){try{[void][Runtime.InteropServices.Marshal]::ReleaseComObject($Value)}catch{}}}
Add-Type @'
using System;using System.Runtime.InteropServices;
public static class ReusableCleanupCalibration {
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint owner);
}
'@
$id=[guid]::NewGuid().ToString('N')
$reportRoot=Join-Path $PWD ('reports/runtime/reusable-cleanup-calibration/'+$id)
$fixtureRoot=Join-Path ([IO.Path]::GetTempPath()) ('invsys-plan022-launcher-red-'+$id)
New-Item -ItemType Directory -Path $reportRoot,$fixtureRoot|Out-Null
Write-Output ('Calibration: '+$reportRoot)
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed})}
$excel=$null;$owner=$null;$forced=$false
try{
    $excel=New-Object -ComObject Excel.Application
    $excel.DisplayAlerts=$false
    [uint32]$ownerId=0;[void][ReusableCleanupCalibration]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$ownerId)
    $owner=Get-Process -Id $ownerId
    $inside=$excel.Workbooks.Add();$inside.SaveAs((Join-Path $fixtureRoot 'inside.xlsb'),50)
    $outside=$excel.Workbooks.Add();$outside.Worksheets.Item(1).Cells.Item(1,1).Value2='Owned calibration sentinel'
    $outside.SaveAs((Join-Path $reportRoot 'outside.xlsb'),50)
    $rejected=$false
    try{[void](Close-ReusableFixtureWorkbooks -Excel $excel -RuntimeRoot $fixtureRoot)}catch{$rejected=$true}
    Check 'Ownership.OutOfRootWorkbookRejected' $rejected
    Check 'Ownership.NoWorkbookClosedBeforeWholeSetValidation' ([int]$excel.Workbooks.Count -eq 2)
    Check 'Ownership.OwnedWorkbookStillOpen' ([string]$inside.Name -ceq 'inside.xlsb')
    Check 'Ownership.ForeignFixtureValuePreserved' ([string]$outside.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq 'Owned calibration sentinel')
    # This outside-root workbook was created by this calibration; close it explicitly.
    $outside.Close($false);Release-ComObject $outside
    $closed=Close-ReusableFixtureWorkbooks -Excel $excel -RuntimeRoot $fixtureRoot
    Check 'Closure.OwnedWorkbookClosed' ($closed.ClosedWorkbooks -eq 1 -and $closed.RemainingWorkbooks -eq 0)
    $excel.Quit();Release-ComObject $excel;$excel=$null
    $exitReceipt=Wait-ReusableAutomationExit -ProcessId $owner.Id -Variables (Get-Variable -Scope Script)
    $exitReceipt|ConvertTo-Json -Depth 4|Set-Content (Join-Path $reportRoot 'exit.json')
    Check 'Closure.NormalExitWithoutReleaseFailure' ($exitReceipt.UnassistedExitObserved -and $null -eq $exitReceipt.Failure -and $exitReceipt.References.ReleaseFailures -eq 0)
}finally{
    if($null -ne $excel){try{$excel.Quit()}catch{};Release-ComObject $excel}
    if($null -ne $owner -and -not $owner.HasExited){$forced=$true;Stop-Process -Id $owner.Id -Force}
    Check 'Closure.NoForcedTermination' (-not $forced)
    $checks|ConvertTo-Json|Set-Content (Join-Path $reportRoot 'checks.json')
}
$checks|Format-Table -AutoSize
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
