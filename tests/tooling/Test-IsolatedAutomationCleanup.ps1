[CmdletBinding()]
param([string]$RepoRoot='.',[switch]$IncludeExpiredNestedReferences)
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before blank automation calibration.'}
. (Join-Path $PSScriptRoot 'IsolatedAutomationCleanup.ps1')
$root=Join-Path $repo ('reports/runtime/isolated-automation-cleanup/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
Write-Output ('Evidence: '+$root)
Add-Type @'
using System; using System.Runtime.InteropServices;
public static class CleanupCalibrationOwner {
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr hwnd,out uint owner);
}
'@
$excel=$null;$ownedProcess=$null;$forced=$false
$checks=[Collections.Generic.List[object]]::new()
try {
    $excel=New-Object -ComObject Excel.Application
    [uint32]$owner=0;[void][CleanupCalibrationOwner]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$owner)
    $ownedProcess=Get-Process -Id $owner
    $excel.DisplayAlerts=$false
    $books=$excel.Workbooks
    $book=$books.Add()
    $sheets=$book.Worksheets
    $sheet=$sheets.Item(1)
    $cell=$sheet.Range('A1')
    $cell.Value2='Disposable cleanup calibration'
    $references=[Collections.Generic.List[object]]::new()
    foreach($entry in @($books,$book,$sheets,$sheet,$cell)){$references.Add($entry)}
    $cycle=@{Alias=$references;Scalar='CLEANUP_REDACTION_SENTINEL'}
    $matrix=New-Object 'object[,]' 2,2
    $matrix[0,0]=$book;$matrix[1,1]=$cell
    $cycle.Matrix=$matrix
    $cycle.Self=$cycle
    $references.Add($cycle)
    if($IncludeExpiredNestedReferences){
        # Smoke retains closed/released peer workbooks inside its collections.
        # Exercise that shape separately from the expired application variable.
        $expiredBook=$books.Add()
        $retired=@{List=[Collections.Generic.List[object]]::new();Map=@{Closed=$expiredBook}}
        $retired.List.Add($expiredBook)
        $references.Add($retired)
        $expiredBook.Close($false)
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($expiredBook)
        $expiredBook=$null
    }
    $book.Close($false);$excel.Quit()
    $expiredApplication=$excel
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel);$excel=$null
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
    $retained=-not $ownedProcess.WaitForExit(2000)
    $checks.Add([pscustomobject]@{Check='Baseline.ReferenceRetentionReproduced';Passed=$retained})
    $release=Release-IsolatedAutomationVariables -Variables (Get-Variable -Scope Script)
    $expectedReferences=if($IncludeExpiredNestedReferences){6}else{5}
    $checks.Add([pscustomobject]@{Check='Cleanup.UniqueReferencesDespiteAliasesAndCycles';Passed=($release.UniqueComReferences -eq $expectedReferences)})
    $checks.Add([pscustomobject]@{Check='Cleanup.NoReleaseFailures';Passed=($release.ReleaseFailures -eq 0)})
    $checks.Add([pscustomobject]@{Check='Cleanup.AlreadyReleasedReferenceSkipped';Passed=(($release.AlreadyReleasedVariables+$release.AlreadyReleased) -ge 1)})
    if($IncludeExpiredNestedReferences){
        $checks.Add([pscustomobject]@{Check='Cleanup.ExpiredNestedReferencesSkipped';Passed=($release.AlreadyReleased -ge 1)})
    }
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
    $checks.Add([pscustomobject]@{Check='Cleanup.NormalProcessExit';Passed=$ownedProcess.WaitForExit(10000)})
    $again=Release-IsolatedAutomationVariables -Variables (Get-Variable -Scope Script)
    $checks.Add([pscustomobject]@{Check='Cleanup.RepeatedReleaseSafe';Passed=($again.ReleaseFailures -eq 0)})
    $safe=$release|ConvertTo-Json
    $checks.Add([pscustomobject]@{Check='Cleanup.MetadataOnly';Passed=($safe -notmatch 'SENTINEL|[A-Za-z]:\\|Exception|Value')})
    $safe|Set-Content (Join-Path $root 'release.json')
} catch {
    $facts=[pscustomobject]@{ExceptionType=$_.Exception.GetType().Name;HResult=$_.Exception.HResult;File=(Split-Path -Leaf $_.InvocationInfo.ScriptName);Line=$_.InvocationInfo.ScriptLineNumber}
    $facts|ConvertTo-Json|Set-Content (Join-Path $root 'failure.json')
    $facts|ConvertTo-Json
    throw
} finally {
    if($null -ne $excel){try{$excel.Quit()}catch{};[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
    if($null -ne $ownedProcess -and -not $ownedProcess.HasExited){$forced=$true;Stop-Process -Id $ownedProcess.Id -Force}
    $checks.Add([pscustomobject]@{Check='Cleanup.NoForcedTermination';Passed=(-not $forced)})
    $checks|ConvertTo-Json|Set-Content (Join-Path $root 'checks.json')
}
$checks|Format-Table -AutoSize
Write-Output ('Evidence: '+$root)
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
