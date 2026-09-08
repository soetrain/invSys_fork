[CmdletBinding()]
param(
    [string]$RepoRoot = '.',
    [string]$DeployRoot = 'deploy/current',
    [ValidateSet('RED','GREEN')][string]$Phase = 'GREEN'
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$repo = (Resolve-Path -LiteralPath $RepoRoot).Path
$deploy = (Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
if (Get-Process EXCEL -ErrorAction SilentlyContinue) { throw 'Close Excel before isolated packaged identity validation.' }
$path = Join-Path $deploy 'invSys.Core.xlam'
$before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
$results = [Collections.Generic.List[object]]::new()
$excel = $null; $core = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false; $excel.DisplayAlerts = $false
    $core = $excel.Workbooks.Open($path)
    $helper = $core.VBProject.VBComponents.Add(1)
    $helper.Name = 'TestSystemKeyGeneration'
    $helper.CodeModule.AddFromString(@'
Option Explicit
Public Function Burst(ByVal resetAmbientRandom As Boolean) As String
    Dim seen As Object, index As Long, value As String, duplicates As Long, blanks As Long, ignored As Single
    Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbBinaryCompare
    For index = 1 To 50000
        If resetAmbientRandom Then
            ignored = Rnd(-1)
            Randomize 1
        End If
        value = modRoleEventWriter.CreateSystemKey()
        If Len(value) = 0 Then blanks = blanks + 1
        If seen.Exists(value) Then
            duplicates = duplicates + 1
        Else
            seen.Add value, True
        End If
    Next index
    Burst = CStr(seen.Count) & "|" & CStr(duplicates) & "|" & CStr(blanks)
End Function
'@)
    foreach ($reset in @($false,$true)) {
        $counts = ([string]$excel.Run("'invSys.Core.xlam'!TestSystemKeyGeneration.Burst",$reset)).Split('|')
        if ($counts.Count -ne 3) { throw 'Identity probe returned an invalid count envelope.' }
        $name = if ($reset) { 'Identity.AmbientRandomIndependent' } else { 'Identity.BurstUnique' }
        $passed = [int]$counts[0] -eq 50000 -and [int]$counts[1] -eq 0 -and [int]$counts[2] -eq 0
        $results.Add([pscustomobject]@{ Check=$name; Passed=$passed; Attempted=50000; Unique=[int]$counts[0]; Duplicates=[int]$counts[1]; Blank=[int]$counts[2] })
        Write-Output ("{0}: {1}; unique={2}, duplicate={3}, blank={4}" -f $name,$passed,$counts[0],$counts[1],$counts[2])
    }
}
catch {
    $results.Add([pscustomobject]@{Check='Harness.Exception';Passed=$false})
    Write-Output ('Harness exception: '+$_.Exception.Message)
}
finally {
    if ($null -ne $core) { try { $core.Close($false) } catch {} }
    if ($null -ne $excel) { try { $excel.Quit() } catch {} }
    foreach ($com in @($core,$excel)) { if ($null -ne $com) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($com) } }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
$after = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
$results.Add([pscustomobject]@{Check='Identity.PackageUnchanged';Passed=($before -ceq $after)})
$output = Join-Path $repo 'reports/runtime/slice4be-system-key'
New-Item -ItemType Directory -Path $output -Force | Out-Null
$results | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $output ($Phase.ToLowerInvariant()+'.json')) -Encoding UTF8
$failed = @($results | Where-Object {-not $_.Passed}).Count
Write-Output ('Passed={0}; Failed={1}' -f ($results.Count-$failed),$failed)
if ($failed -gt 0) { exit 1 }
