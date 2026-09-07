[CmdletBinding()]
param([string]$RepoRoot='', [string]$DeployRoot='deploy/current', [string]$SourceReportPath='')
if([string]::IsNullOrWhiteSpace($RepoRoot)){$RepoRoot=Split-Path -Parent (Split-Path -Parent $PSScriptRoot)}
$deploy=(Resolve-Path -LiteralPath (Join-Path $RepoRoot $DeployRoot)).Path
$ErrorActionPreference='Stop'
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel first.'}
Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class InvSysCompileProcess {
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint processId);
}
'@
$excel=New-Object -ComObject Excel.Application
[uint32]$ownedProcessId=0
[void][InvSysCompileProcess]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$ownedProcessId)
$sourceHashes=@()
try {
    $excel.Visible=$false; $excel.DisplayAlerts=$false; $excel.EnableEvents=$false; $excel.AutomationSecurity=1
    # Probe a role package before preloading Core: preloading can mask a stale
    # absolute dependency path retained when a candidate package is copied.
    $probe=$excel.Workbooks.Open((Join-Path $deploy 'invSys.Operations.xlam'),0,$true)
    foreach($reference in $probe.VBProject.References){
        if($reference.Name -like 'invSys_*' -and -not [string]::Equals(
            (Split-Path -Parent $reference.FullPath),$deploy,[StringComparison]::OrdinalIgnoreCase)){
            throw 'Cold-start Operations dependency resolves outside the tested package directory.'
        }
    }
    Write-Output 'COLD START PASS invSys.Operations.xlam'
    $probe.Close($false)
    $books=@()
    foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')) {
        $path=(Resolve-Path (Join-Path $deploy $name)).Path
        $books+=,$excel.Workbooks.Open($path,0,$true)
    }
    foreach($book in $books) {
        $project=$book.VBProject
        foreach($reference in $project.References){
            if($reference.IsBroken){throw ('Broken reference: '+$book.Name)}
            if($reference.Name -like 'invSys_*'){
                $referenceRoot=Split-Path -Parent $reference.FullPath
                if(-not [string]::Equals($referenceRoot,$deploy,[StringComparison]::OrdinalIgnoreCase)){
                    throw ('Package reference resolves outside the tested package directory: '+$book.Name)
                }
            }
        }
        $excel.VBE.ActiveVBProject=$project
        foreach($component in $project.VBComponents){if($component.Type -eq 1){$component.CodeModule.CodePane.Show();break}}
        Start-Sleep -Milliseconds 300
        $control=$excel.VBE.CommandBars.FindControl(1,578)
        if($null -eq $control){throw 'Compile command unavailable.'}
        if($control.Enabled){$control.Execute()}
        Start-Sleep -Milliseconds 500
        $control=$excel.VBE.CommandBars.FindControl(1,578)
        if($control.Enabled){
            Write-Output ('Enabled caption: '+$control.Caption)
            $pane=$excel.VBE.ActiveCodePane
            [int]$sl=0;[int]$sc=0;[int]$el=0;[int]$ec=0
            $pane.GetSelection([ref]$sl,[ref]$sc,[ref]$el,[ref]$ec)
            Write-Output ('Selection: '+$pane.CodeModule.Name+':'+$sl+':'+$sc)
            throw ('Compile did not finish: '+$book.Name)
        }
        Write-Output ('COMPILE PASS '+$book.Name)
        if($SourceReportPath -ne ''){
            foreach($component in $project.VBComponents){
                $text=''
                if($component.CodeModule.CountOfLines -gt 0){$text=$component.CodeModule.Lines(1,$component.CodeModule.CountOfLines)}
                $sha=[Security.Cryptography.SHA256]::Create()
                try {$hash=[BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes(($text -replace "`r`n","`n")))).Replace('-','').ToLowerInvariant()} finally {$sha.Dispose()}
                $sourceHashes += [pscustomobject]@{Package=$book.Name;Component=$component.Name;CodeSha256=$hash}
            }
        }
    }
    if($SourceReportPath -ne ''){$sourceHashes | ConvertTo-Json | Set-Content -LiteralPath $SourceReportPath -Encoding UTF8}
} finally {
    foreach($book in @($excel.Workbooks)){try{$book.Close($false)}catch{}}
    $empty=($excel.Workbooks.Count -eq 0)
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    # VBE COM references can keep this isolated, empty Excel process alive.
    if($empty -and $ownedProcessId -gt 0){
        $owned=Get-Process -Id $ownedProcessId -ErrorAction SilentlyContinue
        if($null -ne $owned -and -not $owned.WaitForExit(1000)){Stop-Process -Id $ownedProcessId}
    }
}
