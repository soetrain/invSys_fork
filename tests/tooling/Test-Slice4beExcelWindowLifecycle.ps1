# Supplemental host control: synthetic workbook/add-in only; no invSys runtime.
[CmdletBinding()]
param([string]$RepoRoot='.',[ValidateSet('Com','Macro','MacroWithVbe','MacroWithWorkbook')][string]$Mode='Macro')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be closed.'}
. (Join-Path $PSScriptRoot 'Slice4beGuideResourceTrace.ps1')
$root=Join-Path (Resolve-Path $RepoRoot).Path ('reports/runtime/excel-window-lifecycle/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
Write-Output ('Host control: '+$root)
function Mark([string]$Stage){Get-GuideResourceSample $Stage|ConvertTo-Json -Compress|Add-Content (Join-Path $root 'resources.jsonl')}
function Close-Host($App){
    $owned=@(Get-Process EXCEL)
    if($owned.Count -ne 1){throw 'Host ownership unavailable.'}
    $App.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($App)
    [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
    if(-not $owned[0].WaitForExit(30000)){throw 'Host did not close normally.'}
}
$excel=New-Object -ComObject Excel.Application
$excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
$data=Join-Path $root 'synthetic-read.xlsx';$probe=Join-Path $root 'synthetic-probe.xlam'
$book=$excel.Workbooks.Add()
$book.SaveAs($data,51);$book.Close($false)
[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
$book=$excel.Workbooks.Add()
$module=$book.VBProject.VBComponents.Add(1);$module.Name='HostControl'
$module.CodeModule.AddFromString(@'
Option Explicit
Public Function ReadAndClose(ByVal path As String) As Boolean
    Dim book As Workbook
    Set book = Application.Workbooks.Open(path, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
    ReadAndClose = book.ReadOnly
    book.Close SaveChanges:=False
    Set book = Nothing
End Function
'@)
$book.IsAddin=$true;$book.SaveAs($probe,55);$book.Close($false)
[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($module);$module=$null
[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
Close-Host $excel;$excel=$null
$before=(Get-FileHash -LiteralPath $data).Hash
$start=[DateTimeOffset]::UtcNow
$excel=New-Object -ComObject Excel.Application
$excel.Visible=$true;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1
$addin=$null;$held=$null
try {
    if($Mode -ne 'Com'){$addin=$excel.Workbooks.Open($probe,0,$true)}
    if($Mode -eq 'MacroWithVbe'){$vbe=$excel.VBE;[void]$vbe.MainWindow.Visible;[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($vbe);$vbe=$null}
    if($Mode -eq 'MacroWithWorkbook'){
        $operator=Join-Path $root 'synthetic-operator.xlsx'
        Copy-Item -LiteralPath $data -Destination $operator
        $held=$excel.Workbooks.Open($operator,0,$true)
    }
    Mark 'Baseline'
    for($index=1;$index -le 12;$index++){
        if($Mode -eq 'Com'){
            $book=$excel.Workbooks.Open($data,0,$true);$book.Close($false)
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book);$book=$null
        } else {
            $ok=$excel.Run("'synthetic-probe.xlam'!HostControl.ReadAndClose",$data)
            if($ok -isnot [bool] -or -not $ok){throw 'Synthetic read failed.'}
        }
        Mark ('AfterRead'+$index)
    }
    if($null -ne $held){$held.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($held);$held=$null}
    if($null -ne $addin){$addin.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($addin);$addin=$null}
} finally {
    Close-Host $excel;$excel=$null
    [pscustomobject]@{Mode=$Mode;StartUTC=$start.ToString('o');EndUTC=[DateTimeOffset]::UtcNow.ToString('o');ExcelClosed=@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0;SourcePreserved=(Get-FileHash -LiteralPath $data).Hash -ceq $before;ProductAcceptance=$false}|ConvertTo-Json|Set-Content (Join-Path $root 'result.json')
}
