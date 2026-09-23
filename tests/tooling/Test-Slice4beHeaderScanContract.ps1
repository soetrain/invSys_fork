# Generic disposable Excel tables only; no managed inventory or XLAM mutation.
[CmdletBinding()]
param([string]$RepoRoot='.',[ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before calibration.'}
$root=Join-Path $repo ('reports/runtime/header-scan-contract/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
Write-Output ('Report: '+$root)
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_release1_full_chain.ps1'),[ref]$tokens,[ref]$errors)
$definition=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Test-NoRowHeaders'},$false)
if($errors.Count -or $null -eq $definition){throw 'Actual header scan unavailable.'}
. ([scriptblock]::Create($definition.Extent.Text))
$checks=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$checks.Add([pscustomobject]@{Check=$Name;Passed=$Passed});Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))}
function Scan-Is([bool]$Expected){try{return (Test-NoRowHeaders -Workbooks @($first,$last)) -eq $Expected}catch{return $false}}
$excel=$null;$first=$null;$last=$null;$firstSheet=$null;$lastSheet=$null;$safeTable=$null;$otherTable=$null;$target=$null;$column=$null
$start=[DateTimeOffset]::UtcNow
try {
    $excel=New-Object -ComObject Excel.Application
    $excel.Visible=$false;$excel.DisplayAlerts=$false
    $first=$excel.Workbooks.Add();$last=$excel.Workbooks.Add()
    $firstSheet=$first.Worksheets.Item(1);$lastSheet=$last.Worksheets.Item($last.Worksheets.Count)
    $firstSheet.Range('A1').Value2='Allowed';$firstSheet.Range('A2').Value2='Control'
    $safeTable=$firstSheet.ListObjects.Add(1,$firstSheet.Range('A1:A2'),$null,1)
    $lastSheet.Range('A1').Value2='Allowed';$lastSheet.Range('A2').Value2='Control'
    $otherTable=$lastSheet.ListObjects.Add(1,$lastSheet.Range('A1:A2'),$null,1)
    $lastSheet.Range('D1').Value2='Allowed';$lastSheet.Range('E1').Value2='Custom'
    $lastSheet.Range('D2').Value2='Control';$lastSheet.Range('E2').Value2='Preserve'
    $target=$lastSheet.ListObjects.Add(1,$lastSheet.Range('D1:E2'),$null,1)
    $column=$target.ListColumns.Item(2)
    Check 'AllowedHeadersAcrossWorkbooks' (Scan-Is $true)
    $column.Name=' rOw '
    Check 'LastWorkbookLastTableMixedCaseForbiddenHeader' (Scan-Is $false)
    $target.ShowHeaders=$false
    Check 'HiddenForbiddenHeaderStillRejected' (Scan-Is $false)
    $target.ShowHeaders=$true
    $column.Name='Custom_Unknown'
    $target.ShowHeaders=$false
    Check 'HiddenAllowedHeaderAccepted' (Scan-Is $true)
    Check 'ReadDoesNotRevealHiddenHeaders' (-not $target.ShowHeaders)
    $target.ShowHeaders=$true
    Check 'ReadPreservesUnknownHeaderAndValue' ($column.Name -ceq 'Custom_Unknown' -and [string]$column.DataBodyRange.Cells.Item(1,1).Value2 -ceq 'Preserve')
} finally {
    foreach($value in @($column,$target,$otherTable,$safeTable,$lastSheet,$firstSheet)){
        if($null -ne $value){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}
    }
    foreach($book in @($last,$first)){if($null -ne $book){$book.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($book)}}
    if($null -ne $excel){$excel.Quit();[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}
}
[GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
for($i=0;$i -lt 15 -and (Get-Process EXCEL -ErrorAction SilentlyContinue);$i++){Start-Sleep -Seconds 2}
Check 'ExcelClosedWithoutAssistance' (@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -eq 0)
$queryErrors=@()
$events=@(Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1001,1002;StartTime=$start.LocalDateTime;EndTime=([DateTimeOffset]::UtcNow.LocalDateTime)} -ErrorAction SilentlyContinue -ErrorVariable queryErrors)
if(@($queryErrors|Where-Object FullyQualifiedErrorId -NotLike 'NoMatchingEventsFound*').Count){throw 'Event query unavailable.'}
Check 'NoApplicationFailureEvents' ($events.Count -eq 0)
ConvertTo-Json -InputObject @($checks.ToArray())|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
Write-Output ('Report: '+$root)
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
