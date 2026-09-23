# Diagnostic controls for the Release 1 shutdown failure, not slice acceptance.
[CmdletBinding()]
param(
    [ValidateSet('Empty','Packages','TablesCollected','Restart','RestartCollected','RestartConfigured','RestartSnapshot','RestartRefreshed','RestartCustomHeader','RestartHeaders','RestartRead')][string]$Case='Empty',
    [string]$RepoRoot='.',
    [string]$DeployRoot='deploy/validation-settings-diagnostic',
    [switch]$Worker,
    [string]$ReportRoot='',
    [ValidateRange(180,300)][int]$ObservationSeconds=210
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$deploy=(Resolve-Path -LiteralPath (Join-Path $repo $DeployRoot)).Path
function Write-Json([string]$Name,[object]$Value){ConvertTo-Json -InputObject $Value -Depth 5|Set-Content -LiteralPath (Join-Path $ReportRoot $Name)}
function Release-Com([object]$Value){if($null -ne $Value){[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Value)}}
if($Worker){
    if(-not (Test-Path -LiteralPath $ReportRoot -PathType Container)){throw 'Controller report root required.'}
    Add-Type @'
using System;using System.Runtime.InteropServices;
public static class ShutdownControlOwner {
 [DllImport("user32.dll")]public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
}
'@
    $excel=$null;$books=[Collections.Generic.List[object]]::new();$failed=$false
    if($Case.StartsWith('Restart',[StringComparison]::Ordinal)){
        $receipt=Join-Path $repo 'reports/runtime/settings-diagnostic-chain-passive-phase6_live_role_workflow_results.md'
        $body=Get-Content -LiteralPath $receipt -Raw
        $match=[regex]::Match($body,'(?m)^- Runtime root override:\s*(.+?)\s*$')
        if(-not $match.Success){throw 'Generated live fixture reference unavailable.'}
        $source=(Resolve-Path -LiteralPath $match.Groups[1].Value.Trim()).Path
        $systemTemp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
        if(-not $source.StartsWith($systemTemp+'\invsys-phase6-live-',[StringComparison]::OrdinalIgnoreCase)){throw 'Only generated live test data is permitted.'}
        $clone=Join-Path $systemTemp ('invsys-shutdown-control-'+[IO.Path]::GetFileName($ReportRoot))
        if(Test-Path -LiteralPath $clone){throw 'Preserve existing diagnostic fixture.'}
        New-Item -ItemType Directory -Path $clone|Out-Null
        foreach($item in Get-ChildItem -LiteralPath $source){Copy-Item -LiteralPath $item.FullName -Destination $clone -Recurse}
        Write-Json 'fixture-provenance.json' ([pscustomobject]@{SourceReceiptHash=(Get-FileHash -LiteralPath $receipt).Hash;GeneratedFixtureOnly=$true;ClonedFiles=@(Get-ChildItem -LiteralPath $clone -Recurse -File).Count})
        $tokens=$null;$errors=$null
        $harnessPath=Join-Path $repo 'tools/validate_release1_full_chain.ps1'
        $ast=[Management.Automation.Language.Parser]::ParseFile($harnessPath,[ref]$tokens,[ref]$errors)
        if($errors.Count){throw 'Original chain harness does not parse.'}
        $functions=$ast.FindAll({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst]},$false)
        foreach($definition in $functions){
            $text=$definition.Extent.Text
            if($definition.Name -ceq 'Invoke-RestartReconciliation'){
                $created='$localExcel = New-Object -ComObject Excel.Application'
                $quit='try { $localExcel.Quit() } catch {}'
                if([regex]::Matches($text,[regex]::Escape($created)).Count -ne 1 -or [regex]::Matches($text,[regex]::Escape($quit)).Count -ne 1){throw 'Original lifecycle anchors differ.'}
                $text=$text.Replace($created,$created+"`r`n"+'Save-RestartControlOwner $localExcel')
                $text=$text.Replace($quit,'Write-Json ''before-quit.json'' ([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString(''o'');WorkbookCount=[int]$localExcel.Workbooks.Count})'+"`r`n"+$quit+"`r`n"+'Write-Json ''after-quit.json'' ([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString(''o'')})')
                if($Case -ceq 'RestartConfigured'){
                    $cut='$receiveBook = @($operatorBooks | Where-Object { $_.Name -like "*.Receiving.Operator.xlsb" })[0]'
                    if([regex]::Matches($text,[regex]::Escape($cut)).Count -ne 1){throw 'Configured phase anchor differs.'}
                    $text=$text.Replace($cut,'Write-Json ''phase-cut.json'' ([pscustomobject]@{Phase=''AfterLoadConfig'';PackageCount=$packageMap.Count;OperatorCount=$operatorBooks.Count;RestartAssertionsSkipped=$true})'+"`r`n"+'return'+"`r`n"+$cut)
                }
                if($Case -ceq 'RestartSnapshot'){
                    $cut='$refreshOk = $true'
                    if([regex]::Matches($text,[regex]::Escape($cut)).Count -ne 1){throw 'Snapshot phase anchor differs.'}
                    $text=$text.Replace($cut,'Add-Result ''Diagnostic.SnapshotCompleted'' $snapshotOk ''Diagnostic phase only.'''+"`r`n"+'Write-Json ''phase-cut.json'' ([pscustomobject]@{Phase=''AfterSnapshot'';RestartAssertionsSkipped=$true})'+"`r`n"+'return'+"`r`n"+$cut)
                }
                if($Case -ceq 'RestartRefreshed'){
                    $cut='Add-Result "FinalRefresh"'
                    if([regex]::Matches($text,[regex]::Escape($cut)).Count -ne 1){throw 'Refresh phase anchor differs.'}
                    $text=$text.Replace($cut,'Add-Result ''Diagnostic.RefreshCompleted'' ($snapshotOk -and $refreshOk) ''Diagnostic phase only.'''+"`r`n"+'Write-Json ''phase-cut.json'' ([pscustomobject]@{Phase=''AfterRefresh'';RestartAssertionsSkipped=$true})'+"`r`n"+'return'+"`r`n"+$cut)
                }
                if($Case -ceq 'RestartRead'){
                    $cut='$logTable = Get-ListObject -Workbook $inventoryWorkbook'
                    if([regex]::Matches($text,[regex]::Escape($cut)).Count -ne 1){throw 'Read phase anchor differs.'}
                    $text=$text.Replace($cut,'Write-Json ''phase-cut.json'' ([pscustomobject]@{Phase=''BeforeReplay'';LaterRestartAssertionsSkipped=$true})'+"`r`n"+'return'+"`r`n"+$cut)
                }
                if($Case -ceq 'RestartHeaders'){
                    $cut='$entityTable = Get-ListObject -Workbook $inventoryWorkbook'
                    if([regex]::Matches($text,[regex]::Escape($cut)).Count -ne 1){throw 'Header phase anchor differs.'}
                    $text=$text.Replace($cut,'Write-Json ''phase-cut.json'' ([pscustomobject]@{Phase=''AfterHeaders'';LaterRestartAssertionsSkipped=$true})'+"`r`n"+'return'+"`r`n"+$cut)
                }
                if($Case -ceq 'RestartCustomHeader'){
                    $cut='$noRowHeaders = '
                    if([regex]::Matches($text,[regex]::Escape($cut)).Count -ne 1){throw 'Custom header phase anchor differs.'}
                    $text=$text.Replace($cut,'Write-Json ''phase-cut.json'' ([pscustomobject]@{Phase=''BeforeNoRowEnumeration'';LaterRestartAssertionsSkipped=$true})'+"`r`n"+'return'+"`r`n"+$cut)
                }
                Write-Json 'instrumentation.json' ([pscustomobject]@{HarnessHash=(Get-FileHash -LiteralPath $harnessPath).Hash;CreatedAnchors=1;QuitAnchors=1;OriginalAssertionsUnchanged=($Case -cin @('Restart','RestartCollected'));DiagnosticPhaseCut=($Case -cnotin @('Restart','RestartCollected'))})
            }
            . ([scriptblock]::Create($text))
        }
        function Save-RestartControlOwner([object]$Application){
            [uint32]$owner=0
            [void][ShutdownControlOwner]::GetWindowThreadProcessId([IntPtr]$Application.Hwnd,[ref]$owner)
            $process=Get-Process -Id $owner
            Write-Json 'created.json' ([pscustomobject]@{ProcessId=$owner;CreatedUTC=$process.StartTime.ToUniversalTime().ToString('o');UTC=[DateTimeOffset]::UtcNow.ToString('o')})
        }
        $deployPath=$deploy;$results=[Collections.Generic.List[object]]::new()
        try {
            Invoke-RestartReconciliation -LiveResultText ('- Runtime root override: '+$clone)
            Write-Json 'restart-checks.json' @($results|Select-Object Check,Passed)
            $failed=@($results|Where-Object {-not $_.Passed}).Count -gt 0
            if($Case -cne 'Restart'){
                Write-Json 'collection-start.json' ([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');OriginalProcedureReturned=$true})
                [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
                Write-Json 'collection-end.json' ([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o')})
            }
        } catch {
            $failed=$true
            Write-Json 'worker-failure.json' ([pscustomobject]@{Type=$_.Exception.GetBaseException().GetType().FullName;HResult=('0x{0:X8}' -f $_.Exception.GetBaseException().HResult)})
        }
    } else {
    try {
        $excel=New-Object -ComObject Excel.Application
        [uint32]$owner=0
        [void][ShutdownControlOwner]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$owner)
        $process=Get-Process -Id $owner
        Write-Json 'created.json' ([pscustomobject]@{ProcessId=$owner;CreatedUTC=$process.StartTime.ToUniversalTime().ToString('o');UTC=[DateTimeOffset]::UtcNow.ToString('o')})
        $excel.Visible=$false;$excel.DisplayAlerts=$false
        $initialCount=[int]$excel.Workbooks.Count
        if($initialCount -ne 0){throw 'Baseline contains unexpected workbooks.'}
        if($Case -ceq 'Packages'){
            foreach($name in @('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')){
                $book=$excel.Workbooks.Open((Join-Path $deploy $name))
                $books.Add($book)
            }
        }
        if($Case -ceq 'TablesCollected'){
            function Add-TableReferenceControl([object]$Application){
                $book=$Application.Workbooks.Add()
                $books.Add($book)
                $sheet=$book.Worksheets.Item(1)
                $sheet.Range('A1').Value2='Key'
                $sheet.Range('B1').Value2='Value'
                $sheet.Range('A2').Value2='TEST'
                $sheet.Range('B2').Value2=1
                $table=$sheet.ListObjects.Add(1,$sheet.Range('A1:B2'),$null,1)
                $column=$table.ListColumns.Add()
                $column.Name='Custom_Control'
                $column.DataBodyRange.Cells.Item(1,1).Value2='PRESERVE'
                Write-Json 'table-checks.json' ([pscustomobject]@{PureExcel=$true;NoInvSysMacros=$true;ValueReadBack=([string]$table.DataBodyRange.Cells.Item(1,3).Value2 -ceq 'PRESERVE')})
            }
            Add-TableReferenceControl $excel
        }
        Write-Json 'loaded.json' ([pscustomobject]@{Case=$Case;InitialWorkbookCount=$initialCount;ExplicitPackageCount=$books.Count;WorkbookCount=[int]$excel.Workbooks.Count})
    } catch {
        $failed=$true
        Write-Json 'worker-failure.json' ([pscustomobject]@{Type=$_.Exception.GetBaseException().GetType().FullName;HResult=('0x{0:X8}' -f $_.Exception.GetBaseException().HResult)})
    } finally {
        $closing=$books.ToArray();[Array]::Reverse($closing)
        foreach($book in $closing){$book.Close($false);Release-Com $book}
        if($null -ne $excel){
            Write-Json 'before-quit.json' ([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');WorkbookCount=[int]$excel.Workbooks.Count})
            $excel.Quit()
            Release-Com $excel
            Write-Json 'after-quit.json' ([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o')})
        }
    }
    }
    if($Case -ceq 'TablesCollected'){
        Write-Json 'collection-start.json' ([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o')})
        [GC]::Collect();[GC]::WaitForPendingFinalizers();[GC]::Collect()
        Write-Json 'collection-end.json' ([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o')})
    }
    # Keep the original controller alive, but never reattach or call Quit again.
    $until=[DateTimeOffset]::UtcNow.AddSeconds($ObservationSeconds+45)
    while(-not (Test-Path (Join-Path $ReportRoot 'release-worker.txt')) -and [DateTimeOffset]::UtcNow -lt $until){Start-Sleep -Milliseconds 250}
    if($failed){exit 1}
    exit 0
}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Close Excel before a diagnostic control.'}
if($ReportRoot){throw 'Controller creates a unique report root.'}
$ReportRoot=Join-Path $repo ('reports/runtime/slice4be-shutdown-control/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $ReportRoot|Out-Null
$start=[DateTimeOffset]::UtcNow
$packagePins=Get-Content (Join-Path $repo 'reports/runtime/settings-diagnostic-package-pins.json') -Raw|ConvertFrom-Json
foreach($pin in $packagePins){if((Get-FileHash -LiteralPath (Join-Path $deploy $pin.Package)).Hash -cne $pin.Hash){throw 'Frozen candidate differs.'}}
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$settings=Get-InvSysTestSettingsSnapshot
$restorationComplete=$false
$child=$null
try {
Write-Json 'start.json' ([pscustomobject]@{UTC=$start.ToString('o');Case=$Case;ObservationSeconds=$ObservationSeconds;DiagnosticOnly=$true;SourceHash=(Get-FileHash -LiteralPath $PSCommandPath).Hash})
$arguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File',('"'+$PSCommandPath+'"'),'-Worker','-Case',$Case,'-RepoRoot',('"'+$repo+'"'),'-DeployRoot',('"'+$DeployRoot+'"'),'-ReportRoot',('"'+$ReportRoot+'"'),'-ObservationSeconds',$ObservationSeconds)
$child=Start-Process -FilePath powershell.exe -ArgumentList $arguments -WindowStyle Hidden -PassThru -RedirectStandardOutput (Join-Path $ReportRoot 'worker.stdout.log') -RedirectStandardError (Join-Path $ReportRoot 'worker.stderr.log')
$null=$child.Handle
Write-Output ('Report: '+$ReportRoot)
$dispatchDeadline=[DateTimeOffset]::UtcNow.AddSeconds(120)
while(-not (Test-Path (Join-Path $ReportRoot 'after-quit.json')) -and -not $child.HasExited -and [DateTimeOffset]::UtcNow -lt $dispatchDeadline){Start-Sleep -Milliseconds 250}
if(-not (Test-Path (Join-Path $ReportRoot 'after-quit.json'))){throw 'Worker has not completed the original Quit; preserve its process and inspect.'}
$identity=Get-Content (Join-Path $ReportRoot 'created.json') -Raw|ConvertFrom-Json
$quit=Get-Content (Join-Path $ReportRoot 'after-quit.json') -Raw|ConvertFrom-Json
$deadline=[DateTimeOffset]::Parse($quit.UTC).AddSeconds($ObservationSeconds)
$samples=[Collections.Generic.List[object]]::new()
do {
    $process=Get-Process -Id $identity.ProcessId -ErrorAction SilentlyContinue
    $present=$null -ne $process -and $process.StartTime.ToUniversalTime().ToString('o') -ceq $identity.CreatedUTC
    $samples.Add([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');SameProcessPresent=$present})
    if(-not $present){break}
    Start-Sleep -Seconds 2
}while([DateTimeOffset]::UtcNow -lt $deadline)
Write-Json 'native-observation.json' $samples.ToArray()
$closedWhileWorkerAlive=-not $present -and -not $child.HasExited
'Release'|Set-Content -LiteralPath (Join-Path $ReportRoot 'release-worker.txt')
if(-not $child.WaitForExit(30000)){throw 'Worker exit pending; preserve it and inspect.'}
$child.Refresh()
Write-Json 'worker-exit.json' ([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');ExitCode=$child.ExitCode})
for($i=0;$i -lt 15 -and (Get-Process EXCEL -ErrorAction SilentlyContinue);$i++){Start-Sleep -Seconds 2}
if(Get-Process EXCEL -ErrorAction SilentlyContinue){
    Write-Json 'cleanup-wait.json' ([pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');NormalClosureUnavailable=$true})
    Wait-RecordingCleanup -Creator $null -Worker $child
}
$restored=Restore-InvSysTestSettingsSnapshot $settings
$restorationComplete=$true
$audit=[DateTimeOffset]::UtcNow
$errors=@()
$events=@(Get-WinEvent -FilterHashtable @{LogName='Application';Id=1000,1001,1002;StartTime=$start.LocalDateTime;EndTime=$audit.LocalDateTime} -ErrorAction SilentlyContinue -ErrorVariable errors)
if(@($errors|Where-Object FullyQualifiedErrorId -NotLike 'NoMatchingEventsFound*').Count){throw 'Event audit unavailable.'}
foreach($pin in $packagePins){if((Get-FileHash -LiteralPath (Join-Path $deploy $pin.Package)).Hash -cne $pin.Hash){throw 'Frozen package changed.'}}
$result=[pscustomobject]@{Case=$Case;StartUTC=$start.ToString('o');AfterQuitUTC=$quit.UTC;AuditUTC=$audit.ToString('o');WorkerExit=$child.ExitCode;ClosedWhileWorkerAlive=$closedWhileWorkerAlive;ApplicationEvents=$events.Count;ApplicationEventIds=@($events|ForEach-Object Id);SettingsRestored=$restored;PackageHashesPreserved=$true;Passed=($child.ExitCode -eq 0 -and $closedWhileWorkerAlive -and $events.Count -eq 0 -and $restored);DiagnosticOnly=$true;FullSliceAccepted=$false}
Write-Json 'result.json' $result
$result|ConvertTo-Json -Depth 4
if(-not $result.Passed){exit 1}
} finally {
    if(-not $restorationComplete){
        Wait-RecordingCleanup -Creator $null -Worker $child
        $restored=Restore-InvSysTestSettingsSnapshot $settings
        Write-Json 'failure-path-settings-restoration.json' ([pscustomobject]@{Restored=$restored;UTC=[DateTimeOffset]::UtcNow.ToString('o')})
    }
}
