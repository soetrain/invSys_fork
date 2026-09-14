# D18 interruption proof: retain a real unclosed journal across distinct Excel
# processes. Termination is restricted to the current disposable fixture runtime.
# No reset seam, deleted Close record, saved instrumentation or persisted secret.
function Test-Slice4beRecordingRestart($Fixture) {
    if(-not ('RecordingRestartProcess' -as [type])) {
        Add-Type @'
using System;
using System.Runtime.InteropServices;
public static class RecordingRestartProcess {
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    public static uint Owner(long window) {
        if(window == 0) throw new InvalidOperationException("Excel window unavailable.");
        uint process;
        if(GetWindowThreadProcessId(new IntPtr(window), out process) == 0 || process == 0)
            throw new InvalidOperationException("Excel process unavailable.");
        return process;
    }
}
'@
    }
    function CopyRestartProbe([string]$Package,[string]$Module,[string[]]$Procedures) {
        $code=$packages[$Package].VBProject.VBComponents.Item($Module).CodeModule
        $parts=@(foreach($procedure in $Procedures){$code.Lines($code.ProcStartLine($procedure,0),$code.ProcCountLines($procedure,0))})
        [pscustomobject]@{Package=$Package;Module=$Module;Text=($parts -join "`r`n")}
    }
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    SetRecordingPolicy $true
    OpenRecordingViewer
    if((RecordingControl 'Start Recording' 'Click') -cne 'DELIVERED'){throw 'Restart fixture Start control is unavailable.'}
    $action=SaveRecordedSetting '631'
    $sequence=[string]$action.Attempt.SequenceId
    $entries=@(RecordingJournal $sequence|Sort-Object Version)
    $prepared=(HasSequence $action 1) -and (JournalChain $sequence 3) -and
        @($entries|Where-Object RecordType -CEQ 'Close').Count -eq 0 -and
        (RecordingControl 'Stop Recording') -ceq 'True|True'
    Check 'RecordingRestart.RealActionDurableBeforeInterruption' $prepared
    if(-not $prepared){throw 'Refusing process interruption without the real active journal fixture.'}
    $pathId=[string]$entries[0].ActionPathId
    $probes=@(
        (CopyRestartProbe 'invSys.Operations.xlam' 'frmInventoryViewer' @('RecordingControlForTest','RecordingStatusForTest')),
        (CopyRestartProbe 'invSys.Operations.xlam' 'modInventoryViewer' @('RecordingControlForTest','RecordingStatusForTest','RecordingLibraryForTest')),
        (CopyRestartProbe 'invSys.Admin.xlam' 'frmAdminSettings' @('D5TestSave')),
        (CopyRestartProbe 'invSys.Admin.xlam' 'TestD5Commands' @('OpenSettings','CloseSettings','SaveSettings'))
    )
    $rootPath=[IO.Path]::GetFullPath($runRoot).TrimEnd('\')+'\'
    $tempPath=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $rootPath.StartsWith($tempPath,[StringComparison]::OrdinalIgnoreCase) -or
        (Split-Path $runRoot -Leaf) -notlike 'invsys-config-command-*') {throw 'Restart requires the generated disposable fixture root.'}
    $packageNames=@('invSys.Core.xlam','invSys.Inventory.Domain.xlam','invSys.Designs.Domain.xlam','invSys.Operations.xlam','invSys.Admin.xlam')
    $packagePins=@{}
    foreach($name in $packageNames){$path=Join-Path $deploy $name;$packagePins[$path]=(Get-FileHash -LiteralPath $path).Hash}
    # Verify every loaded project and workbook, including hidden add-ins, before
    # intentionally interrupting this runtime. Never terminate an unknown project.
    $projects=$excel.VBE.VBProjects
    if($projects.Count -ne 5){throw 'Restart requires exactly the five candidate projects.'}
    foreach($project in $projects){
        $path=[IO.Path]::GetFullPath([string]$project.FileName)
        if(-not $packagePins.ContainsKey($path)){throw 'Unknown loaded project prevents interruption.'}
    }
    foreach($book in @($excel.Workbooks)){
        $path=[IO.Path]::GetFullPath([string]$book.FullName)
        if($packagePins.ContainsKey($path)){continue}
        if(-not $path.StartsWith($rootPath,[StringComparison]::OrdinalIgnoreCase) -or -not $book.Saved){
            throw 'Unknown or unsaved non-package workbook prevents interruption.'
        }
    }
    $owned=@(Get-Process EXCEL -ErrorAction Stop)
    $owner=[RecordingRestartProcess]::Owner([long]$excel.Hwnd)
    if($owned.Count -ne 1 -or $initialExcelProcessIds.Count -ne 1 -or
        $owned[0].Id -ne $owner -or $owner -ne $initialExcelProcessIds[0]){throw 'Excel ownership changed before interruption.'}
    [void]$owned[0].Handle
    $authorityBefore=RestartPins $Fixture.Root
    $otherBefore=RestartPins $b.Root
    $journalBefore=RestartPins $journalRoot
    # Keep the exact Process handle; PID reuse or a new unrelated instance is
    # not a termination target. No Quit/form close/SignOut is called beforehand.
    if($owned[0].HasExited -or [RecordingRestartProcess]::Owner([long]$excel.Hwnd) -ne $owner -or
        @(Get-Process EXCEL -ErrorAction Stop).Count -ne 1){throw 'Excel ownership changed at the interruption boundary.'}
    $owned[0].Kill()
    if(-not $owned[0].WaitForExit(5000)){throw 'Interrupted fixture process has not exited.'}
    try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)}catch{}
    $script:excel=$null; $script:packages=@{}
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Another Excel process prevents cold restart.'}
    Check 'RecordingRestart.InterruptionPreservesUnclosedJournal' (
        (RestartPinsEqual $journalBefore $journalRoot) -and (JournalChain $sequence 3))
    $state=@{Fixture=$Fixture;OtherFixture=$b;Deploy=$deploy;RunRoot=$runRoot;ReportRoot=$reportRoot;
        PackageNames=$packageNames;PackagePins=$packagePins;Probes=$probes;OldExcelId=$owner;
        ParentControllerId=$PID;Action=$action;Sequence=$sequence;PathId=$pathId;
        AuthorityBefore=$authorityBefore;OtherBefore=$otherBefore;JournalBefore=$journalBefore}
    $workerPath=Join-Path $PSScriptRoot 'Slice4beRecordingRestartWorker.ps1'
    $startInfo=[Diagnostics.ProcessStartInfo]::new()
    $startInfo.FileName=Join-Path $PSHOME 'powershell.exe'
    $startInfo.Arguments='-NoProfile -ExecutionPolicy Bypass -File "'+$workerPath+'"'
    $startInfo.UseShellExecute=$false; $startInfo.CreateNoWindow=$true
    $startInfo.RedirectStandardInput=$true; $startInfo.RedirectStandardOutput=$true; $startInfo.RedirectStandardError=$true
    $worker=[Diagnostics.Process]::new(); $worker.StartInfo=$startInfo
    if(-not $worker.Start()){throw 'Fresh reader controller could not start.'}
    $errors=$worker.StandardError.ReadToEndAsync()
    try {
        # Credentials, paths and fixture observations never enter command-line
        # arguments, files, logs or worker output. Only this private pipe carries them.
        $worker.StandardInput.WriteLine(($state|ConvertTo-Json -Depth 40 -Compress))
        $worker.StandardInput.Close()
        $badProtocol=$false
        while($null -ne ($line=$worker.StandardOutput.ReadLine())){
            try {
                $message=$line|ConvertFrom-Json
                if($message.Type -ceq 'Check' -and $message.Name -match '^RecordingRestart\.[A-Za-z]+$' -and $message.Passed -is [bool]){
                    Check $message.Name $message.Passed
                } elseif($message.Type -ceq 'Stage' -and $message.Name -in @('Pristine Viewer','Install drivers','Read journal')){
                    Write-Output ('Fresh recording reader: '+$message.Name)
                } else {$badProtocol=$true}
            } catch {$badProtocol=$true}
        }
        $worker.WaitForExit()
        if($badProtocol -or $errors.Result.Length -gt 0 -or $worker.ExitCode -ne 0){throw 'Fresh recording reader failed; see its sanitized check/stage evidence.'}
        if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Fresh recording reader left Excel open.'}
    } finally {
        $state=$null
        $worker.Dispose()
    }
}
