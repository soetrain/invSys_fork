# Validate private-pipe handling without starting Excel or creating authority.
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$workerPath=Join-Path $PSScriptRoot 'Slice4beRecordingRestartWorker.ps1'
$sentinel='NOT-A-CREDENTIAL-REDACTION-PROBE-'+[guid]::NewGuid().ToString('N')
$root=Join-Path ([IO.Path]::GetTempPath()) ('invsys-config-command-protocol-'+[guid]::NewGuid().ToString('N'))
$state=@{Fixture=@{Root=(Join-Path $root 'a');Secret=$sentinel};OtherFixture=@{Root=(Join-Path $root 'b')};
    Deploy=$sentinel;RunRoot=$root;ReportRoot=$sentinel;PackageNames=@();PackagePins=@{};Probes=@();
    OldExcelId=1;ParentControllerId=$PID;Action=@{};Sequence=$sentinel;PathId=$sentinel;
    AuthorityBefore=@{};OtherBefore=@{};JournalBefore=@{}}
$before=@(Get-Process EXCEL -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$results=@()
foreach($case in @('Valid','Malformed','MissingField','UnexpectedField','EscapedRoot')){
    $copy=$state.Clone()
    switch($case){
        'MissingField'{$copy.Remove('Action')}
        'UnexpectedField'{$copy.Add('Unexpected',$sentinel)}
        'EscapedRoot'{$copy.Fixture=@{Root=[IO.Path]::GetTempPath();Secret=$sentinel}}
    }
    $inputText=if($case -eq 'Malformed'){'{'+$sentinel}else{$copy|ConvertTo-Json -Depth 10 -Compress}
    $info=[Diagnostics.ProcessStartInfo]::new()
    $info.FileName=Join-Path $PSHOME 'powershell.exe'
    $info.Arguments='-NoProfile -ExecutionPolicy Bypass -File "'+$workerPath+'" -ValidateInputOnly'
    $info.UseShellExecute=$false; $info.CreateNoWindow=$true
    $info.RedirectStandardInput=$true; $info.RedirectStandardOutput=$true; $info.RedirectStandardError=$true
    $worker=[Diagnostics.Process]::new(); $worker.StartInfo=$info
    if(-not $worker.Start()){throw 'Protocol worker could not start.'}
    try{
        $stdoutTask=$worker.StandardOutput.ReadToEndAsync(); $stderrTask=$worker.StandardError.ReadToEndAsync()
        $worker.StandardInput.WriteLine($inputText); $worker.StandardInput.Close()
        if(-not $worker.WaitForExit(15000)){$worker.Kill();throw 'Read-only protocol worker timed out.'}
        $stdout=$stdoutTask.Result; $stderr=$stderrTask.Result
        $reply=$stdout|ConvertFrom-Json
        $valid=$case -eq 'Valid'
        $safe=-not ($stdout.Contains($sentinel) -or $stdout.Contains($root) -or $stderr.Contains($sentinel) -or
            $stderr.Contains($root) -or $info.Arguments.Contains($sentinel))
        $passed=$safe -and $stderr.Length -eq 0 -and $reply.Type -ceq 'Check' -and $reply.Passed -is [bool] -and
            $reply.Passed -eq $valid -and $worker.ExitCode -eq $(if($valid){0}else{1}) -and
            $reply.Name -ceq $(if($valid){'RecordingRestart.WorkerInputValidated'}else{'RecordingRestart.WorkerHarnessFailure'})
        $results+=[pscustomobject]@{Check=('RecordingWorkerProtocol.'+$case);Passed=$passed}
    }finally{$worker.Dispose()}
}
$after=@(Get-Process EXCEL -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$results+=[pscustomobject]@{Check='RecordingWorkerProtocol.NoExcelOrFixtureCreated';Passed=(($before -join ',') -ceq ($after -join ',') -and -not (Test-Path -LiteralPath $root))}
$results|ConvertTo-Json
if(@($results|Where-Object {-not $_.Passed}).Count){exit 1}
