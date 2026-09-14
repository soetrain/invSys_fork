[CmdletBinding()]
param([string]$ClientPipeName='')
$ErrorActionPreference='Stop'
. (Join-Path $PSScriptRoot 'Slice4beRecordingTransfer.ps1')
if($ClientPipeName){
    if($ClientPipeName -notmatch '^invsys-recording-[0-9a-f]{32}$'){throw 'Invalid test pipe.'}
    $client=[IO.Pipes.NamedPipeClientStream]::new('.',$ClientPipeName,[IO.Pipes.PipeDirection]::Out)
    try{$client.Connect(10000);Write-RecordingHandoff $client (Read-RecordingStandardInput)}finally{$client.Dispose()}
    exit 0
}
$name='invsys-recording-'+[guid]::NewGuid().ToString('N')
$pipe=New-RecordingHandoffServer $name
$child=$null;$reader=$null
$before=@(Get-Process EXCEL -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$data=('NOT-A-CREDENTIAL-TRANSFER-PROBE-'+[guid]::NewGuid().ToString('N')+[char]0x263a)*4096
$checks=@()
try{
    $rules=@($pipe.GetAccessControl().GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier]))
    $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User
    $checks+=[pscustomobject]@{Check='RecordingHandoff.CurrentUserOnly';Passed=($rules.Count -eq 1 -and $rules[0].IdentityReference -eq $sid -and $rules[0].AccessControlType -eq 'Allow')}
    $connect=$pipe.WaitForConnectionAsync()
    $info=[Diagnostics.ProcessStartInfo]::new()
    $info.FileName=Join-Path $PSHOME 'powershell.exe'
    $info.Arguments='-NoProfile -ExecutionPolicy Bypass -File "'+$PSCommandPath+'" -ClientPipeName '+$name
    $info.UseShellExecute=$false;$info.CreateNoWindow=$true
    $info.RedirectStandardInput=$true;$info.RedirectStandardOutput=$true;$info.RedirectStandardError=$true
    $child=[Diagnostics.Process]::new();$child.StartInfo=$info
    if(-not $child.Start()){throw 'Transfer test client did not start.'}
    $output=$child.StandardOutput.ReadToEndAsync();$errors=$child.StandardError.ReadToEndAsync()
    if(-not $connect.Wait(10000)){throw 'Transfer test client did not connect.'}
    $reader=[IO.StreamReader]::new($pipe,[Text.Encoding]::UTF8)
    $read=$reader.ReadToEndAsync()
    Write-RecordingHandoff $child.StandardInput.BaseStream $data
    $child.StandardInput.Close()
    $received=$read.GetAwaiter().GetResult()
    $child.WaitForExit()
    $checks+=[pscustomobject]@{Check='RecordingHandoff.LargeUnicodeTransferExact';Passed=($received -ceq $data)}
    $checks+=[pscustomobject]@{Check='RecordingHandoff.ChildExitedBeforeUse';Passed=($child.HasExited -and $child.ExitCode -eq 0 -and $child.Id -ne $PID)}
    $checks+=[pscustomobject]@{Check='RecordingHandoff.NoPublicOutput';Passed=($output.Result.Length -eq 0 -and $errors.Result.Length -eq 0 -and -not $info.Arguments.Contains($data))}
}finally{
    if($reader){$reader.Dispose()};$pipe.Dispose()
    if($child){if(-not $child.HasExited){$child.Kill()};$child.Dispose()}
}
$after=@(Get-Process EXCEL -ErrorAction SilentlyContinue|Select-Object -ExpandProperty Id)
$checks+=[pscustomobject]@{Check='RecordingHandoff.NoExcelCreated';Passed=(($before -join ',') -ceq ($after -join ','))}
$checks|ConvertTo-Json
if(@($checks|Where-Object {-not $_.Passed}).Count){exit 1}
