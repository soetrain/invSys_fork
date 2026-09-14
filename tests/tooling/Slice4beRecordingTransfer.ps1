# Private, current-user-only transfer for disposable recording fixtures.
function New-RecordingHandoffServer([string]$Name){
    if($Name -notmatch '^invsys-recording-[0-9a-f]{32}$'){throw 'Invalid recording pipe name.'}
    $security=[IO.Pipes.PipeSecurity]::new()
    $security.SetAccessRuleProtection($true,$false)
    $sid=[Security.Principal.WindowsIdentity]::GetCurrent().User
    $security.AddAccessRule([IO.Pipes.PipeAccessRule]::new($sid,[IO.Pipes.PipeAccessRights]::ReadWrite,[Security.AccessControl.AccessControlType]::Allow))
    return [IO.Pipes.NamedPipeServerStream]::new($Name,[IO.Pipes.PipeDirection]::In,1,[IO.Pipes.PipeTransmissionMode]::Byte,[IO.Pipes.PipeOptions]::Asynchronous,4096,4096,$security)
}
function Write-RecordingHandoff($Pipe,[string]$Text){
    $writer=[IO.StreamWriter]::new($Pipe,[Text.UTF8Encoding]::new($false),4096,$true)
    try{$writer.Write($Text);$writer.Flush()}finally{$writer.Dispose()}
}
function Read-RecordingStandardInput {
    $reader=[IO.StreamReader]::new([Console]::OpenStandardInput(),[Text.UTF8Encoding]::new($false),$true,4096,$true)
    try{return $reader.ReadToEnd()}finally{$reader.Dispose()}
}
