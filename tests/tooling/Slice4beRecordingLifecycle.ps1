# Test-controller lifecycle only. Sensitive restoration/fixture state stays in
# the caller's memory until every controller and Excel process has exited.
function Get-InvSysTestSettingsSnapshot {
    $root='HKCU:\Software\VB and VBA Program Settings\invSys'; $snapshot=@{}
    if(Test-Path -LiteralPath $root){
        foreach($key in @(Get-Item -LiteralPath $root)+@(Get-ChildItem -LiteralPath $root -Recurse)){
            $values=@{}
            foreach($name in $key.GetValueNames()){$values[$name]=@($key.GetValue($name),$key.GetValueKind($name))}
            $snapshot[$key.Name]=$values
        }
    }
    return $snapshot
}

function Restore-InvSysTestSettingsSnapshot {
    param([hashtable]$Snapshot)
    if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must exit before local settings restoration.'}
    $root='HKCU:\Software\VB and VBA Program Settings\invSys'
    if(Test-Path -LiteralPath $root){
        foreach($key in @(Get-Item -LiteralPath $root)+@(Get-ChildItem -LiteralPath $root -Recurse)){
            foreach($name in $key.GetValueNames()){
                if(-not $Snapshot.ContainsKey($key.Name) -or -not $Snapshot[$key.Name].ContainsKey($name)){
                    Remove-ItemProperty -LiteralPath ('Registry::'+$key.Name) -Name $name
                }
            }
        }
    }
    foreach($path in $Snapshot.Keys){foreach($name in $Snapshot[$path].Keys){
        $saved=$Snapshot[$path][$name]
        New-ItemProperty -LiteralPath ('Registry::'+$path) -Name $name -Value $saved[0] -PropertyType $saved[1] -Force|Out-Null
    }}
    $actual=Get-InvSysTestSettingsSnapshot
    $expectedCount=0; $actualCount=0
    foreach($values in $Snapshot.Values){$expectedCount+=$values.Count}
    foreach($values in $actual.Values){$actualCount+=$values.Count}
    if($actualCount -ne $expectedCount){return $false}
    foreach($path in $Snapshot.Keys){foreach($name in $Snapshot[$path].Keys){
        if(-not $actual.ContainsKey($path) -or -not $actual[$path].ContainsKey($name)){return $false}
        $before=$Snapshot[$path][$name]; $after=$actual[$path][$name]
        if($before[1] -ne $after[1]){return $false}
        $beforeValues=@($before[0]); $afterValues=@($after[0])
        if($beforeValues.Count -ne $afterValues.Count){return $false}
        for($i=0;$i -lt $beforeValues.Count;$i++){if($beforeValues[$i] -cne $afterValues[$i]){return $false}}
    }}
    return $true
}

function Wait-RecordingCleanup {
    param($Creator, $Worker)
    $announced=$false
    while(($null -ne $Creator -and -not $Creator.HasExited) -or
          ($null -ne $Worker -and -not $Worker.HasExited) -or
          @(Get-Process EXCEL -ErrorAction SilentlyContinue).Count -gt 0){
        if(-not $announced){
            Write-Output 'Recording cleanup waiting for remaining processes; restoration state retained in memory.'
            $announced=$true
        }
        Start-Sleep -Milliseconds 500
    }
}

function Get-RecordingFailureFacts {
    param([Management.Automation.ErrorRecord]$Failure, [string]$LastMacro='')
    # Never copy exception text, arguments, fixture state or authentication data.
    # LastMacro is a developer boundary identifier, not proof that it failed.
    if($LastMacro -cnotmatch '^invSys\.(Core|Operations|Admin|Inventory\.Domain|Designs\.Domain)\.xlam\|[A-Za-z_][A-Za-z0-9_]*\.[A-Za-z_][A-Za-z0-9_]*$'){
        $LastMacro='Unavailable'
    }
    [pscustomobject]@{
        LastAttemptedMacro=$LastMacro
        ExceptionType=$Failure.Exception.GetType().FullName
        HResult=$Failure.Exception.HResult
        InnerHResult=if($null -ne $Failure.Exception.InnerException){$Failure.Exception.InnerException.HResult}else{$null}
    }
}
