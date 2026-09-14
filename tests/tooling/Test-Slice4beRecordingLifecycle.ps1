# Isolated control-flow/redaction checks: no Excel, registry or fixture writes.
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
. (Join-Path $PSScriptRoot 'Slice4beRecordingLifecycle.ps1')
$results=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})}
& {
    $creator=[pscustomobject]@{HasExited=$false}
    $worker=[pscustomobject]@{HasExited=$false}
    $script:polls=0; $restored=$false
    function Get-Process { if($script:polls -lt 3){[pscustomobject]@{Id=1}} }
    function Start-Sleep {
        if($restored){throw 'Restoration ran beneath a live process.'}
        $script:polls++
        if($script:polls -ge 1){$creator.HasExited=$true}
        if($script:polls -ge 2){$worker.HasExited=$true}
    }
    $messages=@(Wait-RecordingCleanup -Creator $creator -Worker $worker)
    $restored=$true
    Check 'RecordingLifecycle.WaitRetainsControlUntilAllProcessesExit' ($script:polls -eq 3 -and $restored)
    Check 'RecordingLifecycle.OneFixedWaitingNotice' ($messages.Count -eq 1 -and $messages[0] -ceq 'Recording cleanup waiting for remaining processes; restoration state retained in memory.')
    $script:polls=3
    $messages=@(Wait-RecordingCleanup -Creator $creator -Worker $worker)
    Check 'RecordingLifecycle.ClosedProcessesNeedNoWait' ($script:polls -eq 3 -and $messages.Count -eq 0)
}
& {
    $root='HKEY_CURRENT_USER\Software\VB and VBA Program Settings\invSys'
    $script:fakeKeys=@{}; $script:settingsExcelOpen=$false; $script:settingWrites=0
    function New-FakeKey([string]$Name){
        $key=[pscustomobject]@{Name=$Name;Values=@{};Kinds=@{}}
        $key|Add-Member ScriptMethod GetValueNames {return @($this.Values.Keys)}
        $key|Add-Member ScriptMethod GetValue {param($name) return $this.Values[$name]}
        $key|Add-Member ScriptMethod GetValueKind {param($name) return $this.Kinds[$name]}
        $script:fakeKeys[$Name]=$key
        return $key
    }
    function KeyName([string]$Path){$Path.Replace('Registry::','').Replace('HKCU:','HKEY_CURRENT_USER')}
    function Test-Path {param($LiteralPath) return $script:fakeKeys.ContainsKey((KeyName $LiteralPath))}
    function Get-Item {param($LiteralPath) return $script:fakeKeys[(KeyName $LiteralPath)]}
    function Get-ChildItem {param($LiteralPath,[switch]$Recurse) foreach($key in $script:fakeKeys.Values){if($key.Name -ne (KeyName $LiteralPath)){$key}}}
    function Get-Process {if($script:settingsExcelOpen){[pscustomobject]@{Id=1}}}
    function Remove-ItemProperty {param($LiteralPath,$Name) $key=Get-Item $LiteralPath;$key.Values.Remove($Name);$key.Kinds.Remove($Name);$script:settingWrites++}
    function New-ItemProperty {param($LiteralPath,$Name,$Value,$PropertyType,[switch]$Force) $key=Get-Item $LiteralPath;$key.Values[$Name]=$Value;$key.Kinds[$Name]=$PropertyType;$script:settingWrites++}
    [void](New-FakeKey $root)
    $key=New-FakeKey ($root+'\Runtime')
    $key.Values['UserChoice']='MixedCase';$key.Kinds['UserChoice']=[Microsoft.Win32.RegistryValueKind]::String
    $key.Values['UnknownBinary']=[byte[]]@(0,127,255);$key.Kinds['UnknownBinary']=[Microsoft.Win32.RegistryValueKind]::Binary
    $snapshot=Get-InvSysTestSettingsSnapshot
    $key.Values['UserChoice']='fixture';$key.Values['FixtureOnly']='temporary';$key.Kinds['FixtureOnly']=[Microsoft.Win32.RegistryValueKind]::String
    $key.Values['UnknownBinary']=[byte[]]@(1)
    $script:settingsExcelOpen=$true;$refused=$false
    try{[void](Restore-InvSysTestSettingsSnapshot $snapshot)}catch{$refused=$true}
    Check 'RecordingLifecycle.LiveExcelPreventsSettingWrites' ($refused -and $script:settingWrites -eq 0)
    $script:settingsExcelOpen=$false
    $verified=Restore-InvSysTestSettingsSnapshot $snapshot
    Check 'RecordingLifecycle.SettingsRestoreVerifiesValuesAndTypes' ($verified -and $key.Values['UserChoice'] -ceq 'MixedCase' -and $key.Kinds['UserChoice'] -eq [Microsoft.Win32.RegistryValueKind]::String)
    Check 'RecordingLifecycle.UnknownBinarySettingPreserved' (($key.Values['UnknownBinary'] -join ',') -ceq '0,127,255' -and $key.Kinds['UnknownBinary'] -eq [Microsoft.Win32.RegistryValueKind]::Binary)
    Check 'RecordingLifecycle.FixtureOnlySettingRemoved' (-not $key.Values.ContainsKey('FixtureOnly') -and $key.Values.Count -eq 2)
}
$sentinel='NOT-A-CREDENTIAL-REDACTION-PROBE-'+[guid]::NewGuid().ToString('N')
try{throw [InvalidOperationException]::new($sentinel)}catch{$failure=$_}
$facts=Get-RecordingFailureFacts -Failure $failure -LastMacro 'invSys.Operations.xlam|modInventoryViewer.RecordingControlForTest'
$json=$facts|ConvertTo-Json -Compress
Check 'RecordingLifecycle.FailureFactsExcludeExceptionText' (-not $json.Contains($sentinel) -and @($facts.PSObject.Properties).Count -eq 4 -and $facts.HResult -is [int])
Check 'RecordingLifecycle.DeclaredMacroRetained' ($facts.LastAttemptedMacro -ceq 'invSys.Operations.xlam|modInventoryViewer.RecordingControlForTest')
$facts=Get-RecordingFailureFacts -Failure $failure -LastMacro $sentinel
Check 'RecordingLifecycle.UnregisteredDiagnosticTextExcluded' ($facts.LastAttemptedMacro -ceq 'Unavailable' -and -not ($facts|ConvertTo-Json -Compress).Contains($sentinel))
$results|ConvertTo-Json
if(@($results|Where-Object {-not $_.Passed}).Count){exit 1}
