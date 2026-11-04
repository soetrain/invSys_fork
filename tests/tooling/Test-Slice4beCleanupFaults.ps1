[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$HarnessPath,
    [Parameter(Mandatory=$true)][string]$EvidenceRoot
)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
if(Test-Path -LiteralPath $EvidenceRoot){throw 'Use a new evidence directory; preserve previous results.'}
$source=[IO.File]::ReadAllText((Resolve-Path -LiteralPath $HarnessPath).Path)
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseInput($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Harness parse failure is not behavioral RED.'}
$match=[regex]::Matches($source,"(?ms)^        \`$hasSettingsHostEvidence=.*?(?=^        Write-InventoryShutdownTrace 'BeforeFinalRelease')")
if($match.Count -ne 1){throw 'Cleanup boundaries missing or ambiguous; not behavioral RED.'}
$blockAst=[Management.Automation.Language.Parser]::ParseInput($match[0].Value,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Cleanup block parse failure is not behavioral RED.'}
# Execute only the actual bounded cleanup block against fake objects.
# Never dot-source/run the harness, allocate COM, release COM, or call its GC.
$allowedCommands=@('Get-Command','Write-InventoryShutdownTrace','Write-SettingsHostEvidence','Check')
foreach($command in $blockAst.FindAll({param($n)$n -is [Management.Automation.Language.CommandAst]},$true)){
    if($command.GetCommandName() -notin $allowedCommands){throw 'Unexpected command in isolated cleanup block.'}
}
foreach($call in $blockAst.FindAll({param($n)$n -is [Management.Automation.Language.InvokeMemberExpressionAst]},$true)){
    if($call.Member.Value -notin @('Close','Quit')){throw 'Unexpected method in isolated cleanup block.'}
}
$checkAst=@($ast.FindAll({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Check'},$true))
if($checkAst.Count -ne 1){throw 'Expected the actual harness Check function.'}
$cleanup=[scriptblock]::Create($match[0].Value)
$actualCheck=[scriptblock]::Create($checkAst[0].Extent.Text)
$null=New-Item -ItemType Directory -Path $EvidenceRoot
$assertions=[Collections.Generic.List[object]]::new()
$observations=[Collections.Generic.List[object]]::new()
function Assert-Fixture([string]$Name,[bool]$Passed){
    $assertions.Add([pscustomobject]@{Check=$Name;Passed=$Passed})
    Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))
}
function Invoke-Fixture([string]$Name,[int]$BookCount,[int]$FailedBook,[bool]$QuitFails,[bool]$SettingsEvidence){
    $events=[Collections.Generic.List[string]]::new()
    $results=[Collections.Generic.List[object]]::new()
    $CheckGuideTransfer=$false
    $CheckSettingsDiagnostics=$SettingsEvidence
    . $actualCheck
    function Write-InventoryShutdownTrace([string]$Stage){$events.Add($Stage)}
    function Write-SettingsHostEvidence([string]$Stage,[bool]$ReadExcel=$true){
        $events.Add('Settings.'+$Stage+'.'+$ReadExcel)
    }
    $books=@()
    for($i=1;$i -le $BookCount;$i++){
        $fake=[pscustomobject]@{Index=$i;Fails=($i -eq $FailedBook);Events=$events;Calls=0;SaveArguments=@()}
        $fake|Add-Member -MemberType ScriptMethod -Name Close -Value {
            param($SaveChanges)
            $this.Calls++
            $this.SaveArguments+=,$SaveChanges
            $this.Events.Add('CloseAttempt'+$this.Index)
            if($this.Fails){throw 'SYNTHETIC_PRIVATE close failure C:\private\fixture.xlsx'}
        }
        $books+=,$fake
    }
    $excel=[pscustomobject]@{Workbooks=$books;Events=$events;Calls=0;Fails=$QuitFails}
    $excel|Add-Member -MemberType ScriptMethod -Name Quit -Value {
        $this.Calls++
        $this.Events.Add('QuitAttempt')
        if($this.Fails){throw 'SYNTHETIC_PRIVATE quit failure C:\private\fixture.xlsx'}
    }
    $uncaught=$false
    $output=@()
    try{$output=@(. $cleanup)}catch{$uncaught=$true}
    $expectedChecks=@()
    if($FailedBook -gt 0){$expectedChecks+='Harness.Exception.workbook cleanup'}
    if($QuitFails){$expectedChecks+='Harness.Exception.Excel quit'}
    $actualNames=@($results|ForEach-Object {$_.Check})
    Assert-Fixture ($Name+'.QuitAttemptedExactlyOnce') ($excel.Calls -eq 1)
    Assert-Fixture ($Name+'.ExactTypedFailureChecks') (($actualNames -join '|') -ceq ($expectedChecks -join '|') -and
        @($results|Where-Object {$_.Passed -isnot [bool] -or $_.Passed}).Count -eq 0)
    Assert-Fixture ($Name+'.NoExceptionEscapesToSkipLaterCleanup') (!$uncaught)
    # Preserve the pre-existing close-loop stop-on-first-error policy: this fix
    # changes Quit reachability/reporting, not close retries or COM ownership.
    $closeLimit=if($FailedBook -gt 0){$FailedBook}else{$BookCount}
    $closeValid=$true
    foreach($fake in $books){
        $expectedCalls=[int]($fake.Index -le $closeLimit)
        if($fake.Calls -ne $expectedCalls){$closeValid=$false}
        foreach($save in $fake.SaveArguments){if($save -isnot [bool] -or $save){$closeValid=$false}}
    }
    Assert-Fixture ($Name+'.CloseOrderNoSaveNoRetry') $closeValid
    $expectedEvents=[Collections.Generic.List[string]]::new()
    for($i=1;$i -le $closeLimit;$i++){
        $expectedEvents.Add('BeforeWorkbookClose'+$i)
        $expectedEvents.Add('CloseAttempt'+$i)
        if($i -ne $FailedBook){$expectedEvents.Add('AfterWorkbookClose'+$i)}
    }
    if($SettingsEvidence){$expectedEvents.Add('Settings.BeforeQuit.True')}
    $expectedEvents.Add('BeforeQuit');$expectedEvents.Add('QuitAttempt')
    if(!$QuitFails){
        $expectedEvents.Add('AfterQuit')
        if($SettingsEvidence){$expectedEvents.Add('Settings.QuitReturned.False')}
    }
    Assert-Fixture ($Name+'.ExactLifecycleMarkers') (($events -join '|') -ceq ($expectedEvents -join '|'))
    Assert-Fixture ($Name+'.NoExceptionContentLeak') (($output -join "`n") -notmatch 'SYNTHETIC_PRIVATE|private\\fixture')
    $json=$results.ToArray()|ConvertTo-Json -Depth 4
    $decoded=@()
    if($json){$decoded=$json|ConvertFrom-Json}
    $failed=@($decoded|Where-Object {-not $_.Passed}).Count
    Assert-Fixture ($Name+'.FailureSurvivesReportSerialization') ($failed -eq $expectedChecks.Count)
    $observations.Add([pscustomobject]@{
        Case=$Name;QuitAttempts=$excel.Calls;Events=$events.ToArray()
        HarnessChecks=$results.ToArray();SerializedFailedCount=$failed;Uncaught=$uncaught
    })
}
Invoke-Fixture 'Empty' 0 0 $false $false
Invoke-Fixture 'Happy' 2 0 $false $false
Invoke-Fixture 'HappySettings' 2 0 $false $true
Invoke-Fixture 'FirstCloseThrows' 2 1 $false $false
Invoke-Fixture 'LastCloseThrowsSettings' 2 2 $false $true
Invoke-Fixture 'QuitThrows' 2 0 $true $false
Invoke-Fixture 'CloseAndQuitThrow' 2 1 $true $true
Invoke-Fixture 'EmptyQuitThrowsSettings' 0 0 $true $true
$summary=[pscustomobject]@{
    PowerShellVersion=$PSVersionTable.PSVersion.ToString()
    HarnessSHA256=(Get-FileHash -LiteralPath $HarnessPath).Hash
    TestSHA256=(Get-FileHash -LiteralPath $PSCommandPath).Hash
    Cases=$observations.Count;Passed=@($assertions|Where-Object {$_.Passed}).Count
    Failed=@($assertions|Where-Object {!$_.Passed}).Count
    ActualCleanupBlockExecuted=$true;FullHarnessExecuted=$false
    LiveExcelUsed=$false;CursorCalled=$false;WprCalled=$false;ForcedTermination=$false
}
$utf8=[Text.UTF8Encoding]::new($false)
[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'checks.json'),($assertions.ToArray()|ConvertTo-Json -Depth 4),$utf8)
[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'cases.json'),($observations.ToArray()|ConvertTo-Json -Depth 6),$utf8)
[IO.File]::WriteAllText((Join-Path $EvidenceRoot 'summary.json'),($summary|ConvertTo-Json),$utf8)
$summary|ConvertTo-Json
if($summary.Failed -gt 0){exit 1}
