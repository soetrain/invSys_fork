# B0 runner actions use the operator controls. Only independent proof reads use COM.
# Never synthesize a run, journal, owner event or evaluation to satisfy this gate.
function Test-ReceivingRun($Guide,$Original,$Staging,[string]$InventoryPath,$Fixture,$Other,[ref]$Scope) {
    function RunnerControl([string]$Name,[string]$Action,[string]$Value='') {
        BoundControl $Name $Action $Value 'frmActionPathRun'
    }
    function OpenRunSetup {
        OpenRecordingViewer
        if((BoundLibrary 'Open') -cne 'DELIVERED'){throw 'Accepted recording library is unavailable; not runner RED.'}
        BoundOpen;BoundSelect $Guide
        $delivered=(BoundControl 'btnRunHowTo' 'Click') -ceq 'DELIVERED'
        if($delivered -and (RunnerControl '' 'Count') -cne '1') {
            $stage=Run 'invSys.Core.xlam' 'modExecutionRun.ExecutionSetupStageForTest'
            $notice=BoundControl 'lblPublishedGuideStatus' 'Label'
            [pscustomobject]@{ExecutionSetupStage=$stage;TargetRefusal=$notice.StartsWith('Run requires');Unconfigured=$notice.StartsWith('Execution not configured');PermissionRefusal=$notice.Contains('permission');ProfileRefusal=$notice.Contains('profiles are invalid')}|ConvertTo-Json -Compress|Write-Host
        }
        return ($delivered -and (RunnerControl '' 'Count') -ceq '1')
    }
    function RunFiles {
        if(Test-Path -LiteralPath $runRoot){Get-ChildItem -LiteralPath $runRoot -File -Filter '*.json'}
    }
    function RecordingPins {
        $pins=@{}
        foreach($file in Get-ChildItem -LiteralPath $journalRoot -File -Filter '*.json'){$pins[$file.FullName]=(Get-FileHash -LiteralPath $file.FullName).Hash}
        return $pins
    }
    function BusinessPins {
        $pins=@{}
        foreach($fixturePath in @($InventoryPath,$Fixture.Config,$Other.Config,$otherInventory)) {
            $pins[$fixturePath]=(Get-FileHash -LiteralPath $fixturePath).Hash
        }
        return $pins
    }
    $profileRoot=Join-Path $journalRoot 'ExecutionProfiles'
    $profiles=@(Get-ChildItem -LiteralPath $profileRoot -File -Filter '*.json'|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Sort-Object Version)
    if($profiles.Count -ne 2){throw 'Actual profile authoring prerequisite is incomplete; not runner RED.'}
    $savedProfile=$profiles[-1]
    $profilePins=BoundPins $profileRoot;$guidePins=BoundPins (Join-Path $journalRoot 'Guides')
    $runRoot=Join-Path $journalRoot 'Runs'
    $otherInventory=Join-Path $Other.Root ($Other.Warehouse+'.invSys.Data.Inventory.xlsb')
    $businessBefore=BusinessPins
    $activityBefore=ActivityPins
    $journalsBefore=RecordingPins
    $authority=Open-ReceivingEvidenceBook $InventoryPath
    try {
        $priorEvents=@(Get-ReceivingFixtureRows (Table $authority 'tblAppliedEvents')|ForEach-Object EventID)
        $priorKeys=@(Get-ReceivingFixtureRows (Table $authority 'tblInventoryLog')|ForEach-Object System_Key)
    } finally {$authority.Close($false)}
    $opened=OpenRunSetup
    Check 'ReceivingRun.ActualRunOpensSetup' $opened
    $guideLabel=RunnerControl 'lblRunGuide' 'Label'
    $profileLabel=RunnerControl 'lblRunProfile' 'Label'
    $targetLabel=RunnerControl 'lblRunTarget' 'Label'
    Check 'ReceivingRun.SetupNamesExactGuideProfileTarget' ($opened -and $guideLabel.Contains([string]$Guide.ContentSha256) -and $profileLabel.Contains([string]$savedProfile.ContentSha256) -and $targetLabel.Contains([string]$Fixture.Warehouse) -and $targetLabel.Contains('Training'))
    Check 'ReceivingRun.OpenSetupDoesNotExecute' ($opened -and (BoundSame $businessBefore (BusinessPins)) -and $Staging.ListRows.Count -eq 0 -and (BoundSame $journalsBefore (RecordingPins)))
    # Required source input is transient and must not default to an arbitrary row.
    [void](RunnerControl 'btnStartRun' 'Click')
    Check 'ReceivingRun.MissingEntityRefusesDispatch' ($opened -and (BoundSame $businessBefore (BusinessPins)) -and $Staging.ListRows.Count -eq 0 -and (BoundSame $journalsBefore (RecordingPins)))
    $entities=@((RunnerControl 'lstRunSourceEntities' 'Values') -split "`n"|Where-Object {$_ -cne '' -and $_ -cne 'MISSING'})
    $selected=$false;$selectedKey=''
    if($entities.Count){$selectedKey=$entities[0];$selected=(RunnerControl 'lstRunSourceEntities' 'Select' '0') -ceq 'SELECTED'}
    Check 'ReceivingRun.ExactTargetLocalEntityPrompt' ($opened -and $selected -and $selectedKey -cne '' -and $selectedKey -cin $priorKeys -and (RunnerControl 'lstRunSourceEntities' 'Selected') -ceq $selectedKey)
    # A setup captured in A cannot dispatch into B, even when B is the active target.
    SelectTarget $Other 'config-admin'
    [void](RunnerControl 'btnStartRun' 'Click')
    Check 'ReceivingRun.ContextChangeRefusesDispatch' ($opened -and (BoundSame $businessBefore (BusinessPins)) -and $Staging.ListRows.Count -eq 0 -and (BoundSame $journalsBefore (RecordingPins)))
    CloseRecordingViewer
    SelectTarget $Fixture 'config-admin'
    $opened=OpenRunSetup
    $entities=@((RunnerControl 'lstRunSourceEntities' 'Values') -split "`n"|Where-Object {$_ -cne '' -and $_ -cne 'MISSING'})
    $selected=$false;$selectedKey=''
    if($entities.Count){$selectedKey=$entities[0];$selected=(RunnerControl 'lstRunSourceEntities' 'Select' '0') -ceq 'SELECTED'}
    $mode=(RunnerControl 'cboRunMode' 'Write' 'Run all') -ceq 'DELIVERED'
    # Keep prior identities in memory; runtime evidence emits booleans/counts only.
    $runPins=BoundPins $runRoot
    $started=(RunnerControl 'btnStartRun' 'Click') -ceq 'DELIVERED'
    $newFiles=@(RunFiles|Where-Object {-not $runPins.ContainsKey($_.FullName)})
    $runRows=@($newFiles|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Sort-Object Revision)
    $latest=$null
    if($runRows.Count){$latest=$runRows[-1]}
    $chainFiles=@()
    if($null -ne $latest){$chainFiles=@(RunFiles|Where-Object {$_.Name.StartsWith(([string]$latest.RunId+'.'),[StringComparison]::Ordinal)})}
    $dispatched=$opened -and $selected -and $mode -and $started -and $null -ne $latest -and $latest.State -ceq 'Completed'
    Check 'ReceivingRun.StartDispatchesOrderedGuide' ($dispatched -and (@($latest.Steps|ForEach-Object ControlId) -join '|') -ceq (@($Guide.Steps|ForEach-Object ControlId) -join '|') -and @($latest.Steps|Where-Object State -CNE 'Completed').Count -eq 0)
    . (Join-Path $PSScriptRoot 'Slice4beReceivingRunProof.ps1')
    $proof=$false
    Test-ReceivingRunProof $Guide $savedProfile $Original $latest $chainFiles $priorEvents $priorKeys $InventoryPath $Staging $dispatched ([ref]$proof)
    Check 'ReceivingRun.SourceGuideAndProfileVersionsPreserved' ((BoundSame $profilePins (BoundPins $profileRoot)) -and (BoundSame $guidePins (BoundPins (Join-Path $journalRoot 'Guides'))))
    Check 'ReceivingRun.OtherRuntimePreserved' ((Get-FileHash -LiteralPath $otherInventory).Hash -ceq $businessBefore[$otherInventory] -and (Get-FileHash -LiteralPath $Other.Config).Hash -ceq $businessBefore[$Other.Config])
    $replayActions=@(foreach($file in Get-ChildItem -LiteralPath $activityRoot -File -Filter '*.json') {
        if(-not $activityBefore.ContainsKey($file.Name)){Get-Content -LiteralPath $file.FullName -Raw|ConvertFrom-Json}
    })
    Check 'ReceivingRun.DummyInputsRemainOutsideActivity' ((ConvertTo-Json -InputObject @($replayActions) -Depth 15 -Compress) -notmatch 'B0-REPLAY-REFERENCE|B0-REPLAY-LOCATION|B0-REPLAY-LOT')
    . (Join-Path $PSScriptRoot 'Slice4beReceivingRunControls.ps1')
    Test-ReceivingRunControls
    . (Join-Path $PSScriptRoot 'Slice4beReceivingRunGuards.ps1')
    Test-ReceivingRunGuards
    . (Join-Path $PSScriptRoot 'Slice4beReceivingRunClose.ps1')
    Test-ReceivingRunBusyClose
    $Scope.Value=[pscustomobject]@{ReplayExecuted=$dispatched;FreshReplayProof=$proof;B0Accepted=$false}
}
