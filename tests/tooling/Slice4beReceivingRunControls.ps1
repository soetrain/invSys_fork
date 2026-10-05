# Complement the full B0 proof with actual Step through/Next/Stop and layout checks.
function Test-ReceivingRunControls {
    [void](RunnerControl 'btnCloseRun' 'Click')
    $prior=BoundPins $runRoot
    $business=BusinessPins
    $opened=OpenRunSetup
    $choices=@((RunnerControl 'lstRunSourceEntities' 'Values') -split "`n"|Where-Object {$_ -cne '' -and $_ -cne 'MISSING'})
    $selected=$false
    if($choices.Count){$selected=(RunnerControl 'lstRunSourceEntities' 'Select' '0') -ceq 'SELECTED'}
    foreach($size in @('Minimum','Default','Larger')){
        Check ('ReceivingRun.Layout.'+$size) ($opened -and (RunnerControl '' 'Fit' $size) -ceq 'True')
    }
    $mode=(RunnerControl 'cboRunMode' 'Write' 'Step through') -ceq 'DELIVERED'
    $started=(RunnerControl 'btnStartRun' 'Click') -ceq 'DELIVERED'
    function LatestControlAttempt {
        $rows=@(RunFiles|Where-Object {-not $prior.ContainsKey($_.FullName)}|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Sort-Object Revision)
        if($rows.Count){return $rows[-1]}
        return $null
    }
    $initial=LatestControlAttempt
    $ready=$opened -and $selected -and $mode -and $started -and $null -ne $initial -and $initial.State -ceq 'Running' -and $initial.Mode -ceq 'StepThrough' -and @($initial.Steps).Count -eq 0
    Check 'ReceivingRun.StepThroughWaitsForNext' ($ready -and (BoundSame $business (BusinessPins)))
    $next=(RunnerControl 'btnNextRunStep' 'Click') -ceq 'DELIVERED'
    $one=LatestControlAttempt
    $oneStep=$ready -and $next -and $null -ne $one -and $one.State -ceq 'Running' -and @($one.Steps).Count -eq 1 -and $one.Steps[0].ControlId -ceq 'RECEIVING_OPEN' -and $one.Steps[0].State -ceq 'Completed'
    Check 'ReceivingRun.NextExecutesExactlyOneOwner' $oneStep
    $beforeRepeat=BoundPins $runRoot
    $repeat=RunnerControl 'btnStartRun' 'Click'
    Check 'ReceivingRun.ActiveAttemptCannotStartAgain' ($oneStep -and $repeat -ceq 'DISABLED' -and (BoundSame $beforeRepeat (BoundPins $runRoot)))
    $stopped=(RunnerControl 'btnStopRun' 'Click') -ceq 'DELIVERED'
    $last=LatestControlAttempt
    $stoppedAtOne=$oneStep -and $stopped -and $null -ne $last -and $last.State -ceq 'Stopped' -and @($last.Steps).Count -eq 1
    Check 'ReceivingRun.StopRetainsPartialAttempt' ($stoppedAtOne -and (BoundSame $business (BusinessPins)))
    $beforeNext=BoundPins $runRoot
    $again=RunnerControl 'btnNextRunStep' 'Click'
    Check 'ReceivingRun.StoppedAttemptCannotDispatchLaterStep' ($stoppedAtOne -and $again -ceq 'DISABLED' -and (BoundSame $beforeNext (BoundPins $runRoot)))
    $evaluationRoot=Join-Path $journalRoot 'Evaluations';$evaluations=BoundPins $evaluationRoot
    $verify=(RunnerControl 'btnVerifyRun' 'Click') -ceq 'DELIVERED'
    $fresh=@()
    if(Test-Path -LiteralPath $evaluationRoot){$fresh=@(Get-ChildItem -LiteralPath $evaluationRoot -File -Filter '*.json'|Where-Object {-not $evaluations.ContainsKey($_.FullName)})}
    $evaluation=$null
    if($fresh.Count -eq 1){$evaluation=Get-Content -LiteralPath $fresh[0].FullName -Raw|ConvertFrom-Json}
    Check 'ReceivingRun.StoppedRunCannotBorrowPriorSuccess' ($stoppedAtOne -and $verify -and $null -ne $evaluation -and $evaluation.ActionPathId -ceq $last.Recording.ActionPathId -and $evaluation.ResultState -ceq 'Failed' -and @($evaluation.TerminalSources).Count -eq 0)
    $preserved=$true
    foreach($path in $prior.Keys){if(-not (Test-Path -LiteralPath $path) -or (Get-FileHash -LiteralPath $path).Hash -cne $prior[$path]){$preserved=$false}}
    Check 'ReceivingRun.NewAttemptPreservesPriorRunRevisions' $preserved
    [void](RunnerControl 'btnCloseRun' 'Click')
}
