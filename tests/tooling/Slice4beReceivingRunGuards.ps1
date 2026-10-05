# Boundary fixtures run only after the original full replay and Stop proof.
function Test-ReceivingRunGuards {
    $traceReplacement=$false
    function CloseGuardRun { [void](RunnerControl 'btnCloseRun' 'Click') }
    function SetOwnRecording([bool]$Enabled) {
        try {
            [void](Run 'invSys.Admin.xlam' 'TestD5Commands.OpenSettings')
            if(-not [bool](Run 'invSys.Admin.xlam' 'TestD5Commands.RecordingOwnUserForTest' @($Enabled))){throw 'Actual user recording policy save failed; not replay RED.'}
        } finally {[void](Run 'invSys.Admin.xlam' 'TestD5Commands.CloseSettings')}
    }
    function LatestGuardRun($Before) {
        $rows=@(RunFiles|Where-Object {-not $Before.ContainsKey($_.FullName)}|ForEach-Object {Get-Content -LiteralPath $_.FullName -Raw|ConvertFrom-Json}|Sort-Object Revision)
        if($rows.Count){return $rows[-1]}
        return $null
    }
    function StartGuardRun {
        if($traceReplacement){Write-Host 'Replacement.Setup.Before'}
        if(-not (OpenRunSetup)){throw 'Accepted guard setup is unavailable; not a new boundary RED.'}
        if($traceReplacement){Write-Host 'Replacement.Setup.Opened'}
        if((RunnerControl 'lstRunSourceEntities' 'Select' '0') -cne 'SELECTED'){throw 'Guard source entity unavailable.'}
        [void](RunnerControl 'cboRunMode' 'Write' 'Step through')
        if($traceReplacement){Write-Host 'Replacement.Start.Before'}
        [void](RunnerControl 'btnStartRun' 'Click')
        if($traceReplacement){Write-Host 'Replacement.Start.Returned'}
    }
    # Purpose mutations are invalid disposable Config fixtures, never a product relabel command.
    $configBytes=[IO.File]::ReadAllBytes($Fixture.Config)
    $configHash=(Get-FileHash -LiteralPath $Fixture.Config).Hash
    try {
        foreach($purpose in @('Operational','','Unknown')) {
            CloseGuardRun
            $config=$excel.Workbooks.Open($Fixture.Config,0,$false)
            try {
                $table=Table $config 'tblWarehouseConfig'
                $table.DataBodyRange.Cells.Item(1,$table.ListColumns.Item('WarehousePurpose').Index).Value2=$purpose
                $config.Save()
            } finally {$config.Close($false)}
            $before=BusinessPins;$runs=BoundPins $runRoot;$journals=RecordingPins
            $opened=OpenRunSetup
            $case=if($purpose -ceq ''){'Missing'}else{$purpose}
            Check ('ReceivingRun.Guard.Purpose.'+$case) (-not $opened -and (BoundSame $before (BusinessPins)) -and (BoundSame $runs (BoundPins $runRoot)) -and (BoundSame $journals (RecordingPins)))
            [IO.File]::WriteAllBytes($Fixture.Config,$configBytes)
        }
    } finally {[IO.File]::WriteAllBytes($Fixture.Config,$configBytes)}
    Check 'ReceivingRun.Guard.PurposeFixtureRestored' ((Get-FileHash -LiteralPath $Fixture.Config).Hash -ceq $configHash)

    # Receiving rights permit guide use without granting authoring or Admin rights.
    CloseRecordingViewer;SelectTarget $Fixture 'config-reader'
    $canReceive=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('RECEIVE_POST','config-reader',$Fixture.Warehouse,'S1')
    $canAuthor=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ACTION_PATH_MAINT','config-reader',$Fixture.Warehouse,'S1')
    $canAdmin=Run 'invSys.Core.xlam' 'modAuth.CanPerform' @('ADMIN_MAINT','config-reader',$Fixture.Warehouse,'S1')
    if($canReceive -isnot [bool] -or -not $canReceive -or $canAuthor -or $canAdmin){throw 'Receiving-only role fixture is invalid.'}
    $before=BoundPins $runRoot;$business=BusinessPins
    [void](Run 'invSys.Core.xlam' 'modExecutionRun.ArmRunEntryProbeForTest')
    StartGuardRun
    $initial=LatestGuardRun $before
    $started=$null -ne $initial -and $initial.State -ceq 'Running' -and $initial.CreatedByUserId -ceq 'config-reader'
    Check 'ReceivingRun.Guard.NonAuthorCanStart' ($started -and (BoundControl 'btnConfigureExecution' 'State') -ceq 'True|False')
    [void](RunnerControl 'btnNextRunStep' 'Click')
    $one=LatestGuardRun $before
    $oneStep=$started -and $one.State -ceq 'Running' -and @($one.Steps).Count -eq 1 -and $one.Steps[0].State -ceq 'Completed'
    Check 'ReceivingRun.Guard.NonAuthorUsesOrdinaryOwner' ($oneStep -and (BoundSame $business (BusinessPins)))
    Check 'ReceivingRun.Guard.NestedStartSetupAndNextRefused' ($oneStep -and (Run 'invSys.Core.xlam' 'modExecutionRun.RunEntryProbeResultForTest') -ceq '1|1|True')

    # Revoke the fixture's ordinary permission after its first real step.
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $authBytes=[IO.File]::ReadAllBytes($authPath);$authHash=(Get-FileHash -LiteralPath $authPath).Hash
    try {
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        try {
            $caps=Table $auth 'tblCapabilities';$revoked=0
            foreach($row in $caps.ListRows){
                if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-reader' -and $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'RECEIVE_POST'){
                    $row.Range.Cells.Item(1,$caps.ListColumns.Item('Status').Index).Value2='Inactive';$revoked++
                }
            }
            $auth.Save()
        } finally {$auth.Close($false)}
        if($revoked -ne 1){throw 'Receiving permission revocation fixture was not unique.'}
        [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
        $activity=ActivityPins
        [void](RunnerControl 'btnNextRunStep' 'Click')
        $blocked=LatestGuardRun $before
        Check 'ReceivingRun.Guard.LostPermissionStopsLaterDispatch' ($oneStep -and $blocked.State -ceq 'Blocked' -and @($blocked.Steps).Count -eq 1 -and (BoundSame $business (BusinessPins)) -and (BoundSame $activity (ActivityPins)))
    } finally {
        [IO.File]::WriteAllBytes($authPath,$authBytes)
        [void](Run 'invSys.Core.xlam' 'modAuth.ReloadAuth' @($Fixture.Warehouse))
        CloseGuardRun;CloseRecordingViewer;SelectTarget $Fixture 'config-admin'
    }
    Check 'ReceivingRun.Guard.PermissionFixtureRestored' ((Get-FileHash -LiteralPath $authPath).Hash -ceq $authHash)

    foreach($case in @('RecordingStopped','PolicyDisabled','UserDisabled')) {
        if($TraceReceivingRunCloseForTest -and $case -ceq 'PolicyDisabled'){
            foreach($package in @('invSys.Operations.xlam','invSys.Admin.xlam')){
                [void](Run $package 'TestRunGuardTrace.InitializeForTest' @((Join-Path $reportRoot ($package+'.guard-callbacks.log'))))
            }
        }
        $before=BoundPins $runRoot
        StartGuardRun
        [void](RunnerControl 'btnNextRunStep' 'Click')
        $one=LatestGuardRun $before
        if($null -eq $one -or $one.State -cne 'Running' -or @($one.Steps).Count -ne 1){throw 'Capture guard requires an actual first step.'}
        try {
            if($case -ceq 'RecordingStopped'){
                # Reopening the ordinary Viewer refreshes its recording controls.
                OpenRecordingViewer
                if((RecordingControl 'Stop Recording' 'Click') -cne 'DELIVERED'){throw 'Actual recording Stop control unavailable.'}
            }elseif($case -ceq 'UserDisabled'){SetOwnRecording $false}else{SetRecordingPolicy $false}
            $business=BusinessPins;$activity=ActivityPins
            [void](RunnerControl 'btnNextRunStep' 'Click')
            $blocked=LatestGuardRun $before
            Check ('ReceivingRun.Guard.'+$case+'StopsLaterDispatch') ($blocked.State -ceq 'Blocked' -and @($blocked.Steps).Count -eq 1 -and (BoundSame $business (BusinessPins)) -and (BoundSame $activity (ActivityPins)))
            if($case -ceq 'UserDisabled'){
                CloseGuardRun
                . (Join-Path $PSScriptRoot 'Slice4beReceivingUserAudit.ps1')
                Test-ReceivingUserAudit $Fixture
            }
        } finally {
            Write-Host ($case+'.CloseRun.Before')
            CloseGuardRun
            Write-Host ($case+'.CloseRun.Returned')
            if($case -ceq 'UserDisabled'){SetOwnRecording $true}
            if($case -ceq 'PolicyDisabled'){
                Write-Host 'PolicyDisabled.Restore.Before'
                SetRecordingPolicy $true
                Write-Host 'PolicyDisabled.Restore.Returned'
            }
        }
    }

    # A closed/reopened workbook has a different object even when path/name match.
    $traceReplacement=$true
    $before=BoundPins $runRoot
    StartGuardRun
    Write-Host 'Replacement.OpenOwner.Before'
    [void](RunnerControl 'btnNextRunStep' 'Click')
    Write-Host 'Replacement.OpenOwner.Returned'
    $one=LatestGuardRun $before
    if($null -eq $one -or $one.State -cne 'Running' -or @($one.Steps).Count -ne 1){throw 'Replacement guard requires an actual captured workbook.'}
    $label=RunnerControl 'lblRunWorkbook' 'Label'
    $name=$label.Substring('Captured Receiving workbook: '.Length)
    $captured=$excel.Workbooks.Item($name);$path=$captured.FullName
    # Retain a disposable host as in the existing Receiving lifecycle fixture.
    # This does not bypass the observed close-time fault, which must be resolved.
    $hostBook=$excel.Workbooks.Add()
    $business=BusinessPins;$activity=ActivityPins;$operatorHash=Get-ReceivingFixtureHash $path
    Write-Host 'Replacement.WorkbookClose.Before'
    $captured.Close($false)
    Write-Host 'Replacement.WorkbookClose.Returned'
    $replacement=$excel.Workbooks.Open($path,0,$false)
    Write-Host 'Replacement.WorkbookOpen.Returned'
    try {
        Write-Host 'Replacement.Next.Before'
        $next=RunnerControl 'btnNextRunStep' 'Click'
        Write-Host 'Replacement.Next.Returned'
        $blocked=LatestGuardRun $before
        $steps=@($blocked.Steps)
        $partial=$steps.Count -eq 1 -or ($steps.Count -eq 2 -and $steps[1].State -cne 'Completed')
        $terminal=$blocked.State -cin @('Stopped','Blocked','Failed','Unknown')
        Check 'ReceivingRun.Guard.SameNameReplacementCannotRetarget' ($terminal -and $partial -and (BoundSame $business (BusinessPins)) -and (BoundSame $activity (ActivityPins)) -and (Get-ReceivingFixtureHash $path) -ceq $operatorHash)
        Check 'ReceivingRun.Guard.ReplacementPreventsFurtherNext' ($terminal -and (RunnerControl 'btnNextRunStep' 'Click') -cin @('DISABLED','MISSING'))
        [pscustomobject]@{RunTerminal=$terminal;PartialStepsPreserved=$partial;RunnerMissing=($next -ceq 'MISSING');NoBusinessChange=(BoundSame $business (BusinessPins));NoOwnerDispatch=(BoundSame $activity (ActivityPins))}|ConvertTo-Json|Set-Content -LiteralPath (Join-Path $reportRoot 'replacement-facts.json')
    } finally {
        CloseGuardRun;$replacement.Close($false)
        if($null -ne $hostBook){CloseRecordingViewer;$hostBook.Close($false)}
    }
}
