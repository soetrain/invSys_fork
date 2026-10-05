# Exercise both existing authorization boundaries without replacing either result.
function Invoke-CompleteWorksheetPermissionNotice([ValidateSet('Permission','Submission')][string]$Kind='Permission') {
    $token=[guid]::NewGuid().ToString('N')
    $ready=Join-Path $runRoot ('complete-worksheet-notice-'+$token+'.ready')
    $stop=Join-Path $runRoot ('complete-worksheet-notice-'+$token+'.stop')
    $observer=Join-Path $repo 'tools/plan022-dialog-observer.ps1'
    $handle=[long]$excel.Hwnd
    $job=Start-Job -ArgumentList $observer,$handle,$ready,$stop,$Kind -ScriptBlock {
        param($observer,$handle,$ready,$stop,$kind)
        $ErrorActionPreference='Stop'
        . $observer
        Invoke-Plan022NativeDialogObservation -ProcessId 0 -TimeoutSeconds 0
        [uint32]$owner=0
        [void][Plan022NativeDialogs]::GetWindowThreadProcessId([IntPtr]$handle,[ref]$owner)
        if(-not $owner){throw 'Owned completion process unavailable.'}
        $created=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
        [IO.File]::WriteAllText($ready,'Ready')
        $role=$false;$completion=$false;$until=[DateTime]::UtcNow.AddSeconds(45)
        while([DateTime]::UtcNow -lt $until -and -not(Test-Path -LiteralPath $stop)){
            if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $created){throw 'Completion process identity changed.'}
            # Raw text remains in this worker's memory; only fixed booleans leave it.
            foreach($text in @([Plan022NativeDialogs]::Poll($owner))){
                if($kind -ceq 'Submission' -and $text -ceq 'WINDOW|Production Complete Run|#32770'){$completion=$true}
                if($text -like 'WINDOW_ELEMENT|*|ControlType.Text|Current user lacks PROD_POST capability.*'){
                    if($text -like 'WINDOW_ELEMENT|Production Complete Run|*'){$completion=$true}else{$role=$true}
                }
            }
            Start-Sleep -Milliseconds 100
        }
        [pscustomobject]@{RoleNotice=$role;CompletionNotice=$completion}
    }
    try{
        for($i=0;$i -lt 100 -and -not(Test-Path -LiteralPath $ready);$i++){Start-Sleep -Milliseconds 100}
        if(-not(Test-Path -LiteralPath $ready)){throw 'Owned dialog observer unavailable; not product RED.'}
        $returned=[bool](Probe 'CompleteEntryAct' @(''))
        [IO.File]::WriteAllText($stop,'Stop')
        [void](Wait-Job $job -Timeout 5)
        $notice=Receive-Job $job -ErrorAction Stop
        if($null -eq $notice){throw 'Owned dialog receipt unavailable; not product RED.'}
        [pscustomobject]@{Returned=$returned;RoleNotice=$notice.RoleNotice;CompletionNotice=$notice.CompletionNotice}
    }finally{if($job.State -eq 'Running'){Stop-Job $job};Remove-Job $job -Force}
}

function Test-ProductionCompleteWorksheetPermission($Fixture,$Book,$Decoy,[string]$Canary){
    $authPath=Join-Path $Fixture.Root ($Fixture.Warehouse+'.invSys.Auth.xlsb')
    $savedAuth=[IO.File]::ReadAllBytes($authPath);$authHash=(Get-FileHash $authPath).Hash
    $context=[string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext')
    $label='CompleteActivity.Worksheet.AdminOnly'
    try{
        $auth=$excel.Workbooks.Open($authPath,0,$false)
        try{
            $caps=Table $auth 'tblCapabilities';$changed=0
            foreach($row in $caps.ListRows){
                if($row.Range.Cells.Item(1,$caps.ListColumns.Item('UserId').Index).Value2 -ceq 'config-producer' -and
                   $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2 -ceq 'PROD_POST'){
                    $row.Range.Cells.Item(1,$caps.ListColumns.Item('Capability').Index).Value2='ADMIN_MAINT';$changed++
                }
            }
            if($changed -ne 1){throw 'Exact disposable capability fixture unavailable.'}
            $auth.Save()
        }finally{$auth.Close($false)}
        if(-not [bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CompleteWorksheetAdminOnlyForTest') -or
           [string](Run 'invSys.Core.xlam' 'modActivity.CaptureContext') -cne $context){throw 'Distinct live permission gates unavailable; not product RED.'}
        $Decoy.Activate();$before=@(Get-Slice4beActivityFiles $Fixture)
        $notice=Invoke-CompleteWorksheetPermissionNotice
        Check ($label+'.ActualHandlerReturned') $notice.Returned
        Check ($label+'.BothNativeRefusalMessages') ($notice.RoleNotice -and $notice.CompletionNotice)
        Check ($label+'.GuardsRestored') ([bool](Probe 'CompleteEntryFact' @('GuardsRestored')))
        Check ($label+'.NoInventorySubmission') ([string](Probe 'CompleteWorksheetEvents') -ceq '')
        foreach($fact in @('Unsubmitted','PermissionMessage','OutputCustom','CheckCustom')){
            Check ($label+'.'+$fact) ([bool](Probe 'CompleteWorksheetFact' @($fact,$Canary)))
        }
        Pair $before 'DENIED' $label
        CaptureOwnedFormByCaptionEvidence 'Production' 'complete-activity-worksheet-admin-only.png'
        Check ($label+'.DecoyPreserved') ($Decoy.Worksheets.Count -eq 1 -and $Decoy.Worksheets.Item(1).Cells.Item(1,1).Value2 -ceq $Canary)
    }finally{
        [IO.File]::WriteAllBytes($authPath,$savedAuth)
        if((Get-FileHash $authPath).Hash -cne $authHash){throw 'Disposable authorization restore failed.'}
    }
    # Keep the same captured session/form; the next normal click proves recovery.
    if([bool](Run 'invSys.Core.xlam' 'TestShippingCatalog.CompleteWorksheetAdminOnlyForTest')){throw 'Live authorization did not recover.'}
    Check ($label+'.ExactInputUnchanged') ([bool](Probe 'CompleteWorksheetFact' @('InputUnchanged',$Canary)))
}
