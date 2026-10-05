# Next Batch is local preparation. Real completion is a reusable fixture prerequisite.
function Get-NextPathAuthorityPins($Fixture){
    $pins=@{}
    foreach($file in Get-ChildItem $Fixture.Root -Recurse -File|Where-Object {$_.Extension -in '.xlsb','.xlsm' -and $_.Name -notlike '*.Snapshot.*' -and -not $_.Name.StartsWith('~$')}){$pins[$file.FullName]=SavedHash $file.FullName}
    if($pins.Count -lt 2){throw 'Next Batch authority fixture unavailable; not product RED.'}
    return $pins
}
function Prepare-NextPath($Book,[string]$Canary,[string]$Mode){
    if($Mode -ceq 'Reusable'){
        if(-not [bool](Probe 'CompleteBaselinePrepare' @($true)) -or -not [bool](Probe 'CompleteBaselineAct') -or -not [bool](Probe 'CompleteBaselineCompleted')){throw 'Actual completed batch prerequisite unavailable; not product RED.'}
    }else{
        # Only this generated operator fixture's staging is reset between recordings.
        if(@($Book.Worksheets|Where-Object Name -CEQ 'Production').Count -and -not [bool](Probe 'NextActivityMissingWorksheet')){throw 'Owned staging reset failed.'}
        if(-not [bool](Probe 'NextActivityWorksheetPrepare' @($Canary))){throw 'Next Batch worksheet preparation unavailable; not product RED.'}
    }
    [void](Probe 'RunLocalShowAndCapture' @($Book.Name,'CHECK_IN'))
}
function Invoke-NextPath($Fixture,[string]$Canary,[string]$Mode,[string]$Label,$Before){
    if($Mode -ceq 'Reusable'){
        PathCheck ('InstructionPaths.'+$Label+'.ActualHandlerAndGuards') ([bool](Probe 'NextActivityAct' @('')))
        PathCheck ('InstructionPaths.'+$Label+'.OwnerAdvancedOnce') ([int](Probe 'NextBaselineEntries') -eq 1 -and [bool](Probe 'NextBaselineReady'))
    }else{
        $tokens=$null;$errors=$null
        $ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'),[ref]$tokens,[ref]$errors)
        $definition=$ast.Find({param($n) $n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -ceq 'Start-DialogCaptureAndDismiss'},$false)
        if($errors.Count -or $null -eq $definition){throw 'Native notification observer unavailable; not product RED.'}
        . ([scriptblock]::Create($definition.Extent.Text))
        $processes=@(Get-Process EXCEL);if($processes.Count -ne 1){throw 'Isolated Excel required.'}
        $stop=Join-Path $runRoot ('next-path-notice-'+$Label)
        $observer=Start-DialogCaptureAndDismiss -ExcelProcessId $processes[0].Id -TimeoutSeconds 30 -StopPath $stop
        try{PathCheck ('InstructionPaths.'+$Label+'.ActualHandlerAndGuards') ([bool](Probe 'NextActivityAct' @('')))}finally{
            [IO.File]::WriteAllText($stop,'');Wait-Job $observer -Timeout 35|Out-Null
            $notice=@(Receive-Job $observer -ErrorAction SilentlyContinue) -join "`n"
            PathCheck ('InstructionPaths.'+$Label+'.NativeNotice') ($observer.State -eq 'Completed' -and $notice.Contains('Next Batch ready. Inventory selections cleared for unchecked processes.'))
            if($observer.State -ne 'Completed'){Stop-Job $observer};Remove-Job $observer
        }
        foreach($fact in @('Output','Identity','Custom','Palette')){PathCheck ('InstructionPaths.'+$Label+'.'+$fact) ([bool](Probe 'NextActivityWorksheetFact' @($fact,$Canary)))}
    }
    $after=Get-NextPathAuthorityPins $Fixture;$same=$Before.Count -eq $after.Count
    foreach($file in $Before.Keys){$same=$same -and $after.ContainsKey($file) -and $after[$file] -ceq $Before[$file]}
    PathCheck ('InstructionPaths.'+$Label+'.CanonicalBytesPreservedByNextBatch') $same
}
