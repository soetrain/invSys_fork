# Native keyboard-to-painted-list timing for the two disposable guide searches.
# The evidence contains only stage names, elapsed time and row counts.
function Test-GuideSearchNative([string]$Form,[string]$Caption,[string]$Search,[string]$List,[string]$Prefix,[string]$First,[string]$Second,[string]$Third,[bool]$ExpectUnselected) {
    . (Join-Path $PSScriptRoot 'Slice4beGuideInputNative.ps1')
    if(-not ('WorksheetInput' -as [type])){. (Join-Path $PSScriptRoot 'Slice4beWorksheetInput.ps1')}
    Initialize-GuideInputNative
    $window=[InvSysSettingsCapture]::OwnedVisibleForm($Caption,[IntPtr]$excel.Hwnd)
    [uint32]$excelPid=0;[uint32]$formPid=0
    [void][GuideInputNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelPid)
    [void][GuideInputNative]::GetWindowThreadProcessId($window,[ref]$formPid)
    if($window -eq [IntPtr]::Zero -or $excelPid -eq 0 -or $formPid -ne $excelPid){
        throw 'Exact test-owned guide search form unavailable; not behavioral RED.'
    }
    $probe={param($name,$action,$value) [string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($Form,$name,$action,$value))}
    if((& $probe $Search 'Text' '') -cne ''){throw 'Guide search must start empty for native timing.'}
    $allRows=[int](& $probe $List 'Rows' '')
    if($allRows -lt 1){throw 'Guide search has no disposable rows for native timing.'}
    if(-not [WorksheetInput]::Activate($window)){
        [void][InvSysSettingsCapture]::ActivateByCaptionClick($window,[IntPtr]$excel.Hwnd)
    }
    if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
        throw 'Guide search window is not foreground; no key sent.'
    }
    if((& $probe $Search 'Focus' '') -cne 'FOCUSED'){
        throw 'Guide search textbox could not take focus; no key sent.'
    }
    $sender=New-Object -ComObject WScript.Shell
    $samples=New-Object System.Collections.ArrayList
    try {
        $steps=@(
            [pscustomobject]@{Name='FirstAfterSelection';Keys=$First;Length=1;Rows='Some';LimitMs=1500},
            [pscustomobject]@{Name='Second';Keys=$Second;Length=2;Rows='Some';LimitMs=750},
            [pscustomobject]@{Name='Third';Keys=$Third;Length=3;Rows='Some';LimitMs=750},
            [pscustomobject]@{Name='Backspace';Keys='{BACKSPACE}';Length=2;Rows='Some';LimitMs=750},
            [pscustomobject]@{Name='NoMatch';Keys='^aQQQQ';Length=4;Rows='None';LimitMs=1500},
            [pscustomobject]@{Name='Clear';Keys='^a{DEL}';Length=0;Rows='All';LimitMs=750}
        )
        foreach($step in $steps){
            if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window -or
               (& $probe '' 'Active' '') -cne $Search){
                throw ('Native guide search lost focus before '+$step.Name+'; no key sent.')
            }
            $clock=[Diagnostics.Stopwatch]::StartNew()
            $sender.SendKeys($step.Keys)
            $delivered=$false;$rows=-1;$length=-1
            for($attempt=0;$attempt -lt 60;$attempt++){
                $length=(& $probe $Search 'Text' '').Length
                $rows=[int](& $probe $List 'PaintRows' '')
                $match=if($step.Rows -ceq 'All'){$rows -eq $allRows}elseif($step.Rows -ceq 'None'){$rows -eq 0}else{$rows -gt 0}
                if($length -eq $step.Length -and $match){$delivered=$true;break}
                Start-Sleep -Milliseconds 25
            }
            $clock.Stop()
            if(-not $delivered){throw ('Guide search native result unavailable at '+$step.Name+'; text length='+$length+'; rows='+$rows)}
            $sample=[pscustomobject]@{Form=$Form;Stage=$step.Name;ElapsedMs=[math]::Round($clock.Elapsed.TotalMilliseconds,1);Rows=$rows;LimitMs=$step.LimitMs}
            [void]$samples.Add($sample)
            Check ($Prefix+'.NativeSearch.'+$step.Name+'.Responsive') ($clock.Elapsed.TotalMilliseconds -le $step.LimitMs)
            if($ExpectUnselected){
                Check ($Prefix+'.NativeSearch.'+$step.Name+'.Unselected') ((& $probe $List 'Selected' '') -ceq '')
            }
        }
    } finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sender)
        [void](& $probe $Search 'Write' '')
    }
    @($samples) | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path $reportRoot ('native-'+$Form+'-search.json')) -Encoding UTF8
}
