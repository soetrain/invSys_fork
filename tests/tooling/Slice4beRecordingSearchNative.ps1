# Native keyboard latency through the packaged Action Paths form. This fixture
# records only timings and counts; the disposable search characters stay local.
function Test-RecordingSearchNative([string]$PathId) {
    . (Join-Path $PSScriptRoot 'Slice4beGuideInputNative.ps1')
    . (Join-Path $PSScriptRoot 'Slice4beWorksheetInput.ps1')
    Initialize-GuideInputNative
    $window=[InvSysSettingsCapture]::OwnedVisibleForm('Action Paths',[IntPtr]$excel.Hwnd)
    if($window -eq [IntPtr]::Zero){throw 'Owned visible Action Paths form unavailable; not behavioral RED.'}
    [uint32]$excelOwner=0;[uint32]$formOwner=0
    [void][GuideInputNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelOwner)
    [void][GuideInputNative]::GetWindowThreadProcessId($window,[ref]$formOwner)
    if($excelOwner -eq 0 -or $formOwner -ne $excelOwner){throw 'Action Paths form belongs to another process; not behavioral RED.'}
    [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Search',''))
    $selected=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Select',$PathId))
    if($selected -cne 'SELECTED'){throw 'Disposable recording selection unavailable; not behavioral RED.'}
    if(-not [WorksheetInput]::Activate($window)){throw 'Action Paths form could not be foregrounded; not behavioral RED.'}
    $focus=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchFocus',''))
    if($focus -cne 'txtPathSearch'){throw 'Search textbox did not take focus; not behavioral RED.'}
    if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
        throw 'Action Paths form is not the native keyboard target; not behavioral RED.'
    }
    $sender=New-Object -ComObject WScript.Shell
    $metrics=New-Object System.Collections.ArrayList
    try {
        foreach($step in @(
            [pscustomobject]@{Name='FirstAfterSelection';Keys='R';Text='R';Rows=1;LimitMs=1500},
            [pscustomobject]@{Name='Second';Keys='e';Text='Re';Rows=1;LimitMs=750},
            [pscustomobject]@{Name='Third';Keys='c';Text='Rec';Rows=1;LimitMs=750},
            [pscustomobject]@{Name='Backspace';Keys='{BACKSPACE}';Text='Re';Rows=1;LimitMs=750},
            [pscustomobject]@{Name='Clear';Keys='^a{DEL}';Text='';Rows=1;LimitMs=750},
            [pscustomobject]@{Name='NoMatch';Keys='X';Text='X';Rows=0;LimitMs=750}
        )) {
            if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
                throw 'Native keyboard target changed during Search diagnostic; not behavioral RED.'
            }
            $active=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchActive',''))
            if($active -cne 'txtPathSearch'){throw 'Native Search textbox lost focus; not behavioral RED.'}
            $clock=[Diagnostics.Stopwatch]::StartNew()
            $sender.SendKeys($step.Keys)
            $delivered=$false;$rows=-1
            for($attempt=0;$attempt -lt 60;$attempt++) {
                $value=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchText',''))
                $rows=[int](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchPaint',''))
                if($value -ceq $step.Text -and $rows -eq $step.Rows){$delivered=$true;break}
                Start-Sleep -Milliseconds 25
            }
            $clock.Stop()
            if(-not $delivered){throw ('Native Search did not reach the expected textbox/list state at '+$step.Name+'; not behavioral RED.')}
            $unselected=([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('ListIndex','')) -ceq '-1')
            $sample=[pscustomobject]@{Stage=$step.Name;ElapsedMs=[math]::Round($clock.Elapsed.TotalMilliseconds,1);Rows=$rows;Unselected=$unselected;LimitMs=$step.LimitMs}
            [void]$metrics.Add($sample)
            Check ('RecordingRead.NativeSearch.'+$step.Name+'.Unselected') $unselected
            Check ('RecordingRead.NativeSearch.'+$step.Name+'.Responsive') ($clock.Elapsed.TotalMilliseconds -le $step.LimitMs)
        }
    } finally {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sender)
        [void](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Search',''))
    }
    $file=Join-Path $reportRoot 'native-recording-search.json'
    @($metrics) | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $file -Encoding UTF8
}
