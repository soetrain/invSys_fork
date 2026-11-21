# Native keyboard latency through the packaged Action Paths form. This fixture
# records only timings and counts; the disposable search characters stay local.
function Focus-RecordingSearchNative([IntPtr]$window) {
    $parts=([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchBounds',''))).Split('|')
    if($parts.Count -ne 6){throw 'Exact Action Paths Search geometry unavailable; not behavioral RED.'}
    $numbers=@($parts | ForEach-Object {[double]::Parse($_,[Globalization.CultureInfo]::CurrentCulture)})
    if($numbers[2] -le 0 -or $numbers[3] -le 0 -or $numbers[4] -le 0 -or $numbers[5] -le 0){
        throw 'Action Paths Search client geometry unavailable; not behavioral RED.'
    }
    $rect=New-Object GuideInputNative+Rect
    if(-not [GuideInputNative]::GetClientRect($window,[ref]$rect)){throw 'Action Paths client rect unavailable; not behavioral RED.'}
    $point=New-Object GuideInputNative+Point
    $point.X=[int][Math]::Round(($numbers[0]+$numbers[2]/2)*($rect.Right-$rect.Left)/$numbers[4])
    $point.Y=[int][Math]::Round(($numbers[1]+$numbers[3]/2)*($rect.Bottom-$rect.Top)/$numbers[5])
    if(-not [GuideInputNative]::ClientToScreen($window,[ref]$point)){
        throw 'Action Paths Search screen point unavailable; not behavioral RED.'
    }
    $original=New-Object GuideInputNative+Point
    if(-not [GuideInputNative]::GetCursorPos([ref]$original)){throw 'Cursor read failed before Search focus.'}
    $left=[GuideInputNative]::GetSystemMetrics(76);$top=[GuideInputNative]::GetSystemMetrics(77)
    $width=[GuideInputNative]::GetSystemMetrics(78);$height=[GuideInputNative]::GetSystemMetrics(79)
    if($width -le 0 -or $height -le 0 -or $original.X -lt $left -or $original.X -ge $left+$width -or
       $original.Y -lt $top -or $original.Y -ge $top+$height -or $point.X -lt $left -or $point.X -ge $left+$width -or
       $point.Y -lt $top -or $point.Y -ge $top+$height){
        throw 'Search focus point or cursor outside current display; no input attempted.'
    }
    $inputError=''
    try {
        if(-not [WorksheetInput]::Activate($window)){
            [void][InvSysSettingsCapture]::ActivateByCaptionClick($window,[IntPtr]$excel.Hwnd)
        }
        if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
            throw 'Owned Action Paths window could not be foregrounded; not behavioral RED.'
        }
        [WorksheetInput]::Click($window,$point.X,$point.Y)
        Start-Sleep -Milliseconds 150
    } catch {
        $inputError=$_.Exception.GetBaseException().Message
    } finally {
        $restored=[GuideInputNative]::SetCursorPos($original.X,$original.Y)
        $restoreError=if($restored){0}else{[Runtime.InteropServices.Marshal]::GetLastWin32Error()}
        $actual=New-Object GuideInputNative+Point
        $readRestored=[GuideInputNative]::GetCursorPos([ref]$actual)
    }
    if(-not $restored -or -not $readRestored -or $actual.X -ne $original.X -or $actual.Y -ne $original.Y){
        throw ("Search cursor restoration failed; SetCursorPos=$restored; Win32 error=$restoreError; prior input error=$inputError")
    }
    if($inputError){throw ("Search textbox native focus failed: $inputError")}
    $active=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchActive',''))
    if($active -cne 'txtPathSearch'){throw 'Native click did not focus Action Paths Search; not behavioral RED.'}
}

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
    if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
        if(-not [WorksheetInput]::Activate($window)){
            # The owned-caption fallback is already used by the guide's native
            # handler test when Windows rejects SetForegroundWindow alone.
            [void][InvSysSettingsCapture]::ActivateByCaptionClick($window,[IntPtr]$excel.Hwnd)
        }
    }
    $focus=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchFocus',''))
    if($focus -cne 'txtPathSearch'){throw 'Search textbox did not take focus; not behavioral RED.'}
    $loadedRows=[int](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('Rows',''))
    if($loadedRows -lt 1){throw 'Disposable recording rows unavailable; not behavioral RED.'}
    if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
        throw 'Action Paths form is not the native keyboard target; not behavioral RED.'
    }
    $sender=New-Object -ComObject WScript.Shell
    $metrics=New-Object System.Collections.ArrayList
    try {
        foreach($step in @(
            [pscustomobject]@{Name='FirstAfterSelection';Keys='R';Text='R';Rows=$loadedRows;LimitMs=1500},
            [pscustomobject]@{Name='Second';Keys='e';Text='Re';Rows=$loadedRows;LimitMs=750},
            [pscustomobject]@{Name='Third';Keys='c';Text='Rec';Rows=$loadedRows;LimitMs=750},
            [pscustomobject]@{Name='Backspace';Keys='{BACKSPACE}';Text='Re';Rows=$loadedRows;LimitMs=750},
            [pscustomobject]@{Name='Clear';Keys='^a{DEL}';Text='';Rows=$loadedRows;LimitMs=750},
            [pscustomobject]@{Name='NoMatch';Keys='X';Text='X';Rows=0;LimitMs=750}
        )) {
            if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
                throw 'Native keyboard target changed during Search diagnostic; not behavioral RED.'
            }
            $active=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchActive',''))
            if($active -cne 'txtPathSearch'){throw 'Native Search textbox lost focus; not behavioral RED.'}
            if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
                throw ('Native Search form lost foreground after focus read at '+$step.Name+'; no key sent.')
            }
            if($step.Name -ceq 'FirstAfterSelection'){
                Focus-RecordingSearchNative $window
            }
            $clock=[Diagnostics.Stopwatch]::StartNew()
            $sender.SendKeys($step.Keys)
            $delivered=$false;$rows=-1;$textLength=-1
            for($attempt=0;$attempt -lt 60;$attempt++) {
                $value=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchText',''))
                $textLength=$value.Length
                $rows=[int](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchPaint',''))
                if($value -ceq $step.Text -and $rows -eq $step.Rows){$delivered=$true;break}
                Start-Sleep -Milliseconds 25
            }
            $clock.Stop()
            if(-not $delivered){
                $activeAfter=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.RecordingLibraryForTest' @('SearchActive',''))
                $foregroundOwned=([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -eq $window)
                throw ('Native Search did not reach expected state at '+$step.Name+'; exact text='+($value -ceq $step.Text)+'; text length='+$textLength+'; rows='+$rows+'; expected rows='+$step.Rows+'; active='+$activeAfter+'; foreground owned='+$foregroundOwned+'; not behavioral RED.')
            }
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
