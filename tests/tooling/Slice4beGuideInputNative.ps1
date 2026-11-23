# Native, process-bound actions for the disposable guide editor fixture only.
function Initialize-GuideInputNative {
    if(-not ('GuideInputNative' -as [type])){
        Add-Type @'
using System; using System.Runtime.InteropServices;
public static class GuideInputNative {
 [StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
 [DllImport("user32.dll",SetLastError=true)] public static extern bool GetCursorPos(out Point point);
 [DllImport("user32.dll",SetLastError=true)] public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll",SetLastError=true)] public static extern bool PostMessage(IntPtr h,uint message,IntPtr w,IntPtr l);
 [DllImport("user32.dll",SetLastError=true)] public static extern bool GetClientRect(IntPtr h,out Rect rect);
 [DllImport("user32.dll",SetLastError=true)] public static extern bool ClientToScreen(IntPtr h,ref Point point);
 [DllImport("user32.dll")] public static extern int GetSystemMetrics(int n);
 [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint process);
 [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
}
'@
    }
    Initialize-SettingsCapture
}

function Get-OwnedGuideInputWindow([string]$caption='Action Path guide') {
    Initialize-GuideInputNative
    $window=[InvSysSettingsCapture]::OwnedVisibleForm($caption,[IntPtr]$excel.Hwnd)
    [uint32]$excelProcess=0;[uint32]$formProcess=0
    [void][GuideInputNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelProcess)
    [void][GuideInputNative]::GetWindowThreadProcessId($window,[ref]$formProcess)
    if($window -eq [IntPtr]::Zero -or $excelProcess -eq 0 -or $formProcess -ne $excelProcess){
        throw 'Exact owned guide window unavailable; not behavioral RED.'
    }
    return $window
}

function Move-GuideFocusByTabNative {
    $window=Get-OwnedGuideInputWindow
    if(-not ('WorksheetInput' -as [type])){. (Join-Path $PSScriptRoot 'Slice4beWorksheetInput.ps1')}
    $before=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @('frmActionPathGuide','','Active',''))
    $activation=$null
    try {
        if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
            if(-not [WorksheetInput]::Activate($window)){
                [void][InvSysSettingsCapture]::ActivateByCaptionClick($window,[IntPtr]$excel.Hwnd)
            }
        }
        if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
            throw 'Owned guide form is not foreground for Tab; not behavioral RED.'
        }
        $activation=New-Object -ComObject WScript.Shell
        $activation.SendKeys('{TAB}')
        for($attempt=0;$attempt -lt 20;$attempt++){
            Start-Sleep -Milliseconds 50
            $after=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @('frmActionPathGuide','','Active',''))
            if($after -cne $before){return $after}
        }
        throw ("Native Tab did not move guide focus from $before; not behavioral RED.")
    } finally {
        if($null -ne $activation){[void][Runtime.InteropServices.Marshal]::ReleaseComObject($activation)}
    }
}

function Click-GuideControlNative([string]$form,[string]$caption,[string]$name) {
    $window=Get-OwnedGuideInputWindow $caption
    if(-not ('WorksheetInput' -as [type])){. (Join-Path $PSScriptRoot 'Slice4beWorksheetInput.ps1')}
    $state=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($form,$name,'State',''))
    if($state -cne 'True|True'){throw ("Exact enabled guide control unavailable: $name; not behavioral RED.")}
    $parts=([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @($form,$name,'Bounds',''))).Split('|')
    if($parts.Count -ne 6){throw ("Exact guide control geometry unavailable: $name; not behavioral RED.")}
    $numbers=@($parts | ForEach-Object {[double]::Parse($_,[Globalization.CultureInfo]::CurrentCulture)})
    if($numbers[4] -le 0 -or $numbers[5] -le 0){throw 'Guide client geometry unavailable; not behavioral RED.'}
    $rect=New-Object GuideInputNative+Rect
    if(-not [GuideInputNative]::GetClientRect($window,[ref]$rect)){throw 'Guide client rect unavailable; not behavioral RED.'}
    $point=New-Object GuideInputNative+Point
    $point.X=[int][Math]::Round(($numbers[0]+$numbers[2]/2)*($rect.Right-$rect.Left)/$numbers[4])
    $point.Y=[int][Math]::Round(($numbers[1]+$numbers[3]/2)*($rect.Bottom-$rect.Top)/$numbers[5])
    if(-not [GuideInputNative]::ClientToScreen($window,[ref]$point)){throw 'Guide control screen point unavailable; not behavioral RED.'}
    $x=$point.X;$y=$point.Y
    $original=New-Object GuideInputNative+Point
    if(-not [GuideInputNative]::GetCursorPos([ref]$original)){throw 'Cursor read failed before guide focus.'}
    $left=[GuideInputNative]::GetSystemMetrics(76);$top=[GuideInputNative]::GetSystemMetrics(77)
    $width=[GuideInputNative]::GetSystemMetrics(78);$height=[GuideInputNative]::GetSystemMetrics(79)
    if($width -le 0 -or $height -le 0 -or $original.X -lt $left -or $original.X -ge $left+$width -or
       $original.Y -lt $top -or $original.Y -ge $top+$height -or $x -lt $left -or $x -ge $left+$width -or
       $y -lt $top -or $y -ge $top+$height){
        throw ("Guide focus input geometry outside display; control=$name; bounds=$left,$top,$width,$height; cursor=$($original.X),$($original.Y); target=$x,$y; no input attempted.")
    }
    $inputError=''
    try {
        if(-not [WorksheetInput]::Activate($window)){
            [void][InvSysSettingsCapture]::ActivateByCaptionClick($window,[IntPtr]$excel.Hwnd)
        }
        if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
            throw 'Owned guide window could not be foregrounded; not behavioral RED.'
        }
        [WorksheetInput]::Click($window,$x,$y)
        # SendInput queues the click; let its mouse-up finish before restoring the pointer.
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
        throw ("Cursor restoration failed after guide focus; control=$name; SetCursorPos=$restored; Win32 error=$restoreError; original=$($original.X),$($original.Y); actual=$($actual.X),$($actual.Y); prior input error=$inputError")
    }
    if($inputError){throw ("Guide control input failed for ${name}: $inputError")}
}

function Focus-GuideControlNative([string]$name) {
    Click-GuideControlNative 'frmActionPathGuide' 'Action Path guide' $name
    $active=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @('frmActionPathGuide','','Active',''))
    if($active -cne $name){throw ("Native guide focus missed ${name}; active=$active; not behavioral RED.")}
    return 'FOCUSED'
}

function Type-GuideTextNative([string]$name,[string]$value) {
    if($value -cnotmatch '^[A-Za-z ,]+$'){throw 'Native guide typing fixture contains unsupported SendKeys characters.'}
    $window=Get-OwnedGuideInputWindow
    if(-not [WorksheetInput]::Activate($window)){
        [void][InvSysSettingsCapture]::ActivateByCaptionClick($window,[IntPtr]$excel.Hwnd)
    }
    if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
        throw 'Owned guide form is not foreground for native typing; not behavioral RED.'
    }
    $sender=New-Object -ComObject WScript.Shell
    try {$sender.SendKeys($value)} finally {[void][Runtime.InteropServices.Marshal]::ReleaseComObject($sender)}
    for($attempt=0;$attempt -lt 30;$attempt++){
        if(([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @('frmActionPathGuide',$name,'Text',''))) -ceq $value){return 'DELIVERED'}
        Start-Sleep -Milliseconds 50
    }
    throw ("Native guide typing did not reach exact ${name}; not behavioral RED.")
}

function Click-GuideCancelNative {
    $window=Get-OwnedGuideInputWindow
    if(-not ('WorksheetInput' -as [type])){. (Join-Path $PSScriptRoot 'Slice4beWorksheetInput.ps1')}
    $state=[string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @('frmActionPathGuide','btnCancelGuide','State',''))
    if($state -cne 'True|True'){throw 'Exact enabled guide Cancel button unavailable; not behavioral RED.'}
    $parts=([string](Run 'invSys.Operations.xlam' 'modInventoryViewer.GuideDraftControlForTest' @('frmActionPathGuide','btnCancelGuide','Bounds',''))).Split('|')
    if($parts.Count -ne 6){throw 'Exact guide Cancel geometry unavailable; not behavioral RED.'}
    $numbers=@($parts | ForEach-Object {[double]::Parse($_,[Globalization.CultureInfo]::CurrentCulture)})
    if($numbers[4] -le 0 -or $numbers[5] -le 0){throw 'Guide client geometry unavailable; not behavioral RED.'}
    $rect=New-Object GuideInputNative+Rect
    if(-not [GuideInputNative]::GetClientRect($window,[ref]$rect)){throw 'Guide client rect unavailable; not behavioral RED.'}
    $point=New-Object GuideInputNative+Point
    $point.X=[int][Math]::Round(($numbers[0]+$numbers[2]/2)*($rect.Right-$rect.Left)/$numbers[4])
    $point.Y=[int][Math]::Round(($numbers[1]+$numbers[3]/2)*($rect.Bottom-$rect.Top)/$numbers[5])
    if(-not [GuideInputNative]::ClientToScreen($window,[ref]$point)){throw 'Guide client screen point unavailable; not behavioral RED.'}
    $x=$point.X;$y=$point.Y
    $original=New-Object GuideInputNative+Point
    if(-not [GuideInputNative]::GetCursorPos([ref]$original)){throw 'Cursor read failed before guide Cancel.'}
    $left=[GuideInputNative]::GetSystemMetrics(76);$top=[GuideInputNative]::GetSystemMetrics(77)
    $width=[GuideInputNative]::GetSystemMetrics(78);$height=[GuideInputNative]::GetSystemMetrics(79)
    if($width -le 0 -or $height -le 0 -or $original.X -lt $left -or $original.X -ge $left+$width -or
       $original.Y -lt $top -or $original.Y -ge $top+$height -or $x -lt $left -or $x -ge $left+$width -or
       $y -lt $top -or $y -ge $top+$height){
        throw ("Guide Cancel input geometry outside display; bounds=$left,$top,$width,$height; cursor=$($original.X),$($original.Y); target=$x,$y; no input attempted.")
    }
    $inputError=''
    try {
        if(-not [WorksheetInput]::Activate($window)){
            [void][InvSysSettingsCapture]::ActivateByCaptionClick($window,[IntPtr]$excel.Hwnd)
        }
        if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
            throw 'Owned guide window could not be foregrounded; not behavioral RED.'
        }
        [WorksheetInput]::Click($window,$x,$y)
    } catch {
        $inputError=$_.Exception.GetBaseException().Message
    } finally {
        $restored=[GuideInputNative]::SetCursorPos($original.X,$original.Y)
        $restoreError=if($restored){0}else{[Runtime.InteropServices.Marshal]::GetLastWin32Error()}
        $actual=New-Object GuideInputNative+Point
        $readRestored=[GuideInputNative]::GetCursorPos([ref]$actual)
    }
    if(-not $restored -or -not $readRestored -or $actual.X -ne $original.X -or $actual.Y -ne $original.Y){
        throw ("Cursor restoration failed; SetCursorPos=$restored; Win32 error=$restoreError; GetCursorPos=$readRestored; original=$($original.X),$($original.Y); actual=$($actual.X),$($actual.Y); target=$x,$y; bounds=$left,$top,$width,$height; prior input error=$inputError")
    }
    if($inputError){throw ('Guide Cancel input failed: '+$inputError)}
    return $true
}

function Close-GuideWindowXNative {
    $window=Get-OwnedGuideInputWindow
    Add-Type -AssemblyName UIAutomationClient
    Add-Type -AssemblyName UIAutomationTypes
    $root=[Windows.Automation.AutomationElement]::FromHandle($window)
    $pattern=$root.GetCurrentPattern([Windows.Automation.WindowPattern]::Pattern)
    if($null -eq $pattern){throw 'Owned guide window Close pattern unavailable; not behavioral RED.'}
    $pattern.Close()
    $deadline=[DateTime]::UtcNow.AddSeconds(10)
    while([GuideInputNative]::IsWindow($window) -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
    return -not [GuideInputNative]::IsWindow($window)
}
