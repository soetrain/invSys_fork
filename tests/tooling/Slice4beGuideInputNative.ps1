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
 [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint process);
 [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
}
'@
    }
    Initialize-SettingsCapture
}

function Get-OwnedGuideInputWindow {
    Initialize-GuideInputNative
    $window=[InvSysSettingsCapture]::OwnedVisibleForm('Action Path guide',[IntPtr]$excel.Hwnd)
    [uint32]$excelProcess=0;[uint32]$formProcess=0
    [void][GuideInputNative]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$excelProcess)
    [void][GuideInputNative]::GetWindowThreadProcessId($window,[ref]$formProcess)
    if($window -eq [IntPtr]::Zero -or $excelProcess -eq 0 -or $formProcess -ne $excelProcess){
        throw 'Exact owned guide window unavailable; not behavioral RED.'
    }
    return $window
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
    try {
        if(-not [WorksheetInput]::Activate($window)){
            [void][InvSysSettingsCapture]::ActivateByCaptionClick($window,[IntPtr]$excel.Hwnd)
        }
        if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){
            throw 'Owned guide window could not be foregrounded; not behavioral RED.'
        }
        [WorksheetInput]::Click($window,$x,$y)
    } finally {
        if(-not [GuideInputNative]::SetCursorPos($original.X,$original.Y)){
            throw ('Cursor restoration failed; Win32 error '+[Runtime.InteropServices.Marshal]::GetLastWin32Error())
        }
    }
    return $true
}

function Close-GuideWindowXNative {
    $window=Get-OwnedGuideInputWindow
    if(-not [GuideInputNative]::PostMessage($window,0x10,[IntPtr]::Zero,[IntPtr]::Zero)){
        throw ('Owned guide X delivery failed; Win32 error '+[Runtime.InteropServices.Marshal]::GetLastWin32Error())
    }
    $deadline=[DateTime]::UtcNow.AddSeconds(10)
    while([GuideInputNative]::IsWindow($window) -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
    return -not [GuideInputNative]::IsWindow($window)
}
