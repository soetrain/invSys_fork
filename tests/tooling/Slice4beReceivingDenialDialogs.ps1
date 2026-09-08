# Rendered evidence for fixed notices in an owned disposable Excel process.
# Only the existing observer may dismiss a sole OK; no multi-button approval.
function Start-ReceivingDenialDialogEvidence {
    $directory=Join-Path $reportRoot 'launcher-denial-dialogs'
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $ready=Join-Path $runRoot 'denial-dialog-ready'
    $stop=Join-Path $runRoot 'denial-dialog-stop'
    $job=Start-Job -ArgumentList $repo,([long]$excel.Hwnd),$directory,$ready,$stop -ScriptBlock {
        param($repo,$handle,$directory,$ready,$stop)
        $ErrorActionPreference='Stop'
        . (Join-Path $repo 'tools/plan022-dialog-observer.ps1')
        Invoke-Plan022NativeDialogObservation -ProcessId 0 -TimeoutSeconds 0
        Add-Type -ReferencedAssemblies System.Drawing @'
using System; using System.Text; using System.Drawing; using System.Runtime.InteropServices;
public static class ReceivingDenialDialog {
    public delegate bool EnumProc(IntPtr h,IntPtr ignored);
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f,IntPtr p);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h,EnumProc f,IntPtr p);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h,StringBuilder s,int n);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h,StringBuilder s,int n);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h,out Rect r);
    [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr h,IntPtr dc,uint flags);
    public static uint Owner(IntPtr h) { uint p; GetWindowThreadProcessId(h,out p); return p; }
    static string Class(IntPtr h) { var s=new StringBuilder(128); GetClassName(h,s,s.Capacity); return s.ToString(); }
    static string Text(IntPtr h) { var s=new StringBuilder(2048); GetWindowText(h,s,s.Capacity); return s.ToString(); }
    public static IntPtr Find(uint process,string expected) {
        IntPtr found=IntPtr.Zero;
        EnumWindows((h,ignored)=>{
            if(Owner(h)!=process || !IsWindowVisible(h) || Class(h)!="#32770") return true;
            bool matches=false;
            EnumChildWindows(h,(child,unused)=>{
                if(Owner(child)==process && IsWindowVisible(child) && Class(child)=="Static" && Text(child).Contains(expected)) matches=true;
                return true;
            },IntPtr.Zero);
            if(matches) found=h;
            return found==IntPtr.Zero;
        },IntPtr.Zero);
        return found;
    }
    public static bool Save(IntPtr h,uint process,string path) {
        Rect r; if(Owner(h)!=process || !IsWindowVisible(h) || !GetWindowRect(h,out r)) return false;
        if(r.Right<=r.Left || r.Bottom<=r.Top || r.Right-r.Left>2000 || r.Bottom-r.Top>2000) return false;
        using(var bitmap=new Bitmap(r.Right-r.Left,r.Bottom-r.Top)) {
            using(var graphics=Graphics.FromImage(bitmap)) {
                var dc=graphics.GetHdc(); bool ok;
                try { ok=PrintWindow(h,dc,2); } finally { graphics.ReleaseHdc(dc); }
                if(!ok) return false;
            }
            bitmap.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
        return true;
    }
}
'@
        $owned=[ReceivingDenialDialog]::Owner([IntPtr]$handle)
        if($owned -eq 0){throw 'Owned denial evidence process unavailable.'}
        $observed=@{DeniedVisible=$false;DeniedCaptured=$false;TrackingVisible=$false;TrackingCaptured=$false}
        Set-Content -LiteralPath $ready -Value 'ready'
        $deadline=[DateTime]::UtcNow.AddSeconds(60)
        while([DateTime]::UtcNow -lt $deadline -and -not (Test-Path -LiteralPath $stop)) {
            foreach($case in @(@('Denied','Current user does not have RECEIVE_POST for this warehouse/station.'),@('Tracking','Tracking unavailable'))) {
                if($observed[$case[0]+'Visible']){continue}
                $window=[ReceivingDenialDialog]::Find($owned,$case[1])
                if($window -ne [IntPtr]::Zero) {
                    $observed[$case[0]+'Visible']=$true
                    try {$observed[$case[0]+'Captured']=[ReceivingDenialDialog]::Save($window,$owned,(Join-Path $directory ($case[0]+'.png')))} catch {}
                }
            }
            # Raw modal text remains in memory and is discarded, never reported.
            $null=[Plan022NativeDialogs]::Poll($owned)
            if($observed.DeniedVisible -and $observed.TrackingVisible){break}
            Start-Sleep -Milliseconds 100
        }
        [pscustomobject]$observed
    }
    $deadline=[DateTime]::UtcNow.AddSeconds(12)
    while(-not (Test-Path -LiteralPath $ready) -and $job.State -eq 'Running' -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
    if(-not (Test-Path -LiteralPath $ready)){Stop-Job $job; Remove-Job $job; throw 'Denial dialog observer did not become ready.'}
    return $job
}

function Complete-ReceivingDenialDialogEvidence($Job) {
    [void](Wait-Job $Job -Timeout 3)
    Set-Content -LiteralPath (Join-Path $runRoot 'denial-dialog-stop') -Value 'stop'
    [void](Wait-Job $Job -Timeout 5)
    try {
        if($Job.State -ne 'Completed'){throw 'Denial dialog observer did not complete.'}
        $observed=Receive-Job $Job
        foreach($name in @('DeniedVisible','DeniedCaptured','TrackingVisible','TrackingCaptured')){Check ('LauncherDenial.Dialog.'+$name) ([bool]$observed.$name)}
        $observed | Select-Object DeniedVisible,DeniedCaptured,TrackingVisible,TrackingCaptured | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $reportRoot 'launcher-denial-dialogs/results.json')
    } finally {if($Job.State -eq 'Running'){Stop-Job $Job}; Remove-Job $Job}
}
