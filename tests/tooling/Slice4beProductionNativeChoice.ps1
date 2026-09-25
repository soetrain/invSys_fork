# Native evidence is limited to this generated Excel process and four fixed questions.
function Invoke-ProductionNativeCancellation {
    param([ValidateSet('Process','Recipe')][string]$Designer,
          [ValidateSet('Release','Obsolete')][string]$Action,[string]$CaptureName)
    $handle=[long]$excel.Hwnd
    $ready=Join-Path $runRoot ('production-dialog-'+[guid]::NewGuid().ToString('N')+'.ready')
    $capture=Join-Path $reportRoot $CaptureName
    $job=Start-Job -ArgumentList $handle,$Designer,$Action,$ready,$capture -ScriptBlock {
        param($handle,$designer,$action,$ready,$capture)
        $ErrorActionPreference='Stop'
        Add-Type -ReferencedAssemblies System.Drawing @'
using System;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
public static class ProductionCancelDialog {
    public delegate bool EnumProc(IntPtr h,IntPtr p);
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int L,T,R,B; }
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc f,IntPtr p);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr h,EnumProc f,IntPtr p);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h,StringBuilder s,int n);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr h,StringBuilder s,int n);
    [DllImport("user32.dll")] static extern IntPtr GetDlgItem(IntPtr h,int id);
    [DllImport("user32.dll")] static extern bool PostMessage(IntPtr h,uint msg,IntPtr w,IntPtr l);
    [DllImport("user32.dll")] static extern IntPtr SendMessage(IntPtr h,uint msg,IntPtr w,IntPtr l);
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h,out Rect r);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
    [DllImport("user32.dll")] static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    sealed class PhysicalPixels : IDisposable {
        readonly IntPtr previous;
        public PhysicalPixels() {
            previous=SetThreadDpiAwarenessContext(new IntPtr(-4));
            if(previous==IntPtr.Zero)throw new Exception("Physical capture coordinates unavailable.");
        }
        public void Dispose() {
            if(SetThreadDpiAwarenessContext(previous)==IntPtr.Zero)
                throw new Exception("Capture DPI context could not be restored.");
        }
    }
    public static uint Owner(IntPtr h) { uint p; GetWindowThreadProcessId(h,out p); return p; }
    static string Text(IntPtr h) { var s=new StringBuilder(256); GetWindowText(h,s,s.Capacity); return s.ToString(); }
    static string Class(IntPtr h) { var s=new StringBuilder(128); GetClassName(h,s,s.Capacity); return s.ToString(); }
    public static bool Exact(IntPtr h,uint owner,string title,string prompt) {
        if(Owner(h)!=owner || !IsWindowVisible(h) || Class(h)!="#32770" || Text(h)!=title) return false;
        bool matched=false; int buttons=0;
        EnumChildWindows(h,(child,p)=>{
            if(!IsWindowVisible(child)) return true;
            if(Class(child)=="Static" && Text(child)==prompt) matched=true;
            if(Class(child)=="Button") buttons++;
            return true;
        },IntPtr.Zero);
        var yes=GetDlgItem(h,6); var no=GetDlgItem(h,7);
        return matched && buttons==2 && Owner(yes)==owner && Owner(no)==owner &&
               Text(yes).Replace("&","")=="Yes" && Text(no).Replace("&","")=="No";
    }
    public static IntPtr Find(uint owner,string title,string prompt) {
        IntPtr found=IntPtr.Zero;
        EnumWindows((h,p)=>{ if(Exact(h,owner,title,prompt)) found=h; return found==IntPtr.Zero; },IntPtr.Zero);
        return found;
    }
    public static bool DefaultNo(IntPtr h) {
        return (SendMessage(h,0x0400,IntPtr.Zero,IntPtr.Zero).ToInt64() & 0xffff)==7;
    }
    public static bool Capture(IntPtr h,uint owner,string title,string prompt,string path) {
        // Match window bounds to CopyFromScreen pixels without changing process DPI.
        using(var pixels=new PhysicalPixels()) {
        Rect r;
        if(!Exact(h,owner,title,prompt) || GetForegroundWindow()!=h || !GetWindowRect(h,out r) ||
           r.R<=r.L || r.B<=r.T || r.R-r.L>1600 || r.B-r.T>1000) return false;
        int left=GetSystemMetrics(76),top=GetSystemMetrics(77);
        if(r.L<left || r.T<top || r.R>left+GetSystemMetrics(78) || r.B>top+GetSystemMetrics(79)) return false;
        using(var b=new Bitmap(r.R-r.L,r.B-r.T)) {
            using(var g=Graphics.FromImage(b)) { g.CopyFromScreen(r.L,r.T,0,0,b.Size); }
            if(GetForegroundWindow()!=h || !Exact(h,owner,title,prompt)) return false;
            int ink=0,minLight=765,maxLight=0;
            for(int y=b.Height/4;y<b.Height*63/100;y++) for(int x=b.Width/10;x<b.Width*92/100;x++) {
                Color c=b.GetPixel(x,y); if(c.R<220 || c.G<220 || c.B<220) ink++;
                int light=c.R+c.G+c.B; minLight=Math.Min(minLight,light); maxLight=Math.Max(maxLight,light);
            }
            if(ink<100 || maxLight-minLight<120) return false;
            b.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
        return true;
        }
    }
    public static bool Focus(IntPtr h,uint owner,string title,string prompt) {
        return Exact(h,owner,title,prompt) && SetForegroundWindow(h);
    }
    public static bool Decline(IntPtr h,uint owner,string title,string prompt) {
        return Exact(h,owner,title,prompt) && PostMessage(GetDlgItem(h,7),0x00F5,IntPtr.Zero,IntPtr.Zero);
    }
}
'@
        $owner=[ProductionCancelDialog]::Owner([IntPtr]$handle)
        if($owner -eq 0){throw 'Generated Excel owner is unavailable.'}
        $created=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
        $title=$designer+' Designer'
        $prompt=if($action -ceq 'Release'){'Release this immutable '+$designer+' version?'}else{'Obsolete this '+$designer+' version?'}
        [IO.File]::WriteAllText($ready,'Ready')
        $until=[DateTime]::UtcNow.AddSeconds(30)
        while([DateTime]::UtcNow -lt $until){
            if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $created){throw 'Generated process identity changed.'}
            $dialog=[ProductionCancelDialog]::Find($owner,$title,$prompt)
            if($dialog -ne [IntPtr]::Zero){
                $captured=$false;$clicked=$false;$defaultNo=[ProductionCancelDialog]::DefaultNo($dialog)
                try {
                    for($paint=0;$paint -lt 3 -and -not $captured;$paint++){
                        [void][ProductionCancelDialog]::Focus($dialog,$owner,$title,$prompt)
                        Start-Sleep -Milliseconds 300
                        try {$captured=[ProductionCancelDialog]::Capture($dialog,$owner,$title,$prompt,$capture)}catch {$captured=$false}
                    }
                } finally {
                    # Decline even if capture fails; never confirm, retry the action or dismiss another dialog.
                    $clicked=[ProductionCancelDialog]::Decline($dialog,$owner,$title,$prompt)
                }
                return [pscustomobject]@{ExactQuestion=$true;DefaultNo=$defaultNo;Captured=$captured;NoClickDelivered=$clicked}
            }
            Start-Sleep -Milliseconds 100
        }
        throw 'Expected generated Production question did not appear.'
    }
    try {
        for($i=0;$i -lt 100 -and -not (Test-Path -LiteralPath $ready);$i++){Start-Sleep -Milliseconds 100}
        if(-not(Test-Path -LiteralPath $ready)){throw 'Production confirmation observer did not become ready.'}
        $returned=[string](Run 'invSys.Operations.xlam' 'TestProductionDesigner.NativeCancel' @($Designer,$Action))
        [void](Wait-Job $job -Timeout 5)
        $observed=Receive-Job $job -ErrorAction Stop
        if($null -eq $observed){throw 'Native confirmation evidence unavailable.'}
        return [pscustomobject]@{HandlerReturned=$returned -ceq 'RETURNED';ExactQuestion=$observed.ExactQuestion;DefaultNo=$observed.DefaultNo;Captured=$observed.Captured;NoClickDelivered=$observed.NoClickDelivered}
    } finally {
        if($job.State -eq 'Running'){Stop-Job $job}
        Remove-Job $job
    }
}
