# D18 Admin UOM coverage through the actual packaged form handlers.
function Install-AdminUomActivityProbe {
    $form=$packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('frmAdminSettings').CodeModule
    $form.AddFromString(@'
Public Function UomActivityActionForTest(ByVal action As String, ByVal value As String) As String
    Dim i As Long
    Select Case action
        Case "Add"
            mTxtUom.Value = value
            mBtnUomAdd_Click
        Case "Remove"
            mLstUoms.ListIndex = -1
            For i = 0 To mLstUoms.ListCount - 1
                If CStr(mLstUoms.List(i, 0)) = value Then mLstUoms.ListIndex = i: Exit For
            Next i
            mBtnUomRemove_Click
        Case "Reset"
            mBtnUomReset_Click
        Case Else
            Err.Raise 5, , "Unknown fixture action"
    End Select
    UomActivityActionForTest = mLblStatus.Caption
End Function
Public Function UomSurfaceForTest() As Boolean
    Dim control As Variant
    mPages.Value = 0
    Me.Repaint
    If mLstConfig.ListCount = 0 Or mLstUoms.ListCount = 0 Then Exit Function
    For Each control In Array(mTxtUom, mLstUoms, mBtnUomAdd, mBtnUomRemove, mBtnUomReset)
        If Not control.Visible Or control.Left < 0 Or control.Top < 0 Then Exit Function
        If control.Left + control.Width > mPages.Width - 12 Then Exit Function
        If control.Top + control.Height > mPages.Height - 24 Then Exit Function
    Next control
    UomSurfaceForTest = True
End Function
'@)
    $packages['invSys.Admin.xlam'].VBProject.VBComponents.Item('TestD5Commands').CodeModule.AddFromString(@'
Public Function UomActivityAction(ByVal action As String, ByVal value As String) As String
    UomActivityAction = mForm.UomActivityActionForTest(action, value)
End Function
Public Function UomPublishForTest() As Boolean
    Dim report As String
    UomPublishForTest = modAdminConsole.GenerateInventorySnapshot("config-admin", "", Nothing, "", Nothing, report)
End Function
Public Function UomSurfaceForTest() As Boolean
    UomSurfaceForTest = mForm.UomSurfaceForTest()
End Function
'@)
}

function Invoke-AdminUomResetChoice([ValidateSet('Yes','No')][string]$Choice,[string]$CaptureName) {
    # The only click is the caller-selected response to this exact disposable
    # fixture's existing UOM reset question. No unrelated dialogs are dismissed.
    $handle=[long]$excel.Hwnd
    $ready=Join-Path $runRoot ('uom-dialog-'+[guid]::NewGuid().ToString('N')+'.ready')
    $capture=Join-Path $reportRoot $CaptureName
    $job=Start-Job -ArgumentList $handle,$Choice,$ready,$capture -ScriptBlock {
        param($handle,$choice,$ready,$capture)
        $ErrorActionPreference='Stop'
        Add-Type -ReferencedAssemblies System.Drawing @'
using System;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
public static class UomResetDialog {
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
    [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr h,out Rect r);
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] static extern int GetSystemMetrics(int index);
    public static uint Owner(IntPtr h) { uint p; GetWindowThreadProcessId(h,out p); return p; }
    static string Text(IntPtr h) { var s=new StringBuilder(256); GetWindowText(h,s,s.Capacity); return s.ToString(); }
    static string Class(IntPtr h) { var s=new StringBuilder(128); GetClassName(h,s,s.Capacity); return s.ToString(); }
    public static bool Exact(IntPtr h,uint owner) {
        if(Owner(h)!=owner || !IsWindowVisible(h) || Class(h)!="#32770" || Text(h)!="invSys Settings") return false;
        bool prompt=false;
        int buttons=0;
        EnumChildWindows(h,(child,p)=>{
            if(!IsWindowVisible(child)) return true;
            if(Class(child)=="Static" && Text(child)=="Reset the warehouse UOM catalog to defaults?") prompt=true;
            if(Class(child)=="Button") buttons++;
            return true;
        },IntPtr.Zero);
        var yes=GetDlgItem(h,6); var no=GetDlgItem(h,7);
        return prompt && buttons==2 && Owner(yes)==owner && Owner(no)==owner &&
               Text(yes).Replace("&","")=="Yes" && Text(no).Replace("&","")=="No";
    }
    public static IntPtr Find(uint owner) {
        IntPtr found=IntPtr.Zero;
        EnumWindows((h,p)=>{ if(Exact(h,owner)) found=h; return found==IntPtr.Zero; },IntPtr.Zero);
        return found;
    }
    public static bool Capture(IntPtr h,uint owner,string path) {
        Rect r; if(!Exact(h,owner) || GetForegroundWindow()!=h || !GetWindowRect(h,out r) || r.R<=r.L || r.B<=r.T || r.R-r.L>1600 || r.B-r.T>1000) return false;
        int left=GetSystemMetrics(76),top=GetSystemMetrics(77);
        if(r.L<left || r.T<top || r.R>left+GetSystemMetrics(78) || r.B>top+GetSystemMetrics(79)) return false;
        using(var b=new Bitmap(r.R-r.L,r.B-r.T)) {
            using(var g=Graphics.FromImage(b)) { g.CopyFromScreen(r.L,r.T,0,0,b.Size); }
            if(GetForegroundWindow()!=h || !Exact(h,owner)) return false;
            int ink=0;
            for(int y=b.Height/4;y<b.Height*63/100;y++) for(int x=b.Width/10;x<b.Width*92/100;x++) {
                Color c=b.GetPixel(x,y); if(c.R<220 || c.G<220 || c.B<220) ink++;
            }
            if(ink<100) return false;
            b.Save(path,System.Drawing.Imaging.ImageFormat.Png);
        }
        return true;
    }
    public static bool Focus(IntPtr h,uint owner) { return Exact(h,owner) && SetForegroundWindow(h); }
    public static bool Choose(IntPtr h,uint owner,bool yes) {
        return Exact(h,owner) && PostMessage(GetDlgItem(h,yes?6:7),0x00F5,IntPtr.Zero,IntPtr.Zero);
    }
}
'@
        $owner=[UomResetDialog]::Owner([IntPtr]$handle)
        if($owner -eq 0){throw 'Generated Excel owner is unavailable.'}
        $created=(Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks
        [IO.File]::WriteAllText($ready,'Ready')
        $until=[DateTime]::UtcNow.AddSeconds(30)
        while([DateTime]::UtcNow -lt $until){
            if((Get-Process -Id $owner).StartTime.ToUniversalTime().Ticks -ne $created){throw 'Generated process identity changed.'}
            $dialog=[UomResetDialog]::Find($owner)
            if($dialog -ne [IntPtr]::Zero){
                # Retry only observation of the same exact native question.
                # Never replay Reset. Reject blank/covered/offscreen captures.
                $captured=$false
                for($paint=0;$paint -lt 3 -and -not $captured;$paint++){
                    [void][UomResetDialog]::Focus($dialog,$owner)
                    Start-Sleep -Milliseconds 300
                    $captured=[UomResetDialog]::Capture($dialog,$owner,$capture)
                }
                $clicked=[UomResetDialog]::Choose($dialog,$owner,($choice -ceq 'Yes'))
                return [pscustomobject]@{ExactQuestion=$true;Captured=$captured;RequestedChoice=$choice;ClickDelivered=$clicked}
            }
            Start-Sleep -Milliseconds 100
        }
        throw 'Expected generated UOM confirmation did not appear.'
    }
    try {
        for($i=0;$i -lt 100 -and -not (Test-Path -LiteralPath $ready);$i++){Start-Sleep -Milliseconds 100}
        if(-not (Test-Path -LiteralPath $ready)){throw 'UOM confirmation observer did not become ready.'}
        $status=[string](Run 'invSys.Admin.xlam' 'TestD5Commands.UomActivityAction' @('Reset',''))
        [void](Wait-Job $job -Timeout 5)
        $observation=Receive-Job $job -ErrorAction Stop
        if($null -eq $observation -or -not $observation.ExactQuestion -or -not $observation.Captured -or -not $observation.ClickDelivered -or $observation.RequestedChoice -cne $Choice){throw 'Actual native confirmation evidence is incomplete.'}
        return $status
    } finally {
        if($job.State -eq 'Running'){Stop-Job $job}
        Remove-Job $job
    }
}
