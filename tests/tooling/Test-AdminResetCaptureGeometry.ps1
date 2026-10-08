# Admin Reset capture geometry; actual packaged carrier/UOM handlers remain separate gates.
[CmdletBinding()]
param([ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$root=Join-Path (Get-Location) ('reports/runtime/admin-reset-geometry/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$source=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'Slice4beAdminUomProbe.ps1') -Raw
$match=[regex]::Match($source,"(?s)Add-Type -ReferencedAssemblies System.Drawing @'\r?\n(.*?)\r?\n'@")
if(-not $match.Success){throw 'Actual native observer source unavailable.'}
Add-Type -ReferencedAssemblies System.Drawing -TypeDefinition $match.Groups[1].Value
Add-Type -AssemblyName System.Windows.Forms
Add-Type -ReferencedAssemblies System.Drawing @'
using System;using System.Drawing;using System.Runtime.InteropServices;
public static class NativeGeometryReference {
 [StructLayout(LayoutKind.Sequential)]struct Rect{public int L,T,R,B;}
 [DllImport("user32.dll")]static extern bool GetWindowRect(IntPtr h,out Rect r);
 [DllImport("user32.dll")]public static extern IntPtr GetThreadDpiAwarenessContext();
 [DllImport("user32.dll")]public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
 [DllImport("user32.dll")]public static extern bool AreDpiAwarenessContextsEqual(IntPtr a,IntPtr b);
 public static Bitmap PhysicalReference(IntPtr h) {
  IntPtr prior=SetThreadDpiAwarenessContext(new IntPtr(-4));
  if(prior==IntPtr.Zero)throw new Exception("Physical reference unavailable.");
  try {Rect r;if(!GetWindowRect(h,out r))throw new Exception("Reference bounds unavailable.");
   var b=new Bitmap(r.R-r.L,r.B-r.T);
   using(var g=Graphics.FromImage(b)){g.CopyFromScreen(r.L,r.T,0,0,b.Size);}return b;
  }finally{SetThreadDpiAwarenessContext(prior);}
 }
 public static bool SameInterior(Bitmap actual,Bitmap expected) {
  if(actual.Width!=expected.Width || actual.Height!=expected.Height)return false;
  int same=0,total=0;
  for(int y=actual.Height/4;y<actual.Height*3/4;y+=7)for(int x=actual.Width/8;x<actual.Width*7/8;x+=7){
   total++;if(actual.GetPixel(x,y).ToArgb()==expected.GetPixel(x,y).ToArgb())same++;
  }return total>0 && same>=total*0.98;
 }
}
'@
[UomResetDialog]::ExpectedPrompt='Reset shipping carriers to defaults?'
$rows=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$rows.Add([pscustomobject]@{Check=$Name;Passed=$Passed});Write-Host ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))}
$script:observerFailure=$false;$script:observed=$false
$timer=New-Object Windows.Forms.Timer;$timer.Interval=700
$timer.add_Tick({
    $timer.Stop();$dialog=[IntPtr]::Zero;$actual=$null;$expected=$null;$prior=[IntPtr]::Zero
    try {
        $dialog=[UomResetDialog]::Find([uint32]$PID)
        if($dialog -eq [IntPtr]::Zero){throw 'Synthetic native question unavailable.'}
        $script:observed=$true
        $prior=[NativeGeometryReference]::SetThreadDpiAwarenessContext([IntPtr](-1))
        if($prior -eq [IntPtr]::Zero){throw 'DPI-unaware fixture context unavailable.'}
        $before=[NativeGeometryReference]::GetThreadDpiAwarenessContext()
        [void][UomResetDialog]::Focus($dialog,[uint32]$PID)
        $path=Join-Path $root 'question.png'
        $captured=[UomResetDialog]::Capture($dialog,[uint32]$PID,$path)
        Check 'NativeCapture.ExactOwnedQuestionCaptured' $captured
        Check 'NativeCapture.CallerDpiRestoredAfterSuccess' ([NativeGeometryReference]::AreDpiAwarenessContextsEqual($before,[NativeGeometryReference]::GetThreadDpiAwarenessContext()))
        $expected=[NativeGeometryReference]::PhysicalReference($dialog)
        if($captured){$actual=[Drawing.Bitmap]::new($path)}
        Check 'NativeCapture.PhysicalPixelDimensions' ($null -ne $actual -and $actual.Width -eq $expected.Width -and $actual.Height -eq $expected.Height)
        Check 'NativeCapture.ExactQuestionScreenPixels' ($null -ne $actual -and [NativeGeometryReference]::SameInterior($actual,$expected))
        $invalid=Join-Path $root 'invalid.png'
        $refused=-not [UomResetDialog]::Capture([IntPtr]::Zero,[uint32]$PID,$invalid)
        Check 'NativeCapture.InvalidWindowRefused' ($refused -and -not (Test-Path -LiteralPath $invalid))
        Check 'NativeCapture.CallerDpiRestoredAfterFailure' ([NativeGeometryReference]::AreDpiAwarenessContextsEqual($before,[NativeGeometryReference]::GetThreadDpiAwarenessContext()))
    } catch {$script:observerFailure=$true}
    finally {
        if($null -ne $actual){$actual.Dispose()};if($null -ne $expected){$expected.Dispose()}
        if($prior -ne [IntPtr]::Zero){[void][NativeGeometryReference]::SetThreadDpiAwarenessContext($prior)}
        if($dialog -ne [IntPtr]::Zero){[void][UomResetDialog]::Choose($dialog,[uint32]$PID,$false)}
    }
})
try {
    $timer.Start()
    $choice=[Windows.Forms.MessageBox]::Show('Reset shipping carriers to defaults?','invSys Settings',[Windows.Forms.MessageBoxButtons]::YesNo,[Windows.Forms.MessageBoxIcon]::Question,[Windows.Forms.MessageBoxDefaultButton]::Button2)
    Check 'NativeCapture.OnlyNoChosen' ($choice -eq [Windows.Forms.DialogResult]::No)
    Check 'NativeCapture.ObserverCompleted' ($script:observed -and -not $script:observerFailure)
} finally {
    $timer.Stop();$timer.Dispose()
    ConvertTo-Json -InputObject @($rows.ToArray())|Set-Content -LiteralPath (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
}
Write-Output ('REPORT_ROOT='+$root)
if(@($rows|Where-Object {-not $_.Passed}).Count){exit 1}
