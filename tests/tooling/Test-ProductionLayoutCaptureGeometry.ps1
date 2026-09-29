# Developer capture calibration; packaged geometry and visual review remain separate.
[CmdletBinding()]
param([ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop';Set-StrictMode -Version Latest
$repo=[IO.Path]::GetFullPath((Join-Path $PSScriptRoot '../..'))
$root=Join-Path $repo ('reports/runtime/production-layout-capture/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$path=Join-Path $repo 'tools/validate_slice9_production_layout.ps1'
$source=Get-Content -LiteralPath $path -Raw
$native=[regex]::Match($source,'(?s)Add-Type -TypeDefinition @"\r?\n(.*?)\r?\n"@')
if(-not $native.Success){throw 'Actual layout native helper unavailable.'}
Add-Type -TypeDefinition $native.Groups[1].Value
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($path,[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Layout tool does not parse.'}
$function=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Save-WindowScreenshot'},$false)
if($null -eq $function){throw 'Actual layout capture function unavailable.'}
. ([scriptblock]::Create($function.Extent.Text))
Add-Type -AssemblyName System.Windows.Forms
Add-Type @'
using System;using System.Runtime.InteropServices;
public static class LayoutCaptureReference {
 [StructLayout(LayoutKind.Sequential)]public struct Rect{public int Left,Top,Right,Bottom;}
 [DllImport("dwmapi.dll")]public static extern int DwmGetWindowAttribute(IntPtr h,uint attribute,out Rect r,int size);
 [DllImport("user32.dll")]public static extern IntPtr GetThreadDpiAwarenessContext();
 [DllImport("user32.dll")]public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
 [DllImport("user32.dll")]public static extern bool AreDpiAwarenessContextsEqual(IntPtr a,IntPtr b);
}
'@
$rows=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$rows.Add([pscustomobject]@{Check=$Name;Passed=$Passed});Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))}
$form=$null;$bitmap=$null
$previous=[LayoutCaptureReference]::SetThreadDpiAwarenessContext([IntPtr](-1))
if($previous -eq [IntPtr]::Zero){throw 'DPI-unaware fixture context unavailable.'}
try {
    $form=New-Object Windows.Forms.Form
    $form.Text='invSys layout capture fixture'
    $form.FormBorderStyle=[Windows.Forms.FormBorderStyle]::None
    $form.StartPosition=[Windows.Forms.FormStartPosition]::Manual
    $form.Location=New-Object Drawing.Point(40,40)
    $form.Size=New-Object Drawing.Size(320,160)
    $form.BackColor=[Drawing.Color]::RoyalBlue
    $form.ShowInTaskbar=$false
    $form.Show();[Windows.Forms.Application]::DoEvents()
    $window=$form.Handle;$physical=New-Object LayoutCaptureReference+Rect
    if([LayoutCaptureReference]::DwmGetWindowAttribute($window,9,[ref]$physical,[Runtime.InteropServices.Marshal]::SizeOf($physical)) -ne 0){throw 'Independent physical bounds unavailable.'}
    $before=[LayoutCaptureReference]::GetThreadDpiAwarenessContext()
    $imagePath=Join-Path $root 'fixture.png'
    Save-WindowScreenshot -Handle $window -Path $imagePath
    $bitmap=[Drawing.Bitmap]::new($imagePath)
    Check 'LayoutCapture.PhysicalWidth' ($bitmap.Width -eq $physical.Right-$physical.Left)
    Check 'LayoutCapture.PhysicalHeight' ($bitmap.Height -eq $physical.Bottom-$physical.Top)
    Check 'LayoutCapture.CallerContextAfterSuccess' ([LayoutCaptureReference]::AreDpiAwarenessContextsEqual($before,[LayoutCaptureReference]::GetThreadDpiAwarenessContext()))
    $invalid=Join-Path $root 'invalid.png';$refused=$false
    try {Save-WindowScreenshot -Handle ([IntPtr]::Zero) -Path $invalid}
    catch {$refused=$true}
    Check 'LayoutCapture.InvalidHandleRefused' ($refused -and -not (Test-Path -LiteralPath $invalid))
    Check 'LayoutCapture.CallerContextAfterInvalidHandle' ([LayoutCaptureReference]::AreDpiAwarenessContextsEqual($before,[LayoutCaptureReference]::GetThreadDpiAwarenessContext()))
    $failed=$false
    try {Save-WindowScreenshot -Handle $window -Path (Join-Path $root 'missing/fixture.png')}
    catch {$failed=$true}
    Check 'LayoutCapture.WriteFailureReported' $failed
    Check 'LayoutCapture.CallerContextAfterWriteFailure' ([LayoutCaptureReference]::AreDpiAwarenessContextsEqual($before,[LayoutCaptureReference]::GetThreadDpiAwarenessContext()))
} finally {
    if($null -ne $bitmap){$bitmap.Dispose()}
    if($null -ne $form){$form.Close();$form.Dispose()}
    [void][LayoutCaptureReference]::SetThreadDpiAwarenessContext($previous)
    ConvertTo-Json -InputObject @($rows.ToArray())|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
}
Write-Output ('REPORT_ROOT='+$root)
if(@($rows|Where-Object {-not $_.Passed}).Count){exit 1}
