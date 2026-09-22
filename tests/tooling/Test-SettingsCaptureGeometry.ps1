# Capture-harness geometry only; this does not establish product RED/GREEN.
[CmdletBinding()]
param([string]$RepoRoot='.',[ValidateSet('RED','GREEN')][string]$Phase='GREEN')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$root=Join-Path $repo ('reports/runtime/capture-geometry/'+[Guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
$tokens=$null;$parseErrors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tests/tooling/Test-Slice4beConfigCommands.ps1'),[ref]$tokens,[ref]$parseErrors)
if($parseErrors.Count){throw 'Capture controller does not parse.'}
$definition=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq 'Initialize-SettingsCapture'},$false)
if($null -eq $definition){throw 'Actual capture helper is unavailable.'}
. ([scriptblock]::Create($definition.Extent.Text))
Initialize-SettingsCapture
Add-Type -AssemblyName System.Windows.Forms
Add-Type @'
using System;using System.Runtime.InteropServices;
public static class CaptureGeometryReference {
 [StructLayout(LayoutKind.Sequential)]public struct Rect{public int Left,Top,Right,Bottom;}
 [DllImport("dwmapi.dll")]public static extern int DwmGetWindowAttribute(IntPtr h,uint attribute,out Rect r,int size);
 [DllImport("user32.dll")]public static extern IntPtr GetThreadDpiAwarenessContext();
 [DllImport("user32.dll")]public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
 [DllImport("user32.dll")]public static extern bool AreDpiAwarenessContextsEqual(IntPtr first,IntPtr second);
}
'@
$rows=[Collections.Generic.List[object]]::new()
function Check([string]$Name,[bool]$Passed){$rows.Add([pscustomobject]@{Check=$Name;Passed=$Passed});Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))}
$form=$null;$bitmap=$null
$previous=[CaptureGeometryReference]::SetThreadDpiAwarenessContext([IntPtr](-1))
if($previous -eq [IntPtr]::Zero){throw 'Cannot establish the DPI-unaware fixture context.'}
try {
    # A borderless, disposable solid-color window gives an independent physical
    # rectangle without reading or capturing any workbook or user content.
    $form=New-Object Windows.Forms.Form
    $form.Text='invSys capture geometry fixture'
    $form.FormBorderStyle=[Windows.Forms.FormBorderStyle]::None
    $form.StartPosition=[Windows.Forms.FormStartPosition]::Manual
    $form.Location=New-Object Drawing.Point(40,40)
    $form.Size=New-Object Drawing.Size(320,160)
    $form.BackColor=[Drawing.Color]::RoyalBlue
    $form.ShowInTaskbar=$false
    $form.Show()
    [Windows.Forms.Application]::DoEvents()
    $window=$form.Handle
    $physical=New-Object CaptureGeometryReference+Rect
    if([CaptureGeometryReference]::DwmGetWindowAttribute($window,9,[ref]$physical,[Runtime.InteropServices.Marshal]::SizeOf($physical)) -ne 0 -or $physical.Right -le $physical.Left -or $physical.Bottom -le $physical.Top){throw 'Independent physical fixture bounds unavailable.'}
    $before=[CaptureGeometryReference]::GetThreadDpiAwarenessContext()
    $imagePath=Join-Path $root 'owned-fixture.png'
    [InvSysSettingsCapture]::SaveWindow($window,$imagePath)
    $bitmap=[Drawing.Image]::FromFile($imagePath)
    Check 'Capture.PixelDimensionsMatchPhysicalWindow' ($bitmap.Width -eq ($physical.Right-$physical.Left) -and $bitmap.Height -eq ($physical.Bottom-$physical.Top))
    Check 'Capture.DpiContextRestoredAfterSuccess' ([CaptureGeometryReference]::AreDpiAwarenessContextsEqual($before,[CaptureGeometryReference]::GetThreadDpiAwarenessContext()))
    $refused=$false
    try {[InvSysSettingsCapture]::SaveWindow([IntPtr]::Zero,(Join-Path $root 'invalid.png'))}
    catch {$refused=$_.Exception.GetBaseException().Message -ceq 'Requested form window unavailable.'}
    Check 'Capture.InvalidWindowRefused' ($refused -and -not (Test-Path (Join-Path $root 'invalid.png')))
    Check 'Capture.DpiContextRestoredAfterFailure' ([CaptureGeometryReference]::AreDpiAwarenessContextsEqual($before,[CaptureGeometryReference]::GetThreadDpiAwarenessContext()))
    [pscustomobject]@{ExpectedWidth=$physical.Right-$physical.Left;ExpectedHeight=$physical.Bottom-$physical.Top;ImageWidth=$bitmap.Width;ImageHeight=$bitmap.Height}|ConvertTo-Json|Set-Content (Join-Path $root geometry.json)
}finally {
    if($null -ne $bitmap){$bitmap.Dispose()}
    if($null -ne $form){$form.Close();$form.Dispose()}
    [void][CaptureGeometryReference]::SetThreadDpiAwarenessContext($previous)
    ConvertTo-Json -InputObject @($rows.ToArray())|Set-Content (Join-Path $root ($Phase.ToLowerInvariant()+'.json'))
}
Write-Output ('REPORT_ROOT='+$root)
if(@($rows|Where-Object {-not $_.Passed}).Count){exit 1}
