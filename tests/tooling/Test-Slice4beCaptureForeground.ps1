# Developer-only calibration of the existing capture helper with an empty fixture.
# This script opens no invSys package and changes none; this is not product RED/GREEN.
[CmdletBinding()]
param([string]$RepoRoot='.')
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
# Imported capture helpers leave packaged guide diagnostics disabled here.
$GuideCaptureVisibleExcelForTest=$false
$GuideCaptureSavedWorkbookForTest=$false
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
if(Get-Process EXCEL -ErrorAction SilentlyContinue){throw 'Excel must be closed before disposable capture calibration.'}
$reportRoot=Join-Path $repo ('reports/runtime/capture-foreground-calibration/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $reportRoot|Out-Null
$tokens=$null;$parseErrors=$null
$tree=[Management.Automation.Language.Parser]::ParseFile((Join-Path $repo 'tests/tooling/Test-Slice4beConfigCommands.ps1'),[ref]$tokens,[ref]$parseErrors)
if($parseErrors.Count){throw 'Existing capture harness does not parse.'}
foreach($name in @('Initialize-SettingsCapture','CaptureFormEvidence','CaptureOwnedFormEvidence')){
    $function=$tree.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -ceq $name},$true)
    if($null -eq $function){throw ('Required capture function missing: '+$name)}
    Invoke-Expression $function.Extent.Text
}
Initialize-SettingsCapture
Add-Type @'
using System;using System.Runtime.InteropServices;
public static class ForegroundFixtureOwner {
 [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window,out uint process);
}
'@
$excel=$null;$books=$null;$book=$null;$components=$null;$form=$null;$module=$null;$process=$null
$observations=[Collections.Generic.List[object]]::new()
$started=[DateTimeOffset]::UtcNow.ToString('o')
$failure='';$normalQuit=$false
function Observe-Capture([string]$Case,[bool]$Visible){
    $excel.Visible=$Visible
    $actual=$excel.Visible
    if($actual -isnot [bool] -or $actual -ne $Visible){throw 'Fixture Excel visibility was not verified.'}
    [void]$excel.Run(("'"+$book.Name+"'!modForegroundFixture.ShowFixture"))
    $window=[InvSysSettingsCapture]::OwnedVisibleForm('invSys disposable capture fixture',[IntPtr]$excel.Hwnd)
    if($window -eq [IntPtr]::Zero){throw 'Fixture form is not uniquely owned and visible.'}
    $before=[InvSysSettingsCapture]::ForegroundIdentity($window)
    $result='Captured'
    try {CaptureOwnedFormEvidence 'invSys disposable capture fixture' ($Case+'.png') $window.ToInt64()}
    catch {
        if($_.Exception.GetBaseException().Message -cne 'Requested form is not in the foreground.'){throw}
        $result='ForegroundRejected'
    }
    $after=[InvSysSettingsCapture]::ForegroundIdentity($window)
    $observations.Add([pscustomobject]@{Case=$Case;ExcelVisible=$actual;OwnedFormVerified=$true;Result=$result;Before=$before;After=$after})
    Write-Output ($Case+': '+$result)
}
try {
    $excel=New-Object -ComObject Excel.Application
    $window=$excel.Hwnd
    if($null -eq $window -or [long]$window -eq 0){throw 'Owned Excel window unavailable.'}
    [uint32]$ownerId=0
    [void][ForegroundFixtureOwner]::GetWindowThreadProcessId([IntPtr]$window,[ref]$ownerId)
    if($ownerId -eq 0){throw 'Owned Excel process unavailable.'}
    $process=Get-Process -Id $ownerId
    if($process.ProcessName -cne 'EXCEL'){throw 'Fixture owner is not Excel.'}
    $books=$excel.Workbooks
    if($null -eq $books -or $books.Count -isnot [int] -or $books.Count -ne 0){throw 'Disposable instance was not empty.'}
    $excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=1;$excel.Visible=$false
    $book=$books.Add();$components=$book.VBProject.VBComponents
    $form=$components.Add(3);$form.Name='frmForegroundFixture'
    $form.CodeModule.AddFromString(@'
Option Explicit
Private Sub UserForm_Initialize()
    Dim text As MSForms.Label
    Me.Caption = "invSys disposable capture fixture"
    Me.Width = 540: Me.Height = 220
    Set text = Me.Controls.Add("Forms.Label.1", "lblFixture", True)
    text.Move 20, 30, 480, 100
    text.Caption = "Blank workbook and disposable test form. No inventory, credentials or operational workbook data."
    text.WordWrap = True
End Sub
'@)
    $module=$components.Add(1);$module.Name='modForegroundFixture'
    $module.CodeModule.AddFromString(@'
Option Explicit
Private mFixture As frmForegroundFixture
Public Sub ShowFixture()
    If mFixture Is Nothing Then Set mFixture = New frmForegroundFixture
    If Not mFixture.Visible Then mFixture.Show vbModeless
    mFixture.Repaint
    DoEvents
End Sub
Public Sub CloseFixture()
    If mFixture Is Nothing Then Exit Sub
    Unload mFixture: Set mFixture = Nothing
End Sub
'@)
    Observe-Capture 'hidden-first' $false
    Observe-Capture 'visible' $true
    Observe-Capture 'hidden-restored' $false
} catch {
    $failure=$_.Exception.GetType().FullName
    throw
} finally {
    if($null -ne $book){
        try {[void]$excel.Run(("'"+$book.Name+"'!modForegroundFixture.CloseFixture"))}catch{}
        try {$book.Close($false)}catch{}
    }
    if($null -ne $excel -and $null -ne $process){
        try {
            $remaining=$excel.Workbooks.Count
            if($remaining -is [int] -and $remaining -eq 0){$excel.Quit();$normalQuit=$true}
        }catch{}
    }
    foreach($value in @($module,$form,$components,$book,$books,$excel)){
        if($null -ne $value){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($value)}catch{}}
    }
    [GC]::Collect();[GC]::WaitForPendingFinalizers()
    $closed=$false
    if($null -ne $process){$closed=$process.WaitForExit(10000);$process.Dispose()}
    [pscustomobject]@{StartedUTC=$started;FinishedUTC=[DateTimeOffset]::UtcNow.ToString('o');Observations=@($observations.ToArray());FailureType=$failure;NormalQuitRequested=$normalQuit;OwnedExcelClosed=$closed;ForcedTermination=$false;ProductAcceptance=$false}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $reportRoot 'observations.json')
    Write-Output ('Calibration report: '+$reportRoot)
}
