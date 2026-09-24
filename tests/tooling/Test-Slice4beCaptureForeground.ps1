# Developer-only calibration of the existing capture helper with an empty fixture.
# This script opens no invSys package and changes none; this is not product RED/GREEN.
[CmdletBinding()]
param([string]$RepoRoot='.',[ValidateRange(0,200)][int]$WorkbookLifecycleIterations=0)
$ErrorActionPreference='Stop'
Set-StrictMode -Version Latest
# Imported capture helpers leave packaged guide diagnostics disabled here.
$GuideCaptureVisibleExcelForTest=$false
$GuideCaptureSavedWorkbookForTest=$false
$TraceGuideResourcesForTest=$false
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
 [DllImport("user32.dll")] public static extern uint GetGuiResources(IntPtr process,uint flags);
 [DllImport("user32.dll",EntryPoint="GetWindowLongPtrW")] static extern IntPtr GetLong64(IntPtr window,int index);
 [DllImport("user32.dll",EntryPoint="GetWindowLongW")] static extern int GetLong32(IntPtr window,int index);
 public static bool Topmost(IntPtr window){return ((IntPtr.Size==8 ? GetLong64(window,-20).ToInt64() : GetLong32(window,-20)) & 8)!=0;}
}
'@
$excel=$null;$books=$null;$book=$null;$components=$null;$form=$null;$module=$null;$process=$null
$observations=[Collections.Generic.List[object]]::new()
$started=[DateTimeOffset]::UtcNow.ToString('o')
$failure='';$normalQuit=$false
function Observe-Capture([string]$Case,[bool]$Visible,[bool]$RequireVisibility=$true,[bool]$ExpectCaptionRejection=$false,[bool]$DirectCaptionInput=$false){
    $excel.Visible=$Visible
    $actual=$excel.Visible
    if($RequireVisibility -and ($actual -isnot [bool] -or $actual -ne $Visible)){throw 'Fixture Excel visibility was not verified.'}
    [void]$excel.Run(("'"+$book.Name+"'!modForegroundFixture.ShowFixture"))
    $window=[InvSysSettingsCapture]::OwnedVisibleForm('invSys disposable capture fixture',[IntPtr]$excel.Hwnd)
    if($window -eq [IntPtr]::Zero){throw 'Fixture form is not uniquely owned and visible.'}
    $wasTopmost=[ForegroundFixtureOwner]::Topmost($window)
    if($ExpectCaptionRejection -or $DirectCaptionInput){
        [void][InvSysSettingsCapture]::SetForegroundWindow([IntPtr]$excel.Hwnd)
        Start-Sleep -Milliseconds 150
        if([InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -eq $window){throw 'Native caption calibration requires the separate owned workbook foreground.'}
    }
    $before=[InvSysSettingsCapture]::ForegroundIdentity($window)
    $result=if($ExpectCaptionRejection){'CaptionActivationReturned'}else{'Captured'}
    try {
        if($ExpectCaptionRejection){[void][InvSysSettingsCapture]::ActivateByCaptionClick($window,[IntPtr]$excel.Hwnd)}
        elseif($DirectCaptionInput){
            $clicked=[InvSysSettingsCapture]::ActivateByCaptionClick($window,[IntPtr]$excel.Hwnd)
            if(-not $clicked -or [InvSysSettingsCapture]::GetAncestor([InvSysSettingsCapture]::GetForegroundWindow(),2) -ne $window){throw 'Native caption activation was not established.'}
            [InvSysSettingsCapture]::SaveVisibleWindow($window,(Join-Path $reportRoot ($Case+'.png')))
        }
        else {CaptureOwnedFormEvidence 'invSys disposable capture fixture' ($Case+'.png') $window.ToInt64()}
    }
    catch {
        if($ExpectCaptionRejection -and $_.Exception.GetBaseException().Message -ceq 'No uncovered owned form caption point is available.'){
            $facts=[InvSysSettingsCapture]::CaptionFacts()
            if(@($facts|Where-Object {$_ -like 'Candidate|*|OnMonitor=False'}).Count -ne 3 -or @($facts|Where-Object {$_ -like 'Hit|*'}).Count -ne 0 -or @($facts|Where-Object {$_ -like 'Resources|*'}).Count -ne 1){throw 'Failure-point physical diagnostics are incomplete.'}
            $result='CaptionRejectedWithDiagnostics'
        }else{
            if($_.Exception.GetBaseException().Message -cne 'Requested form is not in the foreground.'){throw}
            $result='ForegroundRejected'
        }
    }
    $after=[InvSysSettingsCapture]::ForegroundIdentity($window)
    $topmostRestored=[ForegroundFixtureOwner]::Topmost($window) -eq $wasTopmost
    if(-not $topmostRestored){throw 'Caption calibration changed the original topmost state.'}
    $process.Refresh()
    $observations.Add([pscustomobject]@{Case=$Case;RequestedVisible=$Visible;ExcelVisible=$actual;VisibilityVerified=($actual -is [bool] -and $actual -eq $Visible);OwnedFormVerified=$true;TopmostRestored=$topmostRestored;Result=$result;Before=$before;After=$after;PhysicalCaptionFacts=[InvSysSettingsCapture]::CaptionFacts();Handles=$process.HandleCount;Gdi=[ForegroundFixtureOwner]::GetGuiResources($process.Handle,0);User=[ForegroundFixtureOwner]::GetGuiResources($process.Handle,1)})
    if($ExpectCaptionRejection -and $result -cne 'CaptionRejectedWithDiagnostics'){throw 'Offscreen fixture did not reproduce caption rejection.'}
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
Private mLeft As Single, mTop As Single
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
Public Sub MoveFixtureForTest(ByVal restore As Boolean)
    If restore Then
        mFixture.Left = mLeft: mFixture.Top = mTop
    Else
        mLeft = mFixture.Left: mTop = mFixture.Top
        mFixture.Left = -5000: mFixture.Top = -5000
    End If
End Sub
'@)
    Observe-Capture 'hidden-first' $false
    Observe-Capture 'visible' $true
    Observe-Capture 'hidden-restored' $false
    if($WorkbookLifecycleIterations -gt 0){
        # Simulate closing the previous visible workbook/form while an add-in
        # retains its project, followed by blank workbook creation/closure and reopening.
        [void]$excel.Run(("'"+$book.Name+"'!modForegroundFixture.CloseFixture"))
        $book.IsAddin=$true
        for($cycle=0;$cycle -lt $WorkbookLifecycleIterations;$cycle++){
            $transient=$books.Add()
            try {$transient.Close($false)} finally {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($transient)}
        }
        Observe-Capture 'addin-after-workbook-cycles' $false
        [void]$excel.Run(("'"+$book.Name+"'!modForegroundFixture.CloseFixture"))
        $surface=$books.Add()
        try {Observe-Capture 'reopened-with-workbook' $true}
        finally {$surface.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($surface)}
        # Observe add-in-only readback rather than treating a visibility mismatch
        # as evidence that the owned native form cannot be captured.
        Observe-Capture 'reopened-after-workbook-close' $false $false
        # Deliberately offscreen disposable form exercises rejection diagnostics;
        # no click can reach an unrelated window and no fallback capture is used.
        $anchor=$books.Add();$excel.Visible=$true
        Observe-Capture 'native-caption-activation' $true $true $false $true
        [void]$excel.Run(("'"+$book.Name+"'!modForegroundFixture.MoveFixtureForTest"),$false)
        try {Observe-Capture 'offscreen-rejection' $true $true $true}
        finally {
            [void]$excel.Run(("'"+$book.Name+"'!modForegroundFixture.MoveFixtureForTest"),$true)
            $anchor.Close($false);[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($anchor)
        }
    }
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
    [pscustomobject]@{StartedUTC=$started;FinishedUTC=[DateTimeOffset]::UtcNow.ToString('o');WorkbookLifecycleIterations=$WorkbookLifecycleIterations;Observations=@($observations.ToArray());FailureType=$failure;NormalQuitRequested=$normalQuit;OwnedExcelClosed=$closed;ForcedTermination=$false;ProductAcceptance=$false}|ConvertTo-Json -Depth 5|Set-Content (Join-Path $reportRoot 'observations.json')
    Write-Output ('Calibration report: '+$reportRoot)
}
