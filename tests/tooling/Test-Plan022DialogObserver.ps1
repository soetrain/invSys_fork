[CmdletBinding()]
param([string]$RepoRoot='.', [ValidateSet('RED','GREEN')][string]$Phase='GREEN')
Set-StrictMode -Version Latest
$ErrorActionPreference='Stop'
$repo=(Resolve-Path -LiteralPath $RepoRoot).Path
$tokens=$null; $errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile(
    (Join-Path $repo 'tools/validate_plan022_packaged_launchers.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count){throw 'Launcher harness parse failed.'}
$definition=$ast.FindAll({param($n)
    $n -is [Management.Automation.Language.FunctionDefinitionAst] -and
    $n.Name -eq 'Start-DialogCaptureAndDismiss'
},$true)[0]
. ([scriptblock]::Create($definition.Extent.Text))
$root=Join-Path ([IO.Path]::GetTempPath()) ('invsys-dialog-observer-'+[guid]::NewGuid().ToString('N'))
$reportRoot=Join-Path $repo 'reports/runtime/slice4be-dialog-observer'
New-Item -ItemType Directory -Path $root,$reportRoot -Force | Out-Null
$results=[Collections.Generic.List[object]]::new()
$workers=@(); $observer=$null; $observerStop=Join-Path $root 'observer-stop'
function Check([string]$Name,[bool]$Passed){
    $results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})
    Write-Output ($Name+': '+$(if($Passed){'PASS'}else{'FAIL'}))
}
Add-Type @'
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class ObserverFixtureWindows {
    public delegate bool EnumProc(IntPtr h,IntPtr l);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc p,IntPtr l);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
    [DllImport("user32.dll",CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr h,StringBuilder s,int n);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr h,uint m,IntPtr w,IntPtr l);
    public static IntPtr[] Dialogs(uint wanted){var found=new List<IntPtr>();EnumWindows((h,l)=>{
        uint p;GetWindowThreadProcessId(h,out p);var c=new StringBuilder(128);GetClassName(h,c,128);
        if(p==wanted && IsWindowVisible(h) && c.ToString()=="#32770")found.Add(h);return true;
    },IntPtr.Zero);return found.ToArray();}
    public static void CloseDialogs(uint wanted){foreach(var h in Dialogs(wanted)){
        uint p;GetWindowThreadProcessId(h,out p);if(p==wanted)PostMessage(h,0x10,IntPtr.Zero,IntPtr.Zero);
    }}
}
'@
function Start-Worker([string]$Name){
    $prefix=Join-Path $root $Name
    $job=Start-Job -ArgumentList $prefix -ScriptBlock {
        param($prefix)
        Add-Type -ReferencedAssemblies System.Windows.Forms,System.Drawing @'
using System; using System.IO; using System.Diagnostics; using System.Windows.Forms;
public static class ObserverFixtureHost {
    public static void Run(string prefix){
        using(var form=new Form())using(var timer=new Timer()){
            form.Text="invSys observer ordinary form";form.Width=340;form.Height=180;
            var button=new Button{Text="OK",Left=100,Top=60};form.Controls.Add(button);
            button.Click+=(s,e)=>File.WriteAllText(prefix+".clicked","clicked");
            form.Shown+=(s,e)=>File.WriteAllText(prefix+".ready",Process.GetCurrentProcess().Id.ToString());
            timer.Interval=100;timer.Tick+=(s,e)=>{
                if(File.Exists(prefix+".stop")){form.Close();return;}
                if(File.Exists(prefix+".show")){
                    File.Delete(prefix+".show");
                    MessageBox.Show(form,"Fixture notification retained exactly.","invSys observer modal",MessageBoxButtons.OK);
                }
                if(File.Exists(prefix+".confirm")){
                    File.Delete(prefix+".confirm");
                    File.WriteAllText(prefix+".confirmation-shown","shown");
                    var choice=MessageBox.Show(form,"Fixture confirmation requires an explicit choice.","invSys observer confirmation",MessageBoxButtons.OKCancel);
                    File.WriteAllText(prefix+".confirmation-result",choice.ToString());
                }
            };
            timer.Start();Application.Run(form);
        }
    }
}
'@
        [ObserverFixtureHost]::Run($prefix)
    }
    [pscustomobject]@{Job=$job;Prefix=$prefix;Process=0}
}
try{
    $workers=@((Start-Worker 'target'),(Start-Worker 'other'))
    foreach($worker in $workers){
        $deadline=[DateTime]::UtcNow.AddSeconds(20)
        while(-not (Test-Path ($worker.Prefix+'.ready')) -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
        if(-not (Test-Path ($worker.Prefix+'.ready'))){throw 'Observer fixture did not become ready.'}
        $worker.Process=[uint32](Get-Content ($worker.Prefix+'.ready'))
        [IO.File]::WriteAllText($worker.Prefix+'.show','')
        $deadline=[DateTime]::UtcNow.AddSeconds(10)
        while([ObserverFixtureWindows]::Dialogs($worker.Process).Count -ne 1 -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
        if([ObserverFixtureWindows]::Dialogs($worker.Process).Count -ne 1){throw 'Native modal fixture unavailable.'}
    }
    $observer=Start-DialogCaptureAndDismiss -ExcelProcessId $workers[0].Process -TimeoutSeconds 20 -StopPath $observerStop
    $deadline=[DateTime]::UtcNow.AddSeconds(12)
    while([ObserverFixtureWindows]::Dialogs($workers[0].Process).Count -ne 0 -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
    $notificationDismissed=([ObserverFixtureWindows]::Dialogs($workers[0].Process).Count -eq 0)
    if($notificationDismissed){
        [IO.File]::WriteAllText($workers[0].Prefix+'.confirm','')
        $deadline=[DateTime]::UtcNow.AddSeconds(10)
        while(-not (Test-Path ($workers[0].Prefix+'.confirmation-shown')) -and [DateTime]::UtcNow -lt $deadline){Start-Sleep -Milliseconds 100}
        if(-not (Test-Path ($workers[0].Prefix+'.confirmation-shown'))){throw 'Confirmation fixture unavailable.'}
    }
    Start-Sleep -Milliseconds 1000
    [IO.File]::WriteAllText($observerStop,'')
    Wait-Job $observer -Timeout 25 | Out-Null
    Check 'Observer.CooperativeCompletion' ($observer.State -eq 'Completed')
    $captured=@(Receive-Job $observer -ErrorAction SilentlyContinue) -join "`n"
    # Only fixed synthetic fixture captions/messages enter this diagnostic file.
    $captured | Set-Content -LiteralPath (Join-Path $reportRoot ($Phase.ToLowerInvariant()+'-captured.txt')) -Encoding UTF8
    Check 'Observer.OwnedModalDismissed' $notificationDismissed
    Check 'Observer.ExactNotificationCaptured' ($captured.Contains('Fixture notification retained exactly.'))
    Check 'Observer.ConfirmationNotAccepted' ($notificationDismissed -and -not (Test-Path ($workers[0].Prefix+'.confirmation-result')) -and [ObserverFixtureWindows]::Dialogs($workers[0].Process).Count -eq 1)
    Check 'Observer.OrdinaryFormCommandUntouched' (-not (Test-Path ($workers[0].Prefix+'.clicked')))
    Check 'Observer.OtherProcessModalUntouched' ([ObserverFixtureWindows]::Dialogs($workers[1].Process).Count -eq 1)
    Check 'Observer.OtherProcessCommandUntouched' (-not (Test-Path ($workers[1].Prefix+'.clicked')))
}catch{
    Check 'Observer.HarnessException' $false
    Write-Output $_.Exception.Message
}finally{
    [IO.File]::WriteAllText($observerStop,'')
    if($null -ne $observer){
        Wait-Job $observer -Timeout 25 | Out-Null
        if($observer.State -in @('Completed','Failed','Stopped')){Remove-Job $observer}
    }
    foreach($worker in $workers){
        if($worker.Process -gt 0){[ObserverFixtureWindows]::CloseDialogs($worker.Process)}
        [IO.File]::WriteAllText($worker.Prefix+'.stop','')
        Wait-Job $worker.Job -Timeout 10 | Out-Null
        if($worker.Job.State -notin @('Completed','Failed','Stopped')){throw 'Owned observer fixture has not exited.'}
        Remove-Job $worker.Job
    }
    $resolved=[IO.Path]::GetFullPath($root)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if(-not $resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -or (Split-Path $resolved -Leaf) -notlike 'invsys-dialog-observer-*'){throw 'Fixture cleanup root changed.'}
    Remove-Item -LiteralPath $resolved -Recurse -Force
    $results | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $reportRoot ($Phase.ToLowerInvariant()+'.json')) -Encoding UTF8
}
$failed=@($results|Where-Object {-not $_.Passed}).Count
Write-Output ($Phase+': PASS='+($results.Count-$failed)+' FAIL='+$failed)
if($failed){exit 1}
