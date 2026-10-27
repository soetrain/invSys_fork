# Observation only: fixed phase labels, counts and native process metadata.
# No COM attachment, object release, collection, termination or desktop rebinding.
function Write-ExcelShutdownTiming {
    param(
        [ValidateSet('Assertion','CleanupStart','BeforeEnumeration','BeforeWorkbookClose','AfterWorkbookClose',
            'PackagesCloseViaQuit','BeforeQuit','AfterQuit','BeforeComRelease','AfterComRelease',
            'BeforeGC','AfterGC','BeforeFinalizerWait','AfterFinalizerWait','WorkerExit')][string]$Stage,
        [ValidateSet('None','Workbook','XLAM')][string]$Kind='None',
        [ValidateRange(0,2147483647)][int]$Ordinal=0
    )
    $row=[pscustomobject]@{UTC=[DateTimeOffset]::UtcNow.ToString('o');Ticks=[Diagnostics.Stopwatch]::GetTimestamp();
        Frequency=[Diagnostics.Stopwatch]::Frequency;Stage=$Stage;Kind=$Kind;Ordinal=$Ordinal}
    [IO.File]::AppendAllText((Join-Path $script:ExcelShutdownTimingRoot 'phases.jsonl'),
        (($row|ConvertTo-Json -Compress)+[Environment]::NewLine),[Text.UTF8Encoding]::new($false))
}

function Initialize-ExcelExitObserver {
    if('InvSysExcelExitObserver' -as [type]){return}
    Add-Type -ReferencedAssemblies System.Web.Extensions @'
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Web.Script.Serialization;
public static class InvSysExcelExitObserver {
    [DllImport("user32.dll", SetLastError=true)] static extern uint GetWindowThreadProcessId(IntPtr window, out uint pid);
    [DllImport("kernel32.dll", SetLastError=true)] static extern IntPtr OpenProcess(uint access, bool inherit, uint pid);
    [DllImport("kernel32.dll", SetLastError=true)] static extern bool GetProcessTimes(IntPtr handle, out long created, out long exited, out long kernel, out long user);
    [DllImport("kernel32.dll", SetLastError=true)] static extern uint WaitForSingleObject(IntPtr handle, uint timeout);
    [DllImport("kernel32.dll")] static extern bool CloseHandle(IntPtr handle);
    static IntPtr Open(uint pid, out long created) {
        IntPtr handle=OpenProcess(0x00101000, false, pid);
        if(handle==IntPtr.Zero) throw new System.ComponentModel.Win32Exception(Marshal.GetLastWin32Error());
        long exited,kernel,user;
        if(!GetProcessTimes(handle,out created,out exited,out kernel,out user)) {
            int error=Marshal.GetLastWin32Error();CloseHandle(handle);
            throw new System.ComponentModel.Win32Exception(error);
        }
        return handle;
    }
    static void Save(string path, object value) {
        File.WriteAllText(path+".tmp",new JavaScriptSerializer().Serialize(value));
        File.Move(path+".tmp",path);
    }
    public static void IdentifyWindow(long window, string root) {
        uint pid;
        if(GetWindowThreadProcessId(new IntPtr(window),out pid)==0 || pid==0)
            throw new InvalidOperationException("Owned Excel window could not be identified.");
        long created;IntPtr handle=Open(pid,out created);
        try {
            Save(Path.Combine(root,"owner.json"),new {ProcessId=pid,CreatedFileTime=created,
                CreatedUTC=DateTime.FromFileTimeUtc(created).ToString("o"),WorkerId=Process.GetCurrentProcess().Id});
        } finally {CloseHandle(handle);}
    }
    public static void Watch(uint pid,long expectedCreated,string path) {
        long created;IntPtr handle=Open(pid,out created);
        if(created!=expectedCreated) {CloseHandle(handle);throw new InvalidOperationException("Process identity changed.");}
        // The controller holds this exact native handle throughout the workload.
        // A worker GC/finalizer wait cannot prevent the controller observing exit.
        Thread watcher=new Thread(delegate() {
            try {
                uint wait=WaitForSingleObject(handle,0xffffffff);
                int waitError=wait==0xffffffff?Marshal.GetLastWin32Error():0;
                long c,e,k,u;bool times=GetProcessTimes(handle,out c,out e,out k,out u);
                int timesError=times?0:Marshal.GetLastWin32Error();
                Save(path,new {ProcessId=pid,CreatedFileTime=created,WaitResult=wait,
                    WaitError=waitError,TimesAvailable=times,TimesError=timesError,
                    NativeExitUTC=times && e>0?DateTime.FromFileTimeUtc(e).ToString("o"):null,
                    ObservedUTC=DateTime.UtcNow.ToString("o"),ObservedTicks=Stopwatch.GetTimestamp()});
            } catch { // Fixed failure label; never serialize exception text.
                try {Save(path,new {ObserverFailed=true});} catch {}
            } finally {CloseHandle(handle);}
        });
        watcher.IsBackground=true;watcher.Start();
    }
}
'@
}

function ConvertTo-ExcelShutdownTimedWorker {
    param([string]$Source,[string]$OriginalToolingRoot,[string]$TimingModule,[string]$TimingRoot)
    $tokens=$null;$errors=$null
    $ast=[Management.Automation.Language.Parser]::ParseInput($Source,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw 'Worker source does not parse.'}
    $main=@($ast.EndBlock.Statements|Where-Object {$_ -is [Management.Automation.Language.TryStatementAst] -and
        $null -ne $_.Finally -and $_.Finally.Extent.Text.Contains('[GC]::WaitForPendingFinalizers()')})
    if($main.Count -ne 1){throw 'Unique worker shutdown boundary required.'}
    $original=$main[0].Finally.Extent.Text;$body=$original
    $edits=@(
        @('    if($null -ne $excel) {',("    Write-ExcelShutdownTiming CleanupStart`r`n"+'    if($null -ne $excel) {')),
        @('            foreach($book in @($excel.Workbooks)){',("            Write-ExcelShutdownTiming BeforeEnumeration`r`n"+'            foreach($book in @($excel.Workbooks)){')),
        @('                $book.Close($false)',@'
                $timingKind=if([bool]$book.IsAddin){'XLAM'}else{'Workbook'}
                Write-ExcelShutdownTiming BeforeWorkbookClose $timingKind
                $book.Close($false)
                Write-ExcelShutdownTiming AfterWorkbookClose $timingKind
'@),
        @('            $excel.Quit()',@'
            Write-ExcelShutdownTiming PackagesCloseViaQuit XLAM
            Write-ExcelShutdownTiming BeforeQuit
            $excel.Quit()
            Write-ExcelShutdownTiming AfterQuit
'@),
        @('        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)',@'
        Write-ExcelShutdownTiming BeforeComRelease
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
        Write-ExcelShutdownTiming AfterComRelease
'@),
        @('    [GC]::Collect(); [GC]::WaitForPendingFinalizers()',@'
    Write-ExcelShutdownTiming BeforeGC
    [GC]::Collect()
    Write-ExcelShutdownTiming AfterGC
    Write-ExcelShutdownTiming BeforeFinalizerWait
    [GC]::WaitForPendingFinalizers()
    Write-ExcelShutdownTiming AfterFinalizerWait
'@)
    )
    foreach($edit in $edits){
        if([regex]::Matches($body,[regex]::Escape($edit[0])).Count -ne 1){throw 'Shutdown anchor drift; no workload dispatched.'}
        $body=$body.Replace($edit[0],$edit[1])
    }
    $Source=$Source.Replace($original,$body)
    foreach($edit in @(
        @('    $results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})',("    Write-ExcelShutdownTiming Assertion None (`$results.Count+1)`r`n"+'    $results.Add([pscustomobject]@{Check=$Name;Passed=$Passed})')),
        @('    $initialExcelWindow=[long]$excel.Hwnd',@'
    $initialExcelWindow=[long]$excel.Hwnd
    Initialize-ExcelExitObserver
    [InvSysExcelExitObserver]::IdentifyWindow($initialExcelWindow,$script:ExcelShutdownTimingRoot)
'@)
    )){
        if([regex]::Matches($Source,[regex]::Escape($edit[0])).Count -ne 1){throw 'Worker observation anchor drift.'}
        $Source=$Source.Replace($edit[0],$edit[1])
    }
    # Keep every existing helper lookup at its original source directory.
    $Source=$Source.Replace('$PSScriptRoot',("'"+$OriginalToolingRoot.Replace("'","''")+"'"))
    $ast=[Management.Automation.Language.Parser]::ParseInput($Source,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw ('Timed worker does not parse: '+$errors[0].ErrorId+' at line '+$errors[0].Extent.StartLineNumber)}
    $setup="`r`n"+'. '''+$TimingModule.Replace("'","''")+"'`r`n"+
        '$script:ExcelShutdownTimingRoot = '''+$TimingRoot.Replace("'","''")+"'`r`n"
    $Source=$Source.Insert($ast.EndBlock.Statements[0].Extent.StartOffset,$setup)
    [void][Management.Automation.Language.Parser]::ParseInput($Source,[ref]$tokens,[ref]$errors)
    if($errors.Count){throw 'Timed worker setup does not parse.'}
    return $Source
}
