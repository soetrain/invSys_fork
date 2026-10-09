# Standalone, two-phase environment diagnostic; not invSys acceptance evidence.
# Only bounded cursor movement: no clicks, keys, Excel, desktop rebinding or keep-awake.
[CmdletBinding()]
param([ValidateRange(1,120)][int]$WaitMinutes=120)
$ErrorActionPreference='Stop'
$repo=Split-Path -Parent $PSScriptRoot
$root=Join-Path $repo ('reports/runtime/cursor-handoff/'+[guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $root|Out-Null
Add-Type -ReferencedAssemblies System.Web.Extensions @'
using System;using System.IO;using System.Text;using System.Diagnostics;
using System.Threading;using System.Runtime.InteropServices;using System.Collections.Generic;
using System.Web.Script.Serialization;
public static class CursorHandoffDiagnostic {
 [StructLayout(LayoutKind.Sequential)]public struct Point{public int X,Y;}
 [StructLayout(LayoutKind.Sequential)]public struct Rect{public int Left,Top,Right,Bottom;}
 [StructLayout(LayoutKind.Sequential)]public struct MonitorInfo{public int Size;public Rect Bounds,Work;public uint Flags;}
 public class ReadResult{public bool ReturnValue;public int? FailureWin32Error;public Point? Coordinates;}
 public class WriteResult{public bool ReturnValue;public int? FailureWin32Error;}
 public delegate bool MonitorCallback(IntPtr monitor,IntPtr dc,ref Rect rect,IntPtr data);
 [DllImport("kernel32.dll")]static extern uint GetCurrentThreadId();
 [DllImport("kernel32.dll")]static extern void SetLastError(uint error);
 [DllImport("user32.dll",SetLastError=true)]static extern bool GetCursorPos(out Point point);
 [DllImport("user32.dll",SetLastError=true)]static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll",SetLastError=true)]static extern IntPtr GetThreadDesktop(uint thread);
 [DllImport("user32.dll",SetLastError=true)]static extern IntPtr OpenInputDesktop(uint flags,bool inherit,uint access);
 [DllImport("user32.dll",SetLastError=true)]static extern bool CloseDesktop(IntPtr desktop);
 [DllImport("user32.dll",EntryPoint="GetUserObjectInformationW",CharSet=CharSet.Unicode,SetLastError=true)]
 static extern bool ObjectName(IntPtr handle,int index,StringBuilder value,uint size,out uint needed);
 [DllImport("user32.dll",EntryPoint="GetUserObjectInformationW",SetLastError=true)]
 static extern bool ObjectInput(IntPtr handle,int index,out int value,uint size,out uint needed);
 [DllImport("user32.dll")]static extern IntPtr GetProcessWindowStation();
 [DllImport("user32.dll")]static extern IntPtr GetThreadDpiAwarenessContext();
 [DllImport("user32.dll",SetLastError=true)]static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
 [DllImport("user32.dll")]static extern bool AreDpiAwarenessContextsEqual(IntPtr first,IntPtr second);
 [DllImport("user32.dll")]static extern int GetSystemMetrics(int index);
 [DllImport("user32.dll")]static extern short GetAsyncKeyState(int key);
 [DllImport("user32.dll",SetLastError=true)]static extern bool EnumDisplayMonitors(IntPtr dc,IntPtr clip,MonitorCallback callback,IntPtr data);
 [DllImport("user32.dll",SetLastError=true)]static extern bool GetMonitorInfo(IntPtr monitor,ref MonitorInfo info);
 static bool error5;
 static int? Failure(bool ok,int error){if(!ok && error==5)error5=true;return ok?(int?)null:error;}
 static Dictionary<string,object> Name(IntPtr handle){
  var value=new StringBuilder(512);uint needed;SetLastError(0);
  bool ok=ObjectName(handle,2,value,1024,out needed);int error=Marshal.GetLastWin32Error();
  return new Dictionary<string,object>{{"ReturnValue",ok},{"FailureWin32Error",Failure(ok,error)},{"Name",ok?value.ToString():null}};
 }
 static Dictionary<string,object> Desktop(){
  uint thread=GetCurrentThreadId();SetLastError(0);IntPtr current=GetThreadDesktop(thread);int currentError=Marshal.GetLastWin32Error();
  SetLastError(0);IntPtr active=OpenInputDesktop(0,false,1);int inputError=Marshal.GetLastWin32Error();
  var report=new Dictionary<string,object>{{"NativeThreadId",thread},{"ThreadDesktopError",Failure(current!=IntPtr.Zero,currentError)},
   {"OpenInputDesktopError",Failure(active!=IntPtr.Zero,inputError)},{"WindowStation",Name(GetProcessWindowStation())}};
  try {
   report["ThreadDesktop"]=current==IntPtr.Zero?null:Name(current);
   report["ActiveInputDesktop"]=active==IntPtr.Zero?null:Name(active);
   int input=0;uint needed=0;SetLastError(0);
   bool ok=current!=IntPtr.Zero && ObjectInput(current,6,out input,4,out needed);int error=Marshal.GetLastWin32Error();
   report["ThreadDesktopUoiIoReturn"]=ok;report["ThreadDesktopUoiIoError"]=Failure(ok,error);
   report["ThreadDesktopIsActiveInputDesktop"]=ok?(object)(input!=0):null;
  }finally{if(active!=IntPtr.Zero){SetLastError(0);bool closed=CloseDesktop(active);int error=Marshal.GetLastWin32Error();report["CloseInputDesktopError"]=Failure(closed,error);}}
  return report;
 }
 static List<MonitorInfo> Monitors(){
  var result=new List<MonitorInfo>();
  MonitorCallback callback=delegate(IntPtr h,IntPtr dc,ref Rect r,IntPtr state){
   var info=new MonitorInfo();info.Size=Marshal.SizeOf(typeof(MonitorInfo));SetLastError(0);
   bool ok=GetMonitorInfo(h,ref info);int error=Marshal.GetLastWin32Error();Failure(ok,error);
   if(!ok)throw new InvalidOperationException("Monitor bounds unavailable.");result.Add(info);return true;
  };
  SetLastError(0);bool enumerated=EnumDisplayMonitors(IntPtr.Zero,IntPtr.Zero,callback,IntPtr.Zero);int enumerationError=Marshal.GetLastWin32Error();Failure(enumerated,enumerationError);
  if(!enumerated || result.Count==0)throw new InvalidOperationException("Display enumeration unavailable.");return result;
 }
 static ReadResult Read(){
  Point point;SetLastError(0);bool ok=GetCursorPos(out point);int error=Marshal.GetLastWin32Error();
  return new ReadResult{ReturnValue=ok,FailureWin32Error=Failure(ok,error),Coordinates=ok?(Point?)point:null};
 }
 static WriteResult Move(Point point){
  SetLastError(0);bool ok=SetCursorPos(point.X,point.Y);int error=Marshal.GetLastWin32Error();
  return new WriteResult{ReturnValue=ok,FailureWin32Error=Failure(ok,error)};
 }
 static bool Same(ReadResult result,Point point){return result.ReturnValue && result.Coordinates.Value.X==point.X && result.Coordinates.Value.Y==point.Y;}
 static Dictionary<string,object> Measure(string phase){
  error5=false;var process=Process.GetCurrentProcess();var report=new Dictionary<string,object>{
   {"Phase",phase},{"Utc",DateTimeOffset.UtcNow.ToString("o")},{"ProcessId",process.Id},{"SessionId",process.SessionId},{"NativeThreadId",GetCurrentThreadId()},
   {"Library","PowerShell Add-Type / .NET P/Invoke; no third-party mouse library"},{"MouseApis","user32.dll!SetCursorPos, user32.dll!GetCursorPos"},
   {"ReferenceHelper","InvSysSettingsCapture in tests/tooling/Test-Slice4beConfigCommands.ps1"},{"CoordinateSpace","Physical pixels; temporary per-monitor-v2 DPI scope as in capture helper"}};
  IntPtr originalDpi=GetThreadDpiAwarenessContext(),previous=IntPtr.Zero;ReadResult before=null;bool attempted=false;
  try {
   report["DesktopBefore"]=Desktop();if(error5){report["Diagnosis"]="Desktop API error5; stopped before movement";return report;}
   SetLastError(0);previous=SetThreadDpiAwarenessContext(new IntPtr(-4));int dpiError=Marshal.GetLastWin32Error();
   report["DpiScopeError"]=Failure(previous!=IntPtr.Zero,dpiError);
   if(previous==IntPtr.Zero)throw new InvalidOperationException("Physical coordinate scope unavailable.");
   var monitors=Monitors();report["DisplaysBefore"]=monitors;
   report["VirtualDisplayBefore"]=new Rect{Left=GetSystemMetrics(76),Top=GetSystemMetrics(77),Right=GetSystemMetrics(76)+GetSystemMetrics(78),Bottom=GetSystemMetrics(77)+GetSystemMetrics(79)};
   report["RemoteSessionMetric"]=GetSystemMetrics(4096);
   before=Read();report["CursorBefore"]=before;
   if(!before.ReturnValue || error5){report["Diagnosis"]="GetCursorPos API failure; no movement attempted";return report;}
   foreach(int button in new int[]{1,2,4,5,6})if((GetAsyncKeyState(button)&0x8000)!=0){report["Diagnosis"]="Mouse button held; movement skipped to prevent dragging";return report;}
   Point target=before.Coordinates.Value;bool found=false;
   foreach(var monitor in monitors){var b=monitor.Bounds;
    if(target.X>=b.Left && target.X<b.Right && target.Y>=b.Top && target.Y<b.Bottom){
     if(target.X+8<b.Right){target.X+=8;found=true;}else if(target.X-8>=b.Left){target.X-=8;found=true;}break;
    }
   }
   if(!found){report["Diagnosis"]="No valid small movement within current display; skipped";return report;}
   report["Target"]=target;attempted=true;var moved=Move(target);report["SetCursorPosMove"]=moved;
   if(error5){report["Diagnosis"]="SetCursorPos API failure: error5";return report;}
   var after=Read();report["CursorAfterImmediate"]=after;
   if(error5){report["Diagnosis"]="GetCursorPos API failure after movement: error5";return report;}
   Thread.Sleep(50);var settled=Read();report["CursorAfter50ms"]=settled;
   report["Diagnosis"]=!moved.ReturnValue?"SetCursorPos API failure":
    (!after.ReturnValue || !settled.ReturnValue)?"SetCursorPos succeeded; coordinate read failed":
    Same(after,target)&&Same(settled,target)?"SetCursorPos succeeded; exact target reached":
    Same(after,before.Coordinates.Value)&&Same(settled,before.Coordinates.Value)?"SetCursorPos succeeded; coordinates unchanged":
    "SetCursorPos succeeded; unexpected coordinates";
  }catch(Exception ex){report["DiagnosticException"]=ex.GetType().Name;}
  finally {
   if(attempted && before!=null && before.ReturnValue && !error5){
    report["SetCursorPosRestore"]=Move(before.Coordinates.Value);
    if(!error5){var restored=Read();report["CursorAfterRestore"]=restored;report["OriginalPositionRestored"]=Same(restored,before.Coordinates.Value);}
   }else report["RestoreDisposition"]=error5?"Skipped after error5 stop":"No movement attempted";
   if(!error5){report["DesktopAfter"]=Desktop();if(!error5)report["DisplaysAfter"]=Monitors();}
   if(previous!=IntPtr.Zero){SetLastError(0);IntPtr restored=SetThreadDpiAwarenessContext(previous);int error=Marshal.GetLastWin32Error();report["DpiRestoreError"]=Failure(restored!=IntPtr.Zero,error);}
   report["DpiContextRestored"]=AreDpiAwarenessContextsEqual(originalDpi,GetThreadDpiAwarenessContext());
   report["Error5Observed"]=error5;report["EndUtc"]=DateTimeOffset.UtcNow.ToString("o");
  }
  return report;
 }
 static void Save(string path,object value){var json=new JavaScriptSerializer().Serialize(value);File.WriteAllText(path+".tmp",json);File.Move(path+".tmp",path);}
 public static void Run(string root,int waitMinutes){
  Save(Path.Combine(root,"identity.json"),new {ProcessId=Process.GetCurrentProcess().Id,SessionId=Process.GetCurrentProcess().SessionId,NativeThreadId=GetCurrentThreadId(),DeadlineUtc=DateTimeOffset.UtcNow.AddMinutes(waitMinutes).ToString("o")});
  Save(Path.Combine(root,"before.json"),Measure("BeforeDisconnect"));
  if(error5){File.WriteAllText(Path.Combine(root,"stopped.txt"),"error5 during baseline");return;}
  Console.WriteLine("BEFORE_REPORT="+Path.Combine(root,"before.json"));
  Console.WriteLine("Waiting passively on the same native thread for after.request; no cursor polling or keep-awake.");
  DateTime deadline=DateTime.UtcNow.AddMinutes(waitMinutes);
  while(DateTime.UtcNow<deadline){
   if(File.Exists(Path.Combine(root,"stop.request"))){File.WriteAllText(Path.Combine(root,"stopped.txt"),"stop requested");return;}
   if(File.Exists(Path.Combine(root,"after.request"))){Save(Path.Combine(root,"after.json"),Measure("AfterDisconnect"));File.WriteAllText(Path.Combine(root,"stopped.txt"),error5?"error5 after disconnect":"comparison complete");return;}
   Thread.Sleep(500);
  }
  File.WriteAllText(Path.Combine(root,"stopped.txt"),"passive wait expired");
 }
}
'@
$pointer=Join-Path $repo 'reports/runtime/tscon-cursor-comparison-current.json'
if(Test-Path -LiteralPath $pointer){throw 'Comparison pointer exists; inspect the existing helper before starting another.'}
[pscustomobject]@{Root=$root;ProcessId=$PID;WaitMinutes=$WaitMinutes}|ConvertTo-Json|Set-Content -LiteralPath $pointer
[CursorHandoffDiagnostic]::Run($root,$WaitMinutes)
