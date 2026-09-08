# Developer validation only. Observe native modal dialogs for one owned process;
# never inspect ordinary form controls or accept a multi-button confirmation.
function Invoke-Plan022NativeDialogObservation {
    param([uint32]$ProcessId, [int]$TimeoutSeconds, [string]$StopPath)
    Add-Type @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;
public static class Plan022NativeDialogs {
    public delegate bool EnumProc(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc callback, IntPtr parameter);
    [DllImport("user32.dll")] public static extern bool EnumChildWindows(IntPtr parent, EnumProc callback, IntPtr parameter);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] public static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")] public static extern IntPtr GetParent(IntPtr window);
    [DllImport("user32.dll")] public static extern int GetDlgCtrlID(IntPtr window);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetClassName(IntPtr window, StringBuilder text, int capacity);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern int GetWindowText(IntPtr window, StringBuilder text, int capacity);
    [DllImport("user32.dll")] public static extern bool PostMessage(IntPtr window, uint message, IntPtr wParam, IntPtr lParam);
    private static readonly HashSet<string> seen = new HashSet<string>();
    private static readonly HashSet<IntPtr> dismissed = new HashSet<IntPtr>();
    private static string Class(IntPtr window){var text=new StringBuilder(128);GetClassName(window,text,text.Capacity);return text.ToString();}
    private static string Text(IntPtr window){var text=new StringBuilder(4096);GetWindowText(window,text,text.Capacity);return text.ToString();}
    private static bool Owned(IntPtr window,uint wanted){uint actual;GetWindowThreadProcessId(window,out actual);return actual==wanted;}
    private static void Emit(List<string> output,string text){if(seen.Add(text))output.Add(text);}
    public static string[] Poll(uint wanted){
        var output=new List<string>();var visible=new HashSet<IntPtr>();
        EnumWindows((window,unused)=>{
            if(!Owned(window,wanted)||!IsWindowVisible(window)||Class(window)!="#32770")return true;
            visible.Add(window);var title=Text(window);var buttons=new List<IntPtr>();
            Emit(output,"WINDOW|"+title+"|#32770");
            EnumChildWindows(window,(child,ignored)=>{
                if(!Owned(child,wanted)||!IsWindowVisible(child))return true;
                var kind=Class(child);var text=Text(child);
                if(kind=="Static"&&text.Length>0)Emit(output,"WINDOW_ELEMENT|"+title+"|ControlType.Text|"+text);
                if(kind=="Button"){
                    buttons.Add(child);
                    Emit(output,"WINDOW_ELEMENT|"+title+"|ControlType.Button|"+text);
                }
                return true;
            },IntPtr.Zero);
            // An informational OK is the only permitted automatic dismissal.
            // Recheck owner, parent and identity immediately before posting.
            if(buttons.Count==1 && !dismissed.Contains(window)){
                var button=buttons[0];
                var id=GetDlgCtrlID(button);
                // Windows may use IDCANCEL for the sole OK (Escape-dismissable).
                if((id==1 || id==2) && Text(button).Replace("&","")=="OK" &&
                   IsWindowEnabled(button) && Owned(window,wanted) && Owned(button,wanted) &&
                   GetParent(button)==window && Class(window)=="#32770"){
                    if(PostMessage(window,0x111,(IntPtr)id,button))dismissed.Add(window);
                }
            }
            return true;
        },IntPtr.Zero);
        dismissed.RemoveWhere(window=>!visible.Contains(window));
        return output.ToArray();
    }
}
'@
    $stopAt=[DateTime]::UtcNow.AddSeconds($TimeoutSeconds)
    while([DateTime]::UtcNow -lt $stopAt){
        if($StopPath -ne '' -and [IO.File]::Exists($StopPath)){break}
        [Plan022NativeDialogs]::Poll($ProcessId)
        Start-Sleep -Milliseconds 200
    }
}
