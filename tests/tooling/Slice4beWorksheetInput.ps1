# Developer-only input for isolated Excel fixtures. No operational workbook use.
Add-Type @'
using System;using System.Runtime.InteropServices;
public static class WorksheetInput {
 [StructLayout(LayoutKind.Sequential)] public struct Point { public int X,Y; }
 [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
 [DllImport("user32.dll")] public static extern bool GetWindowRect(IntPtr h,out Rect r);
 [DllImport("user32.dll")] public static extern uint GetDpiForWindow(IntPtr h);
 [StructLayout(LayoutKind.Sequential)] public struct Mouse { public int X,Y; public uint Data,Flags,Time; public UIntPtr Extra; }
 [StructLayout(LayoutKind.Sequential)] public struct Input { public uint Type; public Mouse Mouse; }
 [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
 [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")] static extern IntPtr GetAncestor(IntPtr h,uint flags);
 [DllImport("user32.dll")] static extern IntPtr WindowFromPoint(Point p);
 [DllImport("user32.dll")] static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")] static extern uint SendInput(uint n,Input[] inputs,int size);
 public static uint Owner(IntPtr h) {uint p;GetWindowThreadProcessId(h,out p);return p;}
 public static IntPtr Root(IntPtr h) {return GetAncestor(h,2);}
 public static bool Activate(IntPtr h) {h=Root(h);SetForegroundWindow(h);return h!=IntPtr.Zero && GetForegroundWindow()==h;}
 public static void Click(IntPtr h,int x,int y) {
  uint owner=Owner(h);
  if(owner==0 || GetForegroundWindow()!=Root(h) || Root(WindowFromPoint(new Point {X=x,Y=y}))!=Root(h))
   throw new Exception("Input target is not foreground owned Excel.");
  if(!SetCursorPos(x,y))throw new Exception("Cursor positioning unavailable.");
  Input[] inputs={new Input {Mouse=new Mouse {Flags=2}},new Input {Mouse=new Mouse {Flags=4}}};
  if(SendInput(2,inputs,Marshal.SizeOf(typeof(Input)))!=2)throw new Exception("Input delivery incomplete.");
 }
}
'@
function Get-WorksheetButtonPoint($Excel,$Button,$Workbook,$Sheet) {
    if($null -eq $Excel.ActiveWorkbook -or $Excel.ActiveWorkbook.FullName -cne $Workbook.FullName -or $Excel.ActiveSheet.Name -cne $Sheet.Name){throw 'Native fixture workbook and sheet are not active.'}
    if($Excel.ActiveWindow.ScrollRow -ne 1 -or $Excel.ActiveWindow.ScrollColumn -ne 1){throw 'Native fixture requires an unscrolled worksheet.'}
    $scale=[WorksheetInput]::GetDpiForWindow([IntPtr]$Workbook.Windows.Item(1).Hwnd)/72.0*$Excel.ActiveWindow.Zoom/100.0
    # Calibrated against the rendered fixture: use sheet origin plus scaled points.
    [pscustomobject]@{
        X=$Excel.ActiveWindow.PointsToScreenPixelsX(0)+[int][Math]::Round(($Button.Left+$Button.Width/2)*$scale)
        Y=$Excel.ActiveWindow.PointsToScreenPixelsY(0)+[int][Math]::Round(($Button.Top+$Button.Height/2)*$scale)
    }
}

function Save-WorksheetInputCapture($Excel,$Point,[string]$Path) {
    Add-Type -AssemblyName System.Drawing
    $bounds=New-Object WorksheetInput+Rect
    if(-not [WorksheetInput]::GetWindowRect([WorksheetInput]::Root([IntPtr]$Excel.ActiveWindow.Hwnd),[ref]$bounds)){throw 'Owned worksheet bounds unavailable.'}
    $capture=New-Object Drawing.Bitmap(($bounds.Right-$bounds.Left),($bounds.Bottom-$bounds.Top))
    $graphics=[Drawing.Graphics]::FromImage($capture)
    try {
        $graphics.CopyFromScreen($bounds.Left,$bounds.Top,0,0,$capture.Size)
        $graphics.DrawEllipse([Drawing.Pens]::Red,$Point.X-$bounds.Left-5,$Point.Y-$bounds.Top-5,10,10)
        $capture.Save($Path)
    } finally {$graphics.Dispose();$capture.Dispose()}
}
