# Calibrate real window ancestry before testing the native input helper. Window
# handles stay in memory; reports contain fixed check names and booleans only.
function Initialize-ReceivingNavigationTargetProbe {
    if(-not ('ReceivingNavigationTargetProbe' -as [type])) {
        Add-Type @'
using System; using System.Runtime.InteropServices;
public static class ReceivingNavigationTargetProbe {
    [StructLayout(LayoutKind.Sequential)] public struct Rect { public int Left,Top,Right,Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct Gui { public int Size,Flags; public IntPtr Active,Focus,Capture,Menu,Move,Caret; public Rect CaretRect; }
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h,out uint p);
    [DllImport("user32.dll")] static extern bool GetGUIThreadInfo(uint t,ref Gui g);
    [DllImport("user32.dll")] static extern bool IsChild(IntPtr parent,IntPtr child);
    public static bool WithinForm(IntPtr form) {
        uint owner; uint thread=GetWindowThreadProcessId(form,out owner);
        var state=new Gui(); state.Size=Marshal.SizeOf(state);
        return owner!=0 && GetGUIThreadInfo(thread,ref state) && state.Focus!=IntPtr.Zero &&
            (state.Focus==form || IsChild(form,state.Focus));
    }
}
'@
    }
}

function Test-ReceivingNavigationTarget($Fixture,$Other,[bool]$PreparedFocus) {
    $label='Navigation.NativeTarget'
    $form=[IntPtr][long][double](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationHandle')
    Check ($label+'.PreparedFocusWithinForm') $PreparedFocus
    if(-not $PreparedFocus){throw 'Prepared navigation focus calibration failed; not a product RED.'}
    $outside=-not [ReceivingNavigationTargetProbe]::WithinForm($form)
    Check ($label+'.UnrelatedFocusCalibrated') $outside
    if(-not $outside){throw 'Unrelated navigation focus calibration failed; not a product RED.'}
    $before=@(Get-Slice4beActivityFiles $Fixture)
    $otherHash=Get-ReceivingFixtureHash $Other.FullName
    $refused=$false
    try { [ReceivingNavigationInput]::Key($form,[IntPtr]$excel.Hwnd,37) }
    catch { $refused=$_.Exception.ToString().Contains('Owned focused navigation control unavailable.') }
    Check ($label+'.UnrelatedFocusRejected') $refused
    Check ($label+'.NoFormCallback') ([string](Run 'invSys.Operations.xlam' 'TestReceivingActivity.NavigationTrace') -ceq '')
    Check ($label+'.NoActivity') (@(Get-Slice4beActivityFiles $Fixture).Count -eq $before.Count)
    Check ($label+'.UnrelatedBytesPreserved') ($otherHash -ceq (Get-ReceivingFixtureHash $Other.FullName))
}
