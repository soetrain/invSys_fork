# Native, observational resource counts only. No COM, captions, arguments or data.
if(-not ('InvSysGuideResources' -as [type])){
Add-Type @'
using System; using System.Text; using System.Runtime.InteropServices;
public static class InvSysGuideResources {
    public delegate bool EnumProc(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc callback, IntPtr parameter);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr window, out uint process);
    [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetClassName(IntPtr window, StringBuilder value, int length);
    [DllImport("user32.dll")] public static extern uint GetGuiResources(IntPtr process, uint flag);
    public static int ExcelWindowCount(uint owner) {
        int count=0;
        EnumWindows(delegate(IntPtr window, IntPtr ignored) {
            uint process; GetWindowThreadProcessId(window, out process);
            if(process==owner){var name=new StringBuilder(128); GetClassName(window,name,name.Capacity); if(name.ToString()=="XLMAIN")count++;}
            return true;
        },IntPtr.Zero);
        return count;
    }
}
'@
}
function Get-GuideResourceSample([string]$Boundary,[switch]$AllowExited) {
    $owners=@(Get-Process EXCEL -ErrorAction SilentlyContinue)
    if($AllowExited -and $owners.Count -eq 0){return}
    if($owners.Count -ne 1){throw 'Resource trace requires one isolated Excel process.'}
    $owner=$owners[0]
    try {
    [pscustomobject]@{
        UTC=[DateTimeOffset]::UtcNow.ToString('o');Boundary=$Boundary
        ProcessId=$owner.Id;StartUTC=$owner.StartTime.ToUniversalTime().ToString('o')
        Gdi=[InvSysGuideResources]::GetGuiResources($owner.Handle,0)
        User=[InvSysGuideResources]::GetGuiResources($owner.Handle,1)
        GdiPeak=[InvSysGuideResources]::GetGuiResources($owner.Handle,2)
        XLMAIN=[InvSysGuideResources]::ExcelWindowCount($owner.Id)
        PrivateBytes=$owner.PrivateMemorySize64
    }
    } catch {
        if($AllowExited -and $owner.HasExited){return}
        throw
    }
}
function Write-GuideResourceMark([string]$Boundary) {
    $sample=Get-GuideResourceSample $Boundary
    $sample|ConvertTo-Json -Compress|Add-Content (Join-Path $reportRoot 'guide-resources.jsonl')
    # A diagnostic stop, not a product acceptance threshold. Keep cleanup enabled.
    if($sample.Gdi -ge 3000 -or $sample.User -ge 3000 -or $sample.XLMAIN -ge 80){
        $script:TraceGuideResourcesForTest=$false
        throw 'Diagnostic resource bound reached; preserve trace and enter ordinary cleanup.'
    }
}
