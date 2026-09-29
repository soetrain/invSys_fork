// Diagnostic only: x64 exception metadata, no memory dump, payload or debug strings.
// Attach changes process timing; a passing observed run is not a cold-run repair.
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;

public static class NativeExceptionObserver
{
    [StructLayout(LayoutKind.Explicit, Size = 176)]
    private struct DebugEvent
    {
        [FieldOffset(0)] public uint Kind;
        [FieldOffset(4)] public uint ProcessId;
        [FieldOffset(8)] public uint ThreadId;
        [FieldOffset(16)] public uint ExceptionCode;
        [FieldOffset(16)] public IntPtr FileHandle;
        [FieldOffset(32)] public IntPtr ExceptionAddress;
        [FieldOffset(168)] public uint FirstChance;
    }
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool DebugActiveProcess(uint processId);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool DebugActiveProcessStop(uint processId);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool DebugSetProcessKillOnExit(bool kill);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool WaitForDebugEvent(out DebugEvent data, uint milliseconds);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool ContinueDebugEvent(uint processId, uint threadId, uint status);
    [DllImport("kernel32.dll")] private static extern bool CloseHandle(IntPtr handle);
    [DllImport("kernel32.dll")] private static extern void RaiseException(uint code, uint flags, uint count, IntPtr args);
    [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] private static extern void OutputDebugString(string value);

    private static readonly string[] AllowedModules = {
        "ntdll.dll", "kernelbase.dll", "kernel32.dll", "excel.exe", "vbe7.dll",
        "fm20.dll", "oleaut32.dll", "ole32.dll", "combase.dll", "rpcrt4.dll",
        "mso.dll", "mso20win32client.dll", "user32.dll", "win32u.dll",
        "ucrtbase.dll", "msvcrt.dll", "clr.dll", "mscoreei.dll", "shlwapi.dll"
    };

    private static void WriteState(StreamWriter log, string state, int error)
    {
        log.WriteLine("{\"UTC\":\"" + DateTime.UtcNow.ToString("o") +
            "\",\"State\":\"" + state + "\",\"Win32Error\":" + error + "}");
        log.Flush();
    }

    private static void WriteException(StreamWriter log, DebugEvent data)
    {
        string module = "unknown";
        string offset = "unavailable";
        try
        {
            using (Process target = Process.GetProcessById((int)data.ProcessId))
            {
                long address = data.ExceptionAddress.ToInt64();
                foreach (ProcessModule item in target.Modules)
                {
                    long start = item.BaseAddress.ToInt64();
                    if (address < start || address - start >= item.ModuleMemorySize) continue;
                    string name = item.ModuleName.ToLowerInvariant();
                    if (Array.IndexOf(AllowedModules, name) >= 0)
                    {
                        module = name;
                        offset = (address - start).ToString("X");
                    }
                    else module = "other";
                    break;
                }
            }
        }
        catch { } // No exception text or raw address is written.
        log.WriteLine("{\"UTC\":\"" + DateTime.UtcNow.ToString("o") +
            "\",\"Code\":\"" + data.ExceptionCode.ToString("X8") +
            "\",\"FirstChance\":" + (data.FirstChance != 0 ? "true" : "false") +
            ",\"Module\":\"" + module + "\",\"Offset\":\"" + offset + "\"}");
        log.Flush();
    }

    private static int Observe(uint processId, string root, int seconds, bool faultsOnly)
    {
        if (IntPtr.Size != 8 || seconds < 1 || seconds > 900) return 2;
        using (StreamWriter log = new StreamWriter(Path.Combine(root, "native-events.jsonl"), false))
        {
            if (!DebugActiveProcess(processId))
            {
                WriteState(log, "AttachFailed", Marshal.GetLastWin32Error());
                return 3;
            }
            bool exited = false;
            try
            {
                if (!DebugSetProcessKillOnExit(false))
                {
                    WriteState(log, "DisableKillFailed", Marshal.GetLastWin32Error());
                    return 4;
                }
                WriteState(log, "Attached", 0);
                WriteState(log, faultsOnly ? "FaultsOnly" : "AllExceptions", 0);
                bool initialBreak = true;
                DateTime deadline = DateTime.UtcNow.AddSeconds(seconds);
                while (DateTime.UtcNow < deadline && !File.Exists(Path.Combine(root, "stop")))
                {
                    DebugEvent data;
                    if (!WaitForDebugEvent(out data, 250))
                    {
                        int error = Marshal.GetLastWin32Error();
                        if (error == 121) continue;
                        WriteState(log, "WaitFailed", error);
                        return 5;
                    }
                    uint disposition = 0x00010002;
                    bool ready = false;
                    if (data.Kind == 1)
                    {
                        if (initialBreak && data.ExceptionCode == 0x80000003)
                        {
                            initialBreak = false;
                            ready = true;
                        }
                        else
                        {
                            // Reduce observation overhead while retaining every second-chance
                            // exception and all codes except these observed first-chance notices.
                            bool filtered = faultsOnly && data.FirstChance != 0 &&
                                (data.ExceptionCode == 0xE06D7363 || data.ExceptionCode == 0x40080201 || data.ExceptionCode == 5);
                            if (!filtered) WriteException(log, data);
                            disposition = 0x80010001; // Preserve ordinary exception handling.
                        }
                    }
                    if ((data.Kind == 3 || data.Kind == 6) && data.FileHandle != IntPtr.Zero)
                        CloseHandle(data.FileHandle);
                    if (data.Kind == 5) exited = true;
                    if (!ContinueDebugEvent(data.ProcessId, data.ThreadId, disposition))
                    {
                        WriteState(log, "ContinueFailed", Marshal.GetLastWin32Error());
                        return 6;
                    }
                    if (ready)
                    {
                        WriteState(log, "Ready", 0);
                        File.WriteAllText(Path.Combine(root, "ready"), "Ready");
                    }
                    if (exited) { WriteState(log, "Exited", 0); return 0; }
                }
                WriteState(log, "StopRequestedOrDeadline", 0);
                return 0;
            }
            finally
            {
                if (!exited)
                {
                    bool detached = DebugActiveProcessStop(processId);
                    WriteState(log, detached ? "Detached" : "DetachFailed", detached ? 0 : Marshal.GetLastWin32Error());
                }
            }
        }
    }

    // Disposable calibration target; no Office or project data is loaded.
    private static int Calibrate(string root)
    {
        DateTime deadline = DateTime.UtcNow.AddSeconds(30);
        while (!File.Exists(Path.Combine(root, "ready")))
        {
            if (DateTime.UtcNow > deadline) return 7;
            Thread.Sleep(25);
        }
        OutputDebugString("OBSERVER_REDACTION_SENTINEL_DO_NOT_RECORD");
        try { RaiseException(0xE042BEEF, 0, 0, IntPtr.Zero); }
        catch (SEHException) { File.WriteAllText(Path.Combine(root, "handled"), "Handled"); }
        try { RaiseException(5, 0, 0, IntPtr.Zero); }
        catch (SEHException) { File.WriteAllText(Path.Combine(root, "filtered-handled"), "Handled"); }
        File.WriteAllText(Path.Combine(root, "continued"), "Continued");
        if (File.Exists(Path.Combine(root, "test-detach")))
        {
            while (!File.Exists(Path.Combine(root, "target-exit")))
            {
                if (DateTime.UtcNow > deadline) return 8;
                Thread.Sleep(25);
            }
        }
        return 0;
    }

    public static int Main(string[] args)
    {
        try
        {
            if (args.Length == 2 && args[0] == "calibrate") return Calibrate(args[1]);
            if (args.Length != 3 && !(args.Length == 4 && args[3] == "faults-only")) return 9;
            return Observe(UInt32.Parse(args[0]), args[1], Int32.Parse(args[2]), args.Length == 4);
        }
        catch { return 10; } // Caller treats absent/incomplete receipts as diagnostic failure.
    }
}
