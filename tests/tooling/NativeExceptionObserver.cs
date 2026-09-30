// Diagnostic only: x64 exception metadata, no memory dump, payload or debug strings.
// Attach changes process timing; a passing observed run is not a cold-run repair.
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Collections.Generic;

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
    [DllImport("kernel32.dll", SetLastError = true)] private static extern IntPtr OpenThread(uint access, bool inherit, uint id);
    [DllImport("kernel32.dll", SetLastError = true)] private static extern bool GetThreadContext(IntPtr thread, IntPtr context);
    [DllImport("dbghelp.dll", CharSet = CharSet.Ansi, SetLastError = true)] private static extern bool SymInitialize(IntPtr process, string searchPath, bool invade);
    [DllImport("dbghelp.dll")] private static extern bool SymCleanup(IntPtr process);
    [DllImport("dbghelp.dll")] private static extern uint SymSetOptions(uint options);
    [DllImport("dbghelp.dll")] private static extern IntPtr SymFunctionTableAccess64(IntPtr process, ulong address);
    [DllImport("dbghelp.dll")] private static extern ulong SymGetModuleBase64(IntPtr process, ulong address);
    private delegate IntPtr FunctionTableAccess(IntPtr process, ulong address);
    private delegate ulong ModuleBaseAccess(IntPtr process, ulong address);
    [DllImport("dbghelp.dll")] private static extern bool StackWalk64(uint machine, IntPtr process, IntPtr thread,
        ref StackFrame frame, IntPtr context, IntPtr readMemory, FunctionTableAccess functionTable, ModuleBaseAccess moduleBase, IntPtr translate);
    [StructLayout(LayoutKind.Sequential)]
    private struct Address64 { public ulong Offset; public ushort Segment; public uint Mode; }
    // x64 STACKFRAME64 is 264 bytes, including the current 112-byte KDHELP64.
    // Unused return/argument/kernel fields remain internal and are never serialized.
    [StructLayout(LayoutKind.Explicit, Size = 264)]
    private struct StackFrame
    {
        [FieldOffset(0)] public Address64 PC;
        [FieldOffset(32)] public Address64 Frame;
        [FieldOffset(48)] public Address64 Stack;
    }

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

    private static string SafeFrame(Process target, ulong address, int index)
    {
        string module = "unknown", offset = "unavailable";
        foreach (ProcessModule item in target.Modules)
        {
            ulong start = (ulong)item.BaseAddress.ToInt64();
            if (address < start || address - start >= (ulong)item.ModuleMemorySize) continue;
            string name = item.ModuleName.ToLowerInvariant();
            module = Array.IndexOf(AllowedModules, name) >= 0 ? name : "other";
            if (module != "other") offset = (address - start).ToString("X");
            break;
        }
        return "{\"Index\":" + index + ",\"Module\":\"" + module + "\",\"Offset\":\"" + offset + "\"}";
    }

    private static void WriteFaultStack(string root, DebugEvent data)
    {
        // Only the observed fatal code, second-chance faults and the disposable
        // calibration exception. No routine first-chance stack surveillance.
        if (data.FirstChance != 0 && data.ExceptionCode != 0xC0000028 && data.ExceptionCode != 0xE042BEEF) return;
        List<string> frames = new List<string>();
        string state = "Unavailable";
        int error = 0;
        IntPtr allocation = IntPtr.Zero, thread = IntPtr.Zero;
        try
        {
            using (Process target = Process.GetProcessById((int)data.ProcessId))
            {
                thread = OpenThread(0x0008, false, data.ThreadId); // THREAD_GET_CONTEXT
                if (thread == IntPtr.Zero) { error = Marshal.GetLastWin32Error(); }
                else
                {
                    // CONTEXT_AMD64: 1232 bytes with 16-byte alignment. The debug
                    // event already suspends this thread; no context is written back.
                    allocation = Marshal.AllocHGlobal(1248);
                    IntPtr context = new IntPtr((allocation.ToInt64() + 15) & ~15L);
                    Marshal.Copy(new byte[1232], 0, context, 1232);
                    Marshal.WriteInt32(context, 48, 0x10000B); // CONTEXT_FULL
                    if (!GetThreadContext(thread, context)) error = Marshal.GetLastWin32Error();
                    else
                    {
                        // Deferred loading, ignore symbol-path environment, no prompts.
                        // No symbol server or symbol/parameter names are requested.
                        SymSetOptions(0x00000004 | 0x00000200 | 0x00001000 | 0x00080000);
                        if (!SymInitialize(target.Handle, "", true)) error = Marshal.GetLastWin32Error();
                        else
                        {
                            try
                            {
                                StackFrame frame = new StackFrame();
                                frame.PC.Offset = (ulong)Marshal.ReadInt64(context, 248);
                                frame.Frame.Offset = (ulong)Marshal.ReadInt64(context, 160);
                                frame.Stack.Offset = (ulong)Marshal.ReadInt64(context, 152);
                                frame.PC.Mode = frame.Frame.Mode = frame.Stack.Mode = 3;
                                ulong lastPC = 0, lastSP = 0;
                                for (int i = 0; i < 32; i++)
                                {
                                    if (!StackWalk64(0x8664, target.Handle, thread, ref frame, context, IntPtr.Zero,
                                        SymFunctionTableAccess64, SymGetModuleBase64, IntPtr.Zero)) break;
                                    if (frame.PC.Offset == 0 || (frame.PC.Offset == lastPC && frame.Stack.Offset == lastSP)) break;
                                    frames.Add(SafeFrame(target, frame.PC.Offset, frames.Count));
                                    lastPC = frame.PC.Offset; lastSP = frame.Stack.Offset;
                                }
                                if (frames.Count > 0) state = "Captured";
                            }
                            finally { SymCleanup(target.Handle); }
                        }
                    }
                }
            }
        }
        catch { state = "Unavailable"; frames.Clear(); } // No exception payload.
        finally
        {
            if (thread != IntPtr.Zero) CloseHandle(thread);
            if (allocation != IntPtr.Zero) Marshal.FreeHGlobal(allocation);
        }
        File.AppendAllText(Path.Combine(root, "native-stacks.jsonl"),
            "{\"UTC\":\"" + DateTime.UtcNow.ToString("o") + "\",\"Code\":\"" + data.ExceptionCode.ToString("X8") +
            "\",\"FirstChance\":" + (data.FirstChance != 0 ? "true" : "false") + ",\"State\":\"" + state +
            "\",\"Win32Error\":" + error + ",\"Frames\":[" + String.Join(",", frames.ToArray()) + "]}" + Environment.NewLine);
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
                            WriteFaultStack(root, data);
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
