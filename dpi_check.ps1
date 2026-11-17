# Режим DPI процесса: 0 = не знает о DPI (Windows растягивает), 1 = по системе, 2 = по монитору.
param([int]$ProcessId)
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class D {
  [DllImport("kernel32.dll")] public static extern IntPtr OpenProcess(uint a, bool i, int pid);
  [DllImport("kernel32.dll")] public static extern bool CloseHandle(IntPtr h);
  [DllImport("shcore.dll")] public static extern int GetProcessDpiAwareness(IntPtr h, out int v);
  public static int Get(int pid) { IntPtr h = OpenProcess(0x1000, false, pid); int v = -1; GetProcessDpiAwareness(h, out v); CloseHandle(h); return v; }
}
"@
[D]::Get($ProcessId)
