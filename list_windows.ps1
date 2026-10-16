# Список видимых окон процесса: используется проверкой запуска в GitHub Actions.
param([int]$ProcessId)
Add-Type @"
using System; using System.Text; using System.Collections.Generic; using System.Runtime.InteropServices;
public static class W {
  public delegate bool P(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] static extern bool EnumWindows(P f, IntPtr l);
  [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll", CharSet=CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  public static List<string> Titles(uint target) {
    var r = new List<string>();
    EnumWindows((h, l) => { uint pid; GetWindowThreadProcessId(h, out pid);
      if (pid == target && IsWindowVisible(h)) { var sb = new StringBuilder(512); GetWindowText(h, sb, 512); if (sb.Length > 0) r.Add(sb.ToString()); }
      return true; }, IntPtr.Zero);
    return r;
  }
}
"@
[W]::Titles([uint32]$ProcessId)
