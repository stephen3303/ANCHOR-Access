# Watches an automated Access instance for modal dialogs (VBA compile errors, runtime
# error boxes, MsgBox calls) and hangs, so a script driving Access over COM reports them
# instead of blocking forever behind a window nobody can see.
#
# Dot-source it, then wrap the work:
#
#   . "$PSScriptRoot\tools\AccessWatch.ps1"
#   $acc = New-Object -ComObject Access.Application
#   $watch = Start-AccessWatch $acc -TimeoutSec 600
#   try { ...$acc.Run(...)... }
#   finally { Stop-AccessWatch $watch $acc }      # prints any dialog, always ends the process
#
# When a dialog appears the watcher records its title and text, then ends that Access
# process, so the blocked COM call returns with an error and the script moves on.
# Use -ReportOnly to record dialogs and leave them open (for looking at them by hand).

if (-not ('KitAccessWatch' -as [type])) {
    Add-Type @'
using System; using System.Text; using System.Threading; using System.Diagnostics;
using System.Collections.Generic; using System.Runtime.InteropServices;
public class KitAccessWatch {
    delegate bool EnumProc(IntPtr h, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumWindows(EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern bool EnumChildWindows(IntPtr p, EnumProc cb, IntPtr l);
    [DllImport("user32.dll")] static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] static extern int GetClassName(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr h);

    public readonly uint Pid;
    public readonly List<string> Found = new List<string>();
    public bool ReportOnly, Killed, TimedOut;
    readonly int timeoutMs;
    readonly HashSet<IntPtr> seen = new HashSet<IntPtr>();
    volatile bool stop;
    Thread thread;

    public KitAccessWatch(uint pid, int timeoutSec, bool reportOnly) {
        Pid = pid; timeoutMs = timeoutSec * 1000; ReportOnly = reportOnly;
    }

    public static uint PidOf(IntPtr hwnd) { uint p; GetWindowThreadProcessId(hwnd, out p); return p; }

    static string Text(IntPtr h) { var s = new StringBuilder(4096); GetWindowText(h, s, 4096); return s.ToString(); }
    static string Cls(IntPtr h) { var s = new StringBuilder(256); GetClassName(h, s, 256); return s.ToString(); }

    // Every visible standard dialog (#32770) owned by the process: VBA's compile and
    // runtime error boxes, MsgBox, InputBox. Access's own Office-style alerts are NUIDialog.
    List<IntPtr> Dialogs() {
        var r = new List<IntPtr>();
        EnumWindows((h, l) => {
            uint p; GetWindowThreadProcessId(h, out p);
            if (p == Pid && IsWindowVisible(h)) { var c = Cls(h); if (c == "#32770" || c == "NUIDialog") r.Add(h); }
            return true;
        }, IntPtr.Zero);
        return r;
    }

    static string Describe(IntPtr h) {
        var parts = new List<string>();
        EnumChildWindows(h, (ch, l) => {
            var c = Cls(ch); var t = Text(ch).Trim();
            if (t.Length > 0 && (c == "Static" || c == "Edit" || c == "RichEdit20W")) parts.Add(t);
            return true;
        }, IntPtr.Zero);
        if (parts.Count == 0) {
            // Office-style alerts (NUIDialog) draw their text without child windows.
            try {
                var el = System.Windows.Automation.AutomationElement.FromHandle(h);
                var cond = new System.Windows.Automation.PropertyCondition(
                    System.Windows.Automation.AutomationElement.ControlTypeProperty, System.Windows.Automation.ControlType.Text);
                foreach (System.Windows.Automation.AutomationElement t in el.FindAll(System.Windows.Automation.TreeScope.Descendants, cond)) {
                    var n = t.Current.Name.Trim();
                    if (n.Length > 0 && n != Text(h)) parts.Add(n);
                }
            } catch { }
        }
        var body = string.Join(" | ", parts).Replace("\r\n", " ").Replace("\n", " ");
        while (body.Contains("  ")) body = body.Replace("  ", " ");
        return "[" + Text(h) + "] " + (body.Length > 0 ? body : "(no text)");
    }

    public void Start() {
        thread = new Thread(Loop); thread.IsBackground = true; thread.Start();
    }

    void Loop() {
        var sw = Stopwatch.StartNew();
        while (!stop) {
            Process proc = null;
            try { proc = Process.GetProcessById((int)Pid); } catch { return; }
            if (proc.HasExited) return;
            bool hit = false;
            foreach (var h in Dialogs()) {
                if (seen.Add(h)) { lock (Found) Found.Add(Describe(h)); }
                hit = true;
            }
            if (!hit && timeoutMs > 0 && sw.ElapsedMilliseconds > timeoutMs) {
                lock (Found) Found.Add("[timeout] Access still busy after " + (timeoutMs / 1000) + "s");
                TimedOut = true; hit = true;
            }
            if (hit && !ReportOnly) {
                // Give a second dialog that follows the first a moment to show, then end it.
                Thread.Sleep(300);
                foreach (var h in Dialogs()) if (seen.Add(h)) { lock (Found) Found.Add(Describe(h)); }
                try { proc.Kill(); Killed = true; } catch { }
                return;
            }
            Thread.Sleep(250);
        }
    }

    public string[] Stop() {
        stop = true;
        if (thread != null) thread.Join(2000);
        lock (Found) return Found.ToArray();
    }
}
'@ -ReferencedAssemblies UIAutomationClient, UIAutomationTypes
}

function Get-AccessPid($acc) {
    [KitAccessWatch]::PidOf([IntPtr][long]$acc.hWndAccessApp())
}

function Start-AccessWatch {
    param([Parameter(Mandatory)] $Acc, [int]$TimeoutSec = 0, [switch]$ReportOnly)
    $w = New-Object KitAccessWatch -ArgumentList (Get-AccessPid $Acc), $TimeoutSec, ([bool]$ReportOnly)
    $w.Start()
    $w
}

# Compiles every module. Under automation a failed compile shows only a generic
# "Cannot Compile Project." alert (which the watcher records) and leaves IsCompiled =
# False, without naming the module or line. On failure this re-runs the compile from
# the VBA editor, which shows the error and selects the offending line, reads both,
# then ends the process.
# Returns $null when the project compiles, else "module line N: <code> -- <error>".
function Test-AccessCompiles {
    param([Parameter(Mandatory)] $Acc, [Parameter(Mandatory)] $Watch)
    $Acc.DoCmd.RunCommand(126)   # acCmdCompileAndSaveAllModules
    if ($Acc.IsCompiled) { return $null }
    $Watch.ReportOnly = $true
    $n0 = $Watch.Found.Count
    $ctl = $Acc.VBE.CommandBars.FindControl(1, 578)   # Debug > Compile <project>
    $ps = [powershell]::Create().AddScript({ param($c) try { $c.Execute() } catch {} }).AddArgument($ctl)
    $null = $ps.BeginInvoke()
    for ($i = 0; $i -lt 40 -and $Watch.Found.Count -eq $n0; $i++) { Start-Sleep -Milliseconds 250 }
    $msg = if ($Watch.Found.Count -gt $n0) { $Watch.Found[$Watch.Found.Count - 1] } else { 'no error shown' }
    $where = ''
    try {
        $cp = $Acc.VBE.ActiveCodePane
        $sl = 0; $sc = 0; $el = 0; $ec = 0; $cp.GetSelection([ref]$sl, [ref]$sc, [ref]$el, [ref]$ec)
        $where = "$($cp.CodeModule.Name) line ${sl}: $($cp.CodeModule.Lines($sl, 1).Trim()) -- "
    } catch {}
    Stop-Process -Id $Watch.Pid -Force -ErrorAction SilentlyContinue
    $Watch.Killed = $true
    "$where$msg"
}

# Stops watching, quits Access, and makes sure the process is gone (a hung or killed
# instance otherwise lingers invisibly, holding the .accdb open). Writes each dialog
# found as a warning and returns them; returns nothing when Access stayed clean.
function Stop-AccessWatch {
    param([Parameter(Mandatory)] $Watch, $Acc, [switch]$KeepOpen)
    if ($Acc -and -not $Watch.Killed -and -not $KeepOpen) {
        try { $Acc.CloseCurrentDatabase() } catch {}
        try { $Acc.Quit() } catch {}
        try { [Runtime.InteropServices.Marshal]::ReleaseComObject($Acc) | Out-Null } catch {}
    }
    $found = $Watch.Stop()
    if (-not $KeepOpen) {
        $p = Get-Process -Id $Watch.Pid -ErrorAction SilentlyContinue
        if ($p) { if (-not $p.WaitForExit(5000)) { $p.Kill() } }
    }
    foreach ($f in $found) { Write-Warning "Access dialog: $f" }
    $found
}
