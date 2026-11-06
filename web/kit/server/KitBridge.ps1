<#
Hidden-Access bridge (spike). Instead of a PowerShell copy of a VBA routine,
a route can call the routine itself: the server keeps one hidden Access
instance open on the app's "logic" file (the app's VBA and linked tables, no
start-up) and calls its entry points over COM with Application.Run.

Contract for an entry point: a Public Function that takes plain values and
returns one JSON string {"ok","errNumber","source","message","value"}, traps
every error, and never shows a dialog (see apps\anchor\bridge\modWebBridge.bas).

What the server does when that contract is broken: every call runs under a
watcher (AccessWatch.ps1). If Access shows any dialog (a VBA error box, a
MsgBox, an Access prompt) or is still busy after TimeoutSec, the watcher
records the dialog's text and ends the Access process; the blocked COM call
returns, the route answers with {ok:false, kind:'dialog'|'timeout'}, and a new
instance is started for the next call.

Settings, per app in web.config.json:
  "bridge": { "logic": "apps\\anchor\\bridge\\ANCHOR_logic.accdb", "runtime": false, "timeoutSec": 10, "tests": false }
  runtime  open the logic file with MSACCESS.EXE /runtime (as the free Access
           Runtime would) instead of a full Access automation instance
  tests    adds the spike's fault-injection routes (never in real use)
#>

. (Join-Path $PSScriptRoot 'AccessWatch.ps1')

function Use-KitBridge([string]$LogicFile, [bool]$Runtime = $false, [int]$TimeoutSec = 10, [bool]$Tests = $false) {
    $Kit.Current.Bridge = @{
        LogicFile = $LogicFile; Runtime = $Runtime; TimeoutSec = $TimeoutSec; Tests = $Tests
        Acc = $null; Pid = 0; StartMs = 0; FirstStartMs = 0; Starts = 0; Restarts = 0; Calls = 0; LastProblem = ''
        PidFile = "$LogicFile.pid"
    }
}

function Get-KitAccessExe {
    $p = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\MSACCESS.EXE' -ErrorAction SilentlyContinue).'(default)'
    if (-not $p) { throw 'MSACCESS.EXE is not registered (no Access or Access Runtime on this PC).' }
    return $p.Trim('"')
}

function Start-KitBridge($app) {
    $b = $app.Bridge
    if (-not (Test-Path $b.LogicFile)) { throw "The logic file is missing: $($b.LogicFile)" }
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $acc = $null; $procId = 0
    if ($b.Runtime) {
        $proc = Start-Process (Get-KitAccessExe) -ArgumentList @('/runtime', "`"$($b.LogicFile)`"") -PassThru -WindowStyle Hidden
        $procId = $proc.Id
        $w = New-Object KitAccessWatch -ArgumentList ([uint32]$procId), 60, $false
        $w.Start()
        try {
            # The opened file registers itself; GetObject(path) returns its Application.
            for ($i = 0; $i -lt 240 -and -not $acc; $i++) {
                if ($proc.HasExited) { break }
                try { $acc = [Runtime.InteropServices.Marshal]::BindToMoniker($b.LogicFile) } catch { Start-Sleep -Milliseconds 250 }
            }
        } finally { $found = $w.Stop() }
        if (-not $acc) { throw ("Access /runtime did not open the logic file. " + ($found -join '; ')) }
        try { $acc.Visible = $false } catch { }
    } else {
        $acc = New-Object -ComObject Access.Application
        $procId = [int](Get-AccessPid $acc)
        $w = New-Object KitAccessWatch -ArgumentList ([uint32]$procId), 60, $false
        $w.Start()
        try {
            $acc.Visible = $false
            $acc.AutomationSecurity = 1     # run the logic file's code without a trust prompt
            $acc.OpenCurrentDatabase($b.LogicFile)
        } finally { $found = $w.Stop() }
        if ($w.Killed) { throw ("Access could not open the logic file: " + ($found -join '; ')) }
    }
    $b.Acc = $acc; $b.Pid = $procId; $b.Starts++
    Set-Content -Path $b.PidFile -Value $procId -ErrorAction SilentlyContinue
    $b.StartMs = [int]$sw.ElapsedMilliseconds
    if (-not $b.FirstStartMs) { $b.FirstStartMs = $b.StartMs }
    Write-Host ("{0} bridge: Access {1} pid {2} ready in {3} ms" -f $app.Key, $(if ($b.Runtime) { '/runtime' } else { 'automation' }), $procId, $b.StartMs)
}

function Stop-KitBridge($app) {
    $b = $app.Bridge
    if (-not $b) { return }
    if ($b.Acc) {
        try { $b.Acc.CloseCurrentDatabase() } catch { }
        try { $b.Acc.Quit() } catch { }
        try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($b.Acc) } catch { }
    }
    $b.Acc = $null
    if ($b.Pid) {
        $p = Get-Process -Id $b.Pid -ErrorAction SilentlyContinue
        if ($p -and -not $p.WaitForExit(5000)) { try { $p.Kill() } catch { } }
    }
    $b.Pid = 0
    Remove-Item $b.PidFile -ErrorAction SilentlyContinue
}

# Call one entry point. Always returns an object, never throws:
#   ok, errNumber, source, message, value   from the entry point
#   kind     'result' | 'dialog' | 'timeout' | 'com' | 'start'
#   dialogs  what the watcher saw;  ms  call time;  restartMs  when Access was restarted
function Invoke-KitBridge($app, [string]$Name, [object[]]$Arguments = @()) {
    $b = $app.Bridge
    $out = [ordered]@{ ok = $false; errNumber = 0; source = ''; message = ''; value = $null; kind = 'result'; dialogs = @(); ms = 0; restartMs = $null }
    if (-not $b.Acc) {
        try { Start-KitBridge $app } catch { $out.kind = 'start'; $out.errNumber = -1; $out.message = $_.Exception.Message; return [pscustomobject]$out }
    }
    $b.Calls++
    $w = New-Object KitAccessWatch -ArgumentList ([uint32]$b.Pid), $b.TimeoutSec, $false
    $w.Start()
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $raw = $null; $err = $null
    try {
        # Application.Run takes its arguments by reference.
        $a = @($Arguments + @($null, $null, $null, $null))
        $a1 = $a[0]; $a2 = $a[1]; $a3 = $a[2]; $a4 = $a[3]
        switch ($Arguments.Count) {
            0 { $raw = $b.Acc.Run($Name) }
            1 { $raw = $b.Acc.Run($Name, [ref]$a1) }
            2 { $raw = $b.Acc.Run($Name, [ref]$a1, [ref]$a2) }
            3 { $raw = $b.Acc.Run($Name, [ref]$a1, [ref]$a2, [ref]$a3) }
            default { $raw = $b.Acc.Run($Name, [ref]$a1, [ref]$a2, [ref]$a3, [ref]$a4) }
        }
    } catch { $err = $_ }
    $out.ms = [int]$sw.ElapsedMilliseconds
    $found = $w.Stop()
    $out.dialogs = @($found)

    if ($w.Killed) {
        $out.kind = $(if ($w.TimedOut) { 'timeout' } else { 'dialog' })
        $out.errNumber = -1
        $out.message = $(if ($w.TimedOut) { "Access was still busy after $($b.TimeoutSec) s, so it was restarted." } else { 'Access showed a dialog, so it was restarted: ' + ($found -join '; ') })
    } elseif ($err) {
        $out.kind = 'com'
        $out.errNumber = $err.Exception.HResult
        $out.message = $err.Exception.Message
    } else {
        try {
            $res = [string]$raw | ConvertFrom-Json
            $out.ok = [bool]$res.ok; $out.errNumber = [int]$res.errNumber; $out.source = [string]$res.source
            $out.message = [string]$res.message; $out.value = $res.value
        } catch {
            $out.kind = 'com'; $out.errNumber = -2; $out.message = "The entry point did not return JSON: $raw"
        }
        return [pscustomobject]$out
    }

    # Anything but a clean result: make sure the instance is usable, else replace it.
    $alive = $false
    if (-not $w.Killed -and (Get-Process -Id $b.Pid -ErrorAction SilentlyContinue)) {
        try { $alive = (([string]$b.Acc.Run('WebPing') | ConvertFrom-Json).ok) } catch { $alive = $false }
    }
    if (-not $alive) {
        $b.LastProblem = $out.message
        Stop-KitBridge $app
        $b.Restarts++
        try { Start-KitBridge $app; $out.restartMs = $b.StartMs } catch { $out.message += " Restart failed: $($_.Exception.Message)" }
    }
    return [pscustomobject]$out
}

function Get-KitBridgeStatus($app) {
    $b = $app.Bridge
    $p = $null; if ($b.Pid) { $p = Get-Process -Id $b.Pid -ErrorAction SilentlyContinue }
    return [ordered]@{
        runtime = $b.Runtime; pid = $b.Pid; running = [bool]$p
        firstStartMs = $b.FirstStartMs; lastStartMs = $b.StartMs; starts = $b.Starts; restarts = $b.Restarts; calls = $b.Calls
        workingSetMB = $(if ($p) { [math]::Round($p.WorkingSet64 / 1MB, 1) } else { $null })
        privateMB = $(if ($p) { [math]::Round($p.PrivateMemorySize64 / 1MB, 1) } else { $null })
        timeoutSec = $b.TimeoutSec; lastProblem = $b.LastProblem
    }
}
