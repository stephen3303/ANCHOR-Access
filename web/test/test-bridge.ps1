<#
Spike test of the hidden-Access bridge on ANCHOR: the web server runs
modMessages.MarkRead / Acknowledge and modAudit.AuditVerify themselves, in a
hidden Access instance, over COM.

  web\test\test-bridge.ps1 -DataFile <scratch>\ANCHOR_BE.accdb -Logic <dir>\ANCHOR_logic.accdb [-Runtime] [-Report out.json]

Make the logic file first, linked to the SAME scratch data file:
  access\build.ps1 -BackEnd <scratch>\ANCHOR_BE.accdb -OutDir <dir>
  web\apps\anchor\bridge\make-logic.ps1 -FrontEnd <dir>\ANCHOR.accdb -Out <dir>\ANCHOR_logic.accdb

-Runtime opens the logic file with MSACCESS.EXE /runtime (what the free Access
Runtime does). The logic file's folder must then be an Access Trusted Location,
or Access shows a security notice (the test reports it).

It writes to the data file (a read, an acknowledgement, their audit rows):
use a scratch copy.
#>
param(
    [Parameter(Mandatory = $true)][string]$DataFile,
    [Parameter(Mandatory = $true)][string]$Logic,
    [switch]$Runtime,
    [int]$Port = 8789,
    [int]$TimeoutSec = 8,
    [string]$Report = ''
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$web = Split-Path -Parent $here
function Get-FullPath([string]$p) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) }
$DataFile = Get-FullPath $DataFile; $Logic = Get-FullPath $Logic
$base = "http://localhost:$Port"
$fails = 0
$facts = [ordered]@{ mode = $(if ($Runtime) { 'runtime' } else { 'automation' }); computer = $env:COMPUTERNAME }
function Check([string]$name, [bool]$ok, [string]$detail = '') {
    if ($ok) { Write-Host "PASS  $name $detail" -ForegroundColor Green }
    else { Write-Host "FAIL  $name $detail" -ForegroundColor Red; $script:fails++ }
}
function Note([string]$name, $value) { $script:facts[$name] = $value; Write-Host ("INFO  {0}: {1}" -f $name, ($value | ConvertTo-Json -Compress -Depth 6)) }
function Api([string]$method, [string]$path, $headers = @{}) {
    try {
        $r = Invoke-WebRequest -Uri ($base + $path) -Method $method -Headers $headers -UseDefaultCredentials -UseBasicParsing -TimeoutSec 120
        return @{ status = [int]$r.StatusCode; body = ($r.Content | ConvertFrom-Json); raw = $r.Content }
    } catch [System.Net.WebException] {
        $resp = $_.Exception.Response
        if (-not $resp) { throw }
        # PowerShell 5.1 has often read the body already: it is in ErrorDetails.
        $txt = ''
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $txt = $_.ErrorDetails.Message }
        else {
            $st = $resp.GetResponseStream()
            try { if ($st.CanSeek) { $st.Position = 0 } } catch { }
            $txt = (New-Object System.IO.StreamReader($st)).ReadToEnd()
        }
        $body = $null; try { $body = $txt | ConvertFrom-Json } catch { }
        return @{ status = [int]$resp.StatusCode; body = $body; raw = $txt }
    }
}

$cfgFile = Join-Path $env:TEMP "kit-bridge-test-$Port.json"
@{ title = 'Kit bridge test'; personnel = ''; port = $Port
   apps = @(@{ key = 'anchor'; name = 'ANCHOR'; dataFile = $DataFile; module = (Join-Path $web 'apps\anchor\api.ps1')
               bridge = @{ logic = $Logic; runtime = [bool]$Runtime; timeoutSec = $TimeoutSec; tests = $true } }) } |
    ConvertTo-Json -Depth 5 | Set-Content -Encoding UTF8 $cfgFile
$log = Join-Path $env:TEMP "kit-bridge-test-$Port.log"
$server = $null
try {
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $server = Start-Process powershell -PassThru -WindowStyle Hidden -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
        "`"$(Join-Path $web 'kit\server\KitHost.ps1')`"", '-Config', "`"$cfgFile`"", '-DemoSignIn', '-LogFile', "`"$log`"")
    $up = $false
    for ($i = 0; $i -lt 240 -and -not $up; $i++) {
        Start-Sleep -Milliseconds 250
        try { [void](Invoke-WebRequest "$base/api/me" -UseDefaultCredentials -UseBasicParsing); $up = $true } catch { }
    }
    Check 'server (with hidden Access) starts' $up ("in $($sw.ElapsedMilliseconds) ms")
    if (-not $up) { if (Test-Path $log) { Get-Content $log -Tail 30 }; exit 1 }
    $facts.serverReadyMs = [int]$sw.ElapsedMilliseconds

    $me = Api GET '/api/me'
    $csrf = $me.body.csrf
    $people = (Api GET '/api/people').body.people
    $as = $null; $sum = $null
    foreach ($p in $people) {
        $h = @{ 'X-Kit-As' = [string]$p.PersonID }
        $s = Api GET '/api/anchor/summary' $h
        if ($s.status -eq 200 -and @($s.body.ackDue).Count -gt 0) { $as = $h; $sum = $s; break }
    }
    Check 'found a person with an acknowledgement due' ([bool]$as) $(if ($as) { "person $($as['X-Kit-As'])" })
    $w = @{ 'X-Kit-As' = $(if ($as) { $as['X-Kit-As'] } else { [string]$people[0].PersonID }); 'X-Kit-Csrf' = $csrf }

    $st = (Api GET '/api/anchor/bridge/status' $w).body
    Check 'hidden Access is running' ([bool]$st.running) ("pid $($st.pid)")
    Note 'startMs' $st.firstStartMs
    Note 'memoryAtStartMB' @{ workingSet = $st.workingSetMB; private = $st.privateMB }
    $ping = (Api POST '/api/anchor/bridge/test/ping' $w).body.result
    Check 'entry point answers with a result object' ($ping.ok -and $ping.kind -eq 'result') ($ping.value | ConvertTo-Json -Compress)
    Note 'accessInfo' $ping.value

    # Latency of a bare call.
    $times = @(); for ($i = 0; $i -lt 10; $i++) { $times += (Api POST '/api/anchor/bridge/test/ping' $w).body.result.ms }
    Note 'pingMs' @{ min = ($times | Measure-Object -Minimum).Minimum; avg = [math]::Round(($times | Measure-Object -Average).Average, 1); max = ($times | Measure-Object -Maximum).Maximum }

    if ($as) {
        $target = @($sum.body.ackDue)[0]
        $t0 = [Diagnostics.Stopwatch]::StartNew()
        $view = Api GET "/api/anchor/messages/$($target.id)" $w
        Check 'message view (MarkRead in Access)' ($view.status -eq 200) ("$($t0.ElapsedMilliseconds) ms")
        $t0.Restart()
        $ack = Api POST "/api/anchor/messages/$($target.id)/ack" $w
        Check 'acknowledge (modMessages.Acknowledge in Access)' ($ack.status -eq 200 -and $ack.body.ok) ("$($t0.ElapsedMilliseconds) ms, Access $($ack.body.bridge.ms) ms")
        $again = Api POST "/api/anchor/messages/$($target.id)/ack" $w
        Check 'second acknowledge: VBA error 5 relayed as 409 JSON' ($again.status -eq 409 -and $again.body.bridge.errNumber -eq 5) $again.body.error
        $audit = (Api GET '/api/anchor/audit' $w).body
        $mine = @($audit.rows | Where-Object { $_.Action -eq 'Acknowledge' -and $_.Resource -like "*$($target.reference)*" -and $_.Details -eq 'Receipt acknowledged in ANCHOR.' })
        Check 'audit row written by the VBA routine' ($mine.Count -ge 1) ("Device $($mine[0].Device)")
        Check '... and marked as a web action' ($mine.Count -ge 1 -and $mine[0].Device -like '*(web)')
    }

    $st0 = (Api GET '/api/anchor/bridge/status' $w).body
    $tr = (Api POST '/api/anchor/bridge/test/trapped' $w).body
    Check 'trapped error relayed as JSON' (-not $tr.result.ok -and $tr.result.kind -eq 'result' -and $tr.result.errNumber -eq 5) $tr.result.message
    Check '... without restarting Access' ($tr.status.restarts -eq $st0.restarts)

    foreach ($case in @('untrapped', 'msgbox', 'dialog', 'slow')) {
        $before = (Api GET '/api/anchor/bridge/status' $w).body
        $t0 = [Diagnostics.Stopwatch]::StartNew()
        $res = Api POST "/api/anchor/bridge/test/$case" $w
        $wall = $t0.ElapsedMilliseconds
        $r = $res.body.result
        Note "fault_$case" @{ kind = $r.kind; errNumber = $r.errNumber; message = $r.message; dialogs = $r.dialogs; callMs = $r.ms; restartMs = $r.restartMs; wallMs = $wall
                              restarted = ($res.body.status.restarts -gt $before.restarts) }
        Check "${case}: page gets JSON, not a hang" ($res.status -eq 200 -and -not $r.ok) ("kind $($r.kind) in $wall ms")
        $p2 = (Api POST '/api/anchor/bridge/test/ping' $w).body.result
        Check "${case}: bridge works again after" ($p2.ok) $(if ($r.restartMs) { "restarted in $($r.restartMs) ms" } else { 'no restart needed' })
    }

    # A real write after all the restarts.
    if ($as) {
        $other = $null
        foreach ($p in $people) {
            if ([string]$p.PersonID -eq $as['X-Kit-As']) { continue }
            $h = @{ 'X-Kit-As' = [string]$p.PersonID; 'X-Kit-Csrf' = $csrf }
            $s = Api GET '/api/anchor/summary' $h
            if ($s.status -eq 200 -and @($s.body.ackDue).Count -gt 0) { $other = @{ h = $h; id = @($s.body.ackDue)[0].id }; break }
        }
        if ($other) {
            $a2 = Api POST "/api/anchor/messages/$($other.id)/ack" $other.h
            Check 'acknowledge after restarts' ($a2.status -eq 200) "person $($other.h['X-Kit-As'])"
        }
    }

    $v = (Api GET '/api/anchor/bridge/verify' $w).body
    Check 'modAudit.AuditVerify (in Access) passes' ([bool]$v.access.ok -and [bool]$v.access.value.ok) ("$($v.access.value.checked) entries $($v.access.value.problem)")
    Check 'web chain check passes too' ([bool]$v.web.ok) ("$($v.web.checked) entries $($v.web.problem)")

    $st = (Api GET '/api/anchor/bridge/status' $w).body
    Note 'memoryAtEndMB' @{ workingSet = $st.workingSetMB; private = $st.privateMB }
    Note 'restarts' $st.restarts
} finally {
    $leftover = $null
    if ($server -and -not $server.HasExited) {
        Stop-Process -Id $server.Id -Force
        Start-Sleep -Seconds 2
    }
    $pidFile = "$Logic.pid"
    if (Test-Path $pidFile) {
        $bp = [int](Get-Content $pidFile)
        $leftover = Get-Process -Id $bp -ErrorAction SilentlyContinue
        if ($leftover) { Stop-Process -Id $bp -Force -ErrorAction SilentlyContinue }
        Remove-Item $pidFile -ErrorAction SilentlyContinue
    }
    $facts.accessLeftRunningAfterServerKilled = [bool]$leftover
    Write-Host ("INFO  hidden Access left running after the server was killed: {0}" -f [bool]$leftover)
    Remove-Item $cfgFile -ErrorAction SilentlyContinue
    $facts.failures = $fails
    if ($Report) { $facts | ConvertTo-Json -Depth 6 | Set-Content -Encoding UTF8 $Report }
    Write-Host "Server log: $log"
}
if ($fails) { Write-Host "$fails check(s) failed." -ForegroundColor Red; exit 1 }
Write-Host 'All checks passed.' -ForegroundColor Green
