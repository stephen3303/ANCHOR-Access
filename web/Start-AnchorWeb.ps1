<#
Starts ANCHOR's web front end for the person at this PC and opens it in the
browser. This is the everyday way in: each PC runs its own copy (this folder),
listening only on localhost, and every copy works on the one shared data file
on the network drive, the same way each Access front end does.

Settings come from anchor-web.config.json next to this script:
  { "dataFile":  "S:\\Apps\\ANCHOR\\ANCHOR_BE.accdb",
    "personnel": "S:\\Apps\\Personnel\\Personnel_be.accdb",
    "port": 8787 }

  Start-AnchorWeb.ps1          start (or reuse) the server and open the browser
  Start-AnchorWeb.ps1 -Stop    stop this person's server

The server runs hidden; its log is %LOCALAPPDATA%\ANCHOR\web.log.
#>
param([switch]$Stop)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$state = Join-Path $env:LOCALAPPDATA 'ANCHOR'
New-Item -ItemType Directory -Force $state | Out-Null
$pidFile = Join-Path $state 'web.pid'

function Stop-Server {
    if (Test-Path $pidFile) {
        $old = Get-Content $pidFile -ErrorAction SilentlyContinue
        if ($old) { Stop-Process -Id ([int]$old) -Force -ErrorAction SilentlyContinue }
        Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
    }
}
if ($Stop) { Stop-Server; Write-Host 'ANCHOR web stopped.'; return }

$cfgPath = Join-Path $here 'anchor-web.config.json'
if (-not (Test-Path $cfgPath)) { throw "No settings file: $cfgPath (see anchor-web.config.example.json)." }
$cfg = Get-Content -Raw $cfgPath | ConvertFrom-Json
$port = if ($cfg.port) { [int]$cfg.port } else { 8787 }
$url = "http://localhost:$port/"

function Test-Up {
    try { [void](Invoke-WebRequest "$($url)api/me" -UseDefaultCredentials -UseBasicParsing -TimeoutSec 2); return $true } catch { return $false }
}

if (-not (Test-Up)) {
    Stop-Server
    if (-not (Test-Path $cfg.dataFile)) { throw "Can't reach the data file: $($cfg.dataFile). Is the network drive connected?" }
    # The server must be the same bitness as Office, or it can't load the Access engine.
    $ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $platform = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).Platform
    if ($platform -eq 'x86' -and [Environment]::Is64BitOperatingSystem) { $ps = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe' }
    $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Join-Path $here 'server\AnchorWeb.ps1')`"",
              '-DataFile', "`"$($cfg.dataFile)`"", '-Port', $port)
    if ($cfg.personnel) { $psArgs += @('-Personnel', "`"$($cfg.personnel)`"") } else { throw 'Set "personnel" in anchor-web.config.json.' }
    $p = Start-Process $ps -ArgumentList $psArgs -WindowStyle Hidden -PassThru `
        -RedirectStandardOutput (Join-Path $state 'web.log') -RedirectStandardError (Join-Path $state 'web-errors.log')
    Set-Content $pidFile $p.Id
    $up = $false
    for ($i = 0; $i -lt 30 -and -not $up; $i++) { Start-Sleep -Milliseconds 500; $up = Test-Up; if ($p.HasExited) { break } }
    if (-not $up) { throw "The server didn't start. See $(Join-Path $state 'web-errors.log')." }
}
Start-Process $url
