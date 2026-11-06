<#
Starts an app's web front end for the person at this PC and opens it in the
browser. This is the everyday way in: each PC runs its own copy, listening
only on localhost, and every copy works on the one shared data file on the
network drive, the same way each Access front end does.

  Start-KitWeb.ps1 -Config <app folder>\<app>-web.config.json          start (or reuse) and open the browser
  Start-KitWeb.ps1 -Config <app folder>\<app>-web.config.json -Stop    stop this person's server

The settings file (paths may be relative to the settings file):
  { "app": "ANCHOR", "server": "server\\AnchorWeb.ps1",
    "dataFile": "S:\\Apps\\ANCHOR\\ANCHOR_BE.accdb",
    "personnel": "S:\\Apps\\Personnel\\Personnel_be.accdb", "port": 8787 }

The server runs hidden; its log is %LOCALAPPDATA%\<app>\web.log.
#>
param([Parameter(Mandatory = $true)][string]$Config, [switch]$Stop)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Config)) { throw "No settings file: $Config (copy the .example.json next to it and set the paths)." }
$cfgDir = Split-Path -Parent (Resolve-Path $Config).Path
$cfg = Get-Content -Raw $Config | ConvertFrom-Json
foreach ($k in 'app', 'server', 'dataFile', 'personnel') { if (-not $cfg.$k) { throw "Set `"$k`" in $Config." } }
function Full([string]$p) { if ([IO.Path]::IsPathRooted($p)) { return $p } return (Join-Path $cfgDir $p) }

$state = Join-Path $env:LOCALAPPDATA $cfg.app
New-Item -ItemType Directory -Force $state | Out-Null
$pidFile = Join-Path $state 'web.pid'
$log = Join-Path $state 'web.log'
$port = if ($cfg.port) { [int]$cfg.port } else { 8787 }
$url = "http://localhost:$port/"

function Stop-Server {
    if (Test-Path $pidFile) {
        $old = Get-Content $pidFile -ErrorAction SilentlyContinue
        if ($old) { Stop-Process -Id ([int]$old) -Force -ErrorAction SilentlyContinue }
        Remove-Item $pidFile -Force -ErrorAction SilentlyContinue
    }
}
if ($Stop) { Stop-Server; Write-Host "$($cfg.app) web stopped."; return }

function Test-Up {
    try { [void](Invoke-WebRequest "$($url)api/me" -UseDefaultCredentials -UseBasicParsing -TimeoutSec 2); return $true } catch { return $false }
}

if (-not (Test-Up)) {
    Stop-Server
    $data = Full $cfg.dataFile
    if (-not (Test-Path $data)) { throw "Can't reach the data file: $data. Is the network drive connected?" }
    # The server must be the same bitness as Office, or it can't load the Access engine.
    $ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $platform = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).Platform
    if ($platform -eq 'x86' -and [Environment]::Is64BitOperatingSystem) { $ps = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe' }
    $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Full $cfg.server)`"",
                '-DataFile', "`"$data`"", '-Personnel', "`"$(Full $cfg.personnel)`"", '-Port', $port, '-LogFile', "`"$log`"")
    # No output redirection: the server writes its own log, and inherits none
    # of this window's handles, so a script that captures this one's output
    # doesn't wait on the server.
    $p = Start-Process $ps -ArgumentList $psArgs -WindowStyle Hidden -PassThru
    Set-Content $pidFile $p.Id
    $up = $false
    for ($i = 0; $i -lt 30 -and -not $up; $i++) { Start-Sleep -Milliseconds 500; $up = Test-Up; if ($p.HasExited) { break } }
    if (-not $up) { throw "The server didn't start. See $log." }
}
Start-Process $url
