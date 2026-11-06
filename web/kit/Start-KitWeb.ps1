<#
Starts an app's web front end for the person at this PC and opens it in the
browser. This is the everyday way in: each PC runs its own copy, listening
only on localhost, and every copy works on the one shared data file on the
network drive, the same way each Access front end does.

  Start-KitWeb.ps1 -Config web\web.config.json          start (or reuse) and open the browser
  Start-KitWeb.ps1 -Config web\web.config.json -Stop    stop this person's server

The settings file is the one kit\server\KitHost.ps1 reads (title, personnel,
port, apps). The server runs hidden; its log is %LOCALAPPDATA%\<title>\web.log.
#>
param([Parameter(Mandatory = $true)][string]$Config, [switch]$Stop)
$ErrorActionPreference = 'Stop'
if (-not (Test-Path $Config)) { throw "No settings file: $Config (copy the .example.json next to it and set the paths)." }
$cfgDir = Split-Path -Parent (Resolve-Path $Config).Path
$cfg = Get-Content -Raw $Config | ConvertFrom-Json
foreach ($k in 'title', 'personnel', 'apps') { if (-not $cfg.$k) { throw "Set `"$k`" in $Config." } }
function Full([string]$p) { if ([IO.Path]::IsPathRooted($p)) { return $p } return (Join-Path $cfgDir $p) }
$here = Split-Path -Parent $MyInvocation.MyCommand.Path

$state = Join-Path $env:LOCALAPPDATA ($cfg.title -replace '[^\w\- ]', '')
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
if ($Stop) { Stop-Server; Write-Host "$($cfg.title) stopped."; return }

function Test-Up {
    try { [void](Invoke-WebRequest "$($url)api/me" -UseDefaultCredentials -UseBasicParsing -TimeoutSec 2); return $true } catch { return $false }
}

if (-not (Test-Up)) {
    Stop-Server
    if (-not (Test-Path (Full $cfg.personnel))) { throw "Can't reach the Personnel file: $(Full $cfg.personnel). Is the network drive connected?" }
    # The server must be the same bitness as the Access engine (Office's, or
    # the stand-alone engine's), or it can't load it.
    $ps = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $use32 = $false
    $platform = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration' -ErrorAction SilentlyContinue).Platform
    if ($platform) { $use32 = ($platform -eq 'x86') }
    elseif ([Environment]::Is64BitOperatingSystem) {
        # No Click-to-Run Office: look where the engine is registered.
        $ace64 = (Test-Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Microsoft.ACE.OLEDB.16.0') -or (Test-Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0')
        $ace32 = (Test-Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes\Microsoft.ACE.OLEDB.16.0') -or (Test-Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Classes\Microsoft.ACE.OLEDB.12.0')
        $use32 = ($ace32 -and -not $ace64)
    }
    if ($use32 -and [Environment]::Is64BitOperatingSystem) { $ps = Join-Path $env:WINDIR 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe' }
    $psArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$(Join-Path $here 'server\KitHost.ps1')`"",
                '-Config', "`"$((Resolve-Path $Config).Path)`"", '-LogFile', "`"$log`"")
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
