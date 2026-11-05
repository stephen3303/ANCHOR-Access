# Makes a new shared data file (ANCHOR_BE.accdb) with modBuildBackEnd.bas:
# every table, relationship and index, the lookup lists, and the web app's
# fictional demo traffic from backend\seed (tools\make_seed.py).
#
#   backend\new-backend.ps1                  writes build\ANCHOR_BE.accdb (for testing)
#   backend\new-backend.ps1 -Path <file>     writes it there (a full path to a new .accdb)
#   -Empty                                   no demo traffic: tables and lookups only
#
# It never overwrites: if the file is there already it stops and changes nothing.
# To add the demo rows to a data file that exists, use backend\seed-backend.ps1.
# Then link a front end to it with  .\build.ps1 -BackEnd <file>.
# Requires Microsoft Access (2016+ / Microsoft 365).
param([string]$Path = '', [switch]$Empty)
$ErrorActionPreference = 'Stop'
$backend = Split-Path -Parent $MyInvocation.MyCommand.Path
$here = Split-Path -Parent $backend
if (-not $Path) {
    New-Item -ItemType Directory -Force (Join-Path $here 'build') | Out-Null
    $Path = Join-Path $here 'build\ANCHOR_BE.accdb'
}
# Relative to where the script was started, not to the process's folder.
$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
if (Test-Path $Path) { throw "There is already a file at $Path; nothing was changed." }

. (Join-Path $backend 'BackEndBuilder.ps1')
$seed = if ($Empty) { '' } else { Join-Path $backend 'seed' }
$result = Invoke-BackEndBuilder -Entry 'BuildBackEndAt' -Arguments @($Path, $seed)
Write-Host "BuildBackEndAt: $result"
if ($result -ne 'OK') { exit 1 }
Write-Host "Created: $Path"
