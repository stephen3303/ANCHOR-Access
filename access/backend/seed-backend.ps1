# Seeds an EXISTING shared data file (ANCHOR_BE.accdb) with the lookup lists
# and the web app's fictional demo traffic (backend\seed, modBuildBackEnd.bas),
# adding only the rows it doesn't have yet, so running it again adds nothing.
# Every row is matched on its natural key first (a message on its key, a
# template or workflow on its name, a setting on its key), so rows an admin
# changed or added are never touched, and nothing is ever deleted.
#
#   backend\seed-backend.ps1                      seeds build\ANCHOR_BE.accdb
#   backend\seed-backend.ps1 -Path <file>         seeds that data file
#   -Personnel <file>                             also give the made-up people of a
#                                                 DEMO Personnel file (the app's
#                                                 Start with demo personnel) ANCHOR
#                                                 profiles and roles, the seeded
#                                                 routing steps and release authority
#
# What it adds: message types and classes, settings, alert rules and access
# groups; roles; the 16 demo messages with their sections, attachments,
# addressees and acknowledgements; address groups; the subscription catalogue
# and requests; watch notes, topics and the turnover checklist; the four
# routing workflows and four templates; and the audit rows, only while the
# audit trail is empty (it is append-only and hash chained). All of it is
# made up; nothing in it names a real person.
#
# Close every ANCHOR copy that has the data file open first: it is opened
# exclusively, and all rows go in one transaction (all or none).
# It prints  SeedBackEndAt: OK: <n> rows added (<by table>)  and exits 0, or
# the reason it changed nothing and exits 1.
# Requires Microsoft Access (2016+ / Microsoft 365).
param([string]$Path = '', [string]$Personnel = '')
$ErrorActionPreference = 'Stop'
$backend = Split-Path -Parent $MyInvocation.MyCommand.Path
$here = Split-Path -Parent $backend
if (-not $Path) { $Path = Join-Path $here 'build\ANCHOR_BE.accdb' }
# Relative to where the script was started, not to the process's folder.
$Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
if (-not (Test-Path $Path)) {
    throw "There is no data file at $Path. Make one with backend\new-backend.ps1 (-Empty for one with no demo rows)."
}
if ($Personnel) {
    $Personnel = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Personnel)
    if (-not (Test-Path $Personnel)) { throw "There is no Personnel file at $Personnel." }
}

. (Join-Path $backend 'BackEndBuilder.ps1')
$result = Invoke-BackEndBuilder -Entry 'SeedBackEndAt' -Arguments @($Path, (Join-Path $backend 'seed'), $Personnel)
Write-Host "SeedBackEndAt: $result"
if (-not $result.StartsWith('OK')) { exit 1 }
# The front end shows these rows only when it is linked to this file
# (build.ps1 with no -BackEnd makes a copy with no data at all).
Write-Host "A front end shows these rows only when it is linked to this file. If yours isn't yet:"
Write-Host "  .\build.ps1 -BackEnd `"$Path`"   (from the access folder)"
