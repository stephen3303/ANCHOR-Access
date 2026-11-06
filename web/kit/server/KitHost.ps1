<#
Starts the kit web server with every app listed in a settings file: one
server, one UI, each app on its own Access data file.

  web\kit\server\KitHost.ps1 -Config web\web.config.json
  web\kit\server\KitHost.ps1 -Config web\web.config.json -DemoSignIn     (review only: pick a person)

Settings (paths may be relative to the settings file):
  {
    "title": "Navy Apps",
    "personnel": "S:\\Apps\\Personnel\\Personnel_be.accdb",
    "port": 8787,
    "apps": [
      { "key": "anchor",  "name": "ANCHOR",         "dataFile": "S:\\Apps\\ANCHOR\\ANCHOR_BE.accdb",   "module": "apps\\anchor\\api.ps1" },
      { "key": "examgen", "name": "Exam Generator", "dataFile": "S:\\Apps\\ExamGen\\ExamGen_BE.accdb", "module": "apps\\examgen\\api.ps1" }
    ]
  }

  -Port      overrides the settings file's port
  -Lan       also answer other PCs (admin: netsh http add urlacl url=http://+:8787/ user=<account>; firewall rule)
  -LogFile   write the server's output there (the launcher uses it)
#>
param(
    [Parameter(Mandatory = $true)][string]$Config,
    [int]$Port = 0,
    [switch]$Lan,
    [switch]$DemoSignIn,
    [string]$LogFile = ''
)
$ErrorActionPreference = 'Stop'
$kitServer = Split-Path -Parent $MyInvocation.MyCommand.Path
$web = Split-Path -Parent (Split-Path -Parent $kitServer)
. (Join-Path $kitServer 'KitWeb.ps1')

$Config = (Resolve-Path $Config).Path
$cfgDir = Split-Path -Parent $Config
$cfg = Get-Content -Raw $Config | ConvertFrom-Json
function Get-ConfigPath([string]$p) { if ([IO.Path]::IsPathRooted($p)) { return $p } return (Join-Path $cfgDir $p) }

if (-not $Port) { $Port = if ($cfg.port) { [int]$cfg.port } else { 8787 } }
$personnel = ''; if ($cfg.personnel) { $personnel = Get-ConfigPath $cfg.personnel }
$title = if ($cfg.title) { [string]$cfg.title } else { 'Apps' }
Initialize-KitWeb -Title $title -Personnel $personnel -Port $Port -Lan:$Lan -DemoSignIn:$DemoSignIn `
    -UiFile (Join-Path $web 'dist\index.html') -LogFile $LogFile

# Each module is dot-sourced here, at script level, so the functions it
# defines stay available to its routes. Module functions share one scope:
# prefix them with the app's name (Get-AnchorSetting, not Get-Setting).
foreach ($a in $cfg.apps) {
    Use-KitApp -Key $a.key -Name $a.name -DataFile (Get-ConfigPath $a.dataFile)
    . (Get-ConfigPath $a.module)
}
Start-KitWeb
