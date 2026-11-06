<#
ANCHOR web front end: starts the kit web server with ANCHOR's module.

  web\server\AnchorWeb.ps1 -DataFile S:\Apps\ANCHOR\ANCHOR_BE.accdb -Personnel S:\Apps\Personnel\Personnel_be.accdb
  web\server\AnchorWeb.ps1 -DataFile build\ANCHOR_BE.accdb -DemoSignIn      (review only: pick a person)

  -Port 8787     the port (http://localhost:8787/)
  -Lan           also answer other PCs (admin: netsh http add urlacl url=http://+:8787/ user=<account>; firewall rule)
  -DemoSignIn    no Personnel file; the UI offers "Sign in as". Review only.

Most people don't run this directly: Start-AnchorWeb.cmd does, hidden. See web\README.md.
#>
param(
    [Parameter(Mandatory = $true)][string]$DataFile,
    [string]$Personnel = '',
    [int]$Port = 8787,
    [switch]$Lan,
    [switch]$DemoSignIn,
    [string]$LogFile = ''
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$web = Split-Path -Parent $here
. (Join-Path $web 'kit\server\KitWeb.ps1')
Initialize-KitWeb -AppName 'ANCHOR' -DataFile $DataFile -Personnel $Personnel -Port $Port -Lan:$Lan -DemoSignIn:$DemoSignIn `
    -UiFile (Join-Path $web 'dist\index.html') -LogFile $LogFile
. (Join-Path $here 'AnchorApi.ps1')
Start-KitWeb -OnStart $AnchorOnStart
