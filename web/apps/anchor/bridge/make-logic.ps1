<#
Makes ANCHOR_logic.accdb for the web bridge spike: the ANCHOR front end's VBA,
forms, queries and linked tables, with no start-up (no AutoExec, no AutoKeys,
no start-up form), plus modWebBridge. The web server opens it in a hidden
Access instance and calls the bridge's entry points over COM.

  web\apps\anchor\bridge\make-logic.ps1 -FrontEnd access\build\ANCHOR.accdb -Out <folder>\ANCHOR_logic.accdb

-FrontEnd is a front end built linked to the data file (access\build.ps1
-BackEnd <file>). It is read, never opened as the current database, so its
start-up never runs. Requires Microsoft Access (the full product: the free
Runtime can run the logic file but can't make it).
#>
param(
    [Parameter(Mandatory = $true)][string]$FrontEnd,
    [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$repo = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $here)))
$FrontEnd = (Resolve-Path $FrontEnd).Path
$Out = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Out)
if (Test-Path $Out) { throw "There is already a file at $Out; delete it first." }
. (Join-Path $repo 'access\tools\AccessWatch.ps1')

$acc = New-Object -ComObject Access.Application
$watch = Start-AccessWatch $acc -TimeoutSec 600
$failed = $false
try {
    $acc.Visible = $false
    $acc.AutomationSecurity = 1
    $acc.NewCurrentDatabase($Out)
    $acc.VBE.ActiveVBProject.Name = 'ANCHORLogic'
    $db = $acc.CurrentDb()
    $src = $acc.DBEngine.OpenDatabase($FrontEnd, $false, $true)
    $acImport = 0; $acTable = 0; $acQuery = 1; $acForm = 2; $acReport = 3; $acMacro = 4; $acModule = 5
    $n = @{ links = 0; tables = 0; queries = 0; forms = 0; reports = 0; macros = 0; modules = 0 }

    # Tables: links are re-made with the same connect string (password
    # included); local tables are copied.
    foreach ($td in @($src.TableDefs)) {
        if ($td.Name -like 'MSys*' -or $td.Name -like '~*') { continue }
        if ($td.Connect) {
            $nt = $db.CreateTableDef($td.Name)
            $nt.Connect = $td.Connect
            $nt.SourceTableName = $td.SourceTableName
            $db.TableDefs.Append($nt); $n.links++
        } else {
            $acc.DoCmd.TransferDatabase($acImport, 'Microsoft Access', $FrontEnd, $acTable, $td.Name, $td.Name); $n.tables++
        }
    }
    foreach ($qd in @($src.QueryDefs)) {
        if ($qd.Name -like '~*') { continue }
        $acc.DoCmd.TransferDatabase($acImport, 'Microsoft Access', $FrontEnd, $acQuery, $qd.Name, $qd.Name); $n.queries++
    }
    foreach ($kind in @(@('Modules', $acModule, 'modules'), @('Forms', $acForm, 'forms'), @('Reports', $acReport, 'reports'), @('Scripts', $acMacro, 'macros'))) {
        foreach ($d in @($src.Containers($kind[0]).Documents)) {
            if ($kind[1] -eq $acMacro -and ($d.Name -eq 'AutoExec' -or $d.Name -eq 'AutoKeys')) { continue }
            $acc.DoCmd.TransferDatabase($acImport, 'Microsoft Access', $FrontEnd, $kind[1], $d.Name, $d.Name); $n[$kind[2]]++
        }
    }
    # The front end's own properties the code reads (modSignIn.DataLinked, the error log's build).
    foreach ($p in @('HasDataFile', 'AppVersion', 'AppBuild', 'AppTitle')) {
        try {
            $v = $src.Properties($p).Value
            $db.Properties.Append($db.CreateProperty($p, $src.Properties($p).Type, $v))
        } catch { }
    }
    $src.Close()

    # The bridge module.
    $tmp = Join-Path $env:TEMP 'modWebBridge.bas'
    $text = [IO.File]::ReadAllText((Join-Path $here 'modWebBridge.bas')) -replace "`r?`n", "`r`n"
    $text = $text -replace '^Attribute VB_Name = "modWebBridge"\r\n', ''
    [IO.File]::WriteAllText($tmp, $text, [Text.Encoding]::ASCII)
    $acc.LoadFromText($acModule, 'modWebBridge', $tmp)
    Remove-Item $tmp

    $bad = Test-AccessCompiles $acc $watch
    if ($bad) { Write-Host "Compile error: $bad"; throw 'compile failed' }
    Write-Host ("Made {0}: {1} linked tables, {2} local tables, {3} queries, {4} forms, {5} reports, {6} macros, {7} modules + modWebBridge" -f
        $Out, $n.links, $n.tables, $n.queries, $n.forms, $n.reports, $n.macros, $n.modules)
}
catch { $failed = $true; if (-not $watch.Killed) { throw } }
finally {
    $found = Stop-AccessWatch $watch $acc
    if ($found) { $failed = $true }
}
if ($failed) { Remove-Item $Out -ErrorAction SilentlyContinue; exit 1 }
