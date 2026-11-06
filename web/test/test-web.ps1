<#
End-to-end check of the multi-app web server against real data files. Writes a
temporary settings file, starts kit\server\KitHost.ps1 -DemoSignIn in its own
window, drives the API as the signed-in Windows user, and prints PASS/FAIL per
step.

  web\test\test-web.ps1 -DataFile access\build\ANCHOR_BE.accdb
  web\test\test-web.ps1 -DataFile access\build\ANCHOR_BE.accdb -ExamGenDataFile <eg>\access\build\ExamGen_BE.accdb

The ANCHOR file must be seeded with people (backend\seed-backend.ps1 -Personnel).
The Exam Generator file can be a fresh one from its backend\new-backend.ps1.

It writes to the data files: one read and one acknowledgement in ANCHOR, with
their audit rows, and in the Exam Generator file one tblPeople row for the demo
person (removed again at the end). Use scratch copies.
#>
param(
    [Parameter(Mandatory = $true)][string]$DataFile,
    [string]$ExamGenDataFile = '',
    [int]$Port = 8787
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$web = Split-Path -Parent $here
. (Join-Path $web 'kit\server\KitWeb.ps1')   # for Open-KitDb / Invoke-KitNonQuery
function Get-FullPath([string]$p) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) }
$DataFile = Get-FullPath $DataFile
if ($ExamGenDataFile) {
    $ExamGenDataFile = Get-FullPath $ExamGenDataFile
    $Kit.Provider = Find-KitAceProvider $ExamGenDataFile   # this script opens the EG file itself
}
$base = "http://localhost:$Port"
$fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = '') {
    if ($ok) { Write-Host "PASS  $name $detail" -ForegroundColor Green }
    else { Write-Host "FAIL  $name $detail" -ForegroundColor Red; $script:fails++ }
}
function Api([string]$method, [string]$path, $headers = @{}) {
    try {
        $r = Invoke-WebRequest -Uri ($base + $path) -Method $method -Headers $headers -UseDefaultCredentials -UseBasicParsing
        return @{ status = [int]$r.StatusCode; body = ($r.Content | ConvertFrom-Json); raw = $r.Content }
    } catch [System.Net.WebException] {
        $resp = $_.Exception.Response
        if (-not $resp) { throw }
        $txt = (New-Object System.IO.StreamReader($resp.GetResponseStream())).ReadToEnd()
        $body = $null; try { $body = $txt | ConvertFrom-Json } catch { }
        return @{ status = [int]$resp.StatusCode; body = $body; raw = $txt }
    }
}
function Wait-Server {
    for ($i = 0; $i -lt 40; $i++) {
        Start-Sleep -Milliseconds 500
        try { [void](Invoke-WebRequest "$base/api/me" -UseDefaultCredentials -UseBasicParsing); return $true } catch { }
    }
    return $false
}

# Settings file for this run.
$apps = @(@{ key = 'anchor'; name = 'ANCHOR'; dataFile = $DataFile; module = (Join-Path $web 'apps\anchor\api.ps1') })
if ($ExamGenDataFile) {
    $apps += @{ key = 'examgen'; name = 'Exam Generator'; dataFile = $ExamGenDataFile; module = (Join-Path $web 'apps\examgen\api.ps1') }
}
$cfgFile = Join-Path $env:TEMP "kit-web-test-$Port.json"
@{ title = 'Kit web test'; personnel = ''; port = $Port; apps = $apps } | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 $cfgFile

$server = $null; $egInserted = $null
try {
    $server = Start-Process powershell -PassThru -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
        "`"$(Join-Path $web 'kit\server\KitHost.ps1')`"", '-Config', "`"$cfgFile`"", '-DemoSignIn')
    $up = Wait-Server
    Check 'server starts' $up
    if (-not $up) { exit 1 }

    $page = Invoke-WebRequest "$base/" -UseDefaultCredentials -UseBasicParsing
    Check 'serves the UI' ($page.StatusCode -eq 200 -and $page.Content -match '<div id="root">') ("$($page.Content.Length) bytes")
    Check 'sends a CSP' ([bool]$page.Headers['Content-Security-Policy'])

    $me = Api GET '/api/me'
    Check 'Windows sign-in reaches the server' ($me.body.me.account -match '\\') $me.body.me.account
    Check 'lists the registered apps' (@($me.body.apps).Count -eq $apps.Count) ((@($me.body.apps) | ForEach-Object { $_.key }) -join ', ')
    $csrf = $me.body.csrf

    $people = Api GET '/api/people'
    Check 'lists demo people' ($people.body.people.Count -gt 0) ("$($people.body.people.Count) people")
    $who = [int]$people.body.people[0].PersonID
    $as = @{ 'X-Kit-As' = [string]$who }

    $meAs = Api GET '/api/me' $as
    $anchorEntry = @($meAs.body.apps) | Where-Object { $_.key -eq 'anchor' }
    Check 'ANCHOR profile comes from its own data file' ([bool]$anchorEntry.profile) "person $who"

    $sum = Api GET '/api/anchor/summary' $as
    Check 'summary' ($sum.status -eq 200 -and $sum.body.total -gt 0) ("total $($sum.body.total), unread $($sum.body.unread), ack due $($sum.body.ackDueCount)")

    $list = Api GET '/api/anchor/messages?q=cyber' $as
    Check 'search' ($list.status -eq 200 -and $list.body.messages.Count -gt 0) ("$($list.body.messages.Count) match 'cyber'")
    $odd = Api GET ('/api/anchor/messages?q=' + [uri]::EscapeDataString("x' OR 1=1 --")) $as
    Check 'search text is a value, not SQL' ($odd.status -eq 200 -and $odd.body.messages.Count -eq 0)

    $nope = Api GET '/api/nosuchapp/summary' $as
    Check 'unknown app is 404' ($nope.status -eq 404)
    $stranger = Api GET '/api/anchor/summary' @{ 'X-Kit-As' = '999999' }
    Check 'person without a profile is refused' ($stranger.status -eq 403) $stranger.raw

    $target = $sum.body.ackDue | Select-Object -First 1
    if (-not $target) { Write-Host 'No message waits on this person; skipping the acknowledgement steps.' }
    else {
        $d = Api GET "/api/anchor/messages/$($target.id)" $as
        Check 'message view' ($d.status -eq 200 -and $d.body.message.Reference -eq $target.reference) $target.reference
        Check 'message needs ack' ([bool]$d.body.ackDue)

        $noTok = Api POST "/api/anchor/messages/$($target.id)/ack" $as
        Check 'write without token refused' ($noTok.status -eq 403)

        $h = @{ 'X-Kit-As' = $as['X-Kit-As']; 'X-Kit-Csrf' = $csrf }
        $ack = Api POST "/api/anchor/messages/$($target.id)/ack" $h
        Check 'acknowledge' ($ack.status -eq 200) $ack.raw
        $again = Api POST "/api/anchor/messages/$($target.id)/ack" $h
        Check 'second acknowledge refused' ($again.status -eq 409)
        $d2 = Api GET "/api/anchor/messages/$($target.id)" $as
        Check 'ack recorded' ([bool]$d2.body.ackOn) $d2.body.ackOn
    }

    $audit = Api GET '/api/anchor/audit' $as
    Check 'audit chain verifies after web writes' ([bool]$audit.body.chain.ok) ("$($audit.body.chain.checked) entries $($audit.body.chain.problem)")
    $webRows = @($audit.body.rows | Where-Object { $_.Device -like '*(web)' })
    Check 'web actions are audited' ($webRows.Count -gt 0) ("$($webRows.Count) web entries")

    if ($ExamGenDataFile) {
        # Per-app permissions: the same person, no Exam Generator profile, is refused there.
        $c = Open-KitDb $ExamGenDataFile
        try {
            [void](Invoke-KitNonQuery $c 'DELETE FROM tblPeople WHERE PersonID=?' @($who))
        } finally { $c.Close() }
        $eg403 = Api GET '/api/examgen/overview' $as
        Check 'no Exam Generator profile: refused there' ($eg403.status -eq 403) $eg403.raw
        $sum2 = Api GET '/api/anchor/summary' $as
        Check '... while ANCHOR still answers' ($sum2.status -eq 200)

        $c = Open-KitDb $ExamGenDataFile
        try {
            [void](Invoke-KitNonQuery $c 'INSERT INTO tblPeople (PersonID, IsAdmin, IsActive, NotifyDashboard, NotifyEmail) VALUES (?, False, True, True, False)' @($who))
            $egInserted = $who
        } finally { $c.Close() }
        $eg = Api GET '/api/examgen/overview' $as
        Check 'with a profile: Exam Generator overview' ($eg.status -eq 200 -and $null -ne $eg.body.questionsByStatus) `
            ("$(@($eg.body.questionsByStatus).Count) statuses, $(@($eg.body.topics).Count) topics")
        $meEg = Api GET '/api/me' $as
        $egEntry = @($meEg.body.apps) | Where-Object { $_.key -eq 'examgen' }
        Check 'app list shows the Exam Generator profile' ([bool]$egEntry.profile)
    }
} finally {
    if ($server -and -not $server.HasExited) { Stop-Process -Id $server.Id -Force }
    if ($egInserted) {
        $c = Open-KitDb $ExamGenDataFile
        try { [void](Invoke-KitNonQuery $c 'DELETE FROM tblPeople WHERE PersonID=?' @($egInserted)) } finally { $c.Close() }
    }
    Remove-Item $cfgFile -ErrorAction SilentlyContinue
}
if ($fails) { Write-Host "$fails check(s) failed." -ForegroundColor Red; exit 1 }
Write-Host 'All checks passed.' -ForegroundColor Green
