<#
End-to-end check of the web server against a real data file. Starts
AnchorWeb.ps1 -DemoSignIn in its own window, drives the API as the signed-in
Windows user, and prints PASS/FAIL per step.

  web\server\test-web.ps1 -DataFile access\build\ANCHOR_BE.accdb

The data file must be seeded with people (backend\seed-backend.ps1 -Personnel).
It writes to the data file: one read and one acknowledgement, with their
audit rows. Use a scratch copy.
#>
param([Parameter(Mandatory = $true)][string]$DataFile, [int]$Port = 8787)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$DataFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DataFile)
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
        return @{ status = [int]$resp.StatusCode; body = ($txt | ConvertFrom-Json); raw = $txt }
    }
}

$server = Start-Process powershell -PassThru -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
    (Join-Path $here 'AnchorWeb.ps1'), '-DataFile', "`"$DataFile`"", '-Port', $Port, '-DemoSignIn')
try {
    $up = $false
    for ($i = 0; $i -lt 40 -and -not $up; $i++) {
        Start-Sleep -Milliseconds 500
        try { [void](Invoke-WebRequest "$base/api/me" -UseDefaultCredentials -UseBasicParsing); $up = $true } catch { }
    }
    Check 'server starts' $up
    if (-not $up) { exit 1 }

    $page = Invoke-WebRequest "$base/" -UseDefaultCredentials -UseBasicParsing
    Check 'serves the UI' ($page.StatusCode -eq 200 -and $page.Content -match '<div id="root">') ("$($page.Content.Length) bytes")
    Check 'sends a CSP' ([bool]$page.Headers['Content-Security-Policy'])

    $me = Api GET '/api/me'
    Check 'Windows sign-in reaches the server' ($me.body.me.account -match '\\') $me.body.me.account
    $csrf = $me.body.csrf

    $people = Api GET '/api/people'
    Check 'lists demo people' ($people.body.people.Count -gt 0) ("$($people.body.people.Count) people")
    $as = @{ 'X-Anchor-As' = [string]$people.body.people[0].PersonID }

    $sum = Api GET '/api/summary' $as
    Check 'summary' ($sum.status -eq 200 -and $sum.body.total -gt 0) ("total $($sum.body.total), unread $($sum.body.unread), ack due $($sum.body.ackDueCount)")

    $list = Api GET '/api/messages?q=cyber' $as
    Check 'search' ($list.status -eq 200 -and $list.body.messages.Count -gt 0) ("$($list.body.messages.Count) match 'cyber'")
    $odd = Api GET ('/api/messages?q=' + [uri]::EscapeDataString("x' OR 1=1 --")) $as
    Check 'search text is a value, not SQL' ($odd.status -eq 200 -and $odd.body.messages.Count -eq 0)

    $target = $sum.body.ackDue | Select-Object -First 1
    if (-not $target) { Write-Host 'No message waits on this person; skipping the acknowledgement steps.' }
    else {
        $d = Api GET "/api/messages/$($target.id)" $as
        Check 'message view' ($d.status -eq 200 -and $d.body.message.Reference -eq $target.reference) $target.reference
        Check 'message needs ack' ([bool]$d.body.ackDue)

        $noTok = Api POST "/api/messages/$($target.id)/ack" $as
        Check 'write without token refused' ($noTok.status -eq 403)

        $h = @{ 'X-Anchor-As' = $as['X-Anchor-As']; 'X-Anchor-Csrf' = $csrf }
        $ack = Api POST "/api/messages/$($target.id)/ack" $h
        Check 'acknowledge' ($ack.status -eq 200) $ack.raw
        $again = Api POST "/api/messages/$($target.id)/ack" $h
        Check 'second acknowledge refused' ($again.status -eq 409)
        $d2 = Api GET "/api/messages/$($target.id)" $as
        Check 'ack recorded' ([bool]$d2.body.ackOn) $d2.body.ackOn
    }

    $audit = Api GET '/api/audit' $as
    Check 'audit chain verifies after web writes' ([bool]$audit.body.chain.ok) ("$($audit.body.chain.checked) entries $($audit.body.chain.problem)")
    $webRows = @($audit.body.rows | Where-Object { $_.Device -like '*(web)' })
    Check 'web actions are audited' ($webRows.Count -gt 0) ("$($webRows.Count) web entries")
} finally {
    if ($server -and -not $server.HasExited) { Stop-Process -Id $server.Id -Force }
}
if ($fails) { Write-Host "$fails check(s) failed." -ForegroundColor Red; exit 1 }
Write-Host 'All checks passed.' -ForegroundColor Green
