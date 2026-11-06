<#
End-to-end check of the SABER web module (web\apps\saber\api.ps1) against a
real SABER data file. Starts kit\server\KitHost.ps1 -DemoSignIn with SABER
alone, drives the API, and prints PASS/FAIL per step. Same shape as
test-web.ps1.

  web\test\test-saber.ps1 -DataFile <scratch>\SABER_be.accdb

Make the data file with the SABER Access edition's own build (fictional seed
data: five test inspections, admin and the eight tycom_* review roles):

  git clone https://github.com/stephen3303/SABER-Access <scratch>\SABER-Access
  powershell -ExecutionPolicy Bypass -File <scratch>\SABER-Access\access\build.ps1
  copy <scratch>\SABER-Access\SABER_be.accdb <scratch>\SABER_be.accdb

USE A SCRATCH COPY. The test writes to it:
  - team_member.person_id for members that have none (900000 + id), because a
    freshly built file has never met a Personnel file and demo sign-in needs a
    person id. They are put back to Null at the end.
  - grades and grade_history rows on one attribute of an active inspection
    (left in place: grade_history is add-only by design).
#>
param(
    [Parameter(Mandatory = $true)][string]$DataFile,
    [int]$Port = 8791
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$web = Split-Path -Parent $here
. (Join-Path $web 'kit\server\KitWeb.ps1')   # Open-KitDb / Invoke-KitQuery for set-up and checks
function Get-FullPath([string]$p) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) }
$DataFile = Get-FullPath $DataFile
$Kit.Provider = Find-KitAceProvider $DataFile
$fails = 0
function Check([string]$name, [bool]$ok, [string]$detail = '') {
    if ($ok) { Write-Host "PASS  $name $detail" -ForegroundColor Green }
    else { Write-Host "FAIL  $name $detail" -ForegroundColor Red; $script:fails++ }
}
function Api([string]$method, [string]$path, $headers = @{}, $body = $null, [int]$onPort = $Port) {
    $args2 = @{ Uri = "http://localhost:$onPort$path"; Method = $method; Headers = $headers; UseDefaultCredentials = $true; UseBasicParsing = $true }
    if ($null -ne $body) { $args2.Body = (ConvertTo-Json -InputObject $body -Compress); $args2.ContentType = 'application/json' }
    try {
        $r = Invoke-WebRequest @args2
        return @{ status = [int]$r.StatusCode; body = ($r.Content | ConvertFrom-Json); raw = $r.Content }
    } catch [System.Net.WebException] {
        $resp = $_.Exception.Response
        if (-not $resp) { throw }
        $txt = ''
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $txt = $_.ErrorDetails.Message }
        else {
            $st = $resp.GetResponseStream()
            try { if ($st.CanSeek) { $st.Position = 0 } } catch { }
            $txt = (New-Object System.IO.StreamReader($st)).ReadToEnd()
        }
        $b = $null; try { $b = $txt | ConvertFrom-Json } catch { }
        return @{ status = [int]$resp.StatusCode; body = $b; raw = $txt }
    }
}
function Wait-Server([int]$p) {
    for ($i = 0; $i -lt 40; $i++) {
        Start-Sleep -Milliseconds 500
        try { [void](Invoke-WebRequest "http://localhost:$p/api/me" -UseDefaultCredentials -UseBasicParsing); return $true } catch { }
    }
    return $false
}
function Db([scriptblock]$Block) { $c = Open-KitDb $DataFile; try { & $Block $c } finally { $c.Close() } }
function Start-Server([int]$p, [string]$cfg) {
    return Start-Process powershell -PassThru -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
        "`"$(Join-Path $web 'kit\server\KitHost.ps1')`"", '-Config', "`"$cfg`"", '-DemoSignIn', '-Port', $p)
}
function History([int]$eventId, [int]$attrId) {
    Db { param($c) Invoke-KitQuery $c "SELECT id, [action], old_score, new_score, new_categorical, actor_id FROM grade_history WHERE inspection_event_id=? AND kind='grade' AND node_id=? ORDER BY id" @($eventId, $attrId) }
}

# ---- set-up on the scratch file: person ids, and who to act as
$linked = Db { param($c)
    $ids = @((Invoke-KitQuery $c 'SELECT id FROM team_member WHERE person_id Is Null') | ForEach-Object { [int]$_.id })
    foreach ($id in $ids) { [void](Invoke-KitNonQuery $c 'UPDATE team_member SET person_id=? WHERE id=?' @((900000 + $id), $id)) }
    , $ids
}
$plan = Db { param($c)
    $ev = Invoke-KitQuery $c ("SELECT TOP 1 e.id FROM inspection_event AS e WHERE e.lifecycle_status='active' AND " +
        'EXISTS (SELECT 1 FROM inspection_attribute AS a WHERE a.inspection_event_id=e.id AND a.is_active<>0) ORDER BY e.id')
    $closed = Invoke-KitQuery $c ("SELECT TOP 1 e.id FROM inspection_event AS e WHERE e.lifecycle_status IN ('closed','archived') AND " +
        'EXISTS (SELECT 1 FROM inspection_attribute AS a WHERE a.inspection_event_id=e.id AND a.is_active<>0) ORDER BY e.id')
    $p = @{ eventId = $null; closedId = $null }
    if ($ev.Count) {
        $p.eventId = [int]$ev[0].id
        $grader = Invoke-KitQuery $c ('SELECT TOP 1 t.id, t.person_id FROM team_member AS t INNER JOIN inspection_role AS r ON r.team_member_id=t.id ' +
            'WHERE r.inspection_event_id=? AND t.is_active<>0 AND t.is_admin=0 ORDER BY r.hierarchy_level DESC, t.id') @($p.eventId)
        if ($grader.Count) { $p.graderMember = [int]$grader[0].id; $p.graderPerson = [int]$grader[0].person_id }
        $none = Invoke-KitQuery $c ('SELECT TOP 1 t.id, t.person_id FROM team_member AS t WHERE t.is_active<>0 AND t.is_admin=0 AND ' +
            't.id NOT IN (SELECT team_member_id FROM inspection_role WHERE inspection_event_id=?) ORDER BY t.id') @($p.eventId)
        if ($none.Count) { $p.noRolePerson = [int]$none[0].person_id }
        $attrs = Invoke-KitQuery $c 'SELECT TOP 2 id FROM inspection_attribute WHERE inspection_event_id=? AND is_active<>0 ORDER BY id DESC' @($p.eventId)
        $p.attrId = [int]$attrs[0].id; $p.raceAttrId = [int]$attrs[1].id
        $b = Invoke-KitQuery $c 'SELECT TOP 1 label, min_score, max_score FROM inspection_grade_band WHERE inspection_event_id=? ORDER BY min_score DESC' @($p.eventId)
        $p.band = [string]$b[0].label; $p.bandMid = [int][math]::Floor(([double]$b[0].min_score + [double]$b[0].max_score) / 2 + 0.5)
    }
    if ($closed.Count) {
        $p.closedId = [int]$closed[0].id
        $p.closedAttr = [int](Invoke-KitQuery $c 'SELECT TOP 1 id FROM inspection_attribute WHERE inspection_event_id=? AND is_active<>0' @($p.closedId))[0].id
    }
    $p
}

$cfgFile = Join-Path $env:TEMP "kit-saber-test-$Port.json"
@{ title = 'SABER web test'; personnel = ''; port = $Port
   apps = @(@{ key = 'saber'; name = 'SABER'; dataFile = $DataFile; module = (Join-Path $web 'apps\saber\api.ps1') }) } | ConvertTo-Json -Depth 4 | Set-Content -Encoding UTF8 $cfgFile
$servers = @()
try {
    $servers += Start-Server $Port $cfgFile
    $up = Wait-Server $Port
    Check 'server starts with the SABER module' $up
    if (-not $up) { exit 1 }
    $page = Invoke-WebRequest "http://localhost:$Port/" -UseDefaultCredentials -UseBasicParsing
    Check 'serves the UI' ($page.StatusCode -eq 200 -and $page.Content -match 'saber')

    $me = Api GET '/api/me'
    Check 'lists SABER' ((@($me.body.apps) | Where-Object { $_.key -eq 'saber' -and $_.available }).Count -eq 1)
    $csrf = $me.body.csrf
    $people = Api GET '/api/people'
    Check 'demo people come from team_member' (@($people.body.people).Count -gt 0) ("$(@($people.body.people).Count) people, e.g. $($people.body.people[0].Name)")
    Check 'names are SABER''s own, not "Person n"' ($people.body.people[0].Name -notmatch '^Person \d+$')

    if (-not $plan.eventId -or -not $plan.graderPerson) { Check 'test data has an active inspection with a reviewer' $false; exit 1 }
    $as = @{ 'X-Kit-As' = [string]$plan.graderPerson }
    $w = @{ 'X-Kit-As' = [string]$plan.graderPerson; 'X-Kit-Csrf' = $csrf }

    $meAs = Api GET '/api/me' $as
    $sab = @($meAs.body.apps) | Where-Object { $_.key -eq 'saber' }
    Check 'SABER profile from team_member' ($sab.profile.memberId -eq $plan.graderMember) "member $($sab.profile.memberId), $($sab.profile.billet)"

    # ---- Inspections list
    $all = Api GET '/api/saber/inspections' $as
    Check 'inspections list' ($all.status -eq 200 -and @($all.body.inspections).Count -ge 1) ("$(@($all.body.inspections).Count) shown, counts $(($all.body.counts | ConvertTo-Json -Compress))")
    $sum = 0; foreach ($k in 'active', 'configuring', 'closed', 'archived') { $sum += [int]$all.body.counts.$k }
    Check 'status counts add up to All' ($sum -eq [int]$all.body.counts.all)
    $act = Api GET '/api/saber/inspections?status=active' $as
    Check 'status filter' ((@($act.body.inspections) | Where-Object { $_.status -ne 'active' }).Count -eq 0 -and @($act.body.inspections).Count -eq [int]$all.body.counts.active)
    $odd = Api GET ('/api/saber/inspections?q=' + [uri]::EscapeDataString("x' OR 1=1 --")) $as
    Check 'search text is a value, not SQL' ($odd.status -eq 200 -and @($odd.body.inspections).Count -eq 0)
    $pct = @($all.body.inspections) | Where-Object { $_.id -eq $plan.eventId }
    Check 'grading completion shown' ($null -ne $pct -and $null -ne $pct.pctDone) "$($pct.pctDone)%"
    $yr = Api GET '/api/saber/inspections?year=1999' $as
    Check 'year filter' ($yr.status -eq 200 -and @($yr.body.inspections).Count -eq 0)

    # ---- Inspection and tree
    $d = Api GET "/api/saber/inspections/$($plan.eventId)" $as
    Check 'inspection with its tree' ($d.status -eq 200 -and @($d.body.attributes).Count -gt 0) ("$(@($d.body.missionAreas).Count) areas, $(@($d.body.capabilities).Count) capabilities, $(@($d.body.attributes).Count) attributes")
    Check 'reviewer may grade' ([bool]$d.body.you.canGrade) $d.body.you.stageLabel
    Check 'unknown inspection is 404' ((Api GET '/api/saber/inspections/999999' $as).status -eq 404)
    $stranger = Api GET '/api/saber/inspections' @{ 'X-Kit-As' = '999999' }
    Check 'person without a SABER profile is refused' ($stranger.status -eq 403)

    # ---- grade write path
    $e = $plan.eventId; $a = $plan.attrId
    [void](Api DELETE "/api/saber/inspections/$e/grades/$a" $w)                 # start ungraded
    $h0 = @(History $e $a).Count
    Check 'write without token refused' ((Api PUT "/api/saber/inspections/$e/grades/$a" $as @{ score = 85 }).status -eq 403)
    $g1 = Api PUT "/api/saber/inspections/$e/grades/$a" $w @{ score = 85 }
    Check 'grade 85' ($g1.status -eq 200 -and $g1.body.action -eq 'graded') "$($g1.body.categorical)"
    $g2 = Api PUT "/api/saber/inspections/$e/grades/$a" $w @{ score = 85 }
    Check 'same grade again records nothing' ($g2.status -eq 200 -and $g2.body.action -eq 'unchanged')
    $g3 = Api PUT "/api/saber/inspections/$e/grades/$a" $w @{ band = $plan.band }
    Check 'band pick scores at the midpoint' ($g3.status -eq 200 -and $g3.body.score -eq $plan.bandMid) "$($plan.band) = $($g3.body.score)"
    $g4 = Api PUT "/api/saber/inspections/$e/grades/$a" $w @{ score = $null }
    Check 'Not Evaluated' ($g4.status -eq 200 -and $g4.body.categorical -eq 'Not Evaluated')
    Check '85.5 refused' ((Api PUT "/api/saber/inspections/$e/grades/$a" $w @{ score = 85.5 }).status -eq 400)
    Check '101 refused' ((Api PUT "/api/saber/inspections/$e/grades/$a" $w @{ score = 101 }).status -eq 400)
    Check 'text refused' ((Api PUT "/api/saber/inspections/$e/grades/$a" $w @{ score = '85' }).status -eq 400)
    Check 'unknown band refused' ((Api PUT "/api/saber/inspections/$e/grades/$a" $w @{ band = 'No such band' }).status -eq 400)
    $g5 = Api DELETE "/api/saber/inspections/$e/grades/$a" $w
    Check 'clear' ($g5.status -eq 200 -and $g5.body.action -eq 'cleared')
    $hist = @(History $e $a)
    $new = @($hist | Select-Object -Skip $h0)
    Check 'grade_history: graded, changed, changed, cleared' ((($new | ForEach-Object { $_.action }) -join ',') -eq 'graded,changed,changed,cleared') (($new | ForEach-Object { $_.action }) -join ',')
    Check 'history actor is the team_member id' ((@($new | Where-Object { $_.actor_id -ne $plan.graderMember })).Count -eq 0)
    $hApi = Api GET "/api/saber/inspections/$e/attributes/$a/history" $as
    Check 'history route' ($hApi.status -eq 200 -and @($hApi.body.history).Count -eq $hist.Count) "$(@($hApi.body.history).Count) rows"

    # ---- who may grade
    if ($plan.noRolePerson) {
        $nr = Api PUT "/api/saber/inspections/$e/grades/$a" @{ 'X-Kit-As' = [string]$plan.noRolePerson; 'X-Kit-Csrf' = $csrf } @{ score = 70 }
        Check 'member with no role on the inspection refused' ($nr.status -eq 403) $nr.raw
    }
    if ($plan.closedId) {
        $cl = Api PUT "/api/saber/inspections/$($plan.closedId)/grades/$($plan.closedAttr)" $w @{ score = 70 }
        Check 'closed inspection refused' ($cl.status -eq 409) $cl.raw
    }
    Check 'attribute of another inspection refused' ((Api PUT "/api/saber/inspections/$e/grades/999999" $w @{ score = 70 }).status -eq 404)

    # ---- two servers on one file grade the same ungraded attribute at once
    $r = $plan.raceAttrId
    [void](Api DELETE "/api/saber/inspections/$e/grades/$r" $w)
    $hr0 = @(History $e $r).Count
    $servers += Start-Server ($Port + 1) $cfgFile
    if (Wait-Server ($Port + 1)) {
        $csrf2 = (Api GET '/api/me' @{} $null ($Port + 1)).body.csrf
        $job = {
            param($url, $person, $token, $score)
            try {
                $x = Invoke-WebRequest -Uri $url -Method PUT -UseDefaultCredentials -UseBasicParsing -ContentType 'application/json' `
                    -Headers @{ 'X-Kit-As' = $person; 'X-Kit-Csrf' = $token } -Body "{`"score`":$score}"
                [int]$x.StatusCode
            } catch { if ($_.Exception.Response) { [int]$_.Exception.Response.StatusCode } else { 0 } }
        }
        $j1 = Start-Job $job -ArgumentList "http://localhost:$Port/api/saber/inspections/$e/grades/$r", [string]$plan.graderPerson, $csrf, 70
        $j2 = Start-Job $job -ArgumentList "http://localhost:$($Port + 1)/api/saber/inspections/$e/grades/$r", [string]$plan.graderPerson, $csrf2, 80
        $codes = @(Receive-Job $j1, $j2 -Wait -AutoRemoveJob)
        Check 'race: both servers answer' ((@($codes | Where-Object { $_ -eq 200 })).Count -eq 2) ($codes -join ', ')
        $rows = Db { param($c) Invoke-KitQuery $c 'SELECT Count(*) AS n FROM grade WHERE inspection_event_id=? AND attribute_id=?' @($e, $r) }
        Check 'race: one grade row' ([int]$rows[0].n -eq 1)
        $hr = @(History $e $r | Select-Object -Skip $hr0)
        Check 'race: history is graded then changed' ((($hr | ForEach-Object { $_.action }) -join ',') -eq 'graded,changed') (($hr | ForEach-Object { "$($_.action) $($_.new_score)" }) -join '; ')
    } else { Check 'second server starts' $false }
} finally {
    foreach ($s in $servers) { if ($s -and -not $s.HasExited) { Stop-Process -Id $s.Id -Force } }
    Db { param($c) foreach ($id in $linked) { [void](Invoke-KitNonQuery $c 'UPDATE team_member SET person_id=Null WHERE id=?' @($id)) } }
    Remove-Item $cfgFile -ErrorAction SilentlyContinue
}
if ($fails) { Write-Host "$fails check(s) failed." -ForegroundColor Red; exit 1 }
Write-Host 'All checks passed.' -ForegroundColor Green
