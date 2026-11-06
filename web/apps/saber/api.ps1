<#
SABER's web module: its profile rule, its routes, and the business rules those
routes need, over SABER_be.accdb (the 25 tables the Access edition's
modSchemaDDL builds; schema version 8). Routes answer under /api/saber/.
Everything generic (sign-in, the data layer, security, HTTP) is the kit core,
web\kit\server\KitWeb.ps1.

The rules here are ported from the SABER Access edition's VBA
(github.com/stephen3303/SABER-Access, access\src\modules). Each function names
the routine it came from and keeps its checks, in the same order, so the two
can be compared line by line. The web is meant to become the only front end,
so this module is where those rules live once a page is ported.

Phase 1 (this file): sign-in and profile (modAuth, modTeam), the Inspections
list (pgDashboard), the inspection's grading tree (modRollup.BuildTree, without
the weighted rollups yet) and the grade write path (modGrading.SaveGrade and
ClearGrade) with its revision history (grade_history, LogGradeChange). See
apps\saber\PORT-PLAN.md for what follows.

Who is who: the kit maps the Windows account to a Personnel person_id; SABER's
own team_member row (team_member.person_id, schema 7) is the SABER profile and
carries SABER's rights. Every write records team_member.id as the actor, as
the Access edition does, so grade_history reads the same from either front end.

Access SQL notes: [name], [key], [value], [action] are bracketed (reserved
words); Yes/No columns are tested with <>0, which both engines accept.
#>

$SABER_APP_CODE = 'SABER'               # modTeam.APP_CODE, the kit_person_app code
$SABER_NE_LABEL = 'Not Evaluated'       # modGlobals.NE_LABEL
$SABER_STATUSES = @('active', 'configuring', 'closed', 'archived')   # pgDashboard status cards
$SABER_ACCESS_CACHE_SEC = 60            # how long the Personnel "may open SABER" list is reused

# ---------------------------------------------------------------- small helpers

# modGlobals.SystemSetting: an installation-wide setting, or the fallback.
function Get-SaberSetting($conn, [string]$key, [string]$fallback = '') {
    $rows = Invoke-KitQuery $conn 'SELECT [value] FROM system_setting WHERE [key]=?' @($key)
    if ($rows.Count -gt 0 -and $rows[0].value) { return [string]$rows[0].value }
    return $fallback
}

# modInspections.StatusLabel: "active" -> "Active".
function Get-SaberStatusLabel([string]$status) {
    if (-not $status) { return '' }
    return $status.Substring(0, 1).ToUpper() + $status.Substring(1)
}

# "Jane Q. Smith" -> "Smith" (modTeam.LastWord).
function Get-SaberLastWord([string]$s) {
    $s = $s.Trim()
    return $s.Substring($s.LastIndexOf(' ') + 1)
}

# A score from a request body: a whole number 0..100, $null for Not Evaluated.
# ok is $false (with why) for anything else: text, decimals (the Access edition
# rejects 85.5 rather than rounding it), out of range.
function ConvertTo-SaberScore($v) {
    if ($null -eq $v) { return @{ ok = $true; score = $null } }
    if (-not ($v -is [int] -or $v -is [long] -or $v -is [double] -or $v -is [decimal])) { return @{ ok = $false; error = 'Score must be a whole number from 0 to 100.' } }
    $d = [double]$v
    if ($d -ne [math]::Floor($d)) { return @{ ok = $false; error = 'Score must be a whole number from 0 to 100.' } }
    if ($d -lt 0 -or $d -gt 100) { return @{ ok = $false; error = 'Score must be between 0 and 100.' } }
    return @{ ok = $true; score = [int]$d }
}

# A WHERE fragment that matches a column's current value, Null included, for
# "update only if nobody changed it since I read it". Adds the parameter to
# the list it is given when the value isn't Null.
function Get-SaberSameValueSql([string]$column, $value, [System.Collections.Generic.List[object]]$params) {
    if ($null -eq $value) { return "$column Is Null" }
    $params.Add($value)
    return "$column=?"
}

# ---------------------------------------------------------------- Personnel and team_member (modTeam)

# Person ids that may open SABER (kit_person_app "SABER", both rows active;
# modTeam.RefreshSql's can_use), read from the Personnel file and kept for
# $SABER_ACCESS_CACHE_SEC so a revoked person loses SABER within a minute
# without the Personnel file being opened on every request. $null when no
# Personnel file is configured (demo sign-in): team_member.is_active decides.
$script:SaberAccess = @{ At = [datetime]::MinValue; Ids = $null; People = @{} }
function Get-SaberPersonnelAccess([switch]$Fresh) {
    if (-not $Kit.Personnel) { return $null }
    if (-not $Fresh -and ((Get-Date) - $script:SaberAccess.At).TotalSeconds -lt $SABER_ACCESS_CACHE_SEC) { return $script:SaberAccess }
    $c = Open-KitDb $Kit.Personnel (Get-KitPersonnelKey)
    try {
        $rows = Invoke-KitQuery $c ('SELECT p.person_id, p.username, p.first_name, p.last_name, p.rank, p.is_active, a.is_active AS app_active ' +
            'FROM kit_person AS p LEFT JOIN (SELECT person_id, is_active FROM kit_person_app WHERE app_code=?) AS a ON p.person_id=a.person_id') @($SABER_APP_CODE)
    } finally { $c.Close() }
    $ids = @{}; $people = @{}
    foreach ($p in $rows) {
        $people[[int]$p.person_id] = $p
        if ($p.is_active -and $p.app_active) { $ids[[int]$p.person_id] = $true }
    }
    $script:SaberAccess = @{ At = (Get-Date); Ids = $ids; People = $people }
    return $script:SaberAccess
}

# modTeam.EnsureTeamMember: the team_member id for a person, adding the row
# (no SABER rights) when they have none. A placeholder member with their
# sign-in name and last name (made by a package import) becomes theirs instead
# of a second row; a member whose name is someone else's is kept, and the new
# row gets "<username>.<person_id>".
function Add-SaberTeamMember($conn, $tx, $person) {
    $personId = [int]$person.person_id
    $has = Invoke-KitQuery $conn 'SELECT id FROM team_member WHERE person_id=?' @($personId) $tx
    if ($has.Count -gt 0) { return [int]$has[0].id }
    $username = ([string]$person.username).ToLower()
    $fullName = ("$($person.first_name) $($person.last_name)").Trim()
    $id = $null
    $placeholder = Invoke-KitQuery $conn 'SELECT id, full_name FROM team_member WHERE person_id Is Null AND username=?' @($username) $tx
    if ($placeholder.Count -gt 0) {
        if ((Get-SaberLastWord (ConvertTo-KitText $placeholder[0].full_name)) -eq [string]$person.last_name) { $id = [int]$placeholder[0].id }
        else { $username = "$username.$personId" }
    } elseif ((Invoke-KitQuery $conn 'SELECT id FROM team_member WHERE username=?' @($username) $tx).Count -gt 0) {
        $username = "$username.$personId"           # team_member.username is unique: keep both
    }
    if ($null -ne $id) {
        [void](Invoke-KitNonQuery $conn 'UPDATE team_member SET person_id=? WHERE id=? AND person_id Is Null' @($personId, $id) $tx)
        return $id
    }
    [void](Invoke-KitNonQuery $conn ('INSERT INTO team_member (username, full_name, rank, password_hash, is_active, is_staff, is_admin, ' +
        "can_manage_weights, can_manage_library, created_at, person_id) VALUES (?, ?, ?, '', True, False, False, False, False, Now(), ?)") `
        @((Limit-KitText $username 64), (Limit-KitText $fullName 128), (Limit-KitText (ConvertTo-KitText $person.rank) 64), $personId) $tx)
    return [int](Invoke-KitQuery $conn 'SELECT id FROM team_member WHERE person_id=?' @($personId) $tx)[0].id
}

# modTeam.SyncTeamWithPersonnel, run once at start (the Access edition runs it
# at every start-up and after the Personnel page closes):
#   1. a member with a SABER password and no person yet is linked to the
#      person with the same sign-in name and last name. The Access edition
#      also ADDS such members to the Personnel file; the web server never
#      writes another app's file, so they are only listed in the log until
#      the Personnel admin page is ported (phase 4) or an Access copy runs.
#   2. a person who may open SABER and has no team_member row gets one.
#   3. each linked member's copy of name, rank, sign-in name and active is
#      refreshed from the Personnel file.
# All in one transaction on SABER's file; the Personnel file is only read.
function Sync-SaberTeamWithPersonnel($conn) {
    $acc = Get-SaberPersonnelAccess -Fresh
    if (-not $acc) { return }
    $byUser = @{}
    foreach ($p in $acc.People.Values) { $byUser[([string]$p.username).ToLower()] = $p }
    $report = Use-KitTransaction $conn {
        param($tx)
        $linked = 0; $unlinked = New-Object System.Collections.Generic.List[string]; $added = 0; $refreshed = 0
        # 1
        $old = Invoke-KitQuery $conn "SELECT id, username, full_name FROM team_member WHERE person_id Is Null AND Len(password_hash & '')>0" @() $tx
        foreach ($m in $old) {
            $p = $byUser[([string]$m.username).ToLower()]
            if ($p -and (Get-SaberLastWord (ConvertTo-KitText $m.full_name)) -eq [string]$p.last_name -and
                (Invoke-KitQuery $conn 'SELECT id FROM team_member WHERE person_id=?' @([int]$p.person_id) $tx).Count -eq 0) {
                $linked += Invoke-KitNonQuery $conn 'UPDATE team_member SET person_id=? WHERE id=? AND person_id Is Null' @([int]$p.person_id, [int]$m.id) $tx
            } else { $unlinked.Add([string]$m.username) }
        }
        # 2
        $have = @{}
        foreach ($t in (Invoke-KitQuery $conn 'SELECT person_id FROM team_member WHERE person_id Is Not Null' @() $tx)) { $have[[int]$t.person_id] = $true }
        foreach ($personId in @($acc.Ids.Keys)) {
            if (-not $have.ContainsKey($personId)) { [void](Add-SaberTeamMember $conn $tx $acc.People[$personId]); $added++ }
        }
        # 3 (modTeam.RefreshRow: only what changed; a sign-in name another member has stays as it was)
        foreach ($t in (Invoke-KitQuery $conn 'SELECT id, person_id, username, full_name, rank, is_active FROM team_member WHERE person_id Is Not Null' @() $tx)) {
            $p = $acc.People[[int]$t.person_id]
            if (-not $p) { continue }                          # their person was deleted from the file
            $sets = New-Object System.Collections.Generic.List[string]; $vals = New-Object System.Collections.Generic.List[object]
            $fullName = ("$($p.first_name) $($p.last_name)").Trim()
            $active = $acc.Ids.ContainsKey([int]$t.person_id)
            if (-not [string]::Equals($fullName, (ConvertTo-KitText $t.full_name), [StringComparison]::Ordinal)) { $sets.Add('full_name=?'); $vals.Add((Limit-KitText $fullName 128)) }
            if (-not [string]::Equals((ConvertTo-KitText $p.rank), (ConvertTo-KitText $t.rank), [StringComparison]::Ordinal)) { $sets.Add('rank=?'); $vals.Add((ConvertTo-KitText $p.rank)) }
            if ($active -ne [bool]$t.is_active) { $sets.Add('is_active=' + $(if ($active) { 'True' } else { 'False' })) }
            $u = ([string]$p.username).ToLower()
            if ($u -ne (ConvertTo-KitText $t.username)) {
                if ((Invoke-KitQuery $conn 'SELECT id FROM team_member WHERE username=? AND id<>?' @($u, [int]$t.id) $tx).Count -eq 0) { $sets.Add('username=?'); $vals.Add($u) }
            }
            if ($sets.Count) {
                $vals.Add([int]$t.id)
                [void](Invoke-KitNonQuery $conn ('UPDATE team_member SET ' + ($sets -join ', ') + ' WHERE id=?') $vals.ToArray() $tx)
                $refreshed++
            }
        }
        return @{ linked = $linked; unlinked = $unlinked; added = $added; refreshed = $refreshed }
    }
    Write-Host ("SABER: Personnel sync linked {0}, added {1}, refreshed {2} member(s)." -f $report.linked, $report.added, $report.refreshed)
    if ($report.unlinked.Count) {
        Write-Warning ("SABER: {0} member(s) with a SABER password aren't in the Personnel file yet and can't sign in on the web until they are: {1}" -f $report.unlinked.Count, ($report.unlinked -join ', '))
    }
}

# ---------------------------------------------------------------- who may use SABER (modAuth)

# The SABER profile: the person's active team_member row, with SABER's own
# rights (modAuth.BeginSaberSession's TempVars). With a Personnel file the
# person must also be allowed to open SABER there; a person who is allowed
# and has no row yet gets one with no SABER rights, as signing in to the Access
# edition does (EnsureTeamMember).
function Get-SaberProfile($conn, [int]$personId) {
    $sql = 'SELECT t.id, t.username, t.full_name, t.rank, t.is_active, t.is_admin, t.can_manage_weights, t.can_manage_library, c.[name] AS commandName ' +
        'FROM team_member AS t LEFT JOIN command AS c ON t.command_id=c.id WHERE t.person_id=?'
    $acc = Get-SaberPersonnelAccess
    if ($acc -and -not $acc.Ids.ContainsKey($personId)) { return $null }
    $rows = Invoke-KitQuery $conn $sql @($personId)
    if ($rows.Count -eq 0 -and $acc) {
        [void](Use-KitTransaction $conn { param($tx) Add-SaberTeamMember $conn $tx $acc.People[$personId] })
        $rows = Invoke-KitQuery $conn $sql @($personId)
    }
    if ($rows.Count -eq 0) { return $null }
    $t = $rows[0]
    # With a Personnel file it decides (above); without one, the member row does.
    if (-not $acc -and -not $t.is_active) { return $null }
    $billet = if ($t.is_admin) { 'SABER admin' } elseif ($t.can_manage_weights) { 'Senior leadership' } else { 'Team member' }
    return @{
        memberId = [int]$t.id; username = $t.username; fullName = $t.full_name; rank = $t.rank
        isAdmin = [bool]$t.is_admin; canManageWeights = [bool]$t.can_manage_weights; canManageLibrary = [bool]$t.can_manage_library
        billet = $billet; unit = $t.commandName
    }
}
Set-KitProfileResolver { param($conn, [int]$personId) Get-SaberProfile $conn $personId }

Set-KitPeopleLister {
    param($conn)
    $rows = Invoke-KitQuery $conn ('SELECT t.person_id AS PersonID, t.is_admin, c.[name] AS commandName FROM team_member AS t ' +
        'LEFT JOIN command AS c ON t.command_id=c.id WHERE t.person_id Is Not Null AND t.is_active<>0 ORDER BY t.is_admin, t.full_name')
    foreach ($row in $rows) {
        $row['Billet'] = $(if ($row.is_admin) { 'SABER admin' } else { 'Team member' })
        $row['Unit'] = $row.commandName
    }
    return , $rows
}

# modAuth.UserStage: highest hierarchy_level the member holds in the
# inspection (0: no role). Decides who may grade.
function Get-SaberUserStage($conn, [int]$eventId, [int]$memberId, $tx = $null) {
    $r = Invoke-KitQuery $conn 'SELECT Max(hierarchy_level) AS stage FROM inspection_role WHERE inspection_event_id=? AND team_member_id=?' @($eventId, $memberId) $tx
    if ($r.Count -eq 0 -or $null -eq $r[0].stage) { return 0 }
    return [int]$r[0].stage
}

# modAuth.StageLabel: the label of a review stage from the inspection's role definitions.
function Get-SaberStageLabel($conn, [int]$eventId, [int]$stage) {
    if ($stage -eq 0) { return 'Author' }
    $r = Invoke-KitQuery $conn 'SELECT label FROM inspection_role_definition WHERE inspection_event_id=? AND hierarchy_level=? AND is_active<>0' @($eventId, $stage)
    if ($r.Count -gt 0 -and $r[0].label) { return [string]$r[0].label }
    return "Stage $stage"
}

# ---------------------------------------------------------------- Inspections page (pgDashboard)

# pgDashboard.ListSql over FilterWhere(False): every inspection matching the
# search text and year, newest first, with its grading completion (graded
# attributes over active ones, as on the Progress page; null with no
# attributes yet). The status card filter is applied afterwards, so the card
# counts (pgDashboard.CountCards) come from the same rows.
function Get-SaberInspectionRows($conn, [string]$search, [int]$year) {
    $where = New-Object System.Collections.Generic.List[string]; $params = New-Object System.Collections.Generic.List[object]
    if ($search) {
        $where.Add('(e.[name] LIKE ? OR e.inspected_command LIKE ? OR e.inspecting_command LIKE ?)')
        $pat = Get-KitLikePattern $search
        $params.Add($pat); $params.Add($pat); $params.Add($pat)
    }
    # Dates are optional, so an undated inspection counts in the year it was created.
    if ($year) {
        $where.Add('(Year(e.start_date)=? OR Year(e.end_date)=? OR (e.start_date Is Null AND e.end_date Is Null AND Year(e.created_at)=?))')
        $params.Add($year); $params.Add($year); $params.Add($year)
    }
    $sql = 'SELECT e.id, e.[name] AS name, e.inspected_command AS inspectedCommand, e.inspecting_command AS inspectingCommand, ' +
        'e.lifecycle_status AS status, e.boat_type AS boatType, e.start_date AS startDate, e.end_date AS endDate, e.created_at AS createdAt, ' +
        'a.n AS attributes, g.n AS graded FROM (inspection_event AS e ' +
        'LEFT JOIN (SELECT inspection_event_id, Count(*) AS n FROM grade GROUP BY inspection_event_id) AS g ON g.inspection_event_id=e.id) ' +
        'LEFT JOIN (SELECT inspection_event_id, Count(*) AS n FROM inspection_attribute WHERE is_active<>0 GROUP BY inspection_event_id) AS a ' +
        'ON a.inspection_event_id=e.id'
    if ($where.Count) { $sql += ' WHERE ' + ($where -join ' AND ') }
    $rows = Invoke-KitQuery $conn ($sql + ' ORDER BY e.created_at DESC') $params.ToArray()
    foreach ($row in $rows) {
        $n = 0; if ($null -ne $row.attributes) { $n = [int]$row.attributes }
        $g = 0; if ($null -ne $row.graded) { $g = [int]$row.graded }
        $row['attributes'] = $n; $row['graded'] = $g
        $row['pctDone'] = $(if ($n -eq 0) { $null } else { [int][math]::Floor($g * 100 / $n + 0.5) })
        $row['statusLabel'] = Get-SaberStatusLabel $row.status
    }
    return , $rows
}

# pgDashboard.LoadYears: every year an inspection starts, ends or was created
# in, plus this year, newest first.
function Get-SaberInspectionYears($conn) {
    $ys = @{ ([datetime]::Today.Year) = $true }
    foreach ($r in (Invoke-KitQuery $conn 'SELECT Year(start_date) AS y1, Year(end_date) AS y2, Year(created_at) AS y3 FROM inspection_event')) {
        foreach ($y in @($r.y1, $r.y2, $r.y3)) { if ($null -ne $y) { $ys[[int]$y] = $true } }
    }
    return , @($ys.Keys | Sort-Object -Descending)
}

Add-KitRoute GET '^/inspections$' {
    param($r)
    $year = 0
    if ($r.Query['year'] -match '^\d{4}$') { $year = [int]$r.Query['year'] }
    $status = [string]$r.Query['status']
    if ($status -notin $SABER_STATUSES) { $status = 'all' }
    $rows = Get-SaberInspectionRows $r.Conn ([string]$r.Query['q']) $year
    $counts = [ordered]@{ all = $rows.Count; active = 0; configuring = 0; closed = 0; archived = 0 }
    foreach ($row in $rows) { if ($counts.Contains([string]$row.status)) { $counts[[string]$row.status]++ } }
    $shown = @($rows | Where-Object { $status -eq 'all' -or $_.status -eq $status })
    $r.Json(200, @{ inspections = $shown; counts = $counts; years = (Get-SaberInspectionYears $r.Conn); status = $status
                    classification = (Get-SaberSetting $r.Conn 'classification_level' 'UNCLASSIFIED') })
}

# ---------------------------------------------------------------- inspection and grading tree (modRollup.BuildTree)

# The inspection's header and the full Mission Area > Group > Capability >
# Attribute tree (active rows, in display order, as BuildTree reads them),
# with each attribute's grade. Flat lists: the page nests them by parentId.
# The weighted rollups on every level (modRollup.Compute) are phase 2.
Add-KitRoute GET '^/inspections/(\d{1,9})$' {
    param($r)
    $id = [int]$r.Matches.Groups[1].Value
    $c = $r.Conn
    $ev = Invoke-KitQuery $c ('SELECT e.id, e.[name] AS name, e.inspected_command AS inspectedCommand, e.inspecting_command AS inspectingCommand, ' +
        'e.lifecycle_status AS status, e.boat_type AS boatType, e.start_date AS startDate, e.end_date AS endDate, e.closed_at AS closedAt, ' +
        'c.[name] AS conductingCommand FROM inspection_event AS e LEFT JOIN command AS c ON e.conducting_command_id=c.id WHERE e.id=?') @($id)
    if ($ev.Count -eq 0) { $r.Json(404, @{ error = 'No such inspection.' }); return }
    $insp = $ev[0]
    $insp['statusLabel'] = Get-SaberStatusLabel $insp.status
    $stage = Get-SaberUserStage $c $id $r.Me.profile.memberId
    $editable = ($insp.status -eq 'active')                    # modGlobals.EventIsEditable
    $canGrade = $editable -and ($stage -gt 0 -or $r.Me.profile.isAdmin)
    $why = ''
    if (-not $editable) { $why = 'Grading is only open while the inspection is Active.' }
    elseif (-not $canGrade) { $why = 'You have no role on this inspection.' }
    $attrs = Invoke-KitQuery $c ('SELECT a.id, a.capability_id AS parentId, a.title, a.criterion_text AS criterion, g.numerical_score AS score, ' +
        'g.categorical_level AS categorical, g.is_preliminary AS isPreliminary, g.updated_at AS gradedAt, t.full_name AS gradedBy ' +
        'FROM (inspection_attribute AS a LEFT JOIN grade AS g ON g.attribute_id=a.id) LEFT JOIN team_member AS t ON g.graded_by_id=t.id ' +
        'WHERE a.is_active<>0 AND a.inspection_event_id=? ORDER BY a.capability_id, a.display_order, a.title') @($id)
    $graded = @($attrs | Where-Object { $_.categorical }).Count
    $r.Json(200, @{
        inspection     = $insp
        classification = (Get-SaberSetting $c 'classification_level' 'UNCLASSIFIED')
        bands          = (Invoke-KitQuery $c 'SELECT id, label, min_score AS minScore, max_score AS maxScore FROM inspection_grade_band WHERE inspection_event_id=? ORDER BY min_score, display_order' @($id))
        missionAreas   = (Invoke-KitQuery $c 'SELECT id, title, description FROM inspection_mission_area WHERE is_active<>0 AND inspection_event_id=? ORDER BY display_order, title' @($id))
        groups         = (Invoke-KitQuery $c 'SELECT id, mission_area_id AS parentId, title FROM inspection_group WHERE is_active<>0 AND inspection_event_id=? ORDER BY mission_area_id, display_order, title' @($id))
        capabilities   = (Invoke-KitQuery $c 'SELECT id, group_id AS parentId, title, code FROM inspection_capability WHERE is_active<>0 AND inspection_event_id=? ORDER BY group_id, display_order, title' @($id))
        attributes     = $attrs
        progress       = @{ graded = $graded; total = $attrs.Count }
        you            = @{ stage = $stage; stageLabel = (Get-SaberStageLabel $c $id $stage); canGrade = $canGrade; whyNot = $why }
    })
}

# ---------------------------------------------------------------- grades (modGrading)

# The checks SaveGrade and ClearGrade share, in the Access edition's order,
# inside the write's transaction: the attribute belongs to the inspection, the
# inspection is Active (EventIsEditable), and the member has a role on it or
# is an admin. Returns @{ status; error } on refusal, else @{ capId }.
function Test-SaberCanGrade($conn, $tx, $me, [int]$eventId, [int]$attributeId) {
    $a = Invoke-KitQuery $conn 'SELECT capability_id FROM inspection_attribute WHERE id=? AND inspection_event_id=?' @($attributeId, $eventId) $tx
    if ($a.Count -eq 0) { return @{ status = 404; error = 'No such attribute on this inspection.' } }
    $st = Invoke-KitQuery $conn 'SELECT lifecycle_status FROM inspection_event WHERE id=?' @($eventId) $tx
    if ($st.Count -eq 0 -or $st[0].lifecycle_status -ne 'active') { return @{ status = 409; error = 'Grading is only open while the inspection is Active.' } }
    if ((Get-SaberUserStage $conn $eventId $me.profile.memberId $tx) -eq 0 -and -not $me.profile.isAdmin) { return @{ status = 403; error = 'You have no role on this inspection.' } }
    return @{ status = 0; capId = [int]$a[0].capability_id }
}

# modGrading.LogGradeChange: one grade_history row for the signed-in member.
# Rows are only ever added; this is SABER's record of who changed which grade.
function Add-SaberGradeHistory($conn, $tx, $me, [int]$eventId, [int]$attributeId, [string]$action, $oldScore, $oldCat, $newScore, $newCat) {
    $blank = { param($v) if ($null -eq $v -or [string]$v -eq '') { $null } else { $v } }   # BlankToNull
    [void](Invoke-KitNonQuery $conn ('INSERT INTO grade_history (inspection_event_id, kind, node_type, node_id, [action], old_score, old_categorical, ' +
        "new_score, new_categorical, actor_id, occurred_at) VALUES (?, 'grade', 'attribute', ?, ?, ?, ?, ?, ?, ?, Now())") `
        @($eventId, $attributeId, $action, $(if ($null -eq $oldScore) { $null } else { [double]$oldScore }), (& $blank $oldCat),
          $(if ($null -eq $newScore) { $null } else { [double]$newScore }), (& $blank $newCat), [int]$me.profile.memberId) $tx)
}

# modRollup.BandForScore: the band whose min <= score <= max, or $null.
function Get-SaberBandForScore($conn, $tx, [int]$eventId, [int]$score) {
    $b = Invoke-KitQuery $conn 'SELECT TOP 1 id, label FROM inspection_grade_band WHERE inspection_event_id=? AND min_score<=? AND max_score>=? ORDER BY min_score, display_order' @($eventId, $score, $score) $tx
    if ($b.Count -eq 0) { return $null }
    return $b[0]
}

# modGrading.BandMidpoint: the score a band pick stands for.
function Get-SaberBandMidpoint($conn, $tx, [int]$eventId, [string]$label) {
    $b = Invoke-KitQuery $conn 'SELECT min_score, max_score FROM inspection_grade_band WHERE inspection_event_id=? AND label=?' @($eventId, $label) $tx
    if ($b.Count -eq 0) { return $null }
    return [int][math]::Floor(([double]$b[0].min_score + [double]$b[0].max_score) / 2 + 0.5)
}

# modGrading.SaveGrade: upsert an attribute's grade. score $null marks it Not
# Evaluated. The grade is preliminary while any comment on its capability is
# still in review. A new grade is recorded as "graded", a changed one as
# "changed"; saving the same grade again records nothing. The grade and its
# history row are one transaction (the Access edition writes them one after
# the other). The update carries the old score and band in its WHERE, so if
# another PC changed the grade since it was read here, nothing is written and
# the request answers 409.
function Save-SaberGrade($conn, $tx, $me, [int]$eventId, [int]$attributeId, $score) {
    $ok = Test-SaberCanGrade $conn $tx $me $eventId $attributeId
    if ($ok.status) { return $ok }
    $bandId = $null
    if ($null -eq $score) { $cat = $SABER_NE_LABEL }
    else {
        $band = Get-SaberBandForScore $conn $tx $eventId $score
        if (-not $band) { return @{ status = 400; error = "No grade band covers a score of $score." } }
        $bandId = [int]$band.id; $cat = [string]$band.label
    }
    $open = Invoke-KitQuery $conn "SELECT Count(*) AS n FROM comment WHERE inspection_event_id=? AND capability_id=? AND workflow_status<>'final_approved'" @($eventId, $ok.capId) $tx
    $prelim = $(if ([int]$open[0].n -gt 0) { 'True' } else { 'False' })
    $memberId = [int]$me.profile.memberId
    $cur = Invoke-KitQuery $conn 'SELECT id, numerical_score, categorical_level FROM grade WHERE inspection_event_id=? AND attribute_id=?' @($eventId, $attributeId) $tx
    if ($cur.Count -eq 0) {
        # Two PCs grading the same attribute at once: the unique index on
        # (inspection_event_id, attribute_id) refuses the second insert, the
        # kit runs the request again, and it takes the update path below.
        [void](Invoke-KitNonQuery $conn ('INSERT INTO grade (inspection_event_id, attribute_id, numerical_score, categorical_level, grade_band_id, ' +
            "graded_by_id, graded_at, updated_at, is_preliminary) VALUES (?, ?, ?, ?, ?, ?, Now(), Now(), $prelim)") @($eventId, $attributeId, $score, $cat, $bandId, $memberId) $tx)
        Add-SaberGradeHistory $conn $tx $me $eventId $attributeId 'graded' $null $null $score $cat
        return @{ status = 200; ok = $true; action = 'graded'; score = $score; categorical = $cat }
    }
    $g = $cur[0]
    $params = New-Object System.Collections.Generic.List[object]
    foreach ($v in @($score, $cat, $bandId, $memberId)) { $params.Add($v) }
    $params.Add([int]$g.id)
    $same = (Get-SaberSameValueSql 'numerical_score' $g.numerical_score $params) + ' AND ' + (Get-SaberSameValueSql 'categorical_level' $g.categorical_level $params)
    $n = Invoke-KitNonQuery $conn ("UPDATE grade SET numerical_score=?, categorical_level=?, grade_band_id=?, graded_by_id=?, updated_at=Now(), is_preliminary=$prelim " +
        "WHERE id=? AND $same") $params.ToArray() $tx
    if ($n -ne 1) { return @{ status = 409; error = 'Someone changed this grade a moment ago. Reload to see it.' } }
    $unchanged = ((ConvertTo-KitText $g.numerical_score) -eq (ConvertTo-KitText $score)) -and ((ConvertTo-KitText $g.categorical_level) -eq $cat)
    if (-not $unchanged) { Add-SaberGradeHistory $conn $tx $me $eventId $attributeId 'changed' $g.numerical_score $g.categorical_level $score $cat }
    return @{ status = 200; ok = $true; action = $(if ($unchanged) { 'unchanged' } else { 'changed' }); score = $score; categorical = $cat }
}

# modGrading.ClearGrade: remove an attribute's grade (it becomes ungraded, not
# Not Evaluated), recording the grade it had. The Access edition returns
# silently when it can't; the web says why.
function Clear-SaberGrade($conn, $tx, $me, [int]$eventId, [int]$attributeId) {
    $ok = Test-SaberCanGrade $conn $tx $me $eventId $attributeId
    if ($ok.status) { return $ok }
    $cur = Invoke-KitQuery $conn 'SELECT id, numerical_score, categorical_level FROM grade WHERE inspection_event_id=? AND attribute_id=?' @($eventId, $attributeId) $tx
    if ($cur.Count -eq 0) { return @{ status = 200; ok = $true; action = 'unchanged' } }
    $g = $cur[0]
    $params = New-Object System.Collections.Generic.List[object]
    $params.Add([int]$g.id)
    $same = (Get-SaberSameValueSql 'numerical_score' $g.numerical_score $params) + ' AND ' + (Get-SaberSameValueSql 'categorical_level' $g.categorical_level $params)
    $n = Invoke-KitNonQuery $conn "DELETE FROM grade WHERE id=? AND $same" $params.ToArray() $tx
    if ($n -ne 1) { return @{ status = 409; error = 'Someone changed this grade a moment ago. Reload to see it.' } }
    Add-SaberGradeHistory $conn $tx $me $eventId $attributeId 'cleared' $g.numerical_score $g.categorical_level $null $null
    return @{ status = 200; ok = $true; action = 'cleared' }
}

# Grade an attribute. Body: { "score": 0-100 } a number, { "score": null } Not
# Evaluated, or { "band": "<label>" } a band pick, scored at the band's
# midpoint (subTreeGrading.PickBand).
Add-KitRoute PUT '^/inspections/(\d{1,9})/grades/(\d{1,9})$' {
    param($r)
    $eventId = [int]$r.Matches.Groups[1].Value; $attributeId = [int]$r.Matches.Groups[2].Value
    $b = $r.Body
    $hasBand = $b -and $b.PSObject.Properties['band']
    $hasScore = $b -and $b.PSObject.Properties['score']
    if (-not ($hasBand -xor $hasScore)) { $r.Json(400, @{ error = 'Send a score (a number, or null for Not Evaluated) or a band.' }); return }
    $score = $null
    if ($hasScore) {
        $s = ConvertTo-SaberScore $b.score
        if (-not $s.ok) { $r.Json(400, @{ error = $s.error }); return }
        $score = $s.score
    }
    $conn = $r.Conn; $me = $r.Me
    $result = Use-KitTransaction $conn {
        param($tx)
        if ($hasBand) {
            $label = [string]$b.band
            if ($label -eq $SABER_NE_LABEL) { $score = $null }
            else {
                $score = Get-SaberBandMidpoint $conn $tx $eventId $label
                if ($null -eq $score) { return @{ status = 400; error = "This inspection has no grade band called '$(Limit-KitText $label 64)'." } }
            }
        }
        $res = Save-SaberGrade $conn $tx $me $eventId $attributeId $score
        # A refusal (400/403/404/409) returns before anything is written, so
        # committing it commits nothing.
        return $res
    }
    $r.Json($result.status, $result)
}

Add-KitRoute DELETE '^/inspections/(\d{1,9})/grades/(\d{1,9})$' {
    param($r)
    $eventId = [int]$r.Matches.Groups[1].Value; $attributeId = [int]$r.Matches.Groups[2].Value
    $conn = $r.Conn; $me = $r.Me
    $result = Use-KitTransaction $conn {
        param($tx)
        return (Clear-SaberGrade $conn $tx $me $eventId $attributeId)
    }
    $r.Json($result.status, $result)
}

# dlgGradeHistory for one attribute: every change to its grade, newest first,
# with who made it.
Add-KitRoute GET '^/inspections/(\d{1,9})/attributes/(\d{1,9})/history$' {
    param($r)
    $eventId = [int]$r.Matches.Groups[1].Value; $attributeId = [int]$r.Matches.Groups[2].Value
    $a = Invoke-KitQuery $r.Conn 'SELECT id FROM inspection_attribute WHERE id=? AND inspection_event_id=?' @($attributeId, $eventId)
    if ($a.Count -eq 0) { $r.Json(404, @{ error = 'No such attribute on this inspection.' }); return }
    $rows = Invoke-KitQuery $r.Conn ('SELECT h.id, h.[action] AS action, h.old_score AS oldScore, h.old_categorical AS oldCategorical, ' +
        'h.new_score AS newScore, h.new_categorical AS newCategorical, h.occurred_at AS occurredAt, t.full_name AS actor, t.rank AS actorRank ' +
        "FROM grade_history AS h LEFT JOIN team_member AS t ON h.actor_id=t.id WHERE h.inspection_event_id=? AND h.kind='grade' " +
        "AND h.node_type='attribute' AND h.node_id=? ORDER BY h.occurred_at DESC, h.id DESC") @($eventId, $attributeId)
    $r.Json(200, @{ history = $rows })
}

# ---------------------------------------------------------------- start-up

# Without a Personnel file (demo sign-in) the kit has no names, so the demo
# picker and the top bar would say "Person 12": use SABER's own copy of each
# member's name. With a Personnel file the kit reads real names after this,
# and the Personnel sync brings team_member into line first.
Set-KitOnStart {
    param($conn)
    if ($Kit.Personnel) { Sync-SaberTeamWithPersonnel $conn; return }
    foreach ($t in (Invoke-KitQuery $conn 'SELECT person_id, full_name, rank FROM team_member WHERE person_id Is Not Null')) {
        $Kit.Names[[int]$t.person_id] = ("$($t.rank) $($t.full_name)").Trim()
    }
}
