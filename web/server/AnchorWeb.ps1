<#
ANCHOR web front end: a small JSON server over the ANCHOR Access data file.

A browser can't open an .accdb, so this script sits between them. It runs on
a Windows machine that has Microsoft Access (or the free Access Database
Engine), opens the data file with the Access engine (ACE OLE DB) for each
request, and serves:
  - the web UI: one HTML file, web\dist\index.html (built from web\ui)
  - a JSON API under /api/ that the UI calls

The .accdb stays the one source of truth. The Access front end and this
server work on the same tables at the same time; the Access engine's own
record locking (the .laccdb file) keeps them apart, exactly as it does for
two Access front ends.

Every write follows the same rules as the Access front end (modMessages,
modAudit): an acknowledgement and its audit row are one transaction, and
audit rows are sealed onto the hash chain the same way, so Verify on the
Access Audit page still passes after web activity.

Sign-in: Windows (Negotiate/NTLM, built into Windows). The browser sends
the person's Windows account; nobody types a password. The account is looked
up in the shared Personnel directory (kit_person.windows_account) to get the
person_id that ANCHOR's tables use, exactly like the Access sign-in.

  web\server\AnchorWeb.ps1 -DataFile build\ANCHOR_BE.accdb -Personnel S:\Apps\Personnel\Personnel_be.accdb
  web\server\AnchorWeb.ps1 -DataFile build\ANCHOR_BE.accdb -DemoSignIn     (review only: pick a person)

  -Port 8787        the port (http://localhost:8787/)
  -Lan              also listen on this machine's name, for other PCs on the
                    network. Needs a one-time URL reservation by an admin:
                      netsh http add urlacl url=http://+:8787/ user=<DOMAIN\account>
                    and a firewall rule. Off by default: localhost only.
  -DemoSignIn       no Personnel file: the UI shows a "Sign in as" list of
                    the people in tblPeople. For demos and review only.

Needs: Windows PowerShell 5.1 (built in) and the Access engine of the same
bitness as the PowerShell that runs it (64-bit Office: the normal
powershell.exe; 32-bit Office: C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe).
#>
param(
    [Parameter(Mandatory = $true)][string]$DataFile,
    [string]$Personnel = '',
    [int]$Port = 8787,
    [switch]$Lan,
    [switch]$DemoSignIn,
    [string]$UiFile = ''
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Data

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$DataFile = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($DataFile)
if (-not (Test-Path $DataFile)) { throw "No data file at $DataFile" }
if (-not $UiFile) { $UiFile = Join-Path (Split-Path -Parent $here) 'dist\index.html' }
if ($Personnel) {
    $Personnel = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Personnel)
    if (-not (Test-Path $Personnel)) { throw "No Personnel file at $Personnel" }
}
if (-not $Personnel -and -not $DemoSignIn) {
    throw "Give -Personnel <Personnel_be.accdb> (Windows sign-in) or -DemoSignIn (review only)."
}

$GENESIS = 'ANCHOR-AUDIT-GENESIS'          # modAudit.GENESIS
$DEFAULT_PERSONNEL_KEY = 'kit-personnel-1' # modPersonnel.DEFAULT_KEY
$CsrfToken = [guid]::NewGuid().ToString('N')
$Device = ($env:COMPUTERNAME + ' (web)')
if ($Device.Length -gt 64) { $Device = $Device.Substring(0, 64) }

# ---------------------------------------------------------------- database

# The newest Access engine this PowerShell can load.
function Find-AceProvider {
    foreach ($p in 'Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0') {
        $c = New-Object System.Data.OleDb.OleDbConnection("Provider=$p;Data Source=$DataFile;Mode=Share Deny None")
        try { $c.Open(); $c.Close(); return $p } catch { } finally { $c.Dispose() }
    }
    $bits = if ([Environment]::Is64BitProcess) { '64-bit' } else { '32-bit' }
    throw "This $bits PowerShell can't load the Access engine (ACE OLE DB). Run the PowerShell whose bitness matches Office, or install the Microsoft Access Database Engine redistributable of this bitness."
}
$Provider = Find-AceProvider

function Open-Db([string]$path, [string]$password = '') {
    $cs = "Provider=$Provider;Data Source=$path;Mode=Share Deny None"
    if ($password) { $cs += ";Jet OLEDB:Database Password=$password" }
    $c = New-Object System.Data.OleDb.OleDbConnection($cs)
    $c.Open()
    return $c
}

# Run a SELECT. Parameters are positional (?), always passed as values, never
# pasted into the SQL. Returns a list of ordered hashtables; dates as ISO text.
function Invoke-Query($conn, [string]$sql, [object[]]$params = @(), $tx = $null) {
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sql
    if ($tx) { $cmd.Transaction = $tx }
    foreach ($p in $params) { [void]$cmd.Parameters.AddWithValue('?', $(if ($null -eq $p) { [DBNull]::Value } else { $p })) }
    $rows = New-Object System.Collections.Generic.List[object]
    $r = $cmd.ExecuteReader()
    try {
        while ($r.Read()) {
            $row = [ordered]@{}
            for ($i = 0; $i -lt $r.FieldCount; $i++) {
                $v = $r.GetValue($i)
                if ($v -is [DBNull]) { $v = $null }
                elseif ($v -is [datetime]) { $v = $v.ToString('yyyy-MM-ddTHH:mm:ss') }
                $row[$r.GetName($i)] = $v
            }
            $rows.Add($row)
        }
    } finally { $r.Close() }
    return , $rows
}

function Invoke-NonQuery($conn, [string]$sql, [object[]]$params = @(), $tx = $null) {
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sql
    if ($tx) { $cmd.Transaction = $tx }
    foreach ($p in $params) { [void]$cmd.Parameters.AddWithValue('?', $(if ($null -eq $p) { [DBNull]::Value } else { $p })) }
    return $cmd.ExecuteNonQuery()
}

function Get-Setting($conn, [string]$key, [string]$default) {
    $rows = Invoke-Query $conn 'SELECT SettingValue FROM tblSettings WHERE SettingKey=?' @($key)
    if ($rows.Count -gt 0 -and $rows[0].SettingValue) { return [string]$rows[0].SettingValue }
    return $default
}

# ---------------------------------------------------------------- audit (same as modAudit)

function Get-Sha256Hex([string]$s) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($s))
        return (-join ($bytes | ForEach-Object { $_.ToString('x2') }))
    } finally { $sha.Dispose() }
}

function Text($v) { if ($null -eq $v) { return '' } return [string]$v }

# modAudit.RowText: the columns a row's hash covers, in a fixed order and format.
function Get-AuditRowText($row) {
    $logged = ([datetime]$row.LoggedAt).ToString('yyyy-MM-dd HH:mm:ss')
    return ((Text $row.AuditID) + '|' + $logged + '|' + (Text $row.PersonID) + '|' + (Text $row.UserName) + '|' +
        (Text $row.Action) + '|' + (Text $row.Resource) + '|' + (Text $row.Classification) + '|' +
        (Text $row.Result) + '|' + (Text $row.Device) + '|' + (Text $row.Location) + '|' + (Text $row.Details))
}

# modAudit.AuditSealPending: seal every unsealed row, in AuditID order, onto
# the last sealed row before it. Inside the caller's transaction.
function Invoke-AuditSeal($conn, $tx) {
    $first = Invoke-Query $conn 'SELECT Min(AuditID) AS FirstID FROM tblAudit WHERE RowHash Is Null' @() $tx
    $firstId = $first[0].FirstID
    if ($null -eq $firstId) { return 0 }
    $prev = $GENESIS
    $p = Invoke-Query $conn 'SELECT TOP 1 RowHash FROM tblAudit WHERE AuditID<? ORDER BY AuditID DESC' @([int]$firstId) $tx
    if ($p.Count -gt 0 -and $p[0].RowHash) { $prev = $p[0].RowHash }
    $rows = Invoke-Query $conn 'SELECT * FROM tblAudit WHERE AuditID>=? ORDER BY AuditID' @([int]$firstId) $tx
    $n = 0
    foreach ($row in $rows) {
        if (-not $row.RowHash) {
            $h = Get-Sha256Hex ($prev + '|' + (Get-AuditRowText $row))
            [void](Invoke-NonQuery $conn 'UPDATE tblAudit SET PrevHash=?, RowHash=? WHERE AuditID=?' @($prev, $h, [int]$row.AuditID) $tx)
            $row.RowHash = $h
            $n++
        }
        $prev = $row.RowHash
    }
    return $n
}

# modAudit.AuditLog: add one row for the signed-in person and seal it.
function Add-Audit($conn, $tx, $me, [string]$action, [string]$resource, [string]$details, [string]$classif = 'UNCLASSIFIED', [string]$result = 'Success') {
    function Cut([string]$s, [int]$n) { if ($s.Length -gt $n) { return $s.Substring(0, $n) } return $s }
    $who = $null; if ($me.personId) { $who = [int]$me.personId }
    [void](Invoke-NonQuery $conn ('INSERT INTO tblAudit (LoggedAt, PersonID, UserName, Action, Resource, Classification, Result, Device, Location, Details) ' +
        'VALUES (Now(), ?, ?, ?, ?, ?, ?, ?, ?, ?)') @($who, (Cut $me.username 100), (Cut $action 40), (Cut $resource 150),
        (Cut $classif 20), (Cut $result 10), $Device, (Cut (Text $me.unit) 64), $details) $tx)
    [void](Invoke-AuditSeal $conn $tx)
}

# modAudit.AuditVerify: recompute the whole chain.
function Test-AuditChain($conn) {
    $rows = Invoke-Query $conn 'SELECT * FROM tblAudit ORDER BY AuditID'
    $prev = $GENESIS
    foreach ($row in $rows) {
        if (-not $row.RowHash) { return @{ ok = $false; checked = $rows.Count; problem = "Entry $($row.AuditID) isn't sealed yet." } }
        if ((Text $row.PrevHash) -ne $prev) { return @{ ok = $false; checked = $rows.Count; problem = "Entry $($row.AuditID) doesn't follow the entry before it: an entry was removed or changed before it." } }
        if ((Get-Sha256Hex ($prev + '|' + (Get-AuditRowText $row))) -ne $row.RowHash) { return @{ ok = $false; checked = $rows.Count; problem = "Entry $($row.AuditID) was changed after it was recorded." } }
        $prev = $row.RowHash
    }
    return @{ ok = $true; checked = $rows.Count; problem = '' }
}

# ---------------------------------------------------------------- people

$script:Names = @{}       # person_id -> display name (from Personnel)
$script:Accounts = @{}    # windows account (lower case) -> person_id

function Get-PersonnelKey {
    $kf = Join-Path (Split-Path -Parent $Personnel) 'Personnel.key'
    if (Test-Path $kf) {
        $k = (Get-Content -LiteralPath $kf -TotalCount 1)
        if ($k) { $k = $k.Trim() }
        if ($k) { return $k }
    }
    return $DEFAULT_PERSONNEL_KEY
}

# Read the directory once at start and again when an unknown account signs in.
function Import-Personnel {
    if (-not $Personnel) { return }
    $c = Open-Db $Personnel (Get-PersonnelKey)
    try {
        $rows = Invoke-Query $c 'SELECT person_id, windows_account, rank, first_name, last_name, display_name, is_active FROM kit_person'
        $script:Names = @{}; $script:Accounts = @{}
        foreach ($r in $rows) {
            $nm = Text $r.display_name
            if (-not $nm) { $nm = ((Text $r.rank) + ' ' + (Text $r.first_name) + ' ' + (Text $r.last_name)).Trim() }
            $script:Names[[int]$r.person_id] = $nm
            if ($r.windows_account -and $r.is_active) { $script:Accounts[([string]$r.windows_account).ToLower()] = [int]$r.person_id }
        }
    } finally { $c.Close() }
}

function Get-PersonName([int]$personId) {
    if ($script:Names.ContainsKey($personId)) { return $script:Names[$personId] }
    return "Person $personId"
}

# Who is calling: Windows account -> person_id -> their active ANCHOR profile.
# Returns $null when they can't use ANCHOR.
function Resolve-Me($ctx, $conn) {
    $account = $ctx.User.Identity.Name
    $personId = 0
    if ($DemoSignIn) {
        $h = $ctx.Request.Headers['X-Anchor-As']
        if ($h -match '^\d{1,9}$') { $personId = [int]$h }
    } else {
        $key = $account.ToLower()
        if (-not $script:Accounts.ContainsKey($key)) { Import-Personnel }
        if ($script:Accounts.ContainsKey($key)) { $personId = $script:Accounts[$key] }
    }
    if ($personId -eq 0) { return @{ personId = 0; account = $account } }
    $p = Invoke-Query $conn 'SELECT PersonID, Billet, Unit, IsAdmin, IsActive FROM tblPeople WHERE PersonID=?' @($personId)
    if ($p.Count -eq 0 -or -not $p[0].IsActive) { return @{ personId = 0; account = $account } }
    return @{ personId = $personId; account = $account; username = (Get-PersonName $personId); billet = $p[0].Billet
              unit = $p[0].Unit; isAdmin = [bool]$p[0].IsAdmin }
}

# ---------------------------------------------------------------- API

# The list columns (modMessages.MessageListColsSql), for the person pid.
# personId is an integer we resolved ourselves; text values always go as parameters.
function Get-MessageListSql([int]$personId) {
    return "SELECT m.MessageID AS id, IIf(s.ReadState Is Null,'unread',s.ReadState) AS readState, m.Precedence AS precedence, " +
        "m.Dtg AS dtg, m.Reference AS reference, m.Subject AS subject, m.Originator AS originator, m.MsgType AS msgType, " +
        "m.Classification AS classification, m.Lifecycle AS lifecycle, " +
        "IIf(m.AckRequired AND IIf(m.Source Is Null,'',m.Source)<>? AND s.AckOn Is Null,1,0) AS ackDue, " +
        "s.AckOn AS ackOn, m.AckDue AS ackBy, m.ReceivedOn AS receivedOn " +
        "FROM tblMessages AS m LEFT JOIN (SELECT MessageID, ReadState, AckOn FROM tblMessageUserState WHERE PersonID=$personId) AS s " +
        "ON m.MessageID=s.MessageID"
}

# A LIKE pattern (ANSI-92 wildcards, as OLE DB uses) that matches term literally.
function Get-LikePattern([string]$term) {
    $t = $term -replace '\[', '[[]' -replace '%', '[%]' -replace '_', '[_]'
    return "%$t%"
}

function Get-Messages($conn, $me, $query) {
    $src = Get-Setting $conn 'CommandSource' 'COMSUBLANT'
    $sql = Get-MessageListSql $me.personId
    $where = @(); $params = @($src)
    $q = $query['q']
    if ($q) {
        $where += '(m.Subject LIKE ? OR m.Reference LIKE ? OR m.Originator LIKE ? OR m.MsgType LIKE ?)'
        $pat = Get-LikePattern $q.Substring(0, [Math]::Min(100, $q.Length))
        $params += @($pat, $pat, $pat, $pat)
    }
    $prec = $query['precedence']
    if ($prec -in @('FLASH', 'IMMEDIATE', 'PRIORITY', 'ROUTINE')) { $where += 'm.Precedence=?'; $params += $prec }
    if ($query['unread'] -eq '1') { $where += "(s.ReadState Is Null OR s.ReadState='unread')" }
    if ($query['ackDue'] -eq '1') { $where += "(m.AckRequired AND IIf(m.Source Is Null,'',m.Source)<>? AND s.AckOn Is Null)"; $params += $src }
    if ($where.Count -gt 0) { $sql += ' WHERE ' + ($where -join ' AND ') }
    $sql += ' ORDER BY m.ReceivedOn DESC'
    return , (Invoke-Query $conn $sql $params)
}

function Get-Summary($conn, $me) {
    $src = Get-Setting $conn 'CommandSource' 'COMSUBLANT'
    $all = Invoke-Query $conn ((Get-MessageListSql $me.personId) + ' ORDER BY m.ReceivedOn DESC') @($src)
    $ackDue = New-Object System.Collections.Generic.List[object]
    $unread = 0
    foreach ($m in $all) {
        if ($m.ackDue -eq 1) { $ackDue.Add($m) }
        if ($m.readState -eq 'unread') { $unread++ }
    }
    $latest = New-Object System.Collections.Generic.List[object]
    foreach ($m in $all) { if ($latest.Count -lt 8) { $latest.Add($m) } }
    return @{ total = $all.Count; unread = $unread; ackDueCount = $ackDue.Count; ackDue = $ackDue; latest = $latest
              commandName = (Get-Setting $conn 'CommandName' '') }
}

function Get-MessageDetail($conn, $me, [int]$id) {
    $rows = Invoke-Query $conn 'SELECT * FROM tblMessages WHERE MessageID=?' @($id)
    if ($rows.Count -eq 0) { return $null }
    $m = $rows[0]
    Set-MessageRead $conn $me $id ([string]$m.Reference) ([string]$m.Classification)
    $src = Get-Setting $conn 'CommandSource' 'COMSUBLANT'
    $state = Invoke-Query $conn "SELECT ReadState, FirstReadOn, AckOn FROM tblMessageUserState WHERE MessageID=? AND PersonID=$($me.personId)" @($id)
    $ackOn = $null; if ($state.Count -gt 0) { $ackOn = $state[0].AckOn }
    return @{
        message     = $m
        sections    = (Invoke-Query $conn 'SELECT SortOrder, Heading, Body FROM tblMessageSections WHERE MessageID=? ORDER BY SortOrder' @($id))
        attachments = (Invoke-Query $conn 'SELECT FileName, SizeText, Kind FROM tblAttachments WHERE MessageID=? ORDER BY AttachmentID' @($id))
        addressees  = (Invoke-Query $conn 'SELECT AddrName, AddrRole, AddrKind FROM tblAddressees WHERE MessageID=? ORDER BY AddresseeID' @($id))
        receipts    = (Invoke-Query $conn 'SELECT CommandName, ReceivedOn FROM tblAckReceipts WHERE MessageID=? ORDER BY ReceivedOn' @($id))
        ackOn       = $ackOn
        ackDue      = ([bool]$m.AckRequired -and (Text $m.Source) -ne $src -and $null -eq $ackOn)
    }
}

# modMessages.MarkRead: first open inserts the state row and audits the view.
function Set-MessageRead($conn, $me, [int]$id, [string]$ref, [string]$classif) {
    $tx = $conn.BeginTransaction()
    try {
        $n = Invoke-Query $conn "SELECT ReadState FROM tblMessageUserState WHERE MessageID=? AND PersonID=$($me.personId)" @($id) $tx
        if ($n.Count -eq 0) {
            [void](Invoke-NonQuery $conn "INSERT INTO tblMessageUserState (MessageID, PersonID, ReadState, FirstReadOn) VALUES (?, $($me.personId), 'read', Now())" @($id) $tx)
            Add-Audit $conn $tx $me 'View' "Message ($ref)" 'Message opened and read in the web front end.' $classif
        } elseif ($n[0].ReadState -ne 'read') {
            [void](Invoke-NonQuery $conn "UPDATE tblMessageUserState SET ReadState='read' WHERE MessageID=? AND PersonID=$($me.personId)" @($id) $tx)
        }
        $tx.Commit()
    } catch { $tx.Rollback(); throw }
}

# modMessages.Acknowledge: the acknowledgement and its audit row are one
# transaction; an acknowledgement that can't be audited is not recorded.
function Invoke-Acknowledge($conn, $me, [int]$id) {
    $src = Get-Setting $conn 'CommandSource' 'COMSUBLANT'
    $tx = $conn.BeginTransaction()
    try {
        $m = Invoke-Query $conn ("SELECT m.Reference, m.Classification, IIf(m.AckRequired AND IIf(m.Source Is Null,'',m.Source)<>? AND s.AckOn Is Null,1,0) AS due " +
            "FROM tblMessages AS m LEFT JOIN (SELECT MessageID, AckOn FROM tblMessageUserState WHERE PersonID=$($me.personId)) AS s " +
            "ON m.MessageID=s.MessageID WHERE m.MessageID=?") @($src, $id) $tx
        if ($m.Count -eq 0) { $tx.Rollback(); return @{ status = 404; error = 'No such message.' } }
        if ($m[0].due -ne 1) { $tx.Rollback(); return @{ status = 409; error = "This message doesn't need your acknowledgement." } }
        $has = Invoke-Query $conn "SELECT MessageID FROM tblMessageUserState WHERE MessageID=? AND PersonID=$($me.personId)" @($id) $tx
        if ($has.Count -eq 0) {
            [void](Invoke-NonQuery $conn "INSERT INTO tblMessageUserState (MessageID, PersonID, ReadState, FirstReadOn, AckOn) VALUES (?, $($me.personId), 'read', Now(), Now())" @($id) $tx)
        } else {
            [void](Invoke-NonQuery $conn "UPDATE tblMessageUserState SET ReadState='read', AckOn=Now() WHERE MessageID=? AND PersonID=$($me.personId)" @($id) $tx)
        }
        Add-Audit $conn $tx $me 'Acknowledge' "Message ($($m[0].Reference))" 'Receipt acknowledged in the ANCHOR web front end.' ([string]$m[0].Classification)
        $tx.Commit()
        return @{ status = 200; ok = $true }
    } catch { try { $tx.Rollback() } catch { }; throw }
}

function Get-Audit($conn) {
    $rows = Invoke-Query $conn ('SELECT TOP 200 AuditID, LoggedAt, UserName, Action, Resource, Classification, Result, Device, Location, Details, RowHash ' +
        'FROM tblAudit ORDER BY AuditID DESC')
    return @{ rows = $rows; chain = (Test-AuditChain $conn) }
}

function Get-DemoPeople($conn) {
    $rows = Invoke-Query $conn 'SELECT PersonID, Billet, Unit FROM tblPeople WHERE IsActive ORDER BY PersonID'
    foreach ($r in $rows) { $r['Name'] = Get-PersonName ([int]$r.PersonID) }
    return , $rows
}

# ---------------------------------------------------------------- HTTP

function Send($ctx, [int]$status, [string]$body, [string]$type = 'application/json; charset=utf-8') {
    $res = $ctx.Response
    $res.StatusCode = $status
    $res.ContentType = $type
    $res.Headers['X-Content-Type-Options'] = 'nosniff'
    $res.Headers['X-Frame-Options'] = 'DENY'
    $res.Headers['Referrer-Policy'] = 'no-referrer'
    $res.Headers['Cache-Control'] = 'no-store'
    # One self-contained page: inline script and style, nothing from anywhere else.
    $res.Headers['Content-Security-Policy'] = "default-src 'none'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; img-src data:; connect-src 'self'; base-uri 'none'; form-action 'none'; frame-ancestors 'none'"
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
    $res.ContentLength64 = $bytes.Length
    $res.OutputStream.Write($bytes, 0, $bytes.Length)
    $res.OutputStream.Close()
}

function SendJson($ctx, [int]$status, $obj) { Send $ctx $status (ConvertTo-Json -InputObject $obj -Depth 8 -Compress) }

$AllowedHosts = @("localhost:$Port", "127.0.0.1:$Port")
if ($Lan) { $AllowedHosts += @("$($env:COMPUTERNAME):$Port".ToLower(), "$([System.Net.Dns]::GetHostEntry('').HostName):$Port".ToLower()) }

function Handle($ctx) {
    $req = $ctx.Request
    $path = $req.Url.AbsolutePath
    # Only our own host names: stops a web page that rebinds its DNS name to
    # this machine from reading the API with the person's Windows sign-in.
    if ($AllowedHosts -notcontains ([string]$req.Headers['Host']).ToLower()) { SendJson $ctx 421 @{ error = 'Wrong host.' }; return }

    if ($req.HttpMethod -eq 'GET' -and ($path -eq '/' -or $path -eq '/index.html')) {
        if (-not (Test-Path $UiFile)) { Send $ctx 500 'The UI is not built: web\dist\index.html is missing.' 'text/plain; charset=utf-8'; return }
        Send $ctx 200 ([System.IO.File]::ReadAllText($UiFile)) 'text/html; charset=utf-8'
        return
    }
    if (-not $path.StartsWith('/api/')) { SendJson $ctx 404 @{ error = 'Not found.' }; return }

    # Writes need the token /api/me hands out. Another site can make the
    # browser send a request here with the person's Windows sign-in, but it
    # can't read /api/me, so it can't know the token.
    if ($req.HttpMethod -ne 'GET') {
        if ($req.Headers['X-Anchor-Csrf'] -ne $CsrfToken) { SendJson $ctx 403 @{ error = 'Missing or wrong request token. Reload the page.' }; return }
    }

    $conn = Open-Db $DataFile
    try {
        if ($DemoSignIn -and $path -eq '/api/people') { SendJson $ctx 200 @{ people = (Get-DemoPeople $conn) }; return }
        $me = Resolve-Me $ctx $conn
        if ($path -eq '/api/me') {
            SendJson $ctx 200 @{ me = $me; csrf = $CsrfToken; demoSignIn = [bool]$DemoSignIn; server = $env:COMPUTERNAME
                                 dataFile = (Split-Path -Leaf $DataFile) }
            return
        }
        if ($me.personId -eq 0) { SendJson $ctx 403 @{ error = "Your Windows account ($($me.account)) has no active ANCHOR profile." }; return }

        $query = @{}
        foreach ($k in $req.QueryString.AllKeys) { if ($k) { $query[$k] = $req.QueryString[$k] } }

        switch -Regex ($req.HttpMethod + ' ' + $path) {
            '^GET /api/summary$' { SendJson $ctx 200 (Get-Summary $conn $me); return }
            '^GET /api/messages$' { SendJson $ctx 200 @{ messages = (Get-Messages $conn $me $query) }; return }
            '^GET /api/messages/(\d{1,9})$' {
                $d = Get-MessageDetail $conn $me ([int]$Matches[1])
                if ($null -eq $d) { SendJson $ctx 404 @{ error = 'No such message.' } } else { SendJson $ctx 200 $d }
                return
            }
            '^POST /api/messages/(\d{1,9})/ack$' {
                $r = Invoke-Acknowledge $conn $me ([int]$Matches[1])
                SendJson $ctx $r.status $r
                return
            }
            '^GET /api/audit$' { SendJson $ctx 200 (Get-Audit $conn); return }
            default { SendJson $ctx 404 @{ error = 'Not found.' }; return }
        }
    } finally { $conn.Close() }
}

Import-Personnel
$listener = New-Object System.Net.HttpListener
$listener.AuthenticationSchemes = [System.Net.AuthenticationSchemes]::IntegratedWindowsAuthentication
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Prefixes.Add("http://127.0.0.1:$Port/")
if ($Lan) { $listener.Prefixes.Add("http://+:$Port/") }
$listener.Start()
Write-Host "ANCHOR web: http://localhost:$Port/  (data: $DataFile, engine: $Provider)"
if ($DemoSignIn) { Write-Host 'Demo sign-in is ON: anyone who can reach this page can act as any ANCHOR person. Review only.' -ForegroundColor Yellow }
Write-Host 'Ctrl+C stops it.'
try {
    while ($listener.IsListening) {
        $async = $listener.BeginGetContext($null, $null)
        while (-not $async.AsyncWaitHandle.WaitOne(500)) { }
        $ctx = $listener.EndGetContext($async)
        $t0 = Get-Date
        try { Handle $ctx }
        catch {
            Write-Warning ("{0} {1}: {2}" -f $ctx.Request.HttpMethod, $ctx.Request.Url.AbsolutePath, $_.Exception.Message)
            try { SendJson $ctx 500 @{ error = 'The server hit an error. It is in the server window.' } } catch { }
        }
        Write-Host ("{0:HH:mm:ss} {1} {2} {3} ms" -f $t0, $ctx.Request.HttpMethod, $ctx.Request.Url.PathAndQuery, [int]((Get-Date) - $t0).TotalMilliseconds)
    }
} finally { $listener.Stop(); $listener.Close() }
