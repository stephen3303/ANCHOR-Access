<#
ANCHOR's web module: its profile rule, its routes, and the business rules
those routes need, over ANCHOR_BE.accdb. Everything generic (sign-in, the
data layer, security, HTTP) is the kit core, web\kit\server\KitWeb.ps1.
Routes answer under /api/anchor/.

The rules here were ported from the ANCHOR Access front end's VBA; each
names the routine it came from. The web is now the only front end, so this
module is where those rules live.
#>

$ANCHOR_GENESIS = 'ANCHOR-AUDIT-GENESIS'   # modAudit.GENESIS


function Get-AnchorSetting($conn, [string]$key, [string]$default) {
    $rows = Invoke-KitQuery $conn 'SELECT SettingValue FROM tblSettings WHERE SettingKey=?' @($key)
    if ($rows.Count -gt 0 -and $rows[0].SettingValue) { return [string]$rows[0].SettingValue }
    return $default
}

# ---------------------------------------------------------------- who may use ANCHOR

# modSignIn: an active tblPeople row is an ANCHOR profile.
Set-KitProfileResolver {
    param($conn, [int]$personId)
    $p = Invoke-KitQuery $conn 'SELECT Billet, Unit, IsAdmin, IsActive FROM tblPeople WHERE PersonID=?' @($personId)
    if ($p.Count -eq 0 -or -not $p[0].IsActive) { return $null }
    return @{ billet = $p[0].Billet; unit = $p[0].Unit; isAdmin = [bool]$p[0].IsAdmin }
}

Set-KitPeopleLister {
    param($conn)
    $rows = Invoke-KitQuery $conn 'SELECT PersonID, Billet, Unit FROM tblPeople WHERE IsActive ORDER BY PersonID'
    foreach ($row in $rows) { $row['Name'] = Get-KitPersonName ([int]$row.PersonID) }
    return , $rows
}

# ---------------------------------------------------------------- audit (modAudit)

# modAudit.RowText: the columns a row's hash covers, in a fixed order and format.
function Get-AnchorAuditRowText($row) {
    $logged = ([datetime]$row.LoggedAt).ToString('yyyy-MM-dd HH:mm:ss')
    return ((ConvertTo-KitText $row.AuditID) + '|' + $logged + '|' + (ConvertTo-KitText $row.PersonID) + '|' + (ConvertTo-KitText $row.UserName) + '|' +
        (ConvertTo-KitText $row.Action) + '|' + (ConvertTo-KitText $row.Resource) + '|' + (ConvertTo-KitText $row.Classification) + '|' +
        (ConvertTo-KitText $row.Result) + '|' + (ConvertTo-KitText $row.Device) + '|' + (ConvertTo-KitText $row.Location) + '|' + (ConvertTo-KitText $row.Details))
}

# modAudit.AuditSealPending: seal every unsealed row, in AuditID order, onto
# the last sealed row before it. Inside the caller's transaction.
function Invoke-AnchorAuditSeal($conn, $tx) {
    $first = Invoke-KitQuery $conn 'SELECT Min(AuditID) AS FirstID FROM tblAudit WHERE RowHash Is Null' @() $tx
    $firstId = $first[0].FirstID
    if ($null -eq $firstId) { return 0 }
    $prev = $ANCHOR_GENESIS
    $p = Invoke-KitQuery $conn 'SELECT TOP 1 RowHash FROM tblAudit WHERE AuditID<? ORDER BY AuditID DESC' @([int]$firstId) $tx
    if ($p.Count -gt 0 -and $p[0].RowHash) { $prev = $p[0].RowHash }
    $rows = Invoke-KitQuery $conn 'SELECT * FROM tblAudit WHERE AuditID>=? ORDER BY AuditID' @([int]$firstId) $tx
    $n = 0
    foreach ($row in $rows) {
        if (-not $row.RowHash) {
            $h = Get-KitSha256Hex ($prev + '|' + (Get-AnchorAuditRowText $row))
            [void](Invoke-KitNonQuery $conn 'UPDATE tblAudit SET PrevHash=?, RowHash=? WHERE AuditID=?' @($prev, $h, [int]$row.AuditID) $tx)
            $row.RowHash = $h
            $n++
        }
        $prev = $row.RowHash
    }
    return $n
}

# modAudit.AuditLog: add one row for the signed-in person and seal it.
# [Action] and the others are bracketed: Action is reserved in Access SQL.
function Add-AnchorAudit($conn, $tx, $me, [string]$action, [string]$resource, [string]$details, [string]$classif = 'UNCLASSIFIED', [string]$result = 'Success') {
    [void](Invoke-KitNonQuery $conn ('INSERT INTO tblAudit ([LoggedAt], [PersonID], [UserName], [Action], [Resource], [Classification], [Result], [Device], [Location], [Details]) ' +
        'VALUES (Now(), ?, ?, ?, ?, ?, ?, ?, ?, ?)') @([int]$me.personId, (Limit-KitText $me.username 100), (Limit-KitText $action 40), (Limit-KitText $resource 150),
        (Limit-KitText $classif 20), (Limit-KitText $result 10), (Get-KitDevice), (Limit-KitText (ConvertTo-KitText $me.profile.unit) 64), $details) $tx)
    [void](Invoke-AnchorAuditSeal $conn $tx)
}

# modAudit.AuditVerify: recompute the whole chain.
function Test-AnchorAuditChain($conn) {
    $rows = Invoke-KitQuery $conn 'SELECT * FROM tblAudit ORDER BY AuditID'
    $prev = $ANCHOR_GENESIS
    foreach ($row in $rows) {
        if (-not $row.RowHash) { return @{ ok = $false; checked = $rows.Count; problem = "Entry $($row.AuditID) isn't sealed yet." } }
        if ((ConvertTo-KitText $row.PrevHash) -ne $prev) { return @{ ok = $false; checked = $rows.Count; problem = "Entry $($row.AuditID) doesn't follow the entry before it: an entry was removed or changed before it." } }
        if ((Get-KitSha256Hex ($prev + '|' + (Get-AnchorAuditRowText $row))) -ne $row.RowHash) { return @{ ok = $false; checked = $rows.Count; problem = "Entry $($row.AuditID) was changed after it was recorded." } }
        $prev = $row.RowHash
    }
    return @{ ok = $true; checked = $rows.Count; problem = '' }
}

# ---------------------------------------------------------------- messages (modMessages)

# modMessages.MessageListColsSql over MessagesFromSql, for one person.
# personId is an integer the core resolved; text always goes as a parameter.
function Get-AnchorMessageListSql([int]$personId) {
    return "SELECT m.MessageID AS id, IIf(s.ReadState Is Null,'unread',s.ReadState) AS readState, m.Precedence AS precedence, " +
        "m.Dtg AS dtg, m.Reference AS reference, m.Subject AS subject, m.Originator AS originator, m.MsgType AS msgType, " +
        "m.Classification AS classification, m.Lifecycle AS lifecycle, " +
        "IIf(m.AckRequired AND IIf(m.Source Is Null,'',m.Source)<>? AND s.AckOn Is Null,1,0) AS ackDue, " +
        "s.AckOn AS ackOn, m.AckDue AS ackBy, m.ReceivedOn AS receivedOn " +
        "FROM tblMessages AS m LEFT JOIN (SELECT MessageID, ReadState, AckOn FROM tblMessageUserState WHERE PersonID=$personId) AS s " +
        "ON m.MessageID=s.MessageID"
}

# modMessages.MarkRead: the first open records the read and audits the view.
function Set-AnchorMessageRead($conn, $me, [int]$id, [string]$ref, [string]$classif) {
    Use-KitTransaction $conn {
        param($tx)
        $n = Invoke-KitQuery $conn "SELECT ReadState FROM tblMessageUserState WHERE MessageID=? AND PersonID=$($me.personId)" @($id) $tx
        if ($n.Count -eq 0) {
            [void](Invoke-KitNonQuery $conn "INSERT INTO tblMessageUserState (MessageID, PersonID, ReadState, FirstReadOn) VALUES (?, $($me.personId), 'read', Now())" @($id) $tx)
            Add-AnchorAudit $conn $tx $me 'View' "Message ($ref)" 'Message opened and read in the web front end.' $classif
        } elseif ($n[0].ReadState -ne 'read') {
            [void](Invoke-KitNonQuery $conn "UPDATE tblMessageUserState SET ReadState='read' WHERE MessageID=? AND PersonID=$($me.personId)" @($id) $tx)
        }
    } | Out-Null
}

Add-KitRoute GET '^/summary$' {
    param($r)
    $src = Get-AnchorSetting $r.Conn 'CommandSource' 'COMSUBLANT'
    $all = Invoke-KitQuery $r.Conn ((Get-AnchorMessageListSql $r.Me.personId) + ' ORDER BY m.ReceivedOn DESC') @($src)
    $ackDue = New-Object System.Collections.Generic.List[object]
    $latest = New-Object System.Collections.Generic.List[object]
    $unread = 0
    foreach ($m in $all) {
        if ($m.ackDue -eq 1) { $ackDue.Add($m) }
        if ($m.readState -eq 'unread') { $unread++ }
        if ($latest.Count -lt 8) { $latest.Add($m) }
    }
    $r.Json(200, @{ total = $all.Count; unread = $unread; ackDueCount = $ackDue.Count; ackDue = $ackDue; latest = $latest
                    commandName = (Get-AnchorSetting $r.Conn 'CommandName' '') })
}

Add-KitRoute GET '^/messages$' {
    param($r)
    $src = Get-AnchorSetting $r.Conn 'CommandSource' 'COMSUBLANT'
    $sql = Get-AnchorMessageListSql $r.Me.personId
    $where = @(); $params = @($src)
    if ($r.Query['q']) {
        $where += '(m.Subject LIKE ? OR m.Reference LIKE ? OR m.Originator LIKE ? OR m.MsgType LIKE ?)'
        $pat = Get-KitLikePattern $r.Query['q']
        $params += @($pat, $pat, $pat, $pat)
    }
    if ($r.Query['precedence'] -in @('FLASH', 'IMMEDIATE', 'PRIORITY', 'ROUTINE')) { $where += 'm.Precedence=?'; $params += $r.Query['precedence'] }
    if ($r.Query['unread'] -eq '1') { $where += "(s.ReadState Is Null OR s.ReadState='unread')" }
    if ($r.Query['ackDue'] -eq '1') { $where += "(m.AckRequired AND IIf(m.Source Is Null,'',m.Source)<>? AND s.AckOn Is Null)"; $params += $src }
    if ($where.Count -gt 0) { $sql += ' WHERE ' + ($where -join ' AND ') }
    $r.Json(200, @{ messages = (Invoke-KitQuery $r.Conn ($sql + ' ORDER BY m.ReceivedOn DESC') $params) })
}

Add-KitRoute GET '^/messages/(\d{1,9})$' {
    param($r)
    $id = [int]$r.Matches.Groups[1].Value
    $conn = $r.Conn
    $rows = Invoke-KitQuery $conn 'SELECT * FROM tblMessages WHERE MessageID=?' @($id)
    if ($rows.Count -eq 0) { $r.Json(404, @{ error = 'No such message.' }); return }
    $m = $rows[0]
    Set-AnchorMessageRead $conn $r.Me $id ([string]$m.Reference) ([string]$m.Classification)
    $src = Get-AnchorSetting $conn 'CommandSource' 'COMSUBLANT'
    $state = Invoke-KitQuery $conn "SELECT AckOn FROM tblMessageUserState WHERE MessageID=? AND PersonID=$($r.Me.personId)" @($id)
    $ackOn = $null; if ($state.Count -gt 0) { $ackOn = $state[0].AckOn }
    $r.Json(200, @{
        message     = $m
        sections    = (Invoke-KitQuery $conn 'SELECT SortOrder, Heading, Body FROM tblMessageSections WHERE MessageID=? ORDER BY SortOrder' @($id))
        attachments = (Invoke-KitQuery $conn 'SELECT FileName, SizeText, Kind FROM tblAttachments WHERE MessageID=? ORDER BY AttachmentID' @($id))
        addressees  = (Invoke-KitQuery $conn 'SELECT AddrName, AddrRole, AddrKind FROM tblAddressees WHERE MessageID=? ORDER BY AddresseeID' @($id))
        receipts    = (Invoke-KitQuery $conn 'SELECT CommandName, ReceivedOn FROM tblAckReceipts WHERE MessageID=? ORDER BY ReceivedOn' @($id))
        ackOn       = $ackOn
        ackDue      = ([bool]$m.AckRequired -and (ConvertTo-KitText $m.Source) -ne $src -and $null -eq $ackOn)
    })
}

# modMessages.Acknowledge: the acknowledgement and its audit row are one
# transaction; an acknowledgement that can't be audited is not recorded.
Add-KitRoute POST '^/messages/(\d{1,9})/ack$' {
    param($r)
    $id = [int]$r.Matches.Groups[1].Value
    $conn = $r.Conn; $me = $r.Me
    $src = Get-AnchorSetting $conn 'CommandSource' 'COMSUBLANT'
    $result = Use-KitTransaction $conn {
        param($tx)
        $m = Invoke-KitQuery $conn ("SELECT m.Reference, m.Classification, IIf(m.AckRequired AND IIf(m.Source Is Null,'',m.Source)<>? AND s.AckOn Is Null,1,0) AS due " +
            "FROM tblMessages AS m LEFT JOIN (SELECT MessageID, AckOn FROM tblMessageUserState WHERE PersonID=$($me.personId)) AS s " +
            "ON m.MessageID=s.MessageID WHERE m.MessageID=?") @($src, $id) $tx
        if ($m.Count -eq 0) { return @{ status = 404; error = 'No such message.' } }
        if ($m[0].due -ne 1) { return @{ status = 409; error = "This message doesn't need your acknowledgement." } }
        $has = Invoke-KitQuery $conn "SELECT MessageID FROM tblMessageUserState WHERE MessageID=? AND PersonID=$($me.personId)" @($id) $tx
        if ($has.Count -eq 0) {
            [void](Invoke-KitNonQuery $conn "INSERT INTO tblMessageUserState (MessageID, PersonID, ReadState, FirstReadOn, AckOn) VALUES (?, $($me.personId), 'read', Now(), Now())" @($id) $tx)
        } else {
            # AckOn Is Null in the WHERE: if another PC acknowledged since the
            # check above, nothing changes and this one answers 409.
            $n = Invoke-KitNonQuery $conn "UPDATE tblMessageUserState SET ReadState='read', AckOn=Now() WHERE MessageID=? AND PersonID=$($me.personId) AND AckOn Is Null" @($id) $tx
            if ($n -ne 1) { return @{ status = 409; error = "This message doesn't need your acknowledgement." } }
        }
        Add-AnchorAudit $conn $tx $me 'Acknowledge' "Message ($($m[0].Reference))" 'Receipt acknowledged in the ANCHOR web front end.' ([string]$m[0].Classification)
        return @{ status = 200; ok = $true }
    }
    $r.Json($result.status, $result)
}

Add-KitRoute GET '^/audit$' {
    param($r)
    $rows = Invoke-KitQuery $r.Conn ('SELECT TOP 200 [AuditID], [LoggedAt], [UserName], [Action], [Resource], [Classification], [Result], [Device], [Location], [Details], [RowHash] ' +
        'FROM tblAudit ORDER BY AuditID DESC')
    $r.Json(200, @{ rows = $rows; chain = (Test-AnchorAuditChain $r.Conn) })
}

# At start: seal rows a new data file was seeded with (as the Access front
# end's modSignIn.AppStartUp did), so the chain verifies from the start.
Set-KitOnStart {
    param($conn)
    $n = Use-KitTransaction $conn { param($tx) Invoke-AnchorAuditSeal $conn $tx }
    if ($n) { Write-Host "ANCHOR: sealed $n audit row(s) that were waiting." }
}
