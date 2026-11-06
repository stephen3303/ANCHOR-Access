<#
Kit web server core: one small JSON server for every kit app, each app with
its own Access data file. Nothing here is specific to one app.

The host (KitHost.ps1) reads a settings file that lists the apps. For each it
calls Use-KitApp, then dot-sources the app's module, which adds that app's
routes and hooks:

    Set-KitProfileResolver { param($conn, [int]$personId) ... return @{ ... } or $null }
    Set-KitPeopleLister   { param($conn) ... rows with PersonID, Name, Billet, Unit }    # demo sign-in only
    Add-KitRoute GET '^/things$' { param($r) $r.Json(200, @{ things = (Invoke-KitQuery $r.Conn 'SELECT ...') }) }
    Set-KitOnStart        { param($conn) ... }                                          # once, at start

An app's routes are served under /api/<app key>/ (the pattern above answers
/api/<key>/things), and $r.Conn is a connection to THAT app's data file only.

Who may do what:
  - Sign-in is Windows (Negotiate/NTLM, built into Windows). The caller's
    Windows account is looked up once in the shared Personnel directory
    (kit_person.windows_account) for their person_id: one identity for all apps.
  - Each app decides on its own whether that person may use it: its profile
    resolver reads the app's own people table, in the app's own data file. No
    profile: the app isn't listed for them and its routes answer 403.
  - Whatever an app records (audit trails, receipts) goes into its own data
    file through its own routes, so apps' records never mix.

Also: the Access engine (ACE OLE DB) opened shared for each request and closed
after, so Access front ends and other copies of this server work on the same
files at once; parameterized SQL helpers and transactions; a retry when a
write meets another user's record lock; localhost only unless -Lan; only its
own Host names; a per-run token on every write (CSRF); a strict CSP.

Needs Windows PowerShell 5.1 and the Access engine of the same bitness.
#>
Add-Type -AssemblyName System.Data

$script:Kit = @{
    Title = 'Apps'; Personnel = ''; Port = 8787; Lan = $false; DemoSignIn = $false; UiFile = ''; Provider = ''
    Csrf = [guid]::NewGuid().ToString('N'); Apps = [ordered]@{}; Current = $null; Names = @{}; Accounts = @{}; Device = ''
}
$KIT_DEFAULT_PERSONNEL_KEY = 'kit-personnel-1'   # modPersonnel.DEFAULT_KEY

function Initialize-KitWeb {
    param([string]$Title = 'Apps', [string]$Personnel = '', [int]$Port = 8787, [switch]$Lan, [switch]$DemoSignIn,
          [Parameter(Mandatory = $true)][string]$UiFile, [string]$LogFile = '')
    # The launcher runs the server hidden and asks for a log file.
    if ($LogFile) { Start-Transcript -Path $LogFile -Append | Out-Null }
    if ($Personnel) {
        $Personnel = Resolve-KitPath $Personnel
        if (-not (Test-Path $Personnel)) { throw "No Personnel file at $Personnel" }
    }
    if (-not $Personnel -and -not $DemoSignIn) { throw 'Give a Personnel file (Windows sign-in) or -DemoSignIn (review only).' }
    $Kit.Title = $Title; $Kit.Personnel = $Personnel; $Kit.Port = $Port; $Kit.Lan = [bool]$Lan
    $Kit.DemoSignIn = [bool]$DemoSignIn; $Kit.UiFile = Resolve-KitPath $UiFile
    $d = $env:COMPUTERNAME + ' (web)'; if ($d.Length -gt 64) { $d = $d.Substring(0, 64) }
    $Kit.Device = $d
}

function Resolve-KitPath([string]$p) { return $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) }

# Start registering an app. The routes and hooks that follow belong to it.
# A data file that can't be reached doesn't stop the server: the app is listed
# as unavailable, with why, and the other apps carry on.
function Use-KitApp([string]$Key, [string]$Name, [string]$DataFile) {
    if ($Key -notmatch '^[a-z][a-z0-9-]{0,30}$') { throw "App key '$Key': use lower-case letters, digits and dashes." }
    if ($Kit.Apps.Contains($Key)) { throw "App key '$Key' is listed twice." }
    $app = @{ Key = $Key; Name = $Name; DataFile = (Resolve-KitPath $DataFile); Routes = New-Object System.Collections.Generic.List[object]
              ResolveProfile = $null; ListPeople = $null; OnStart = $null; Bridge = $null; Error = '' }
    if (-not (Test-Path $app.DataFile)) { $app.Error = "Can't reach its data file ($($app.DataFile))." }
    elseif (-not $Kit.Provider) {
        try { $Kit.Provider = Find-KitAceProvider $app.DataFile } catch { throw }
    }
    $Kit.Apps[$Key] = $app
    $Kit.Current = $app
}

function Set-KitProfileResolver([scriptblock]$Block) { $Kit.Current.ResolveProfile = $Block }
function Set-KitPeopleLister([scriptblock]$Block) { $Kit.Current.ListPeople = $Block }
function Set-KitOnStart([scriptblock]$Block) { $Kit.Current.OnStart = $Block }

# The computer name with " (web)", for audit columns that record the device.
function Get-KitDevice { return $Kit.Device }

# Add a route to the current app. Pattern is a regex over the path after
# /api/<key>. The handler gets $r with: Conn (this app's data file), Me (with
# Me.profile, this app's profile), App, Query (hashtable), Matches, Body
# (parsed JSON or $null), Ctx, and $r.Json(status, object).
function Add-KitRoute([string]$Method, [string]$Pattern, [scriptblock]$Handler) {
    $Kit.Current.Routes.Add(@{ Method = $Method.ToUpper(); Pattern = $Pattern; Handler = $Handler })
}

# ---------------------------------------------------------------- database

function Find-KitAceProvider([string]$dataFile) {
    foreach ($p in 'Microsoft.ACE.OLEDB.16.0', 'Microsoft.ACE.OLEDB.12.0') {
        $c = New-Object System.Data.OleDb.OleDbConnection("Provider=$p;Data Source=$dataFile;Mode=Share Deny None")
        try { $c.Open(); $c.Close(); return $p } catch { } finally { $c.Dispose() }
    }
    $bits = if ([Environment]::Is64BitProcess) { '64-bit' } else { '32-bit' }
    throw "This $bits PowerShell can't load the Access engine (ACE OLE DB). Run the PowerShell whose bitness matches Office, or install the Microsoft Access Database Engine redistributable of this bitness."
}

function Open-KitDb([string]$path, [string]$password = '') {
    $cs = "Provider=$($Kit.Provider);Data Source=$path;Mode=Share Deny None"
    if ($password) { $cs += ";Jet OLEDB:Database Password=$password" }
    $c = New-Object System.Data.OleDb.OleDbConnection($cs)
    $c.Open()
    return $c
}

function New-KitCommand($conn, [string]$sql, [object[]]$params, $tx) {
    $cmd = $conn.CreateCommand()
    $cmd.CommandText = $sql
    if ($tx) { $cmd.Transaction = $tx }
    foreach ($p in $params) { [void]$cmd.Parameters.AddWithValue('?', $(if ($null -eq $p) { [DBNull]::Value } else { $p })) }
    return $cmd
}

# Run a SELECT. Parameters are positional (?). Returns a list of ordered
# hashtables; dates come back as ISO text (yyyy-MM-ddTHH:mm:ss).
# Bracket column names that are reserved words in Access SQL ([Action], [Note]).
function Invoke-KitQuery($conn, [string]$sql, [object[]]$params = @(), $tx = $null) {
    $cmd = New-KitCommand $conn $sql $params $tx
    $rows = New-Object System.Collections.Generic.List[object]
    $rd = $cmd.ExecuteReader()
    try {
        while ($rd.Read()) {
            $row = [ordered]@{}
            for ($i = 0; $i -lt $rd.FieldCount; $i++) {
                $v = $rd.GetValue($i)
                if ($v -is [DBNull]) { $v = $null }
                elseif ($v -is [datetime]) { $v = $v.ToString('yyyy-MM-ddTHH:mm:ss') }
                $row[$rd.GetName($i)] = $v
            }
            $rows.Add($row)
        }
    } finally { $rd.Close() }
    return , $rows
}

function Invoke-KitNonQuery($conn, [string]$sql, [object[]]$params = @(), $tx = $null) {
    return (New-KitCommand $conn $sql $params $tx).ExecuteNonQuery()
}

# Run $Block with a transaction: committed when it returns, rolled back when it
# throws. $Block gets the transaction; whatever it returns is returned.
function Use-KitTransaction($conn, [scriptblock]$Block) {
    $tx = $conn.BeginTransaction()
    try { $result = & $Block $tx; $tx.Commit(); return $result }
    catch { try { $tx.Rollback() } catch { }; throw }
}

# A LIKE pattern (OLE DB uses ANSI-92 wildcards) that matches term literally.
function Get-KitLikePattern([string]$term, [int]$max = 100) {
    if ($term.Length -gt $max) { $term = $term.Substring(0, $max) }
    $t = $term -replace '\[', '[[]' -replace '%', '[%]' -replace '_', '[_]'
    return "%$t%"
}

function Get-KitSha256Hex([string]$s) {
    $sha = [System.Security.Cryptography.SHA256]::Create()
    try { return (-join ($sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($s)) | ForEach-Object { $_.ToString('x2') })) }
    finally { $sha.Dispose() }
}

# Text of a value ('' for null), and text cut to a column's width.
function ConvertTo-KitText($v) { if ($null -eq $v) { return '' } return [string]$v }
function Limit-KitText([string]$s, [int]$n) { if ($s.Length -gt $n) { return $s.Substring(0, $n) } return $s }

# ---------------------------------------------------------------- people

function Get-KitPersonnelKey {
    $kf = Join-Path (Split-Path -Parent $Kit.Personnel) 'Personnel.key'
    if (Test-Path $kf) {
        $k = Get-Content -LiteralPath $kf -TotalCount 1
        if ($k) { $k = $k.Trim() }
        if ($k) { return $k }
    }
    return $KIT_DEFAULT_PERSONNEL_KEY
}

# Read the directory at start, and again when an unknown account signs in.
function Import-KitPersonnel {
    if (-not $Kit.Personnel) { return }
    $c = Open-KitDb $Kit.Personnel (Get-KitPersonnelKey)
    try {
        $rows = Invoke-KitQuery $c 'SELECT person_id, windows_account, rank, first_name, last_name, display_name, is_active FROM kit_person'
        $Kit.Names = @{}; $Kit.Accounts = @{}
        foreach ($p in $rows) {
            $nm = [string]$p.display_name
            if (-not $nm) { $nm = ("$($p.rank) $($p.first_name) $($p.last_name)").Trim() }
            $Kit.Names[[int]$p.person_id] = $nm
            if ($p.windows_account -and $p.is_active) { $Kit.Accounts[([string]$p.windows_account).ToLower()] = [int]$p.person_id }
        }
    } finally { $c.Close() }
}

function Get-KitPersonName([int]$personId) {
    if ($Kit.Names.ContainsKey($personId)) { return $Kit.Names[$personId] }
    return "Person $personId"
}

# Who is calling, from Windows (or the demo header). personId 0: unknown.
function Resolve-KitPerson($ctx) {
    $account = $ctx.User.Identity.Name
    $personId = 0
    if ($Kit.DemoSignIn) {
        $h = $ctx.Request.Headers['X-Kit-As']
        if ($h -match '^\d{1,9}$') { $personId = [int]$h }
    } else {
        $key = $account.ToLower()
        if (-not $Kit.Accounts.ContainsKey($key)) { Import-KitPersonnel }
        if ($Kit.Accounts.ContainsKey($key)) { $personId = $Kit.Accounts[$key] }
    }
    $name = ''; if ($personId) { $name = Get-KitPersonName $personId }
    return @{ personId = $personId; account = $account; username = $name }
}

# This app's profile for the person, or $null (no access to this app).
function Get-KitAppProfile($app, $conn, [int]$personId) {
    if ($personId -eq 0) { return $null }
    return (& $app.ResolveProfile $conn $personId)
}

# Every app, with the caller's profile in each (for /api/me and the switcher).
function Get-KitAppList([int]$personId) {
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($app in $Kit.Apps.Values) {
        $entry = [ordered]@{ key = $app.Key; name = $app.Name; available = (-not $app.Error); error = $app.Error; profile = $null }
        if (-not $app.Error -and $personId) {
            try {
                $c = Open-KitDb $app.DataFile
                try { $entry.profile = Get-KitAppProfile $app $c $personId } finally { $c.Close() }
            } catch { $entry.available = $false; $entry.error = "Its data file can't be opened right now." ; Write-Warning "$($app.Key): $($_.Exception.Message)" }
        }
        $list.Add($entry)
    }
    return , $list
}

# Demo sign-in: everyone with a profile in any app, once each.
function Get-KitDemoPeople {
    $seen = @{}
    $list = New-Object System.Collections.Generic.List[object]
    foreach ($app in $Kit.Apps.Values) {
        if ($app.Error -or -not $app.ListPeople) { continue }
        $c = Open-KitDb $app.DataFile
        try {
            foreach ($p in (& $app.ListPeople $c)) {
                $id = [int]$p.PersonID
                if ($seen.ContainsKey($id)) { $seen[$id].Apps += ", $($app.Name)"; continue }
                $row = [ordered]@{ PersonID = $id; Name = (Get-KitPersonName $id); Billet = $p.Billet; Unit = $p.Unit; Apps = $app.Name }
                $seen[$id] = $row
                $list.Add($row)
            }
        } finally { $c.Close() }
    }
    return , $list
}

# ---------------------------------------------------------------- HTTP

function Send-Kit($ctx, [int]$status, [string]$body, [string]$type = 'application/json; charset=utf-8') {
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

function Send-KitJson($ctx, [int]$status, $obj) { Send-Kit $ctx $status (ConvertTo-Json -InputObject $obj -Depth 8 -Compress) }

function Test-KitLocked($err) {
    return ($err.Exception.ToString() -match 'currently locked|could not lock|Could not update|already in use|locked by user|placed in a state')
}

# Another PC inserted the same key between this request's check and its
# insert (two servers on one file). The transaction rolled back, so running
# the request again re-reads and finds the other PC's row.
function Test-KitDuplicateKey($err) {
    return ($err.Exception.ToString() -match 'duplicate values in the index, primary key')
}

function Invoke-KitRequest($ctx) {
    $req = $ctx.Request
    $path = $req.Url.AbsolutePath
    # Only our own host names: stops a web page that rebinds its DNS name to
    # this machine from reading the API with the person's Windows sign-in.
    if ($Kit.AllowedHosts -notcontains ([string]$req.Headers['Host']).ToLower()) { Send-KitJson $ctx 421 @{ error = 'Wrong host.' }; return }

    if ($req.HttpMethod -eq 'GET' -and ($path -eq '/' -or $path -eq '/index.html')) {
        if (-not (Test-Path $Kit.UiFile)) { Send-Kit $ctx 500 "The UI is not built: $($Kit.UiFile) is missing." 'text/plain; charset=utf-8'; return }
        Send-Kit $ctx 200 ([System.IO.File]::ReadAllText($Kit.UiFile)) 'text/html; charset=utf-8'
        return
    }
    if (-not $path.StartsWith('/api/')) { Send-KitJson $ctx 404 @{ error = 'Not found.' }; return }
    # Writes need the token /api/me hands out. Another site can make the
    # browser send a request here with the person's Windows sign-in, but it
    # can't read /api/me, so it can't know the token.
    if ($req.HttpMethod -ne 'GET' -and $req.Headers['X-Kit-Csrf'] -ne $Kit.Csrf) {
        Send-KitJson $ctx 403 @{ error = 'Missing or wrong request token. Reload the page.' }; return
    }

    $person = Resolve-KitPerson $ctx
    if ($path -eq '/api/me') {
        Send-KitJson $ctx 200 @{ me = $person; apps = (Get-KitAppList $person.personId); csrf = $Kit.Csrf; demoSignIn = $Kit.DemoSignIn
                                 title = $Kit.Title; server = $env:COMPUTERNAME }
        return
    }
    if ($Kit.DemoSignIn -and $path -eq '/api/people') { Send-KitJson $ctx 200 @{ people = (Get-KitDemoPeople) }; return }

    # /api/<app key>/<rest>
    $m = [regex]::Match($path, '^/api/([a-z][a-z0-9-]{0,30})(/.*)$')
    if (-not $m.Success -or -not $Kit.Apps.Contains($m.Groups[1].Value)) { Send-KitJson $ctx 404 @{ error = 'Not found.' }; return }
    $app = $Kit.Apps[$m.Groups[1].Value]
    $rest = $m.Groups[2].Value
    if ($app.Error) { Send-KitJson $ctx 503 @{ error = "$($app.Name) is unavailable: $($app.Error)" }; return }
    if ($person.personId -eq 0) { Send-KitJson $ctx 403 @{ error = "Your Windows account ($($person.account)) isn't in the Personnel directory." }; return }

    $body = $null
    if ($req.HasEntityBody) {
        $txt = (New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)).ReadToEnd()
        if ($txt.Length -gt 1MB) { Send-KitJson $ctx 413 @{ error = 'Too large.' }; return }
        if ($txt) { try { $body = $txt | ConvertFrom-Json } catch { Send-KitJson $ctx 400 @{ error = 'The request is not valid JSON.' }; return } }
    }
    $query = @{}
    foreach ($k in $req.QueryString.AllKeys) { if ($k) { $query[$k] = $req.QueryString[$k] } }

    for ($attempt = 1; ; $attempt++) {
        $conn = $null
        try {
            # Opening is inside the retry too: while another PC's engine
            # briefly holds the file (error 3734), the open itself fails.
            $conn = Open-KitDb $app.DataFile
            $appProfile = Get-KitAppProfile $app $conn $person.personId
            if ($null -eq $appProfile) { Send-KitJson $ctx 403 @{ error = "You have no $($app.Name) profile. Ask a $($app.Name) admin to add you." }; return }
            $me = @{ personId = $person.personId; account = $person.account; username = $person.username; profile = $appProfile }
            foreach ($route in $app.Routes) {
                if ($route.Method -ne $req.HttpMethod) { continue }
                $rm = [regex]::Match($rest, $route.Pattern)
                if (-not $rm.Success) { continue }
                $r = [pscustomobject]@{ Conn = $conn; Me = $me; App = $app; Query = $query; Matches = $rm; Body = $body; Ctx = $ctx }
                $r | Add-Member -MemberType ScriptMethod -Name Json -Value { param([int]$s, $o) Send-KitJson $this.Ctx $s $o }
                [void](& $route.Handler $r)
                return
            }
            Send-KitJson $ctx 404 @{ error = 'Not found.' }
            return
        } catch {
            # Another PC holds the record, or just inserted the same key: wait and
            # run the request again (its transaction rolled back). A route must
            # answer with $r.Json only after its last write, so a retry never
            # answers twice.
            if ($attempt -lt 4 -and ((Test-KitLocked $_) -or (Test-KitDuplicateKey $_))) { Start-Sleep -Milliseconds (150 * $attempt); continue }
            throw
        } finally { if ($conn) { $conn.Close() } }
    }
}

# Listen until Ctrl+C. Each available app's OnStart runs once first.
function Start-KitWeb {
    if ($Kit.Apps.Count -eq 0) { throw 'No apps are listed.' }
    if (-not $Kit.Provider) { throw 'None of the apps'' data files can be reached, so the Access engine could not be checked.' }
    foreach ($app in $Kit.Apps.Values) {
        if (-not $app.ResolveProfile) { throw "App '$($app.Key)' must call Set-KitProfileResolver." }
        if ($app.Error) { Write-Warning "$($app.Name): $($app.Error)"; continue }
        if ($app.OnStart) { $c = Open-KitDb $app.DataFile; try { & $app.OnStart $c } finally { $c.Close() } }
    }
    $port = $Kit.Port
    $Kit.AllowedHosts = @("localhost:$port", "127.0.0.1:$port")
    if ($Kit.Lan) { $Kit.AllowedHosts += @("$($env:COMPUTERNAME):$port".ToLower(), "$([System.Net.Dns]::GetHostEntry('').HostName):$port".ToLower()) }
    Import-KitPersonnel

    $listener = New-Object System.Net.HttpListener
    $listener.AuthenticationSchemes = [System.Net.AuthenticationSchemes]::IntegratedWindowsAuthentication
    $listener.Prefixes.Add("http://localhost:$port/")
    $listener.Prefixes.Add("http://127.0.0.1:$port/")
    if ($Kit.Lan) { $listener.Prefixes.Add("http://+:$port/") }
    $listener.Start()
    Write-Host "$($Kit.Title): http://localhost:$port/  (engine: $($Kit.Provider))"
    foreach ($app in $Kit.Apps.Values) { Write-Host ("  {0,-12} {1}" -f $app.Key, $(if ($app.Error) { "UNAVAILABLE: $($app.Error)" } else { $app.DataFile })) }
    if ($Kit.DemoSignIn) { Write-Host 'Demo sign-in is ON: anyone who can reach this page can act as anyone. Review only.' -ForegroundColor Yellow }
    Write-Host 'Ctrl+C stops it.'
    try {
        while ($listener.IsListening) {
            $async = $listener.BeginGetContext($null, $null)
            while (-not $async.AsyncWaitHandle.WaitOne(500)) { }
            $ctx = $listener.EndGetContext($async)
            $t0 = Get-Date
            try { Invoke-KitRequest $ctx }
            catch {
                Write-Warning ("{0} {1}: {2}" -f $ctx.Request.HttpMethod, $ctx.Request.Url.AbsolutePath, $_.Exception.Message)
                try { Send-KitJson $ctx 500 @{ error = 'The server hit an error. It is in the server log.' } } catch { }
            }
            Write-Host ("{0:HH:mm:ss} {1} {2} {3} ms" -f $t0, $ctx.Request.HttpMethod, $ctx.Request.Url.PathAndQuery, [int]((Get-Date) - $t0).TotalMilliseconds)
        }
    } finally { $listener.Stop(); $listener.Close() }
}
