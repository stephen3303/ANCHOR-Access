<#
Kit web server core: a small JSON server over an Access data file, shared by
every app. It holds nothing app-specific. An app supplies its routes and
two hooks in its own module (see ..\..\app\server\ in an app) and starts it:

    . (Join-Path $kit 'server\KitWeb.ps1')
    Initialize-KitWeb -AppName 'ANCHOR' -DataFile $DataFile -Personnel $Personnel -Port $Port -DemoSignIn:$DemoSignIn
    Set-KitProfileResolver { param($conn, [int]$personId) ... return @{ billet = ...; unit = ...; isAdmin = ... } or $null }
    Set-KitPeopleLister   { param($conn) ... rows with PersonID, Name, Billet, Unit }    # for -DemoSignIn
    Add-KitRoute GET '^/api/things$' { param($r) $r.Json(200, @{ things = (Invoke-KitQuery $r.Conn 'SELECT ...') }) }
    Start-KitWeb -OnStart { param($conn) ... }

What the core does for every app:
  - Serves the app's one-file UI (dist\index.html) and its /api/ routes.
  - Windows sign-in (Negotiate/NTLM, built into Windows): the caller's
    Windows account is looked up in the shared Personnel directory
    (kit_person.windows_account) for the person_id the app's tables use, then
    the app's profile resolver says whether they may use this app.
  - Opens the data file with the Access engine (ACE OLE DB) for each request,
    shared, and closes it, so Access front ends and other copies of this
    server work on the same file at the same time.
  - Parameterized SQL helpers (Invoke-KitQuery, Invoke-KitNonQuery,
    Use-KitTransaction). Values always go as parameters.
  - Security: localhost only unless -Lan; only its own Host names; a per-run
    token on every write (CSRF); strict Content-Security-Policy; no-store.
  - A write that hits an Access record lock is retried a few times (each
    route runs its writes in a transaction that rolls back on error).

Needs Windows PowerShell 5.1 and the Access engine of the same bitness.
#>
Add-Type -AssemblyName System.Data

$script:Kit = @{
    AppName = 'App'; DataFile = ''; Personnel = ''; Port = 8787; Lan = $false; DemoSignIn = $false
    UiFile = ''; Provider = ''; Csrf = [guid]::NewGuid().ToString('N'); Routes = New-Object System.Collections.Generic.List[object]
    ResolveProfile = $null; ListPeople = $null; Names = @{}; Accounts = @{}; Device = ''
}
$KIT_DEFAULT_PERSONNEL_KEY = 'kit-personnel-1'   # modPersonnel.DEFAULT_KEY

function Initialize-KitWeb {
    param([string]$AppName, [string]$DataFile, [string]$Personnel = '', [int]$Port = 8787, [switch]$Lan,
          [switch]$DemoSignIn, [Parameter(Mandatory = $true)][string]$UiFile, [string]$LogFile = '')
    # The launcher runs the server hidden and asks for a log file.
    if ($LogFile) { Start-Transcript -Path $LogFile -Append | Out-Null }
    $resolve = { param($p) $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($p) }
    $DataFile = & $resolve $DataFile
    if (-not (Test-Path $DataFile)) { throw "No data file at $DataFile" }
    if ($Personnel) {
        $Personnel = & $resolve $Personnel
        if (-not (Test-Path $Personnel)) { throw "No Personnel file at $Personnel" }
    }
    if (-not $Personnel -and -not $DemoSignIn) { throw 'Give -Personnel <Personnel_be.accdb> (Windows sign-in) or -DemoSignIn (review only).' }
    $Kit.AppName = $AppName; $Kit.DataFile = $DataFile; $Kit.Personnel = $Personnel; $Kit.Port = $Port
    $Kit.Lan = [bool]$Lan; $Kit.DemoSignIn = [bool]$DemoSignIn; $Kit.UiFile = & $resolve $UiFile
    $d = $env:COMPUTERNAME + ' (web)'; if ($d.Length -gt 64) { $d = $d.Substring(0, 64) }
    $Kit.Device = $d
    $Kit.Provider = Find-KitAceProvider $DataFile
}

# The computer name with " (web)", for audit columns that record the device.
function Get-KitDevice { return $Kit.Device }

function Set-KitProfileResolver([scriptblock]$Block) { $Kit.ResolveProfile = $Block }
function Set-KitPeopleLister([scriptblock]$Block) { $Kit.ListPeople = $Block }

# Add a route. Pattern is a regex over "<path>"; method GET, POST, PUT or DELETE.
# The handler gets one argument $r with: Conn, Me, Query (hashtable), Matches,
# Body (parsed JSON or $null), Ctx, and $r.Json(status, object).
function Add-KitRoute([string]$Method, [string]$Pattern, [scriptblock]$Handler) {
    $Kit.Routes.Add(@{ Method = $Method.ToUpper(); Pattern = $Pattern; Handler = $Handler })
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

function Open-KitDb([string]$path = '', [string]$password = '') {
    if (-not $path) { $path = $Kit.DataFile }
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

# Who is calling. personId 0: they can't use this app.
function Resolve-KitMe($ctx, $conn) {
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
    $none = @{ personId = 0; account = $account }
    if ($personId -eq 0) { return $none }
    $appProfile = & $Kit.ResolveProfile $conn $personId
    if ($null -eq $appProfile) { return $none }
    $me = @{ personId = $personId; account = $account; username = (Get-KitPersonName $personId) }
    foreach ($k in $appProfile.Keys) { $me[$k] = $appProfile[$k] }
    return $me
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
    $m = $err.Exception.ToString()
    return ($m -match 'currently locked|could not lock|Could not update|already in use|locked by user')
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
    $body = $null
    if ($req.HasEntityBody) {
        $txt = (New-Object System.IO.StreamReader($req.InputStream, [System.Text.Encoding]::UTF8)).ReadToEnd()
        if ($txt.Length -gt 1MB) { Send-KitJson $ctx 413 @{ error = 'Too large.' }; return }
        if ($txt) { try { $body = $txt | ConvertFrom-Json } catch { Send-KitJson $ctx 400 @{ error = 'The request is not valid JSON.' }; return } }
    }

    for ($attempt = 1; ; $attempt++) {
        $conn = Open-KitDb
        try {
            if ($Kit.DemoSignIn -and $path -eq '/api/people') { Send-KitJson $ctx 200 @{ people = (& $Kit.ListPeople $conn) }; return }
            $me = Resolve-KitMe $ctx $conn
            if ($path -eq '/api/me') {
                Send-KitJson $ctx 200 @{ me = $me; csrf = $Kit.Csrf; demoSignIn = $Kit.DemoSignIn; app = $Kit.AppName
                                         server = $env:COMPUTERNAME; dataFile = (Split-Path -Leaf $Kit.DataFile) }
                return
            }
            if ($me.personId -eq 0) { Send-KitJson $ctx 403 @{ error = "Your Windows account ($($me.account)) has no active $($Kit.AppName) profile." }; return }
            $query = @{}
            foreach ($k in $req.QueryString.AllKeys) { if ($k) { $query[$k] = $req.QueryString[$k] } }
            foreach ($route in $Kit.Routes) {
                if ($route.Method -ne $req.HttpMethod) { continue }
                $m = [regex]::Match($path, $route.Pattern)
                if (-not $m.Success) { continue }
                $r = [pscustomobject]@{ Conn = $conn; Me = $me; Query = $query; Matches = $m; Body = $body; Ctx = $ctx }
                $r | Add-Member -MemberType ScriptMethod -Name Json -Value { param([int]$s, $o) Send-KitJson $this.Ctx $s $o }
                [void](& $route.Handler $r)
                return
            }
            Send-KitJson $ctx 404 @{ error = 'Not found.' }
            return
        } catch {
            # Another user (Access or web) holds the record: wait and try again.
            if ($attempt -lt 4 -and (Test-KitLocked $_)) { Start-Sleep -Milliseconds (150 * $attempt); continue }
            throw
        } finally { $conn.Close() }
    }
}

# Listen until Ctrl+C. OnStart runs once with an open connection first.
function Start-KitWeb([scriptblock]$OnStart = $null) {
    if (-not $Kit.ResolveProfile) { throw 'The app must call Set-KitProfileResolver.' }
    if ($Kit.DemoSignIn -and -not $Kit.ListPeople) { throw 'The app must call Set-KitPeopleLister for -DemoSignIn.' }
    $port = $Kit.Port
    $Kit.AllowedHosts = @("localhost:$port", "127.0.0.1:$port")
    if ($Kit.Lan) { $Kit.AllowedHosts += @("$($env:COMPUTERNAME):$port".ToLower(), "$([System.Net.Dns]::GetHostEntry('').HostName):$port".ToLower()) }
    Import-KitPersonnel
    if ($OnStart) { $c = Open-KitDb; try { & $OnStart $c } finally { $c.Close() } }

    $listener = New-Object System.Net.HttpListener
    $listener.AuthenticationSchemes = [System.Net.AuthenticationSchemes]::IntegratedWindowsAuthentication
    $listener.Prefixes.Add("http://localhost:$port/")
    $listener.Prefixes.Add("http://127.0.0.1:$port/")
    if ($Kit.Lan) { $listener.Prefixes.Add("http://+:$port/") }
    $listener.Start()
    Write-Host "$($Kit.AppName) web: http://localhost:$port/  (data: $($Kit.DataFile), engine: $($Kit.Provider))"
    if ($Kit.DemoSignIn) { Write-Host "Demo sign-in is ON: anyone who can reach this page can act as any $($Kit.AppName) person. Review only." -ForegroundColor Yellow }
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
