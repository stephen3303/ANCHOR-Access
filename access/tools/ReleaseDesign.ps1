# Release design files: what a release build writes beside the locked front end so the
# app can be rebuilt or ported without opening Access.
#
#   design\forms\<form>.json   each form of the front end as a JSON design
#                              (FormDesign.ps1, schema access-form-design/1), with
#                              index.json and, given the build palette, tokens.json
#   design\<database>.sqlserver.sql
#                              a build script for each database's own tables in SQL Server
#                              SQL (T-SQL): CREATE TABLE, CREATE INDEX, then the foreign
#                              keys, for moving the data off Access
#   design\<database>.sql      the same tables in Access SQL, only with -Dialect Access
#
# Only a release build makes these (release.ps1 calls Export-ReleaseDesign before it
# compiles the front end), so they always match a tagged version. They are not kept in
# the repository.
#
# Dot-source it next to FormDesign.ps1 and FormSnapshot.ps1 (in this repository they are
# in review-mcp\; an app's copy keeps all three in one folder):
#
#   . "$PSScriptRoot\tools\ReleaseDesign.ps1"
#   Export-ReleaseDesign -Access $acc -FrontEnd $fe -DataFiles $be -ThemeFile $theme -OutDir "$out\design"
#
# -Dialect picks the SQL files: SqlServer (the default, the release format), Access, or
# both.
#
# The Access SQL is ANSI-92, the mode ADO uses (CurrentProject.Connection, or an OLE DB
# connection to the file), which DEFAULT needs. Run each statement on its own, in order:
# Access runs one statement at a time. Each Access script is checked by running it into
# an empty database and reading that back (Test-SchemaSqlRebuild); the release stops if
# the rebuilt tables differ. The SQL Server script can't be run without a SQL Server; it
# is generated from the same DAO schema read, which the tests cover type by type.
#
# Converting a schema to SQL needs no Access, so it can be tested anywhere PowerShell
# runs: see ReleaseDesign.Tests.ps1.

$script:ReleaseDesignDir = $PSScriptRoot
if (Test-Path -LiteralPath (Join-Path $PSScriptRoot 'FormDesign.ps1')) { . (Join-Path $PSScriptRoot 'FormDesign.ps1') }
else { . (Join-Path $PSScriptRoot 'review-mcp\FormDesign.ps1') }

# Tables that are Access's own, not the app's: system tables, temporary tables and the
# hidden tables behind multi-value and attachment fields.
function Test-SchemaSystemTable([string]$Name) {
    $Name -match '^(MSys|USysRibbons$|~)' -or $Name -match '^f_[0-9A-Fa-f]{32}_'
}

# Reads a database's own tables (not linked ones), their indexes and the relationships
# between them through DAO into plain objects, for ConvertTo-SchemaSql. Opens the file
# read-only with the DBEngine of an Access.Application, so no start-up code runs.
function Read-AccessSchema {
    param([Parameter(Mandatory)]$Access, [Parameter(Mandatory)][string]$Path)
    $db = $Access.DBEngine.OpenDatabase([string]$Path, $false, $true)
    try {
        $tables = New-Object System.Collections.Generic.List[object]
        $tds = $db.TableDefs
        for ($i = 0; $i -lt $tds.Count; $i++) {
            $t = $tds.Item($i)
            $name = [string]$t.Name
            if ((Test-SchemaSystemTable $name) -or [string]$t.Connect) { continue }
            $fields = New-Object System.Collections.Generic.List[object]
            for ($j = 0; $j -lt $t.Fields.Count; $j++) {
                $f = $t.Fields.Item($j)
                $expr = ''
                try { $expr = [string]$f.Properties.Item('Expression').Value } catch {}
                $azl = $false
                try { $azl = [bool]$f.AllowZeroLength } catch {}
                $precision = $null; $scale = $null
                if ([int]$f.Type -eq 20) {
                    try { $precision = [int]$f.Properties.Item('Precision').Value; $scale = [int]$f.Properties.Item('Scale').Value } catch {}
                }
                $fields.Add([pscustomobject]@{
                    Name = [string]$f.Name; Ordinal = [int]$f.OrdinalPosition; Type = [int]$f.Type; Size = [int]$f.Size
                    Required = [bool]$f.Required; AllowZeroLength = $azl; Default = [string]$f.DefaultValue
                    AutoIncrement = (([int]$f.Attributes -band 16) -ne 0)   # dbAutoIncrField
                    Expression = $expr; Precision = $precision; Scale = $scale
                })
            }
            $indexes = New-Object System.Collections.Generic.List[object]
            for ($j = 0; $j -lt $t.Indexes.Count; $j++) {
                $ix = $t.Indexes.Item($j)
                $cols = New-Object System.Collections.Generic.List[object]
                for ($k = 0; $k -lt $ix.Fields.Count; $k++) {
                    $c = $ix.Fields.Item($k)
                    $cols.Add([pscustomobject]@{ Name = [string]$c.Name; Descending = (([int]$c.Attributes -band 1) -ne 0) })   # dbDescending
                }
                $indexes.Add([pscustomobject]@{
                    Name = [string]$ix.Name; Primary = [bool]$ix.Primary; Unique = [bool]$ix.Unique
                    Required = [bool]$ix.Required; IgnoreNulls = [bool]$ix.IgnoreNulls; Foreign = [bool]$ix.Foreign
                    Fields = $cols.ToArray()
                })
            }
            $tables.Add([pscustomobject]@{ Name = $name; Fields = @($fields | Sort-Object Ordinal); Indexes = $indexes.ToArray() })
        }
        $relations = New-Object System.Collections.Generic.List[object]
        $rels = $db.Relations
        for ($i = 0; $i -lt $rels.Count; $i++) {
            $r = $rels.Item($i)
            $pairs = New-Object System.Collections.Generic.List[object]
            for ($k = 0; $k -lt $r.Fields.Count; $k++) {
                $c = $r.Fields.Item($k)
                $pairs.Add([pscustomobject]@{ Name = [string]$c.Name; ForeignName = [string]$c.ForeignName })
            }
            $relations.Add([pscustomobject]@{
                Name = [string]$r.Name; Table = [string]$r.Table; ForeignTable = [string]$r.ForeignTable
                Attributes = [int]$r.Attributes; Fields = $pairs.ToArray()
            })
        }
        [pscustomobject]@{ Tables = $tables.ToArray(); Relations = $relations.ToArray() }
    } finally { $db.Close() }
}

function Format-SqlName([string]$Name) { '[' + $Name.Replace(']', ']]') + ']' }

# A field's type as ANSI-92 Access SQL, or $null with -Reason set when SQL can't make it.
function ConvertTo-SqlType($Field, [ref]$Reason) {
    if ($Field.Expression) { $Reason.Value = 'calculated field'; return $null }
    switch ($Field.Type) {
        1 { 'YESNO' }
        2 { 'BYTE' }
        3 { 'SMALLINT' }
        4 { if ($Field.AutoIncrement) { 'COUNTER' } else { 'INTEGER' } }
        5 { 'CURRENCY' }
        6 { 'REAL' }
        7 { 'FLOAT' }
        8 { 'DATETIME' }
        9 { "BINARY($($Field.Size))" }
        10 { "TEXT($($Field.Size))" }
        11 { 'LONGBINARY' }
        12 { 'LONGTEXT' }
        15 { 'UNIQUEIDENTIFIER' }
        16 { 'BIGINT' }
        20 {
            if ($null -eq $Field.Precision) { $Reason.Value = 'decimal field whose precision DAO does not report'; return $null }
            "DECIMAL($($Field.Precision),$($Field.Scale))"
        }
        default { $Reason.Value = Get-SchemaTypeReason $Field.Type; $null }
    }
}

# Why a field type has no column type in either dialect: 101 attachment, 102-109
# multi-value lookups, 26 Date/Time Extended (Access SQL only; SQL Server has DATETIME2).
function Get-SchemaTypeReason([int]$Type) {
    if ($Type -eq 101) { 'attachment field' }
    elseif ($Type -gt 101) { 'multi-value field' }
    elseif ($Type -eq 26) { 'Date/Time Extended field' }
    else { "field type $Type" }
}

# A field's DefaultValue (as DAO returns it) as an ANSI-92 DEFAULT. Text in either kind
# of quotes becomes single-quoted, Yes/No words become -1/0 and a leading = is dropped,
# so a default reads the same however it was set.
function ConvertTo-SqlDefault([string]$Value) {
    $v = $Value.Trim()
    if ($v.StartsWith('=')) { $v = $v.Substring(1).Trim() }
    if (-not $v) { return $null }
    if ($v -match '^"((?:[^"]|"")*)"$') { return "'" + ($Matches[1].Replace('""', '"').Replace("'", "''")) + "'" }
    if ($v -match "^'((?:[^']|'')*)'$") { return $v }
    if ($v -match '^(?i)(yes|true|on)$') { return '-1' }
    if ($v -match '^(?i)(no|false|off)$') { return '0' }
    $v
}

# Turns Read-AccessSchema's objects into a build script. Returns Text (the whole script,
# with a header saying what it leaves out), Statements (each statement on its own, to
# run in order) and Skipped (fields, indexes and relationships SQL can't make).
# Tables, indexes and relationships are sorted by name so the same schema always gives
# the same script.
function ConvertTo-SchemaSql {
    param([Parameter(Mandatory)]$Schema, [string]$Title = 'database')
    $statements = New-Object System.Collections.Generic.List[string]
    $skipped = New-Object System.Collections.Generic.List[string]
    $made = @{}      # table -> names of the fields the script makes
    $tables = @($Schema.Tables | Sort-Object Name)
    foreach ($t in $tables) {
        $cols = New-Object System.Collections.Generic.List[string]
        $made[$t.Name] = @{}
        foreach ($f in $t.Fields) {
            $why = $null
            $type = ConvertTo-SqlType $f ([ref]$why)
            if (-not $type) { $skipped.Add("$(Format-SqlName $t.Name).$(Format-SqlName $f.Name): $why"); continue }
            $col = "    $(Format-SqlName $f.Name) $type"
            # An AutoNumber's default is the next number; DAO reports none.
            $def = if ($f.AutoIncrement) { $null } else { ConvertTo-SqlDefault $f.Default }
            if ($def) { $col += " DEFAULT $def" }
            if ($f.Required) { $col += ' NOT NULL' }
            $cols.Add($col)
            $made[$t.Name][$f.Name] = $true
        }
        if (-not $cols.Count) { $skipped.Add("$(Format-SqlName $t.Name): no field SQL can make"); continue }
        $statements.Add("CREATE TABLE $(Format-SqlName $t.Name) (`r`n" + ($cols -join ",`r`n") + "`r`n)")
    }
    foreach ($t in $tables) {
        if (-not $made.ContainsKey($t.Name) -or -not $made[$t.Name].Count) { continue }
        # The primary key first, so a foreign key can point at it.
        foreach ($ix in @($t.Indexes | Where-Object { -not $_.Foreign } | Sort-Object @{ Expression = { -not $_.Primary } }, Name)) {
            $missing = @($ix.Fields | Where-Object { -not $made[$t.Name].ContainsKey($_.Name) })
            if ($missing.Count) { $skipped.Add("index $(Format-SqlName $ix.Name) on $(Format-SqlName $t.Name): uses a field SQL can't make"); continue }
            $cols = ($ix.Fields | ForEach-Object { (Format-SqlName $_.Name) + $(if ($_.Descending) { ' DESC' } else { '' }) }) -join ', '
            $unique = if ($ix.Unique -and -not $ix.Primary) { 'UNIQUE ' } else { '' }
            # One WITH option at most: PRIMARY covers the other two.
            $with = if ($ix.Primary) { ' WITH PRIMARY' } elseif ($ix.Required) { ' WITH DISALLOW NULL' } elseif ($ix.IgnoreNulls) { ' WITH IGNORE NULL' } else { '' }
            $statements.Add("CREATE ${unique}INDEX $(Format-SqlName $ix.Name) ON $(Format-SqlName $t.Name) ($cols)$with")
        }
    }
    foreach ($r in @($Schema.Relations | Sort-Object Name)) {
        if ($r.Attributes -band 4) { continue }   # dbRelationInherited: belongs to a linked database
        if ((Test-SchemaSystemTable $r.Table) -or (Test-SchemaSystemTable $r.ForeignTable)) { continue }
        $label = "relationship $(Format-SqlName $r.Name) ($(Format-SqlName $r.ForeignTable) to $(Format-SqlName $r.Table))"
        if (-not $made.ContainsKey($r.Table) -or -not $made.ContainsKey($r.ForeignTable)) { $skipped.Add("${label}: a table is not in this database"); continue }
        if ($r.Attributes -band 2) { $skipped.Add("${label}: not enforced, which SQL can't make"); continue }   # dbRelationDontEnforce
        $missing = @($r.Fields | Where-Object { -not $made[$r.Table].ContainsKey($_.Name) -or -not $made[$r.ForeignTable].ContainsKey($_.ForeignName) })
        if ($missing.Count) { $skipped.Add("${label}: uses a field SQL can't make"); continue }
        $fk = ($r.Fields | ForEach-Object { Format-SqlName $_.ForeignName }) -join ', '
        $pk = ($r.Fields | ForEach-Object { Format-SqlName $_.Name }) -join ', '
        $s = "ALTER TABLE $(Format-SqlName $r.ForeignTable) ADD CONSTRAINT $(Format-SqlName $r.Name) FOREIGN KEY ($fk) REFERENCES $(Format-SqlName $r.Table) ($pk)"
        if ($r.Attributes -band 256) { $s += ' ON UPDATE CASCADE' }    # dbRelationUpdateCascade
        if ($r.Attributes -band 4096) { $s += ' ON DELETE CASCADE' }   # dbRelationDeleteCascade
        $statements.Add($s)
    }
    $head = @(
        "-- Tables of $Title as Access SQL (ANSI-92, the mode ADO uses).",
        '-- Run each statement on its own, in order, in an empty database, for example with',
        '-- CurrentProject.Connection.Execute. Statements end with a semicolon.',
        '-- Not included: data, captions, descriptions, formats, input masks, lookups,',
        '-- validation rules, Allow Zero Length and Unicode compression.'
    )
    if ($skipped.Count) {
        $head += '--'
        $head += "-- Not made by this script ($($skipped.Count)):"
        $head += @($skipped | ForEach-Object { "--   $_" })
    }
    $text = (($head -join "`r`n") + "`r`n`r`n" + (($statements | ForEach-Object { "$_;" }) -join "`r`n`r`n") + "`r`n")
    [pscustomobject]@{ Text = $text; Statements = $statements.ToArray(); Skipped = $skipped.ToArray() }
}

# A constraint or index name made from parts, kept to letters, digits and _.
function Format-SqlServerConstraint([string[]]$Parts) { Format-SqlName (($Parts -join '_') -replace '[^A-Za-z0-9_]', '_') }

# A field's type for SQL Server, or $null with -Reason set when there is none.
function ConvertTo-SqlServerType($Field, [ref]$Reason) {
    if ($Field.Expression) { $Reason.Value = 'calculated field'; return $null }
    switch ($Field.Type) {
        1 { 'BIT' }                     # Yes/No
        2 { 'TINYINT' }                 # Byte
        3 { 'SMALLINT' }                # Integer
        4 { 'INT' }                     # Long Integer
        5 { 'MONEY' }                   # Currency
        6 { 'REAL' }                    # Single
        7 { 'FLOAT' }                   # Double
        8 { 'DATETIME2(0)' }            # Date/Time (whole seconds, as Access keeps them)
        9 { "VARBINARY($([Math]::Max(1, $Field.Size)))" }
        10 { "NVARCHAR($([Math]::Max(1, $Field.Size)))" }   # Short Text
        11 { 'VARBINARY(MAX)' }         # OLE Object
        12 { 'NVARCHAR(MAX)' }          # Long Text (Memo, Hyperlink)
        15 { 'UNIQUEIDENTIFIER' }       # Replication ID
        16 { 'BIGINT' }                 # Large Number
        26 { 'DATETIME2(7)' }           # Date/Time Extended
        20 {
            if ($null -eq $Field.Precision) { $Reason.Value = 'decimal field whose precision DAO does not report'; return $null }
            "DECIMAL($($Field.Precision), $($Field.Scale))"
        }
        default { $Reason.Value = Get-SchemaTypeReason $Field.Type; $null }
    }
}

# A field's DefaultValue as a SQL Server default, or $null when it has none or calls a
# function SQL Server doesn't have (the script then says so in a comment).
function ConvertTo-SqlServerDefault($Field) {
    $d = ([string]$Field.Default).Trim()
    if ($d.StartsWith('=')) { $d = $d.Substring(1).Trim() }
    if (-not $d) { return $null }
    if ($Field.Type -eq 1) {
        if ($d -match '^(?i)(True|Yes|On|-1)$') { return '1' }
        if ($d -match '^(?i)(False|No|Off|0)$') { return '0' }
        return $null
    }
    if ($d -match '^(?i)Now\(\)$') { return 'SYSDATETIME()' }
    if ($d -match '^(?i)Date\(\)$') { return 'CAST(GETDATE() AS date)' }
    if ($d -match '^-?\d+(\.\d+)?$') { return $d }
    if ($d -match '^"((?:[^"]|"")*)"$') { return "N'" + $Matches[1].Replace('""', '"').Replace("'", "''") + "'" }
    if ($d -match "^'((?:[^']|'')*)'$") { return "N'" + $Matches[1] + "'" }
    $null
}

# True if a cascade from $Parent to $Child would give SQL Server a loop or a second
# cascade route between two tables (error 1785), given the cascades already made
# ($Graph: table -> tables it cascades to). Access allows both; SQL Server neither.
function Test-SqlServerCascadeClash($Graph, [string]$Parent, [string]$Child) {
    $reach = {
        param([string]$From, [bool]$Back)
        $seen = @{ $From = $true }; $todo = [System.Collections.Generic.Queue[string]]::new(); $todo.Enqueue($From)
        while ($todo.Count) {
            $n = $todo.Dequeue()
            $next = if ($Back) { @($Graph.Keys | Where-Object { $Graph[$_] -contains $n }) } else { @($Graph[$n]) }
            foreach ($m in $next) { if ($m -and -not $seen.ContainsKey($m)) { $seen[$m] = $true; $todo.Enqueue($m) } }
        }
        $seen
    }
    $above = & $reach $Parent $true    # the parent and every table that cascades to it
    $below = & $reach $Child $false    # the child and every table it cascades to
    foreach ($a in $above.Keys) {
        if ($below.ContainsKey($a)) { return $true }   # a loop
        $fromA = & $reach $a $false
        foreach ($b in $below.Keys) { if ($fromA.ContainsKey($b)) { return $true } }   # a second route
    }
    $false
}

# Notes in $Indexed that $Index serves every leading run of its columns.
function Add-SqlServerIndexed($Indexed, [string]$Table, $Index) {
    if (-not $Indexed.ContainsKey($Table)) { $Indexed[$Table] = @{} }
    $names = @($Index.Fields | ForEach-Object { $_.Name })
    for ($n = 1; $n -le $names.Count; $n++) { $Indexed[$Table][($names[0..($n - 1)] -join '|').ToLowerInvariant()] = $true }
}

# The schema (Read-AccessSchema) as a T-SQL script: CREATE TABLE for each table with its
# primary key and defaults, then its other indexes, then the relationships as foreign
# keys with an index each (last, so table order doesn't matter). Returns Text and Skipped, like
# ConvertTo-SchemaSql; what can't be carried over is also a comment where it would go.
function ConvertTo-SqlServerSql {
    param([Parameter(Mandatory)]$Schema, [string]$Title = 'database')
    $body = New-Object System.Collections.Generic.List[string]
    $skipped = New-Object System.Collections.Generic.List[string]
    $made = @{}
    $indexed = @{}   # table -> the leading column lists of its indexes, so a foreign key isn't indexed twice
    foreach ($t in @($Schema.Tables | Sort-Object Name)) {
        $cols = New-Object System.Collections.Generic.List[string]
        $made[$t.Name] = @{}
        # SQL Server refuses a primary key on a column that allows NULL; Access doesn't
        # need Required set on a key field, since the key already forbids empty values.
        $keyNames = @(@($t.Indexes | Where-Object { $_.Primary }) | Select-Object -First 1 | ForEach-Object { $_.Fields } | ForEach-Object { $_.Name })
        foreach ($f in $t.Fields) {
            $why = $null
            $type = ConvertTo-SqlServerType $f ([ref]$why)
            if (-not $type) {
                $skipped.Add("$(Format-SqlName $t.Name).$(Format-SqlName $f.Name): $why")
                $body.Add("-- $(Format-SqlName $t.Name).$(Format-SqlName $f.Name): $why, not created.")
                continue
            }
            $c = "    $(Format-SqlName $f.Name) $type"
            if ($f.AutoIncrement) { $c += ' IDENTITY(1,1)' }
            # Access keeps Yes/No fields filled (never empty), so they are NOT NULL.
            $c += $(if ($f.Required -or $f.AutoIncrement -or $f.Type -eq 1 -or $keyNames -contains $f.Name) { ' NOT NULL' } else { ' NULL' })
            $def = if ($f.AutoIncrement) { $null } else { ConvertTo-SqlServerDefault $f }
            if (-not $def -and $f.Type -eq 1 -and -not $f.AutoIncrement) { $def = '0' }
            if ($def) { $c += " CONSTRAINT $(Format-SqlServerConstraint 'DF', $t.Name, $f.Name) DEFAULT ($def)" }
            elseif ($f.Default -and -not $f.AutoIncrement) {
                $skipped.Add("$(Format-SqlName $t.Name).$(Format-SqlName $f.Name): default $($f.Default) not converted")
                $body.Add("-- $(Format-SqlName $t.Name).$(Format-SqlName $f.Name): Access default $($f.Default) not converted.")
            }
            # Access refuses "" in a required text field unless Allow Zero Length is on.
            if ($f.Type -eq 10 -and $f.Required -and -not $f.AllowZeroLength) { $c += " CHECK (LEN($(Format-SqlName $f.Name)) > 0)" }
            $cols.Add($c)
            $made[$t.Name][$f.Name] = $true
        }
        if (-not $cols.Count) { $skipped.Add("$(Format-SqlName $t.Name): no field SQL can make"); continue }
        $pk = @($t.Indexes | Where-Object { $_.Primary }) | Select-Object -First 1
        if ($pk -and -not @($pk.Fields | Where-Object { -not $made[$t.Name].ContainsKey($_.Name) }).Count) {
            $cols.Add("    CONSTRAINT $(Format-SqlServerConstraint 'PK', $t.Name) PRIMARY KEY ($((@($pk.Fields | ForEach-Object { Format-SqlName $_.Name })) -join ', '))")
            Add-SqlServerIndexed $indexed $t.Name $pk
        }
        $body.Add("CREATE TABLE $(Format-SqlName $t.Name) (`r`n" + ($cols -join ",`r`n") + "`r`n);")
        foreach ($ix in @($t.Indexes | Where-Object { -not $_.Primary -and -not $_.Foreign } | Sort-Object Name)) {
            if (@($ix.Fields | Where-Object { -not $made[$t.Name].ContainsKey($_.Name) }).Count) {
                $skipped.Add("index $(Format-SqlName $ix.Name) on $(Format-SqlName $t.Name): uses a field SQL can't make"); continue
            }
            $ic = @($ix.Fields | ForEach-Object { (Format-SqlName $_.Name) + $(if ($_.Descending) { ' DESC' } else { '' }) }) -join ', '
            $where = ''
            if ($ix.Unique) {
                # Access allows any number of empty values in a unique index; SQL Server one.
                $nullable = @(foreach ($xf in $ix.Fields) {
                    $fd = @($t.Fields | Where-Object { $_.Name -eq $xf.Name }) | Select-Object -First 1
                    if ($fd -and -not ($fd.Required -or $fd.AutoIncrement -or $fd.Type -eq 1)) { (Format-SqlName $xf.Name) + ' IS NOT NULL' }
                })
                if ($nullable.Count) { $where = ' WHERE ' + ($nullable -join ' AND ') }
            }
            $kind = if ($ix.Unique) { 'CREATE UNIQUE INDEX' } else { 'CREATE INDEX' }
            $body.Add("$kind $(Format-SqlServerConstraint 'IX', $t.Name, $ix.Name) ON $(Format-SqlName $t.Name) ($ic)$where;")
            if (-not $where) { Add-SqlServerIndexed $indexed $t.Name $ix }
        }
        $body.Add('GO')
        $body.Add('')
    }
    $rels = New-Object System.Collections.Generic.List[object]
    foreach ($r in @($Schema.Relations | Sort-Object Name)) {
        if ($r.Attributes -band 4) { continue }   # dbRelationInherited: belongs to a linked database
        if ((Test-SchemaSystemTable $r.Table) -or (Test-SchemaSystemTable $r.ForeignTable)) { continue }
        $label = "relationship $(Format-SqlName $r.Name) ($(Format-SqlName $r.ForeignTable) to $(Format-SqlName $r.Table))"
        if (-not $made.ContainsKey($r.Table) -or -not $made.ContainsKey($r.ForeignTable)) { $skipped.Add("${label}: a table is not in this database"); continue }
        if (@($r.Fields | Where-Object { -not $made[$r.Table].ContainsKey($_.Name) -or -not $made[$r.ForeignTable].ContainsKey($_.ForeignName) }).Count) { $skipped.Add("${label}: uses a field SQL can't make"); continue }
        $rels.Add([pscustomobject]@{ R = $r; Label = $label; Enforced = -not ($r.Attributes -band 2) })   # dbRelationDontEnforce
    }
    # SQL Server refuses a second cascade route between two tables, and loops (error 1785);
    # Access allows both. Cascades are kept one step at a time down the chains, so a
    # shortcut (A to C where A to B to C also cascades) is the one that becomes a plain
    # foreign key: deleting A still reaches C through B.
    $depth = @{}
    $cascading = @($rels | Where-Object { $_.Enforced -and ($_.R.Attributes -band (4096 + 256)) })
    for ($pass = 0; $pass -lt $made.Count; $pass++) {   # longest cascade chain to each table; loops stop at the pass limit
        $changed = $false
        foreach ($x in $cascading) {
            $d = 1 + [int]$depth[$x.R.Table]
            if ($d -gt [int]$depth[$x.R.ForeignTable]) { $depth[$x.R.ForeignTable] = $d; $changed = $true }
        }
        if (-not $changed) { break }
    }
    $dropped = @{}
    foreach ($c in @(@{ Bit = 4096; Sql = 'ON DELETE CASCADE' }, @{ Bit = 256; Sql = 'ON UPDATE CASCADE' })) {   # dbRelationDeleteCascade, dbRelationUpdateCascade
        $graph = @{}
        foreach ($x in @($cascading | Where-Object { $_.R.Attributes -band $c.Bit } | Sort-Object @{ Expression = { [int]$depth[$_.R.ForeignTable] - [int]$depth[$_.R.Table] } }, @{ Expression = { $_.R.Name } })) {
            if (Test-SqlServerCascadeClash $graph $x.R.Table $x.R.ForeignTable) { $dropped["$($x.R.Name)|$($c.Bit)"] = $true; continue }
            if (-not $graph.ContainsKey($x.R.Table)) { $graph[$x.R.Table] = New-Object System.Collections.Generic.List[string] }
            $graph[$x.R.Table].Add($x.R.ForeignTable)
        }
    }
    foreach ($x in $rels) {
        $r = $x.R
        $fk = @($r.Fields | ForEach-Object { Format-SqlName $_.ForeignName }) -join ', '
        $pkc = @($r.Fields | ForEach-Object { Format-SqlName $_.Name }) -join ', '
        $st = "ALTER TABLE $(Format-SqlName $r.ForeignTable) ADD CONSTRAINT $(Format-SqlServerConstraint 'FK', $r.Name) FOREIGN KEY ($fk) REFERENCES $(Format-SqlName $r.Table) ($pkc)"
        $notes = New-Object System.Collections.Generic.List[string]
        foreach ($c in @(@{ Bit = 4096; Sql = 'ON DELETE CASCADE' }, @{ Bit = 256; Sql = 'ON UPDATE CASCADE' })) {
            if (-not ($r.Attributes -band $c.Bit)) { continue }
            if ($dropped.ContainsKey("$($r.Name)|$($c.Bit)")) {
                $skipped.Add("$($x.Label): $($c.Sql) left off, a second cascade route or a loop, which SQL Server refuses")
                $notes.Add("-- $(Format-SqlName $r.Name): $($c.Sql) in Access, left off here: SQL Server allows one cascade route between two tables and no loops.")
            } else { $st += " $($c.Sql)" }
        }
        if (-not $x.Enforced) {
            $skipped.Add("$($x.Label): not enforced in Access, kept as a comment")
            $body.Add("-- Not enforced in Access (no referential integrity):`r`n-- $st;")
            continue
        }
        $body.AddRange($notes)
        $body.Add("$st;")
        # Access indexes a foreign key itself; SQL Server doesn't, so add one unless an
        # index already starts with the same columns.
        $key = (@($r.Fields | ForEach-Object { $_.ForeignName }) -join '|').ToLowerInvariant()
        if (-not $indexed.ContainsKey($r.ForeignTable)) { $indexed[$r.ForeignTable] = @{} }
        if (-not $indexed[$r.ForeignTable].ContainsKey($key)) {
            $indexed[$r.ForeignTable][$key] = $true
            $body.Add("CREATE INDEX $(Format-SqlServerConstraint 'IX', $r.ForeignTable, $r.Name) ON $(Format-SqlName $r.ForeignTable) ($fk);")
        }
        $body.Add('GO')
    }
    $head = @(
        "-- Tables of $Title for SQL Server (T-SQL).",
        '-- A unique index on a column that may be empty is filtered (WHERE ... IS NOT NULL),',
        '-- because Access allows any number of empty values in a unique index and SQL Server one.',
        '-- Batches end with GO: run it with sqlcmd -i or in SSMS.',
        '-- Each foreign key gets an index (Access makes one itself). A cascade that would be a',
        '-- second route between two tables, or a loop, is left off: SQL Server refuses those.',
        '-- Not included: data, captions, descriptions, formats, input masks, lookups and validation rules.'
    )
    if ($skipped.Count) {
        $head += '--'
        $head += "-- Not made by this script ($($skipped.Count)):"
        $head += @($skipped | ForEach-Object { "--   $_" })
    }
    [pscustomobject]@{ Text = (($head -join "`r`n") + "`r`n`r`n" + (($body -join "`r`n").TrimEnd() + "`r`n")); Skipped = $skipped.ToArray() }
}

# Runs a build script's statements into a new, empty database in -WorkDir, reads that
# database back and returns the statements that differ (in the script but not rebuilt,
# '-', or rebuilt but not in the script, '+'). Nothing returned means the script
# rebuilds the tables exactly as far as SQL can describe them.
#
# The database is made with DAO and the statements go through their own ADO connection
# (ANSI-92, as the script is written for), not through Access's CurrentProject: a DEFAULT
# that calls a function, such as Now(), made that way crashes Access in VBE7.DLL when
# the database closes. Each statement is printed before it runs, so a hang names it.
function Test-SchemaSqlRebuild {
    param([Parameter(Mandatory)]$Access, [Parameter(Mandatory)]$Sql, [Parameter(Mandatory)][string]$WorkDir)
    $tmp = Join-Path $WorkDir ('rebuild-' + [guid]::NewGuid().ToString('N') + '.accdb')
    $Access.DBEngine.CreateDatabase([string]$tmp, ';LANGID=0x0409;CP=1252;COUNTRY=0', 128).Close()   # dbLangGeneral, dbVersion120
    try {
        $cn = New-Object -ComObject ADODB.Connection
        try {
            $opened = $false
            foreach ($provider in 'Microsoft.ACE.OLEDB.12.0', 'Microsoft.ACE.OLEDB.16.0') {
                try { $cn.Open("Provider=$provider;Data Source=$tmp"); $opened = $true; break } catch {}
            }
            if (-not $opened) { throw "Could not open $tmp with the Access database engine's OLE DB provider (ACE.OLEDB 12.0 or 16.0)." }
            $n = @($Sql.Statements).Count; $i = 0
            foreach ($st in $Sql.Statements) {
                $i++
                Write-Host ("  rebuild {0}/{1}: {2}" -f $i, $n, (($st -split "`r?`n")[0]))
                try { [void]$cn.Execute($st) }
                catch { throw "Rebuilding from the SQL failed at: $st`r`n$($_.Exception.Message)" }
            }
        } finally {
            if ($cn.State) { $cn.Close() }
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($cn)
        }
        $again = ConvertTo-SchemaSql (Read-AccessSchema -Access $Access -Path $tmp)
    } finally { Remove-Item -LiteralPath $tmp -ErrorAction SilentlyContinue }
    @(Compare-Object @($Sql.Statements) @($again.Statements) | ForEach-Object {
        $(if ($_.SideIndicator -eq '<=') { '- ' } else { '+ ' }) + ($_.InputObject -replace "`r`n\s*", ' ')
    })
}

# Saves every form of a front end with SaveAsText into -SnapshotDir (FormSnapshot's
# layout: <form>.txt, or <form>.failed with the error). The forms are imported into a
# blank database with the code off, so nothing in the app runs: not a StartUpForm and
# not an AutoExec macro (the kit's trust check starts the app through one).
function Save-ReleaseFormText {
    param([Parameter(Mandatory)]$Access, [Parameter(Mandatory)][string]$FrontEnd, [Parameter(Mandatory)][string]$SnapshotDir, [Parameter(Mandatory)][string]$WorkDir)
    $blank = Join-Path $WorkDir ('forms-' + [guid]::NewGuid().ToString('N') + '.accdb')
    New-Item -ItemType Directory -Force -Path $SnapshotDir | Out-Null
    $db = $Access.DBEngine.OpenDatabase([string]$FrontEnd, $false, $true)
    try {
        $docs = $db.Containers.Item('Forms').Documents
        $names = @(for ($i = 0; $i -lt $docs.Count; $i++) { [string]$docs.Item($i).Name })
    } finally { $db.Close() }
    $security = $Access.AutomationSecurity
    $failed = New-Object System.Collections.Generic.List[string]
    try { $Access.CloseCurrentDatabase() } catch {}
    $Access.AutomationSecurity = 3   # msoAutomationSecurityForceDisable: no VBA runs
    try {
        $Access.NewCurrentDatabase([string]$blank)
        try {
            foreach ($n in $names) {
                if ($n -match '[\\/:*?"<>|]') { $failed.Add("${n}: the name has a character that can't be used in a file name"); continue }
                try {
                    $Access.DoCmd.TransferDatabase(0, 'Microsoft Access', [string]$FrontEnd, 2, $n, $n)   # acImport, acForm
                    $Access.SaveAsText(2, $n, [string](Join-Path $SnapshotDir "$n.txt"))                   # acForm
                } catch {
                    [IO.File]::WriteAllText((Join-Path $SnapshotDir "$n.failed"), $_.Exception.Message, (New-Object Text.UTF8Encoding($false)))
                    $failed.Add("${n}: $($_.Exception.Message)")
                }
            }
        } finally { $Access.CloseCurrentDatabase() }
    } finally {
        $Access.AutomationSecurity = $security
        Remove-Item -LiteralPath $blank -ErrorAction SilentlyContinue
    }
    [pscustomobject]@{ Forms = $names.Count; Failed = $failed.ToArray() }
}

# Writes a release's design folder (see the top of this file): the front end's forms as
# JSON in <OutDir>\forms and, for the front end's own tables (when it has any) and each
# data file, <database>.sqlserver.sql and, with -Dialect Access, <database>.sql. Each
# Access script is rebuilt into an empty database and read back. Throws if a form can't
# be exported or an Access script doesn't rebuild its tables, so a release never ships a
# partial design. Returns what it wrote, with what SQL can't make (attachment fields and
# the like) in Skipped.
function Export-ReleaseDesign {
    param(
        [Parameter(Mandatory)]$Access,
        [Parameter(Mandatory)][string]$FrontEnd,
        [string[]]$DataFiles,
        [string]$ThemeFile,
        [Parameter(Mandatory)][string]$OutDir,
        [ValidateSet('Access', 'SqlServer')][string[]]$Dialect = @('SqlServer'),
        [string]$WorkDir = $env:TEMP
    )
    if (Test-Path -LiteralPath $OutDir) { throw "$OutDir already exists." }
    $work = Join-Path $WorkDir ('release-design-' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $work, $OutDir | Out-Null
    try {
        $snap = Join-Path $work 'forms'
        $saved = Save-ReleaseFormText -Access $Access -FrontEnd $FrontEnd -SnapshotDir $snap -WorkDir $work
        if ($saved.Failed.Count) { throw "Could not export these forms:`r`n  " + ($saved.Failed -join "`r`n  ") }
        $designs = Export-FormDesignFolder -SnapshotDir $snap -OutDir (Join-Path $OutDir 'forms') -ThemeFile $ThemeFile
        if ($designs.Failed.Count) { throw "Could not write these form designs:`r`n  " + (($designs.Failed | ForEach-Object { "$($_.form): $($_.error)" }) -join "`r`n  ") }

        $scripts = @(); $skipped = @()
        foreach ($file in @($FrontEnd) + @($DataFiles | Where-Object { $_ })) {
            $schema = Read-AccessSchema -Access $Access -Path $file
            if (-not $schema.Tables.Count) { continue }   # a front end with linked tables only
            $leaf = Split-Path -Leaf $file
            $base = [IO.Path]::GetFileNameWithoutExtension($leaf)
            if ($Dialect -contains 'Access') {
                $sql = ConvertTo-SchemaSql $schema -Title $leaf
                $diff = @(Test-SchemaSqlRebuild -Access $Access -Sql $sql -WorkDir $work)
                if ($diff.Count) { throw "The SQL for $leaf does not rebuild its tables:`r`n  " + ($diff -join "`r`n  ") }
                [IO.File]::WriteAllText((Join-Path $OutDir "$base.sql"), $sql.Text, (New-Object Text.UTF8Encoding($false)))
                $scripts += [pscustomobject]@{ File = "$base.sql"; Tables = @($schema.Tables).Count }
                $skipped += @($sql.Skipped | ForEach-Object { "$base.sql: $_" })
            }
            if ($Dialect -contains 'SqlServer') {
                $ss = ConvertTo-SqlServerSql $schema -Title $leaf
                [IO.File]::WriteAllText((Join-Path $OutDir "$base.sqlserver.sql"), $ss.Text, (New-Object Text.UTF8Encoding($false)))
                $scripts += [pscustomobject]@{ File = "$base.sqlserver.sql"; Tables = @($schema.Tables).Count }
                $skipped += @($ss.Skipped | ForEach-Object { "$base.sqlserver.sql: $_" })
            }
        }
        [pscustomobject]@{ Forms = $designs.Written.Count; Scripts = $scripts; Skipped = $skipped }
    } finally { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
}
