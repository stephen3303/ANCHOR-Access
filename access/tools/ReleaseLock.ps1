# Release lock: the start-up settings a released, locked front end (.accde) must have, so
# a user can't get past the app's own start-up into the database behind it.
#
#   AllowBypassKey      False  holding Shift while the file opens no longer skips the
#                              start-up (AutoExec, the start-up form) and shows the
#                              navigation pane
#   StartUpShowDBWindow False  the navigation pane stays hidden behind the start-up
#   AllowSpecialKeys    False  F11 can't show the navigation pane again, and Ctrl+G and
#                              Ctrl+Break can't reach the code
#
# The build sets none of them on the .accdb, which stays the full, editable program a
# developer opens with Shift held. release.ps1 locks the .accde after SysCmd 603 makes it
# and before it is compacted:
#
#   . "$PSScriptRoot\tools\ReleaseLock.ps1"
#   Lock-ReleaseFrontEnd -Access $acc -Path $accde
#
# DAO opens the file without making it the current database, so no start-up code runs.
# Each property is read back, and a value that didn't stick stops the release.

# The settings, name and value. AllowBypassKey is created as a DDL property: only then can
# a user without design rights not set it back from another database.
$script:ReleaseLockSettings = [ordered]@{
    AllowBypassKey      = $false
    StartUpShowDBWindow = $false
    AllowSpecialKeys    = $false
}

# Set a Yes/No database property on an open DAO database, creating it when the database
# doesn't have it yet (3270, property not found).
function Set-DaoBooleanProperty($Database, [string]$Name, [bool]$Value) {
    try {
        $Database.Properties.Item($Name).Value = $Value
    } catch {
        # 1 is dbBoolean; the 4th argument makes it a DDL property.
        $Database.Properties.Append($Database.CreateProperty($Name, 1, $Value, $true))
    }
}

# Lock the front end at -Path (see the top of the file). -Access is the Access.Application
# the release already runs; its DBEngine opens the file exclusively.
function Lock-ReleaseFrontEnd {
    param(
        [Parameter(Mandatory)] $Access,
        [Parameter(Mandatory)] [string]$Path
    )
    if (-not (Test-Path -LiteralPath $Path)) { throw "No front end to lock at $Path." }
    $db = $Access.DBEngine.OpenDatabase([string]$Path, $true, $false)
    try {
        Set-ReleaseLock $db
    } finally {
        $db.Close()
    }
}

# The work of Lock-ReleaseFrontEnd on a database already open, for the tests.
function Set-ReleaseLock($Database) {
    foreach ($name in $script:ReleaseLockSettings.Keys) {
        Set-DaoBooleanProperty $Database $name $script:ReleaseLockSettings[$name]
    }
    foreach ($name in $script:ReleaseLockSettings.Keys) {
        $now = $null
        try { $now = $Database.Properties.Item($name).Value } catch { }
        if ($null -eq $now -or [bool]$now -ne $script:ReleaseLockSettings[$name]) {
            throw "The front end's $name didn't stay $($script:ReleaseLockSettings[$name]); the release is not locked."
        }
    }
}
