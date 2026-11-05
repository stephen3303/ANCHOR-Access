# Shared by new-backend.ps1 and seed-backend.ps1: loads modBuildBackEnd.bas
# into a scratch database (deleted afterwards), compiles it, and runs one of
# its script entry points, stopping on a compile error or any Access dialog
# instead of hanging behind it (tools\AccessWatch.ps1).
#
#   . .\BackEndBuilder.ps1
#   $result = Invoke-BackEndBuilder -Entry 'SeedBackEndAt' -Arguments @($path, $seed, '')
#
# Returns the entry point's text ("OK..." or "ERROR during ..."). Throws when
# the module doesn't compile or Access showed a dialog.
# Requires Microsoft Access (2016+ / Microsoft 365).

function Invoke-BackEndBuilder {
    param(
        [Parameter(Mandatory)][string]$Entry,
        [Parameter(Mandatory)][AllowEmptyString()][string[]]$Arguments,
        [int]$TimeoutSec = 300
    )
    if ($Arguments.Count -lt 2 -or $Arguments.Count -gt 3) { throw "Invoke-BackEndBuilder takes 2 or 3 arguments for $Entry." }
    $backend = $PSScriptRoot
    $here = Split-Path -Parent $backend
    . (Join-Path $here 'tools\AccessWatch.ps1')

    # The builder runs inside a scratch database that is deleted afterwards.
    $scratch = Join-Path $env:TEMP 'ANCHOR_backend_builder.accdb'
    Remove-Item $scratch -ErrorAction SilentlyContinue

    $acc = New-Object -ComObject Access.Application
    $watch = Start-AccessWatch $acc -TimeoutSec $TimeoutSec
    $result = $null; $problem = $null
    try {
        $acc.Visible = $false
        $acc.AutomationSecurity = 1
        $acc.NewCurrentDatabase($scratch)
        # LoadFromText expects Windows line endings.
        $tmp = Join-Path $env:TEMP 'anchor_modBuildBackEnd.bas'
        $text = [IO.File]::ReadAllText((Join-Path $backend 'modBuildBackEnd.bas')) -replace "`r?`n", "`r`n"
        [IO.File]::WriteAllText($tmp, $text, [Text.Encoding]::ASCII)
        $acc.LoadFromText(5, 'modBuildBackEnd', $tmp)
        Remove-Item $tmp

        $bad = Test-AccessCompiles $acc $watch
        if ($bad) { $problem = "Compile error: $bad" }
        else {
            # Access.Run passes each argument ByRef, so each needs its own variable.
            $a1 = [string]$Arguments[0]; $a2 = [string]$Arguments[1]
            if ($Arguments.Count -eq 3) {
                $a3 = [string]$Arguments[2]
                $result = $acc.Run($Entry, [ref]$a1, [ref]$a2, [ref]$a3)
            }
            else { $result = $acc.Run($Entry, [ref]$a1, [ref]$a2) }
        }
    }
    catch { if (-not $watch.Killed) { $problem = $_.Exception.Message } }
    finally {
        $dialogs = Stop-AccessWatch $watch $acc
        Remove-Item $scratch -ErrorAction SilentlyContinue
    }
    if ($dialogs) { throw 'Stopped: Access showed the dialog(s) above.' }
    if ($watch.Killed -and -not $result) { throw "Stopped: $Entry took longer than $TimeoutSec seconds." }
    if ($problem) { throw $problem }
    [string]$result
}
