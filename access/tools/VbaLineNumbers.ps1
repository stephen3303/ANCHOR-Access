# Adds VBA line numbers to the build copy of a module, so Erl (and so the error
# log) says which line failed. Dot-source it from a build script:
#
#   . (Join-Path $repo 'tools\VbaLineNumbers.ps1')
#   $text = Add-VbaLineNumbers $text            # one module's text
#   Add-VbaLineNumbersToFolder $stage\forms     # every .vba/.bas file in a folder
#
# Only the copy loaded into Access is numbered; the source in git never is.
# Each statement is numbered with its own line number in the source file, so
# "line 42" in the error log is line 42 of that file in the repository.
#
# Numbered: executable lines inside a Sub, Function or Property.
# Left alone: the declarations section, procedure headers and End lines, blank
# and comment lines, continuation lines, lines that start with a label (even with
# a statement after it), Dim / Static / Const, #If directives, and block lines
# VBA may not accept a label on (Case, Else, ElseIf, End If, End Select, End With,
# Loop, Next, Wend). An error on one of those is logged at the numbered line
# before it, e.g. the Select Case line for a failing "Case 1: DoThing".
# A module that already has line numbers (numbered by hand, or numbered before)
# is returned unchanged, so generated numbers never collide with its own.

$script:VbaProcStart = '^\s*((Public|Private|Friend)\s+)?(Static\s+)?(Sub|Function|Property\s+(Get|Let|Set))\s+[A-Za-z_]'
$script:VbaProcEnd = '^\s*End\s+(Sub|Function|Property)\b'
$script:VbaSkip = @(
    '^\s*$',                                                    # blank
    "^\s*('|Rem\b)",                                            # comment
    '^\s*[A-Za-z_][A-Za-z0-9_]*:(?!=)',                         # label (Err_Handler:, Retry: DoWork)
    '^\s*(Dim|Static|Const)\s',                                 # declarations
    '^\s*#',                                                    # #If / #Else / #End If
    '^\s*Attribute\s',                                          # VB_ attributes
    '^\s*(Case|Else|ElseIf|Loop|Next|Wend)\b',                  # block lines
    '^\s*End\s+(If|Select|With)\b'
)

# Number one module's text. Accepts CRLF or LF and returns CRLF text.
function Add-VbaLineNumbers([string]$Text) {
    $lines = $Text -split "`r?`n"
    $numbered = @($lines -match '^\s*\d+(\s|:|$)')
    if ($numbered.Count -gt 0) {
        Write-Warning "Line numbers not added to a module that already has some: $($numbered[0].Trim())"
        return ($lines -join "`r`n")
    }
    $inProc = $false
    $continued = $false     # the previous line ended with " _"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        $wasContinued = $continued
        $continued = $line -match '(^|\s)_\s*$'
        if ($wasContinued) { continue }
        if (-not $inProc) {
            if ($line -match $script:VbaProcStart) { $inProc = $true }
            continue
        }
        if ($line -match $script:VbaProcEnd) { $inProc = $false; continue }
        $skip = $false
        foreach ($p in $script:VbaSkip) { if ($line -match $p) { $skip = $true; break } }
        if ($skip) { continue }
        # VBA line numbers stop at 65529 for Erl; modules never get that long.
        $lines[$i] = [string]($i + 1) + ' ' + $line
    }
    return ($lines -join "`r`n")
}

# Number every .bas and .vba file in a folder in place. Only call it on a
# staged copy, never on the source folder. Written as UTF-8 without a BOM, as
# SaveForm reads them (ReadTextFile), so non-ASCII text survives.
function Add-VbaLineNumbersToFolder([string]$Folder) {
    $utf8 = New-Object Text.UTF8Encoding $false
    foreach ($f in @(Get-ChildItem $Folder -Include *.bas, *.vba -Recurse)) {
        $text = Add-VbaLineNumbers ([IO.File]::ReadAllText($f.FullName))
        [IO.File]::WriteAllText($f.FullName, $text, $utf8)
    }
}
