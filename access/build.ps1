# Builds the ANCHOR front end (build\ANCHOR.accdb, next to this
# script) straight from this folder: the kit's modules, forms and text-measure
# adapter (kit\), and the app's own code, forms and help (src\).
#
#   build.ps1                     build a new front end with no data linked (UI work)
#   build.ps1 -BackEnd <path>     build a new front end linked to that shared data file
#   build.ps1 -UiOnly             rebuild code and forms inside the existing front end,
#                                 keeping its links (also -FrontEndOnly)
#   -OutDir <folder>              build somewhere else than .\build
#   -NoLineNumbers                load the code without line numbers (see below)
#   -Release                      also make the locked copy users get,
#                                 ANCHOR.accde (see below)
#
# The code loaded into the front end gets line numbers (tools\VbaLineNumbers.ps1)
# so the error log names the failing line; the files in git are not changed.
# The build stops and reports any Visual Basic or Access dialog (a compile
# error, an error message) instead of hanging behind it.
#
# -Release, after the build: Access makes ANCHOR.accde from the
# .accdb (SysCmd 603: compiled, with no code or form design left in it to
# change), and tools\ReleaseLock.ps1 turns off, in the .accde only,
# AllowBypassKey (Shift held while it opens skips the start-up and sign-in)
# and AllowSpecialKeys (F11 shows the tables, Ctrl+G the code), and keeps
# the navigation pane hidden. Only the .accde goes to users. The .accdb stays
# the developer's copy, unlocked as before.
# Requires Microsoft Access (2016+ / Microsoft 365). Python is only needed to
# regenerate the palettes (python tools\make_themes.py src).
param(
    [Alias('FrontEndOnly')][switch]$UiOnly,
    [string]$BackEnd = '',
    [string]$OutDir = '',
    [switch]$NoLineNumbers,
    [switch]$Release
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $OutDir) { $OutDir = Join-Path $here 'build' }
New-Item -ItemType Directory -Force $OutDir | Out-Null
$OutDir = (Resolve-Path $OutDir).Path
$fe = Join-Path $OutDir 'ANCHOR.accdb'
$accde = Join-Path $OutDir 'ANCHOR.accde'
$ProjectName = 'ANCHORApp'   # the VBA project's name; never 'ANCHOR' (see below)
if ($BackEnd) {
    if (-not (Test-Path $BackEnd)) { throw "No shared data file at $BackEnd" }
    $BackEnd = (Resolve-Path $BackEnd).Path
}

# SaveForm attaches <srcRoot>\forms\<name>.vba, so stage the app's forms and
# the kit's forms (subGrid, dlgAppearance, dlgDatePick and the rest) in one
# folder, with the help topics (LoadHelpTopics). The folder is this build's
# own, so two builds at once don't overwrite each other's files.
$stage = Join-Path $env:TEMP ('ANCHOR_src_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Force (Join-Path $stage 'forms') | Out-Null
Copy-Item (Join-Path $here 'kit\forms\*.vba') (Join-Path $stage 'forms')
Copy-Item (Join-Path $here 'src\forms\*.vba') (Join-Path $stage 'forms')
Copy-Item (Join-Path $here 'src\help') $stage -Recurse
. (Join-Path $here 'tools\VbaLineNumbers.ps1')
if (-not $NoLineNumbers) { Add-VbaLineNumbersToFolder (Join-Path $stage 'forms') }

. (Join-Path $here 'tools\AccessWatch.ps1')
. (Join-Path $here 'tools\ReleaseLock.ps1')
$acc = New-Object -ComObject Access.Application
$watch = Start-AccessWatch $acc -TimeoutSec 600
$failed = $false
try {
    $acc.Visible = $false
    $acc.AutomationSecurity = 1   # let the build's own VBA run
    if ($UiOnly) {
        if (-not (Test-Path $fe)) { throw "No existing front end at $fe (build once without -UiOnly)" }
        $acc.OpenCurrentDatabase($fe)
    } else {
        Remove-Item $fe -ErrorAction SilentlyContinue
        $acc.NewCurrentDatabase($fe)
    }
    # Access names the VBA project after the file, so ANCHOR.accdb would make a
    # project called ANCHOR, which hides the kit's Anchor procedure (VBA ignores
    # case): "Expected variable or procedure, not project". Give it its own name.
    if ($acc.VBE.ActiveVBProject.Name -ne $ProjectName) { $acc.VBE.ActiveVBProject.Name = $ProjectName }

    # The kit first, then its text-measure adapter, the app's modules and the
    # build palette.
    $files = @(Get-ChildItem (Join-Path $here 'kit\modules') -Filter *.bas) +
             @(Get-Item (Join-Path $here 'kit\adapters\modTextMeasure.bas')) +
             @(Get-ChildItem (Join-Path $here 'src\modules') -Filter *.bas) +
             @(Get-Item (Join-Path $here 'src\themes\light\modTheme.bas'))
    foreach ($f in $files) {
        try { $acc.DoCmd.DeleteObject(5, $f.BaseName) } catch {}
        # LoadFromText expects Windows line endings. The code is written out as
        # ASCII, so a character outside it would quietly become "?": stop instead.
        $tmp = Join-Path $stage ("anchor_" + $f.Name)
        $text = [IO.File]::ReadAllText($f.FullName) -replace "`r?`n", "`r`n"
        if ($text -match '[^\x00-\x7F]') { throw "$($f.Name) has a character outside ASCII; write it with ChrW() instead." }
        if (-not $NoLineNumbers) { $text = Add-VbaLineNumbers $text }
        [IO.File]::WriteAllText($tmp, $text, [Text.Encoding]::ASCII)
        $acc.LoadFromText(5, $f.BaseName, $tmp)
        Remove-Item $tmp
    }

    $bad = Test-AccessCompiles $acc $watch
    if ($bad) { Write-Host "Compile error: $bad"; throw 'compile failed' }

    $a1 = [string]$stage; $a2 = [string]$OutDir; $a3 = [bool]$UiOnly; $a4 = [string]$BackEnd
    $result = $acc.Run('BuildAll', [ref]$a1, [ref]$a2, [ref]$a3, [ref]$a4)
    Write-Host "BuildAll: $result"
    if ($result -ne 'OK') { $failed = $true }

    # The forms' code-behind goes in during BuildAll, after the compile above, and
    # Access compiles a form module only when the form first opens. Compile again so
    # an error there (a reserved word as a name, say) stops the build here instead.
    if (-not $failed) {
        $bad = Test-AccessCompiles $acc $watch
        if ($bad) { Write-Host "Compile error after the forms were built: $bad"; throw 'compile failed' }
    }

    # The locked release copy (-Release, see the top of the file). SysCmd 603
    # needs the source closed and no file where the .accde goes.
    if ($Release -and -not $failed) {
        Remove-Item $accde -ErrorAction SilentlyContinue
        if (Test-Path $accde) { throw "Can't replace $accde; close it and build again." }
        $acc.CloseCurrentDatabase()
        # 603 (make ACCDE) is not in the AcSysCmdAction enum, so PowerShell can't pass it
        # through the typed SysCmd method; call it late-bound, as the kit demo's release.ps1 does.
        [void][System.__ComObject].InvokeMember('SysCmd', 'InvokeMethod', $null, $acc, @(603, [string]$fe, [string]$accde))
        if (-not (Test-Path $accde)) { throw "Access didn't make $accde. Open the .accdb and use File > Save As > Make ACCDE to see why." }
        Lock-ReleaseFrontEnd -Access $acc -Path $accde
        Write-Host "Locked release copy: $accde"
    }
}
catch { $failed = $true; if (-not $watch.Killed) { throw } }
finally {
    $dialogs = Stop-AccessWatch $watch $acc
    Remove-Item $stage -Recurse -Force -ErrorAction SilentlyContinue
}
if ($dialogs) { Write-Host "Build stopped: Access showed the dialog(s) above."; exit 1 }
if ($failed) { exit 1 }
Write-Host "Built: $fe"
