Option Compare Database
Option Explicit

' Access UI kit - choosing files: Windows' own Save As and Open windows, so
' users never type a path. Each starts in the folder the user chose last
' time for the same purpose (remembered per Windows user under APP_NAME),
' offers only the file types the app can handle, and returns "" when the
' user cancels.
'
'   path = KitSaveAs("Export messages", "Messages " & Format$(Date, "yyyy-mm-dd") & ".csv", _
'       "CSV file|*.csv|Excel workbook|*.xlsx", "export")
'   If Len(path) > 0 Then n = ExportGrid(Me.sfList.Form, path)
'
' or, for a table, in one step: ExportGridAs Me.sfList.Form, "Messages".
'
' filters alternates descriptions and patterns, split by "|", the patterns of
' one type split by ";": "CSV file|*.csv|Text|*.txt;*.log". purpose names what
' the files are for ("export", "import"), so each purpose remembers its own
' folder; "" shares one folder.

Private Const SETTINGS_KEY As String = "Files"

Private Const OFN_OVERWRITEPROMPT As Long = &H2
Private Const OFN_HIDEREADONLY As Long = &H4
Private Const OFN_NOCHANGEDIR As Long = &H8
Private Const OFN_PATHMUSTEXIST As Long = &H800
Private Const OFN_FILEMUSTEXIST As Long = &H1000
Private Const OFN_EXPLORER As Long = &H80000
Private Const MAX_PATH_W As Long = 1024         ' characters for the chosen path

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As LongPtr
    lpstrCustomFilter As LongPtr
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As LongPtr
    nMaxFile As Long
    lpstrFileTitle As LongPtr
    nMaxFileTitle As Long
    lpstrInitialDir As LongPtr
    lpstrTitle As LongPtr
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As LongPtr
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
    pvReserved As LongPtr
    dwReserved As Long
    FlagsEx As Long
End Type

Private Declare PtrSafe Function GetSaveFileNameW Lib "comdlg32" (ofn As OPENFILENAME) As Long
Private Declare PtrSafe Function GetOpenFileNameW Lib "comdlg32" (ofn As OPENFILENAME) As Long
Private Declare PtrSafe Function CommDlgExtendedError Lib "comdlg32" () As Long

' Windows' Save As window: returns the chosen path, or "" on Cancel. It asks
' before replacing an existing file, and adds the chosen type's extension
' when the user types a name without one. defaultName is the suggested file
' name (made safe with KitSafeFileName).
Public Function KitSaveAs(ByVal title As String, ByVal defaultName As String, ByVal filters As String, _
        Optional ByVal purpose As String = "") As String
    KitSaveAs = FileWindow(True, title, KitSafeFileName(defaultName), filters, purpose)
End Function

' Windows' Open window for one existing file: returns its path, or "" on
' Cancel.
Public Function KitPickFile(ByVal title As String, ByVal filters As String, Optional ByVal purpose As String = "") As String
    KitPickFile = FileWindow(False, title, "", filters, purpose)
End Function

' Windows' folder picker: returns the chosen folder, or "" on Cancel.
Public Function KitPickFolder(ByVal title As String, Optional ByVal purpose As String = "") As String
    Dim start As String
    On Error GoTo Fail
    start = LastFolder(purpose)
    CloseSplash                                 ' never open behind the splash screen
    With Application.FileDialog(4)              ' msoFileDialogFolderPicker
        .Title = title
        If Len(start) > 0 Then .InitialFileName = WithSlash(start)
        If .Show Then KitPickFolder = .SelectedItems(1)
    End With
    If Len(KitPickFolder) > 0 Then RememberFolder purpose, KitPickFolder
    Exit Function
Fail:
    RaiseAgain "modKitFiles.KitPickFolder", title
End Function

' Export a BuildGrid table as the user sees it (ExportGrid) to a CSV file
' or Excel workbook the user picks, named baseName and today's date, then
' confirm with a toast. Returns the path, or "" when the user cancelled.
Public Function ExportGridAs(gridForm As Form, ByVal baseName As String) As String
    Dim path As String, n As Long
    On Error GoTo Fail
    path = KitSaveAs("Export " & baseName, baseName & " " & Format$(Date, "yyyy-mm-dd") & ".csv", _
        "CSV file|*.csv|Excel workbook|*.xlsx", "export")
    If Len(path) = 0 Then Exit Function
    DoCmd.Hourglass True
    n = ExportGrid(gridForm, path)
    DoCmd.Hourglass False
    KitToast "Exported " & n & IIf(n = 1, " row", " rows") & " to " & Mid$(path, InStrRev(path, "\") + 1), "success"
    ExportGridAs = path
    Exit Function
Fail:
    DoCmd.Hourglass False
    RaiseAgain "modKitFiles.ExportGridAs", baseName
End Function

' s with the characters Windows refuses in file names, and "#" (which
' breaks FollowHyperlink), replaced by "_", trimmed; "file" when nothing is
' left. Line breaks and tabs (a subject typed in a multi-line box) become
' spaces, other control characters "_". A name Windows reserves for a
' device (CON, NUL, COM1, LPT1 ...), with or without an extension, gets "_"
' in front.
Public Function KitSafeFileName(ByVal s As String) As String
    Dim ch As Variant, i As Long, base As String
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", "#")
        s = Replace(s, ch, "_")
    Next
    s = Replace(Replace(Replace(Replace(s, vbCrLf, " "), vbCr, " "), vbLf, " "), vbTab, " ")
    For i = 1 To 31
        s = Replace(s, Chr$(i), "_")
    Next
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = Trim$(s)
    Do While Right$(s, 1) = "."                 ' Windows drops a trailing dot
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) = 0 Then s = "file"
    base = UCase$(Trim$(Split(s, ".")(0)))
    If base = "CON" Or base = "PRN" Or base = "AUX" Or base = "NUL" Or base Like "COM[1-9]" Or base Like "LPT[1-9]" Then s = "_" & s
    KitSafeFileName = s
End Function

' "CSV file|*.csv|Excel workbook|*.xlsx" as Windows wants it: each
' description (with its patterns added) and its patterns ended by a null
' character, the list by another. Raises error 5 unless filters is
' description|patterns pairs.
Public Function FileFilterText(ByVal filters As String) As String
    Dim parts() As String, i As Long, out As String
    parts = Split(filters, "|")
    If UBound(parts) < 1 Or UBound(parts) Mod 2 = 0 Then
        Err.Raise 5, "modKitFiles.FileFilterText", "File filter is not Description|*.ext pairs: " & filters
    End If
    For i = 0 To UBound(parts) Step 2
        If Len(parts(i + 1)) = 0 Then Err.Raise 5, "modKitFiles.FileFilterText", "File filter has no pattern for " & parts(i)
        out = out & parts(i) & " (" & parts(i + 1) & ")" & vbNullChar & parts(i + 1) & vbNullChar
    Next
    FileFilterText = out & vbNullChar
End Function

' The folder last chosen for purpose, or "" when none was or it is gone.
Public Function LastFolder(ByVal purpose As String) As String
    LastFolder = GetSetting(APP_NAME, SETTINGS_KEY, "Folder_" & purpose, "")
    If Len(LastFolder) > 0 Then
        If Not FolderExists(LastFolder) Then LastFolder = ""
    End If
End Function

' Remember path's folder (or path, when it is a folder) for purpose.
Public Sub RememberFolder(ByVal purpose As String, ByVal path As String)
    If Not FolderExists(path) Then
        If InStrRev(path, "\") = 0 Then Exit Sub
        path = Left$(path, InStrRev(path, "\") - 1)
    End If
    If path Like "?:" Then path = path & "\"    ' "C:" alone is the drive's current folder, not its root
    If Len(path) > 0 Then SaveSetting APP_NAME, SETTINGS_KEY, "Folder_" & purpose, path
End Sub

' fileName without its extension when that extension is one of filters'
' patterns ("Report.csv" for "CSV file|*.csv" gives "Report").
Public Function WithoutFilterExtension(ByVal fileName As String, ByVal filters As String) As String
    Dim parts() As String, pats() As String, i As Long, j As Long, p As Long
    WithoutFilterExtension = fileName
    p = InStrRev(fileName, ".")
    If p = 0 Then Exit Function
    parts = Split(filters, "|")
    For i = 1 To UBound(parts) Step 2
        pats = Split(parts(i), ";")
        For j = 0 To UBound(pats)
            If StrComp(Trim$(pats(j)), "*" & Mid$(fileName, p), vbTextCompare) = 0 Then
                WithoutFilterExtension = Left$(fileName, p - 1)
                Exit Function
            End If
        Next
    Next
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

' Save As (saving True) or Open, owned by the Access window.
Private Function FileWindow(ByVal saving As Boolean, ByVal title As String, ByVal defaultName As String, _
        ByVal filters As String, ByVal purpose As String) As String
    Dim ofn As OPENFILENAME, buf As String, filt As String, startDir As String, defExt As String, ok As Long, errCode As Long
    On Error GoTo Fail
    filt = FileFilterText(filters)
    startDir = LastFolder(purpose)
    defExt = FirstExtension(filters)
    ' Without its extension the name takes the type the user picks in the
    ' window (Windows adds it), so switching to Excel saves a .xlsx.
    If saving Then defaultName = WithoutFilterExtension(defaultName, filters)
    ' The folder in the name too: Windows can prefer the folder it used last
    ' over lpstrInitialDir, but never over a path in lpstrFile. Only with a
    ' name: a bare "folder\" makes the Open window fail (FNERR_INVALIDFILENAME).
    If Len(startDir) > 0 And Len(defaultName) > 0 Then defaultName = WithSlash(startDir) & defaultName
    buf = Left$(defaultName, MAX_PATH_W - 1) & String$(MAX_PATH_W - Len(Left$(defaultName, MAX_PATH_W - 1)), vbNullChar)
    CloseSplash                                 ' never open behind the splash screen
    With ofn
        .lStructSize = LenB(ofn)
        .hwndOwner = Application.hWndAccessApp
        .lpstrFilter = StrPtr(filt)
        .nFilterIndex = 1
        .lpstrFile = StrPtr(buf)
        .nMaxFile = MAX_PATH_W
        If Len(startDir) > 0 Then .lpstrInitialDir = StrPtr(startDir)
        .lpstrTitle = StrPtr(title)
        If Len(defExt) > 0 Then .lpstrDefExt = StrPtr(defExt)
        .flags = OFN_EXPLORER Or OFN_HIDEREADONLY Or OFN_NOCHANGEDIR Or OFN_PATHMUSTEXIST
        If saving Then .flags = .flags Or OFN_OVERWRITEPROMPT Else .flags = .flags Or OFN_FILEMUSTEXIST
    End With
    If saving Then ok = GetSaveFileNameW(ofn) Else ok = GetOpenFileNameW(ofn)
    If ok = 0 Then                              ' Cancel, unless Windows reports why the window failed
        errCode = CommDlgExtendedError()
        If errCode <> 0 Then Err.Raise vbObjectError + 513, "modKitFiles", "The file window could not open (code &H" & Hex$(errCode) & ")."
        Exit Function
    End If
    FileWindow = Left$(buf, InStr(buf, vbNullChar) - 1)
    RememberFolder purpose, FileWindow
    Exit Function
Fail:
    RaiseAgain "modKitFiles.FileWindow", title
End Function

' folder ending in one "\".
Private Function WithSlash(ByVal folder As String) As String
    If Right$(folder, 1) = "\" Then WithSlash = folder Else WithSlash = folder & "\"
End Function

' The first pattern's extension without "*." ("csv"), or "" for "*.*".
Private Function FirstExtension(ByVal filters As String) As String
    Dim pat As String, p As Long
    pat = Split(Split(filters, "|")(1), ";")(0)
    p = InStrRev(pat, ".")
    If p > 0 Then FirstExtension = Mid$(pat, p + 1)
    If FirstExtension = "*" Then FirstExtension = ""
End Function
