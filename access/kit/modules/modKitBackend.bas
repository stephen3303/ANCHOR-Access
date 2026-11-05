Option Compare Database
Option Explicit

' Access UI kit - a front end linked to a shared back end (the data file on
' a network share that every user's copy links to).
'
' At startup, before the first form reads data:
'   If Not KitEnsureBackend("item", "MyApp_data.accdb") Then DoCmd.Quit
' checks that the linked tables reach their data file. When the file has
' moved (a new share, a renamed folder) it looks for it next to the front
' end, then asks the user to find it in Windows' Open window (KitPickFile),
' and relinks every linked table to it: one window instead of a broken app.
'
' Version checks keep old copies of the front end off a newer data file.
' The app records its data version in the back end its own way; the kit
' compares and words the message:
'   msg = KitVersionProblem(beVersion, SCHEMA_VERSION, minBuild, APP_BUILD, advice)
'   If Len(msg) > 0 Then KitMsgBox msg, vbExclamation: DoCmd.Quit
'
' Only links to Access data files (Connect ";DATABASE=..." or "MS Access;...")
' are touched, and
' a relink moves only the tables linked to the same file as the probe table,
' so a second linked file (the shared Personnel file) keeps its links.

Private Const DB_PREFIX As String = ";DATABASE="

' The data file a linked table points at ("" when it isn't linked to an
' Access data file, or there is no such table).
Public Function KitBackendPath(ByVal probeTable As String) As String
    Dim cn As String
    On Error Resume Next                        ' no such table: ""
    cn = CurrentDb.TableDefs(probeTable).Connect
    On Error GoTo 0
    KitBackendPath = ConnectPath(cn)
End Function

' The tables linked to Access data files, by name: all of them, or only
' those linked to bePath.
Public Function KitLinkedTables(Optional ByVal bePath As String = "") As Collection
    Dim td As DAO.TableDef, p As String
    Set KitLinkedTables = New Collection
    For Each td In CurrentDb.TableDefs
        p = ConnectPath(td.Connect)
        If Len(p) > 0 Then
            If Len(bePath) = 0 Or StrComp(p, bePath, vbTextCompare) = 0 Then KitLinkedTables.Add td.Name
        End If
    Next
End Function

' Raises error 5 unless path is a usable back end: an absolute (drive or
' \\server) path to an existing .accdb that isn't this front end, with no
' ";" (it would break the connection string).
Public Sub KitCheckBackendPath(ByVal path As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If LCase$(fso.GetExtensionName(path)) <> "accdb" Then Err.Raise 5, "modKitBackend", "The data file must be an .accdb file: " & path
    If Not ((Mid$(path, 2, 2) = ":\") Or Left$(path, 2) = "\\") Then Err.Raise 5, "modKitBackend", "The data file path must be a full path: " & path
    If InStr(path, ";") > 0 Then Err.Raise 5, "modKitBackend", "The data file path can't contain "";"": " & path
    If StrComp(fso.GetAbsolutePathName(path), fso.GetAbsolutePathName(CurrentDb.Name), vbTextCompare) = 0 Then
        Err.Raise 5, "modKitBackend", "The front end can't be its own data file."
    End If
    If Not FileExists(path) Then Err.Raise 5, "modKitBackend", "The data file doesn't exist: " & path
End Sub

' Point the tables linked to fromPath (every table linked to an Access data
' file when fromPath is "") at bePath, checking the path first
' (KitCheckBackendPath). Each table keeps its name, source table and
' password. Changes nothing, raising error 5, when bePath lacks any of the
' source tables. Returns how many tables were relinked.
Public Function KitRelinkAll(ByVal bePath As String, Optional ByVal fromPath As String = "") As Long
    Dim db As DAO.Database, nm As Variant, td As DAO.TableDef, be As DAO.Database, missing As String
    Dim n As Long, d As String, src As String, tables As Collection, busy As Boolean
    On Error GoTo Fail
    KitCheckBackendPath bePath
    Set db = CurrentDb
    Set tables = KitLinkedTables(fromPath)
    If tables.Count = 0 Then Exit Function
    KitBusy "Connecting to the data file..."    ' a file on a share can take a while
    busy = True
    ' Shared, read-only, with the links' password: just to look. A file
    ' that won't open (wrong password, not Access, corrupt, in use) is a
    ' refusal like a missing table: error 5, links unchanged.
    On Error Resume Next
    Set be = DBEngine.OpenDatabase(bePath, False, True, ConnectRest(db.TableDefs(tables(1)).Connect))
    If Err.Number <> 0 Then
        d = Err.Description
        On Error GoTo Fail
        Err.Raise 5, "modKitBackend", "That file can't be opened as the data file. " & d
    End If
    On Error GoTo Fail
    For Each nm In tables
        If Not HasTableDef(be, db.TableDefs(nm).SourceTableName) Then missing = missing & IIf(Len(missing) > 0, ", ", "") & nm
    Next
    be.Close
    Set be = Nothing
    If Len(missing) > 0 Then Err.Raise 5, "modKitBackend", "That data file has no " & missing & " table."
    For Each nm In tables
        Set td = db.TableDefs(nm)
        td.Connect = DB_PREFIX & bePath & ConnectRest(td.Connect)
        td.RefreshLink
        KitRelinkAll = KitRelinkAll + 1
    Next
    db.TableDefs.Refresh
    KitBusyEnd
    Exit Function
Fail:
    ' Clean-up resets Err, so keep the error and put it back.
    n = Err.Number: d = Err.Description: src = Err.Source
    If Not be Is Nothing Then
        On Error Resume Next
        be.Close
        On Error GoTo 0
    End If
    If busy Then KitBusyEnd
    Err.Number = n: Err.Description = d: Err.Source = src
    RaiseAgain "modKitBackend.KitRelinkAll", bePath & " (" & nm & ")"
End Function

' Link tables (an array of names) from bePath, replacing a local or linked
' table of the same name; the source table has the same name. pwd is the
' file's password, when it is encrypted. For the build's link step.
' Returns how many were linked.
Public Function KitLinkTables(ByVal bePath As String, ByVal tables As Variant, Optional ByVal pwd As String = "") As Long
    Dim db As DAO.Database, i As Long, td As DAO.TableDef
    On Error GoTo Fail
    KitCheckBackendPath bePath
    Set db = CurrentDb
    For i = LBound(tables) To UBound(tables)
        On Error Resume Next                    ' not there yet
        db.TableDefs.Delete tables(i)
        On Error GoTo Fail
        Set td = db.CreateTableDef(tables(i))
        td.Connect = DB_PREFIX & bePath & IIf(Len(pwd) > 0, ";PWD=" & pwd, "")
        td.SourceTableName = tables(i)
        db.TableDefs.Append td
        KitLinkTables = KitLinkTables + 1
    Next
    db.TableDefs.Refresh
    Exit Function
Fail:
    RaiseAgain "modKitBackend.KitLinkTables", bePath & " (" & tables(i) & ")"
End Function

' True when probeTable's data file can be opened right now.
Public Function KitBackendReachable(ByVal probeTable As String) As Boolean
    Dim path As String, rs As DAO.Recordset
    path = KitBackendPath(probeTable)
    If Len(path) = 0 Then Exit Function
    If Not FileExists(path) Then Exit Function
    On Error GoTo Unreachable
    Set rs = CurrentDb.OpenRecordset("SELECT TOP 1 * FROM [" & probeTable & "]", dbOpenSnapshot)
    rs.Close
    KitBackendReachable = True
Unreachable:
End Function

' Startup: make sure the linked tables reach their data file (see the top
' of the module). backendFile is the data file's name, looked for next to
' the front end when the linked path is gone. A picked file that isn't the
' data file is explained and the window offered again. Returns False,
' having told the user, when they cancel.
Public Function KitEnsureBackend(ByVal probeTable As String, ByVal backendFile As String) As Boolean
    Dim candidate As String, oldPath As String, problem As String
    On Error GoTo Fail
    If KitBackendReachable(probeTable) Then KitEnsureBackend = True: Exit Function
    oldPath = KitBackendPath(probeTable)
    If Len(oldPath) = 0 Then
        If HasTableDef(CurrentDb, probeTable) Then
            KitEnsureBackend = True             ' a local table: a single-file copy, nothing to reconnect
            Exit Function
        End If
        Err.Raise 5, "modKitBackend", "The app has no " & probeTable & " table to check its data file with."
    End If
    candidate = CurrentProject.Path & "\" & backendFile
    If FileExists(candidate) Then
        If Len(TryRelink(candidate, oldPath, probeTable)) = 0 Then KitEnsureBackend = True: Exit Function
    End If
    KitMsgBox "The " & APP_NAME & " data file can't be found where it was:" & vbCrLf & oldPath & vbCrLf & vbCrLf & _
        "Find " & backendFile & " in the next window to reconnect.", vbInformation
    Do
        candidate = KitPickFile("Find " & backendFile, "Access database|*.accdb", "backend")
        If Len(candidate) = 0 Then
            KitMsgBox APP_NAME & " can't open without its data file. Ask your administrator where " & backendFile & _
                " is kept.", vbExclamation
            Exit Function
        End If
        problem = TryRelink(candidate, oldPath, probeTable)
        If Len(problem) = 0 Then KitEnsureBackend = True: Exit Function
        KitMsgBox problem & vbCrLf & vbCrLf & "Pick " & backendFile & " in the next window, or Cancel to close " & _
            APP_NAME & ".", vbExclamation
    Loop
Fail:
    DoCmd.Hourglass False
    HandleError "modKitBackend.KitEnsureBackend", candidate, "The data file couldn't be reconnected."
End Function

' Relink oldPath's tables to candidate. Returns "" when probeTable then
' reaches its data, else what to tell the user (links unchanged when the
' file was refused). Errors other than a refusal are raised.
Private Function TryRelink(ByVal candidate As String, ByVal oldPath As String, ByVal probeTable As String) As String
    On Error GoTo Refused
    DoCmd.Hourglass True
    KitRelinkAll candidate, oldPath
    DoCmd.Hourglass False
    If Not KitBackendReachable(probeTable) Then TryRelink = "That file isn't the " & APP_NAME & " data file, or it can't be opened."
    Exit Function
Refused:
    If Err.Number <> 5 Then RaiseAgain "modKitBackend.TryRelink", candidate   ' the caller's handler clears the hourglass
    TryRelink = Err.Description
    DoCmd.Hourglass False
End Function

' Why this front end can't be used with its data file, or "" when it can.
' beVersion: the data version the back end records; supports: the version
' this front end is built for; minBuild: the oldest front-end build allowed
' ("" for any); feBuild: this copy's build stamp (the app's own); builds are
' "yyyy-mm-dd hh:nn" so they compare as text. advice ends the message (where
' to get the current copy). An older back end is the app's to upgrade before
' calling this; if it still is older, that is reported too.
Public Function KitVersionProblem(ByVal beVersion As Long, ByVal supports As Long, ByVal minBuild As String, _
        ByVal feBuild As String, Optional ByVal advice As String = "") As String
    Dim msg As String
    If Len(advice) = 0 Then advice = "Ask your administrator for the current copy."
    If beVersion > supports Then
        msg = "The shared data file has been updated for a newer version of " & APP_NAME & " (data version " & _
            beVersion & "; this copy supports version " & supports & ")." & vbCrLf & vbCrLf & _
            "This copy of " & APP_NAME & " is out of date. " & advice
    ElseIf beVersion < supports Then
        msg = "The shared data file is an older version (" & beVersion & ") than this copy needs (" & supports & _
            "), and it couldn't be updated. Ask your administrator to update it."
    ElseIf Len(minBuild) > 0 And feBuild < minBuild Then
        msg = "Your administrator requires a copy of " & APP_NAME & " built on or after " & minBuild & _
            ". This copy was built " & IIf(Len(feBuild) > 0, feBuild, "at an unknown time") & "." & vbCrLf & vbCrLf & _
            "This copy of " & APP_NAME & " is out of date. " & advice
    End If
    KitVersionProblem = msg
End Function

' ---------------------------------------------------------------------------
' Helpers
' ---------------------------------------------------------------------------

Private Function HasTableDef(db As DAO.Database, ByVal nm As String) As Boolean
    Dim td As DAO.TableDef
    For Each td In db.TableDefs
        If StrComp(td.Name, nm, vbTextCompare) = 0 Then HasTableDef = True: Exit Function
    Next
End Function

' The file in an Access link's Connect, or "" for any other kind of link.
' Access stores ";DATABASE=path" as given, but a link with a password as
' "MS Access;PWD=...;DATABASE=path", so the parts are looked for by name.
Private Function ConnectPath(ByVal cn As String) As String
    Dim parts() As String, i As Long
    If Not IsAccessConnect(cn) Then Exit Function
    parts = Split(cn, ";")
    For i = 1 To UBound(parts)
        If UCase$(Left$(parts(i), 9)) = "DATABASE=" Then ConnectPath = Mid$(parts(i), 10): Exit Function
    Next
End Function

' The rest of an Access link's Connect without its type and file
' (";PWD=..."), or "".
Private Function ConnectRest(ByVal cn As String) As String
    Dim parts() As String, i As Long
    If Not IsAccessConnect(cn) Then Exit Function
    parts = Split(cn, ";")
    For i = 1 To UBound(parts)
        If Len(parts(i)) > 0 And UCase$(Left$(parts(i), 9)) <> "DATABASE=" Then ConnectRest = ConnectRest & ";" & parts(i)
    Next
End Function

' True for a link to an Access data file: its type part is "" or "MS Access"
' (an ODBC link starts "ODBC;").
Private Function IsAccessConnect(ByVal cn As String) As Boolean
    Dim kind As String
    If InStr(cn, ";") = 0 Then Exit Function
    kind = Left$(cn, InStr(cn, ";") - 1)
    IsAccessConnect = (Len(kind) = 0 Or StrComp(kind, "MS Access", vbTextCompare) = 0)
End Function
